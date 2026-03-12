from __future__ import annotations

import importlib.util
import json
import re
import unicodedata
from datetime import datetime
from functools import lru_cache
from pathlib import Path
from statistics import mean
from typing import Any
from uuid import uuid4


ROOT_DIR = Path(__file__).resolve().parent
DATA_DIR = ROOT_DIR / "data"
BD_FILE = DATA_DIR / "BD-Calibracion.json"
CLIENTS_FILE = DATA_DIR / "Clientes.json"
NORMS_FILE = DATA_DIR / "Normas.json"
USERS_FILE = DATA_DIR / "Usuarios.json"
STATE_FILE = DATA_DIR / "app_state.json"
HISTORY_DIR = DATA_DIR / "historico"
DOCUMENT_MODULE_DIR = ROOT_DIR / "Documentos PDF.py"


def _read_json(path: Path, default: Any) -> Any:
	if not path.exists() or path.stat().st_size == 0:
		return default

	with path.open("r", encoding="utf-8") as handle:
		return json.load(handle)


def _write_json(path: Path, payload: Any) -> None:
	path.parent.mkdir(parents=True, exist_ok=True)
	with path.open("w", encoding="utf-8") as handle:
		json.dump(payload, handle, ensure_ascii=False, indent=2)


def _safe_slug(value: str) -> str:
	slug = re.sub(r"[^a-zA-Z0-9]+", "_", value.strip())
	return slug.strip("_") or "sin_nombre"


def _normalize_person_name(value: str) -> str:
	text = str(value or "").strip().lower()
	if not text:
		return ""

	normalized = unicodedata.normalize("NFKD", text)
	without_accents = "".join(ch for ch in normalized if not unicodedata.combining(ch))
	without_symbols = re.sub(r"[^a-z0-9]+", " ", without_accents)
	return re.sub(r"\s+", " ", without_symbols).strip()


def _extract_norm_token(value: str) -> str | None:
	match = re.search(r"NOM-\d{3}", value.upper())
	return match.group(0) if match else None


def _coerce_score(value: Any) -> float | None:
	if value in (None, ""):
		return None

	try:
		return round(float(value), 2)
	except (TypeError, ValueError):
		return None


def _timestamp() -> str:
	return datetime.now().strftime("%Y-%m-%d %H:%M")


def _default_state() -> dict[str, Any]:
	return {"evaluations": {}, "visits": [], "quarterly_scores": []}


@lru_cache(maxsize=4)
def _load_module(module_path: str):
	path = Path(module_path)
	spec = importlib.util.spec_from_file_location(path.stem, path)
	if spec is None or spec.loader is None:
		raise ImportError(f"No se pudo cargar el modulo: {path}")

	module = importlib.util.module_from_spec(spec)
	spec.loader.exec_module(module)
	return module


class CalibrationController:
	def __init__(self, root_dir: Path | None = None) -> None:
		self.root_dir = root_dir or ROOT_DIR
		self.current_user: dict[str, Any] | None = None
		self.raw_records: list[dict[str, Any]] = []
		self.norms_catalog: list[dict[str, Any]] = []
		self.users_catalog: list[dict[str, Any]] = []
		self.clients_catalog: list[dict[str, Any]] = []
		self.app_state: dict[str, Any] = _default_state()
		self.reload()

	def reload(self) -> None:
		self.raw_records = _read_json(BD_FILE, [])
		self.norms_catalog = _read_json(NORMS_FILE, [])
		self.clients_catalog = _read_json(CLIENTS_FILE, [])

		users_payload = _read_json(USERS_FILE, {"users": []})
		self.users_catalog = users_payload.get("users", [])

		self.app_state = _read_json(STATE_FILE, _default_state())
		self.app_state.setdefault("evaluations", {})
		self.app_state.setdefault("visits", [])
		self.app_state.setdefault("quarterly_scores", [])
		HISTORY_DIR.mkdir(parents=True, exist_ok=True)

	def authenticate(self, username: str, password: str) -> dict[str, Any] | None:
		normalized_user = username.strip().lower()
		normalized_password = password.strip()

		for user in self.users_catalog:
			if user.get("username", "").strip().lower() != normalized_user:
				continue
			if user.get("password", "").strip() != normalized_password:
				continue

			self.current_user = dict(user)
			return self.current_user

		return None

	def logout(self) -> None:
		self.current_user = None

	def is_admin(self, user: dict[str, Any] | None = None) -> bool:
		candidate = user or self.current_user or {}
		return candidate.get("role") == "admin"

	def available_sections(self, user: dict[str, Any] | None = None) -> list[str]:
		if self.is_admin(user):
			return ["Principal", "Dashboard", "Calendario", "Trimestral", "Configuraciones"]
		return ["Principal", "Calendario"]

	def get_record(self, inspector_name: str) -> dict[str, Any] | None:
		target_name = str(inspector_name or "").strip()
		if not target_name:
			return None

		for record in self.raw_records:
			if str(record.get("NOMBRE", "")).strip() == target_name:
				return record

		normalized_target = _normalize_person_name(target_name)
		if not normalized_target:
			return None

		for record in self.raw_records:
			record_name = str(record.get("NOMBRE", "")).strip()
			if _normalize_person_name(record_name) == normalized_target:
				return record
		return None

	def get_catalog_norms(self) -> list[dict[str, str]]:
		catalog: list[dict[str, str]] = []
		seen: set[str] = set()

		for item in self.norms_catalog:
			token = _extract_norm_token(str(item.get("NOM", "")))
			if not token or token in seen:
				continue

			catalog.append(
				{
					"token": token,
					"nom": str(item.get("NOM", token)).strip(),
					"nombre": str(item.get("NOMBRE", "Catalogo no definido")).strip(),
					"capitulo": str(item.get("CAPITULO", "")).strip(),
				}
			)
			seen.add(token)

		for record in self.raw_records:
			for token in self.get_accredited_norms(record):
				if token in seen:
					continue
				catalog.append(
					{
						"token": token,
						"nom": token,
						"nombre": "Catalogo no definido",
						"capitulo": "No especificado",
					}
				)
				seen.add(token)

		return sorted(catalog, key=lambda item: self._norm_sort_key(item["token"]))

	def get_norm_tokens(self) -> list[str]:
		return [item["token"] for item in self.get_catalog_norms()]

	def get_accredited_norms(self, record_or_name: dict[str, Any] | str | None) -> list[str]:
		if record_or_name is None:
			return []

		record = record_or_name
		if isinstance(record_or_name, str):
			record = self.get_record(record_or_name) or {}

		accredited: list[str] = []
		seen: set[str] = set()
		for key, value in record.items():
			if "ACREDITADOS" not in str(key).upper():
				continue
			if value in (None, ""):
				continue
			token = _extract_norm_token(str(value)) or _extract_norm_token(str(key))
			if token and token not in seen:
				accredited.append(token)
				seen.add(token)

		return sorted(accredited, key=self._norm_sort_key)

	def get_dashboard_people(self) -> list[str]:
		names = {
			str(record.get("NOMBRE", "")).strip()
			for record in self.raw_records
			if str(record.get("NOMBRE", "")).strip()
		}
		names.update(
			str(user.get("name", "")).strip()
			for user in self.users_catalog
			if str(user.get("name", "")).strip()
		)
		return sorted(names)

	def get_assignable_inspectors(self) -> list[str]:
		executives = sorted(
			str(user.get("name", "")).strip()
			for user in self.users_catalog
			if user.get("role") == "ejecutivo" and str(user.get("name", "")).strip()
		)
		return executives or self.get_dashboard_people()

	def get_client_names(self) -> list[str]:
		names = {
			str(client.get("CLIENTE", "")).strip()
			for client in self.clients_catalog
			if str(client.get("CLIENTE", "")).strip()
		}
		return sorted(names)

	def get_client_addresses(self, client_name: str) -> list[dict[str, str]]:
		options: list[dict[str, str]] = []
		for client in self.clients_catalog:
			if str(client.get("CLIENTE", "")).strip() != client_name:
				continue

			for index, address in enumerate(client.get("DIRECCIONES", []), start=1):
				street = str(address.get("CALLE Y NO", "")).strip()
				colony = str(address.get("COLONIA O POBLACION", "")).strip()
				municipality = str(address.get("MUNICIPIO O ALCADIA", "")).strip()
				state = str(address.get("CIUDAD O ESTADO", "")).strip()
				postal_code = str(address.get("CP", "")).strip()
				service = str(address.get("SERVICIO", "")).strip()
				address_text = ", ".join(
					part for part in [street, colony, municipality, state, postal_code] if part
				)
				label = f"Sede {index}: {municipality or state or 'Sin ubicacion'} | {service or 'Sin servicio'}"
				options.append(
					{
						"label": label,
						"address": address_text,
						"service": service or "Sin servicio",
					}
				)

		return options

	def get_latest_evaluation(self, inspector_name: str) -> dict[str, Any]:
		latest = self.app_state.get("evaluations", {}).get(inspector_name)
		if latest:
			return latest

		history = self.get_history(inspector_name)
		return history[-1] if history else {}

	def has_completed_form(self, inspector_name: str) -> bool:
		latest = self.get_latest_evaluation(inspector_name)
		return bool(latest.get("form_completed"))

	def get_history(self, inspector_name: str) -> list[dict[str, Any]]:
		history_path = self._history_file(inspector_name)
		history = _read_json(history_path, [])
		return sorted(history, key=lambda item: item.get("saved_at", ""))

	def get_recent_visits(self, inspector_name: str, limit: int = 5) -> list[dict[str, Any]]:
		visits = self.list_visits(name=inspector_name)
		return visits[:limit]

	def get_principal_rows(self, search_text: str = "", status_filter: str = "Todos") -> list[dict[str, Any]]:
		lowered_query = search_text.strip().lower()
		rows: list[dict[str, Any]] = []

		for record in self.raw_records:
			name = str(record.get("NOMBRE", "Sin nombre")).strip()
			accredited_norms = self.get_accredited_norms(record)
			latest = self.get_latest_evaluation(name)
			latest_score = _coerce_score(latest.get("score"))
			form_completed = bool(latest.get("form_completed"))

			if latest_score is None:
				status = "Pendiente"
			elif latest_score < 90:
				status = "En enfoque"
			else:
				status = "Estable"

			row = {
				"name": name,
				"norms": accredited_norms,
				"norms_text": ", ".join(accredited_norms) if accredited_norms else "Sin acreditaciones",
				"norm_count": len(accredited_norms),
				"latest_date": latest.get("visit_date") or latest.get("saved_at") or "--",
				"latest_score": latest_score,
				"latest_score_text": f"{latest_score:.1f}%" if latest_score is not None else "--",
				"status": status,
				"form_completed": form_completed,
				"actions_text": "Formulario / PDFs",
			}

			searchable_blob = f"{name} {row['norms_text']} {status}".lower()
			if lowered_query and lowered_query not in searchable_blob:
				continue
			if status_filter == "Completos" and not row["form_completed"]:
				continue
			if status_filter == "Pendientes" and row["form_completed"]:
				continue
			if status_filter == "En enfoque" and row["status"] != "En enfoque":
				continue

			rows.append(row)

		return sorted(rows, key=lambda item: item["name"].lower())

	def get_overview_metrics(self) -> dict[str, Any]:
		rows = self.get_principal_rows()
		completed = sum(1 for row in rows if row["form_completed"])
		scores = [row["latest_score"] for row in rows if row["latest_score"] is not None]
		alerts = sum(1 for score in scores if score < 90)
		return {
			"inspectors": len(rows),
			"completed_forms": completed,
			"average_score": round(mean(scores), 1) if scores else None,
			"alerts": alerts,
		}

	def get_norm_card_metrics(self) -> list[dict[str, Any]]:
		catalog = self.get_catalog_norms()
		counts = {item["token"]: 0 for item in catalog}
		for record in self.raw_records:
			for token in self.get_accredited_norms(record):
				counts[token] = counts.get(token, 0) + 1

		return [
			{
				"token": item["token"],
				"label": item["nom"],
				"description": item["nombre"],
				"count": counts.get(item["token"], 0),
			}
			for item in catalog
		]

	def get_executive_profile(self, inspector_name: str) -> dict[str, Any]:
		history = self.get_history(inspector_name)
		chart_points: list[dict[str, Any]] = []
		for entry in history:
			score = _coerce_score(entry.get("score"))
			if score is None:
				continue
			chart_points.append(
				{
					"label": entry.get("visit_date") or entry.get("saved_at", ""),
					"score": score,
				}
			)

		scores = [point["score"] for point in chart_points]
		latest_score = scores[-1] if scores else None
		average_score = round(mean(scores), 1) if scores else None
		focus_required = bool(
			(latest_score is not None and latest_score < 90)
			or (average_score is not None and average_score < 90)
		)

		return {
			"name": inspector_name,
			"accredited_norms": self.get_accredited_norms(inspector_name),
			"history": chart_points,
			"latest_score": latest_score,
			"average_score": average_score,
			"focus_required": focus_required,
			"latest_status": history[-1].get("status", "Sin seguimiento") if history else "Sin seguimiento",
			"recent_visits": self.get_recent_visits(inspector_name),
		}

	def save_principal_record(
		self,
		name: str,
		accredited_norms: list[str],
		original_name: str | None = None,
	) -> dict[str, Any]:
		clean_name = name.strip()
		if not clean_name:
			raise ValueError("El nombre del inspector es obligatorio.")

		if original_name:
			record = self.get_record(original_name)
			if record is None:
				raise ValueError("No se encontro el registro a editar.")
		else:
			record = None

		duplicate = self.get_record(clean_name)
		if duplicate and clean_name != original_name:
			raise ValueError("Ya existe un inspector con ese nombre.")

		updated_record = dict(record or {})
		updated_record["NOMBRE"] = clean_name
		selected_tokens = {token for token in accredited_norms if token}

		accredited_keys: dict[str, str] = {}
		for key, value in updated_record.items():
			if "ACREDITADOS" not in str(key).upper():
				continue
			token = _extract_norm_token(str(value)) or _extract_norm_token(str(key))
			if token:
				accredited_keys[token] = key

		for token in selected_tokens:
			accredited_keys.setdefault(token, f"ACREDITADOS {token}")

		for token, key in accredited_keys.items():
			updated_record[key] = token if token in selected_tokens else None

		if record is None:
			self.raw_records.append(updated_record)
		else:
			index = self.raw_records.index(record)
			self.raw_records[index] = updated_record

		if original_name and clean_name != original_name:
			self._rename_related_history(original_name, clean_name)

		_write_json(BD_FILE, self.raw_records)
		self.reload()
		return updated_record

	def delete_principal_record(self, inspector_name: str) -> None:
		record = self.get_record(inspector_name)
		if record is None:
			return

		self.raw_records.remove(record)
		_write_json(BD_FILE, self.raw_records)
		self.reload()

	def save_evaluation(self, inspector_name: str, payload: dict[str, Any]) -> dict[str, Any]:
		clean_norm = str(payload.get("selected_norm", "")).strip()
		clean_client = str(payload.get("client", "")).strip()
		clean_date = str(payload.get("visit_date", "")).strip()
		clean_status = str(payload.get("status", "")).strip() or "En seguimiento"
		clean_observations = str(payload.get("observations", "")).strip()
		clean_actions = str(payload.get("corrective_actions", "")).strip()
		evaluator = str(payload.get("evaluator", "")).strip()
		score = _coerce_score(payload.get("score"))

		if not clean_date:
			raise ValueError("La fecha de seguimiento es obligatoria.")
		if score is None:
			raise ValueError("El puntaje debe ser numerico.")
		if score < 0 or score > 100:
			raise ValueError("El puntaje debe estar entre 0 y 100.")

		evaluation = {
			"inspector_name": inspector_name,
			"selected_norm": clean_norm or "Sin norma",
			"client": clean_client or "Sin cliente",
			"visit_date": clean_date,
			"score": score,
			"status": clean_status,
			"observations": clean_observations,
			"corrective_actions": clean_actions,
			"evaluator": evaluator or (self.current_user or {}).get("name", "Sin evaluador"),
			"saved_at": _timestamp(),
			"form_completed": True,
		}

		history = self.get_history(inspector_name)
		history.append(evaluation)
		self._write_history(inspector_name, history)

		self.app_state["evaluations"][inspector_name] = evaluation
		_write_json(STATE_FILE, self.app_state)
		self.reload()
		return evaluation

	def list_visits(
		self,
		current_user: dict[str, Any] | None = None,
		name: str | None = None,
	) -> list[dict[str, Any]]:
		visits = list(self.app_state.get("visits", []))
		if name:
			visits = [visit for visit in visits if visit.get("inspector") == name]

		candidate = current_user or self.current_user
		if candidate and candidate.get("role") == "ejecutivo":
			visits = [visit for visit in visits if visit.get("inspector") == candidate.get("name")]

		visits.sort(key=lambda item: (item.get("visit_date", ""), item.get("updated_at", "")), reverse=True)
		return visits

	def list_trimestral_scores(
		self,
		inspector_name: str | None = None,
		year: int | None = None,
		quarter: str | None = None,
	) -> list[dict[str, Any]]:
		scores = list(self.app_state.get("quarterly_scores", []))
		if inspector_name:
			target = str(inspector_name).strip()
			scores = [item for item in scores if str(item.get("inspector", "")).strip() == target]
		if year is not None:
			year_value = int(year)
			scores = [item for item in scores if int(item.get("year", 0)) == year_value]
		if quarter:
			quarter_value = str(quarter).strip().upper()
			scores = [item for item in scores if str(item.get("quarter", "")).strip().upper() == quarter_value]

		scores.sort(
			key=lambda item: (
				int(item.get("year", 0)),
				str(item.get("quarter", "")),
				str(item.get("updated_at", "")),
			),
			reverse=True,
		)
		return scores

	def save_trimestral_score(self, payload: dict[str, Any], score_id: str | None = None) -> dict[str, Any]:
		inspector = str(payload.get("inspector", "")).strip()
		quarter = str(payload.get("quarter", "")).strip().upper()
		year_value = str(payload.get("year", "")).strip()
		notes = str(payload.get("notes", "")).strip()
		score = _coerce_score(payload.get("score"))

		if not inspector:
			raise ValueError("Debes seleccionar un ejecutivo.")
		if quarter not in {"T1", "T2", "T3", "T4"}:
			raise ValueError("El trimestre debe ser T1, T2, T3 o T4.")
		if not year_value.isdigit():
			raise ValueError("El anio debe ser numerico.")
		year = int(year_value)
		if year < 2000 or year > 2100:
			raise ValueError("El anio debe estar en un rango valido.")
		if score is None:
			raise ValueError("La calificacion debe ser numerica.")
		if score < 0 or score > 100:
			raise ValueError("La calificacion debe estar entre 0 y 100.")

		scores = list(self.app_state.get("quarterly_scores", []))
		existing = None
		if score_id:
			existing = next((item for item in scores if item.get("id") == score_id), None)

		record = dict(existing or {})
		record["id"] = record.get("id") or uuid4().hex
		record["inspector"] = inspector
		record["quarter"] = quarter
		record["year"] = year
		record["score"] = score
		record["notes"] = notes
		record["evaluator"] = (self.current_user or {}).get("name", "Sistema")
		record["updated_at"] = _timestamp()

		if existing is None:
			scores.append(record)
		else:
			index = scores.index(existing)
			scores[index] = record

		self.app_state["quarterly_scores"] = scores
		_write_json(STATE_FILE, self.app_state)
		self.reload()
		return record

	def delete_trimestral_score(self, score_id: str) -> None:
		scores = list(self.app_state.get("quarterly_scores", []))
		existing = next((item for item in scores if item.get("id") == score_id), None)
		if existing is None:
			return

		scores.remove(existing)
		self.app_state["quarterly_scores"] = scores
		_write_json(STATE_FILE, self.app_state)
		self.reload()

	def save_visit(self, payload: dict[str, Any], visit_id: str | None = None) -> dict[str, Any]:
		inspector = str(payload.get("inspector", "")).strip()
		client = str(payload.get("client", "")).strip()
		address = str(payload.get("address", "")).strip()
		service = str(payload.get("service", "")).strip()
		visit_date = str(payload.get("visit_date", "")).strip()
		status = str(payload.get("status", "")).strip() or "Programada"
		notes = str(payload.get("notes", "")).strip()

		if not inspector:
			raise ValueError("Debes seleccionar un inspector.")
		if not client:
			raise ValueError("Debes seleccionar un cliente.")
		if not address:
			raise ValueError("Debes seleccionar una direccion.")
		if not visit_date:
			raise ValueError("La fecha de visita es obligatoria.")

		visits = list(self.app_state.get("visits", []))
		existing = None
		if visit_id:
			existing = next((item for item in visits if item.get("id") == visit_id), None)

		visit = dict(existing or {})
		visit["id"] = visit.get("id") or uuid4().hex
		visit["inspector"] = inspector
		visit["client"] = client
		visit["address"] = address
		visit["service"] = service or "Sin servicio"
		visit["visit_date"] = visit_date
		visit["status"] = status
		visit["notes"] = notes
		visit["assigned_by"] = (self.current_user or {}).get("name", "Sistema")
		visit["updated_at"] = _timestamp()

		if existing is None:
			visits.append(visit)
		else:
			index = visits.index(existing)
			visits[index] = visit

		self.app_state["visits"] = visits
		_write_json(STATE_FILE, self.app_state)
		self._sync_visit_history(inspector)
		self.reload()
		return visit

	def delete_visit(self, visit_id: str) -> None:
		visits = list(self.app_state.get("visits", []))
		existing = next((item for item in visits if item.get("id") == visit_id), None)
		if existing is None:
			return

		affected_inspector = existing.get("inspector")
		visits.remove(existing)
		self.app_state["visits"] = visits
		_write_json(STATE_FILE, self.app_state)
		if affected_inspector:
			self._sync_visit_history(affected_inspector)
		self.reload()

	def save_norm(self, payload: dict[str, Any], original_nom: str | None = None) -> dict[str, Any]:
		nom = str(payload.get("NOM", "")).strip()
		nombre = str(payload.get("NOMBRE", "")).strip()
		capitulo = str(payload.get("CAPITULO", "")).strip()

		if not nom:
			raise ValueError("La clave de la norma es obligatoria.")
		if not _extract_norm_token(nom):
			raise ValueError("La clave debe incluir un codigo como NOM-004.")

		duplicate = next((item for item in self.norms_catalog if item.get("NOM") == nom), None)
		if duplicate and nom != original_nom:
			raise ValueError("Ya existe una norma con esa clave.")

		updated = {"NOM": nom, "NOMBRE": nombre, "CAPITULO": capitulo}
		if original_nom:
			original = next((item for item in self.norms_catalog if item.get("NOM") == original_nom), None)
			if original is None:
				raise ValueError("No se encontro la norma a editar.")
			index = self.norms_catalog.index(original)
			self.norms_catalog[index] = updated
		else:
			self.norms_catalog.append(updated)

		_write_json(NORMS_FILE, self.norms_catalog)
		self.reload()
		return updated

	def delete_norm(self, nom: str) -> None:
		original = next((item for item in self.norms_catalog if item.get("NOM") == nom), None)
		if original is None:
			return

		self.norms_catalog.remove(original)
		_write_json(NORMS_FILE, self.norms_catalog)
		self.reload()

	def save_user(self, payload: dict[str, Any], original_username: str | None = None) -> dict[str, Any]:
		name = str(payload.get("name", "")).strip()
		username = str(payload.get("username", "")).strip()
		password = str(payload.get("password", "")).strip()
		role = str(payload.get("role", "")).strip().lower() or "ejecutivo"

		if not all([name, username, password]):
			raise ValueError("Nombre, usuario y contrasena son obligatorios.")
		if role not in {"admin", "ejecutivo"}:
			raise ValueError("El rol debe ser admin o ejecutivo.")

		duplicate = next((item for item in self.users_catalog if item.get("username") == username), None)
		if duplicate and username != original_username:
			raise ValueError("Ya existe un usuario con ese nombre.")

		updated = {
			"name": name,
			"username": username,
			"password": password,
			"role": role,
		}

		if original_username:
			original = next(
				(item for item in self.users_catalog if item.get("username") == original_username),
				None,
			)
			if original is None:
				raise ValueError("No se encontro el usuario a editar.")
			index = self.users_catalog.index(original)
			self.users_catalog[index] = updated
		else:
			self.users_catalog.append(updated)

		_write_json(USERS_FILE, {"users": self.users_catalog})
		self.reload()
		return updated

	def delete_user(self, username: str) -> None:
		original = next((item for item in self.users_catalog if item.get("username") == username), None)
		if original is None:
			return

		self.users_catalog.remove(original)
		_write_json(USERS_FILE, {"users": self.users_catalog})
		self.reload()

	def get_default_document_path(self, inspector_name: str, document_kind: str) -> Path:
		documents_dir = self._history_dir(inspector_name) / "documentos"
		documents_dir.mkdir(parents=True, exist_ok=True)
		slug = _safe_slug(inspector_name)
		if document_kind == "formato":
			filename = f"{slug}_formato_supervision.pdf"
		else:
			filename = f"{slug}_criterio_evaluacion_tecnica.pdf"
		return documents_dir / filename

	def generate_document(self, inspector_name: str, document_kind: str, destination: str | Path) -> Path:
		latest = self.get_latest_evaluation(inspector_name)
		if not latest.get("form_completed"):
			raise ValueError("Completa el formulario antes de descargar el documento.")

		output_path = Path(destination)
		output_path.parent.mkdir(parents=True, exist_ok=True)
		payload = dict(latest)
		payload["inspector_name"] = inspector_name
		payload["accredited_norms"] = self.get_accredited_norms(inspector_name)

		if document_kind == "formato":
			module = _load_module(str(DOCUMENT_MODULE_DIR / "FormatoSupervicion.py"))
			builder = getattr(module, "build_formato_supervision_pdf")
		elif document_kind == "criterio":
			module = _load_module(str(DOCUMENT_MODULE_DIR / "CriterioEvaluacionTecnica.py"))
			builder = getattr(module, "build_criterio_evaluacion_pdf")
		else:
			raise ValueError("Tipo de documento no soportado.")

		builder(output_path, payload)
		return output_path

	def _history_dir(self, inspector_name: str) -> Path:
		path = HISTORY_DIR / _safe_slug(inspector_name)
		path.mkdir(parents=True, exist_ok=True)
		return path

	def _history_file(self, inspector_name: str) -> Path:
		return self._history_dir(inspector_name) / "historico.json"

	def _visits_file(self, inspector_name: str) -> Path:
		return self._history_dir(inspector_name) / "visitas.json"

	def _write_history(self, inspector_name: str, history: list[dict[str, Any]]) -> None:
		_write_json(self._history_file(inspector_name), history)

	def _sync_visit_history(self, inspector_name: str) -> None:
		_write_json(self._visits_file(inspector_name), self.list_visits(name=inspector_name))




def _controller_evaluation_key(inspector_name: str, norm_token: str | None = None) -> str:
	normalized_norm = str(norm_token or "SIN_NORMA").strip().upper() or "SIN_NORMA"
	return f"{inspector_name}::{normalized_norm}"


def _controller_rename_related_history(self, original_name: str, new_name: str) -> None:
	evaluations = self.app_state.get("evaluations", {})
	for key, payload in list(evaluations.items()):
		if key == original_name:
			new_key = new_name
		elif key.startswith(f"{original_name}::"):
			suffix = key.split("::", 1)[1]
			new_key = f"{new_name}::{suffix}"
		else:
			continue

		renamed_payload = dict(payload)
		renamed_payload["inspector_name"] = new_name
		evaluations[new_key] = renamed_payload
		if new_key != key:
			evaluations.pop(key, None)

	for visit in self.app_state.get("visits", []):
		if visit.get("inspector") == original_name:
			visit["inspector"] = new_name

	original_dir = HISTORY_DIR / _safe_slug(original_name)
	target_dir = HISTORY_DIR / _safe_slug(new_name)
	if original_dir.exists() and not target_dir.exists():
		original_dir.rename(target_dir)

	_write_json(STATE_FILE, self.app_state)
	self._sync_visit_history(new_name)


def _controller_norm_sort_key(token: str) -> tuple[int, str]:
	match = re.search(r"(\d{3})", token)
	if match:
		return int(match.group(1)), token
	return 999, token


def _controller_get_latest_evaluation(self, inspector_name: str, norm_token: str | None = None) -> dict[str, Any]:
	evaluations = self.app_state.get("evaluations", {})

	if norm_token:
		key = _controller_evaluation_key(inspector_name, norm_token)
		latest = evaluations.get(key)
		if latest:
			return latest

		history = self.get_history(inspector_name)
		normalized_norm = str(norm_token).strip().upper()
		for entry in reversed(history):
			current_norm = str(entry.get("selected_norm", "")).strip().upper()
			if current_norm == normalized_norm:
				return entry
		return {}

	candidates: list[dict[str, Any]] = []
	for key, payload in evaluations.items():
		if key == inspector_name or key.startswith(f"{inspector_name}::"):
			candidates.append(payload)

	if candidates:
		return sorted(candidates, key=lambda item: item.get("saved_at", ""))[-1]

	history = self.get_history(inspector_name)
	return history[-1] if history else {}


def _controller_has_completed_form(self, inspector_name: str, norm_token: str | None = None) -> bool:
	latest = self.get_latest_evaluation(inspector_name, norm_token)
	return bool(latest.get("form_completed"))


def _controller_get_principal_rows(self, search_text: str = "", status_filter: str = "Todos") -> list[dict[str, Any]]:
	lowered_query = search_text.strip().lower()
	rows: list[dict[str, Any]] = []

	for record in self.raw_records:
		name = str(record.get("NOMBRE", "Sin nombre")).strip()
		accredited_norms = self.get_accredited_norms(record)
		latest = self.get_latest_evaluation(name)
		latest_score = _coerce_score(latest.get("score"))
		form_completed = self.has_completed_form(name)

		if latest_score is None:
			status = "Pendiente"
		elif latest_score < 90:
			status = "En enfoque"
		else:
			status = "Estable"

		if self.is_admin():
			actions_text = "Formulario | Editar | Borrar"
		else:
			actions_text = "Formulario"

		norms_text = ", ".join(accredited_norms) if accredited_norms else "Sin acreditacion"
		row = {
			"row_id": name,
			"name": name,
			"norm": norms_text,
			"norms_text": norms_text,
			"norm_count": len(accredited_norms),
			"latest_date": latest.get("visit_date") or latest.get("saved_at") or "--",
			"latest_score": latest_score,
			"latest_score_text": f"{latest_score:.1f}%" if latest_score is not None else "--",
			"status": status,
			"form_completed": form_completed,
			"actions_text": actions_text,
		}

		searchable_blob = f"{name} {norms_text} {status}".lower()
		if lowered_query and lowered_query not in searchable_blob:
			continue
		if status_filter == "Completos" and not row["form_completed"]:
			continue
		if status_filter == "Pendientes" and row["form_completed"]:
			continue
		if status_filter == "En enfoque" and row["status"] != "En enfoque":
			continue

		rows.append(row)

	return sorted(rows, key=lambda item: item["name"].lower())


def _controller_get_overview_metrics(self) -> dict[str, Any]:
	inspector_names = {
		str(record.get("NOMBRE", "")).strip()
		for record in self.raw_records
		if str(record.get("NOMBRE", "")).strip()
	}

	completed = 0
	scores: list[float] = []
	for inspector_name in inspector_names:
		latest = self.get_latest_evaluation(inspector_name)
		if latest.get("form_completed"):
			completed += 1
		score = _coerce_score(latest.get("score"))
		if score is not None:
			scores.append(score)

	alerts = sum(1 for score in scores if score < 90)
	return {
		"inspectors": len(inspector_names),
		"completed_forms": completed,
		"average_score": round(mean(scores), 1) if scores else None,
		"alerts": alerts,
	}


def _controller_save_evaluation(self, inspector_name: str, payload: dict[str, Any]) -> dict[str, Any]:
	clean_norm = str(payload.get("selected_norm", "")).strip()
	clean_client = str(payload.get("client", "")).strip()
	clean_date = str(payload.get("visit_date", "")).strip()
	clean_status = str(payload.get("status", "")).strip() or "En seguimiento"
	clean_observations = str(payload.get("observations", "")).strip()
	clean_actions = str(payload.get("corrective_actions", "")).strip()
	evaluator = str(payload.get("evaluator", "")).strip()
	score = _coerce_score(payload.get("score"))

	if not clean_date:
		raise ValueError("La fecha de seguimiento es obligatoria.")
	if score is None:
		raise ValueError("El puntaje debe ser numerico.")
	if score < 0 or score > 100:
		raise ValueError("El puntaje debe estar entre 0 y 100.")

	evaluation = {
		"inspector_name": inspector_name,
		"selected_norm": clean_norm or "Sin norma",
		"client": clean_client or "Sin cliente",
		"visit_date": clean_date,
		"score": score,
		"status": clean_status,
		"observations": clean_observations,
		"corrective_actions": clean_actions,
		"evaluator": evaluator or (self.current_user or {}).get("name", "Sin evaluador"),
		"saved_at": _timestamp(),
		"form_completed": True,
	}

	history = self.get_history(inspector_name)
	history.append(evaluation)
	self._write_history(inspector_name, history)

	eval_key = _controller_evaluation_key(inspector_name, evaluation["selected_norm"])
	self.app_state["evaluations"][eval_key] = evaluation
	self.app_state["evaluations"][inspector_name] = evaluation
	_write_json(STATE_FILE, self.app_state)
	self.reload()
	return evaluation


def _controller_get_default_document_path(
	self,
	inspector_name: str,
	document_kind: str,
	norm_token: str | None = None,
) -> Path:
	documents_dir = self._history_dir(inspector_name) / "documentos"
	documents_dir.mkdir(parents=True, exist_ok=True)

	inspector_slug = _safe_slug(inspector_name)
	norm_slug = _safe_slug(norm_token or "sin_norma")
	if document_kind == "formato":
		filename = f"{inspector_slug}_{norm_slug}_formato_supervision.pdf"
	else:
		filename = f"{inspector_slug}_{norm_slug}_criterio_evaluacion_tecnica.pdf"
	return documents_dir / filename


def _controller_generate_document(
	self,
	inspector_name: str,
	document_kind: str,
	destination: str | Path,
	norm_token: str | None = None,
) -> Path:
	latest = self.get_latest_evaluation(inspector_name, norm_token)
	if not latest.get("form_completed"):
		raise ValueError("Completa el formulario antes de descargar el documento.")

	output_path = Path(destination)
	output_path.parent.mkdir(parents=True, exist_ok=True)

	payload = dict(latest)
	payload["inspector_name"] = inspector_name
	payload["selected_norm"] = str(norm_token or payload.get("selected_norm", "Sin norma"))
	payload["accredited_norms"] = self.get_accredited_norms(inspector_name)

	if document_kind == "formato":
		module = _load_module(str(DOCUMENT_MODULE_DIR / "FormatoSupervicion.py"))
		builder = getattr(module, "build_formato_supervision_pdf")
	elif document_kind == "criterio":
		module = _load_module(str(DOCUMENT_MODULE_DIR / "CriterioEvaluacionTecnica.py"))
		builder = getattr(module, "build_criterio_evaluacion_pdf")
	else:
		raise ValueError("Tipo de documento no soportado.")

	builder(output_path, payload)
	return output_path


CalibrationController._evaluation_key = staticmethod(_controller_evaluation_key)
CalibrationController._rename_related_history = _controller_rename_related_history
CalibrationController._norm_sort_key = staticmethod(_controller_norm_sort_key)
CalibrationController.get_latest_evaluation = _controller_get_latest_evaluation
CalibrationController.has_completed_form = _controller_has_completed_form
CalibrationController.get_principal_rows = _controller_get_principal_rows
CalibrationController.get_overview_metrics = _controller_get_overview_metrics
CalibrationController.save_evaluation = _controller_save_evaluation
CalibrationController.get_default_document_path = _controller_get_default_document_path
CalibrationController.generate_document = _controller_generate_document
