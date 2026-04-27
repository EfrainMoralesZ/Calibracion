from __future__ import annotations

import importlib.util
import json
import re
import shutil
import unicodedata
from datetime import date, datetime, timedelta
from functools import lru_cache
from pathlib import Path
from statistics import mean
from typing import Any
from uuid import uuid4

from runtime_paths import app_dir, resource_path


ROOT_DIR = app_dir()
DATA_DIR = ROOT_DIR / "data"
BD_FILE = DATA_DIR / "BD-Calibracion.json"
CLIENTS_FILE = DATA_DIR / "Clientes.json"
NORMS_FILE = DATA_DIR / "Normas.json"
USERS_FILE = DATA_DIR / "Usuarios.json"
STATE_FILE = DATA_DIR / "app_state.json"
HISTORY_DIR = DATA_DIR / "historico"
VISITS_DIR = DATA_DIR / "visitas"
TRIMESTRAL_DIR = DATA_DIR / "trimestral"
NORMS_REPORT_FILE = DATA_DIR / "reporte de normas.json"
CRITERIA_ARCHIVE_DIR = DATA_DIR / "clientes"
AGREEMENTS_ARCHIVE_DIR = CRITERIA_ARCHIVE_DIR / "acuerdos"
DOCUMENT_MODULE_DIR = resource_path("Documentos PDF.py")


TRIMESTRAL_MEDAL_RULES: tuple[tuple[float, str, str, str], ...] = (
	(100.0, "ORO", "Excelente", ""),
	(90.0, "PLATA", "Optimo", ""),
	(80.0, "BRONCE", "Aceptable", "favor de reforzar"),
)


def _seed_runtime_data_dir() -> None:
	"""Copy seed JSON data from bundle to runtime dir when needed."""
	bundle_data = resource_path("data")
	if not bundle_data.exists() or not bundle_data.is_dir():
		return

	if not DATA_DIR.exists():
		shutil.copytree(bundle_data, DATA_DIR)
		return

	for source in bundle_data.iterdir():
		target = DATA_DIR / source.name
		if target.exists():
			continue
		if source.is_dir():
			shutil.copytree(source, target)
		else:
			shutil.copy2(source, target)


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


def _safe_folder_name(value: str) -> str:
	# Normaliza el nombre para evitar duplicados por acentos, mayúsculas, símbolos, etc.
	text = str(value or "").strip()
	if not text:
		return "sin_nombre"

	# Usar la misma lógica que _normalize_person_name para la base del nombre
	normalized = unicodedata.normalize("NFKD", text)
	without_accents = "".join(ch for ch in normalized if not unicodedata.combining(ch))
	without_symbols = re.sub(r"[^a-zA-Z0-9]+", "_", without_accents)
	folder = without_symbols.strip("._").strip("_")
	return folder or "sin_nombre"


def _folder_identity(value: str) -> str:
	return _normalize_person_name(str(value or "").replace("_", " "))


def _history_folder_matches_identity(folder_path: Path, target_identity: str) -> bool:
	if not target_identity:
		return False

	if _folder_identity(folder_path.name) == target_identity:
		return True

	history_payload = _read_json(folder_path / "historico.json", [])
	if isinstance(history_payload, list):
		for entry in history_payload:
			if not isinstance(entry, dict):
				continue
			inspector_name = _normalize_person_name(str(entry.get("inspector_name", "")))
			inspector_supervised = _normalize_person_name(str(entry.get("inspector_supervised", "")))
			if inspector_name == target_identity or inspector_supervised == target_identity:
				return True

	visits_payload = _read_json(folder_path / "visitas.json", [])
	if isinstance(visits_payload, list):
		for entry in visits_payload:
			if not isinstance(entry, dict):
				continue
			inspectors = _normalize_visit_inspectors(entry.get("inspectors"), str(entry.get("inspector", "")))
			for inspector_name in inspectors:
				if _normalize_person_name(inspector_name) == target_identity:
					return True

	return False


def _merge_json_list_file(target_path: Path, source_path: Path) -> None:
	target_data = _read_json(target_path, [])
	source_data = _read_json(source_path, [])
	if not isinstance(target_data, list):
		target_data = []
	if not isinstance(source_data, list):
		source_data = []

	seen: set[str] = set()
	merged: list[Any] = []
	for item in [*target_data, *source_data]:
		try:
			key = json.dumps(item, ensure_ascii=False, sort_keys=True)
		except TypeError:
			key = str(item)
		if key in seen:
			continue
		seen.add(key)
		merged.append(item)

	if merged != target_data:
		_write_json(target_path, merged)


def _move_file_with_suffix(source_file: Path, target_dir: Path) -> None:
	target_dir.mkdir(parents=True, exist_ok=True)
	destination = target_dir / source_file.name
	if destination.exists():
		stem = source_file.stem
		suffix = source_file.suffix
		counter = 1
		while True:
			candidate = target_dir / f"{stem}_migrado_{counter}{suffix}"
			if not candidate.exists():
				destination = candidate
				break
			counter += 1

	source_file.replace(destination)


def _merge_history_directories(target_dir: Path, source_dir: Path) -> None:
	if not source_dir.exists() or source_dir.resolve() == target_dir.resolve():
		return

	target_dir.mkdir(parents=True, exist_ok=True)
	_merge_json_list_file(target_dir / "historico.json", source_dir / "historico.json")
	_merge_json_list_file(target_dir / "visitas.json", source_dir / "visitas.json")

	for child in list(source_dir.iterdir()):
		if child.name in {"historico.json", "visitas.json"}:
			continue
		if child.is_file():
			_move_file_with_suffix(child, target_dir)
			continue

		if child.is_dir():
			target_base = target_dir / child.name
			for nested in child.rglob("*"):
				if not nested.is_file():
					continue
				relative = nested.relative_to(child)
				_move_file_with_suffix(nested, target_base / relative.parent)

	shutil.rmtree(source_dir, ignore_errors=True)


def _normalize_person_name(value: str) -> str:
	text = str(value or "").strip().lower()
	if not text:
		return ""

	normalized = unicodedata.normalize("NFKD", text)
	without_accents = "".join(ch for ch in normalized if not unicodedata.combining(ch))
	without_symbols = re.sub(r"[^a-z0-9]+", " ", without_accents)
	return re.sub(r"\s+", " ", without_symbols).strip()


def _normalize_role_name(value: str | None) -> str:
	raw_value = str(value or "").strip()
	if not raw_value:
		return ""

	role_key = _normalize_person_name(raw_value).replace(" ", "_")
	role_aliases = {
		"admin": "admin",
		"administrador": "admin",
		"subgerente": "sub gerente",
		"coordinador_operativo": "coordinador operativo",
		"coordinadora_en_fiabilidad": "coordinadora en fiabilidad",
		"talento_humano": "talento humano",
		"supervisor": "supervisor",
		"ejecutivo": "ejecutivo tecnico",
		"ejecutivo_tecnico": "ejecutivo tecnico",
		"especialista": "especialista",
	}
	return role_aliases.get(role_key, raw_value.lower())


def _extract_norm_token(value: str) -> str | None:
	match = re.search(r"NOM-\d{3}", value.upper())
	return match.group(0) if match else None


def _normalize_norm_key(value: str | None) -> str:
	raw_value = str(value or "").strip()
	if not raw_value:
		return "SIN_NORMA"

	token = _extract_norm_token(raw_value)
	if token:
		return token

	compact = re.sub(r"\s+", " ", raw_value.upper())
	return compact or "SIN_NORMA"


def _coerce_score(value: Any) -> float | None:
	if value in (None, ""):
		return None

	try:
		return round(float(value), 2)
	except (TypeError, ValueError):
		return None


def _normalize_supervision_answers(raw_answers: Any) -> list[dict[str, str]]:
	normalized: list[dict[str, str]] = []
	if not isinstance(raw_answers, list):
		return normalized

	for item in raw_answers:
		if not isinstance(item, dict):
			continue

		activity = str(item.get("activity", "")).strip()
		if not activity:
			continue

		result = str(item.get("result", "")).strip().lower()
		if result not in {"conforme", "no_conforme", "no_aplica"}:
			result = ""

		normalized.append(
			{
				"activity": activity,
				"result": result,
				"observations": str(item.get("observations", "")).strip(),
			}
		)

	return normalized


def _normalize_technical_normative_rows(raw_rows: Any) -> list[dict[str, str]]:
	normalized: list[dict[str, str]] = []
	if not isinstance(raw_rows, list):
		return normalized

	for item in raw_rows:
		if not isinstance(item, dict):
			continue

		sku = str(item.get("sku", "")).strip()
		applicable_norm = str(item.get("applicable_norm", "")).strip()
		observations = str(item.get("observations", "")).strip()
		result = str(item.get("result", "")).strip().lower()
		c_nc = str(item.get("c_nc", "")).strip().upper()

		if result not in {"conforme", "no_conforme"}:
			if c_nc == "C":
				result = "conforme"
			elif c_nc == "NC":
				result = "no_conforme"
			else:
				result = ""

		if not c_nc:
			if result == "conforme":
				c_nc = "C"
			elif result == "no_conforme":
				c_nc = "NC"

		has_any_value = bool(sku or applicable_norm or result or c_nc or observations)
		if not has_any_value:
			continue

		normalized.append(
			{
				"sku": sku,
				"applicable_norm": applicable_norm,
				"result": result,
				"c_nc": c_nc,
				"observations": observations,
			}
		)

	return normalized


def _timestamp() -> str:
	return datetime.now().strftime("%Y-%m-%d %H:%M")


def _format_criterio_resolution_number(counter: int) -> str:
	return f"{max(1, int(counter)):04d}"


def _normalize_visit_date(raw_value: str) -> str:
	value = str(raw_value or "").strip()
	if not value:
		return ""

	for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
		try:
			return datetime.strptime(value, fmt).strftime("%Y-%m-%d")
		except ValueError:
			continue
	return ""


def _normalize_visit_time(raw_value: str) -> str:
	value = str(raw_value or "").strip().upper().replace(".", "")
	if not value:
		return ""

	for fmt in ("%H:%M", "%H%M", "%I:%M %p", "%I:%M%p", "%I %p"):
		try:
			return datetime.strptime(value, fmt).strftime("%H:%M")
		except ValueError:
			continue
	return ""


def _normalize_visit_inspectors(raw_values: Any, fallback: str = "") -> list[str]:
	entries: list[Any] = []
	if fallback:
		entries.append(fallback)
	if isinstance(raw_values, list):
		entries.extend(raw_values)
	elif raw_values not in (None, ""):
		entries.append(raw_values)

	normalized: list[str] = []
	seen: set[str] = set()
	for entry in entries:
		name = str(entry or "").strip()
		if not name:
			continue
		key = name.lower()
		if key in seen:
			continue
		normalized.append(name)
		seen.add(key)
	return normalized


def _normalize_visit_acceptance_responses(raw_value: Any) -> dict[str, dict[str, str]]:
	responses: dict[str, dict[str, str]] = {}
	if not isinstance(raw_value, dict):
		return responses

	for raw_name, payload in raw_value.items():
		name = str(raw_name or "").strip()
		if not name:
			continue

		confirmed_at = ""
		confirmed_by = name
		if isinstance(payload, dict):
			confirmed_at = str(payload.get("confirmed_at", "")).strip()
			confirmed_by = str(payload.get("confirmed_by", name)).strip() or name
		else:
			confirmed_at = str(payload or "").strip()

		if not confirmed_at:
			continue
		responses[name] = {
			"confirmed_at": confirmed_at,
			"confirmed_by": confirmed_by,
		}

	return responses


def _normalize_visit_record(raw_visit: dict[str, Any], include_display: bool = False) -> dict[str, Any]:
	visit = dict(raw_visit)
	normalized_date = _normalize_visit_date(str(visit.get("visit_date", "")))
	if normalized_date:
		visit["visit_date"] = normalized_date

	assignment_time = _normalize_visit_time(str(visit.get("assignment_time", "")))
	departure_time = _normalize_visit_time(str(visit.get("departure_time", "")))
	visit["assignment_time"] = assignment_time
	visit["departure_time"] = departure_time

	status = str(visit.get("status", "")).strip() or "Programada"
	visit["status"] = "Programada" if status == "En ruta" else status

	inspectors = _normalize_visit_inspectors(
		visit.get("inspectors"),
		str(visit.get("inspector", "")),
	)
	visit["inspectors"] = inspectors
	visit["inspector"] = inspectors[0] if inspectors else ""
	visit["group_id"] = str(visit.get("group_id", "")).strip()

	normalized_responses = _normalize_visit_acceptance_responses(visit.get("acceptance_responses"))
	responses_by_identity = {
		_normalize_person_name(name): (name, payload)
		for name, payload in normalized_responses.items()
		if _normalize_person_name(name)
	}
	filtered_responses: dict[str, dict[str, str]] = {}
	for inspector_name in inspectors:
		identity = _normalize_person_name(inspector_name)
		if not identity:
			continue
		mapped = responses_by_identity.get(identity)
		if mapped is None:
			continue
		_, payload = mapped
		filtered_responses[inspector_name] = dict(payload)
	visit["acceptance_responses"] = filtered_responses
	visit["accepted_count"] = len(filtered_responses)
	visit["required_acceptances"] = len(inspectors)
	if filtered_responses:
		visit["accepted_by"] = ", ".join(filtered_responses.keys())
	else:
		visit["accepted_by"] = ""

	if include_display:
		visit["inspectors_text"] = ", ".join(inspectors) if inspectors else "--"
		if assignment_time and departure_time:
			visit["schedule_text"] = f"{assignment_time} - {departure_time}"
		elif assignment_time:
			visit["schedule_text"] = f"{assignment_time} - --"
		elif departure_time:
			visit["schedule_text"] = f"-- - {departure_time}"
		else:
			visit["schedule_text"] = "--"
	return visit


def _visit_group_key(visit: dict[str, Any]) -> tuple[Any, ...]:
	group_id = str(visit.get("group_id", "")).strip()
	if group_id and len(visit.get("inspectors", [])) > 1:
		return ("group", group_id)
	return (
		"legacy",
		visit.get("visit_date", ""),
		visit.get("assignment_time", ""),
		visit.get("departure_time", ""),
		str(visit.get("client", "")).strip().casefold(),
		str(visit.get("address", "")).strip().casefold(),
		str(visit.get("service", "")).strip().casefold(),
		str(visit.get("status", "")).strip().casefold(),
		str(visit.get("notes", "")).strip().casefold(),
		str(visit.get("assigned_by", "")).strip().casefold(),
	)


def _merge_visit_records(raw_visits: list[dict[str, Any]], include_display: bool = False) -> list[dict[str, Any]]:
	grouped: dict[tuple[Any, ...], dict[str, Any]] = {}
	for raw_visit in raw_visits:
		visit = _normalize_visit_record(raw_visit)
		key = _visit_group_key(visit)
		existing = grouped.get(key)
		if existing is None:
			grouped[key] = visit
			continue

		merged_inspectors = list(existing.get("inspectors", []))
		existing_keys = {name.lower() for name in merged_inspectors}
		for inspector_name in visit.get("inspectors", []):
			key_name = inspector_name.lower()
			if key_name in existing_keys:
				continue
			merged_inspectors.append(inspector_name)
			existing_keys.add(key_name)

		existing["inspectors"] = merged_inspectors
		existing["inspector"] = merged_inspectors[0] if merged_inspectors else ""

		existing_responses = _normalize_visit_acceptance_responses(existing.get("acceptance_responses"))
		incoming_responses = _normalize_visit_acceptance_responses(visit.get("acceptance_responses"))
		for inspector_name, payload in incoming_responses.items():
			current_payload = existing_responses.get(inspector_name)
			if current_payload is None or str(payload.get("confirmed_at", "")) >= str(current_payload.get("confirmed_at", "")):
				existing_responses[inspector_name] = payload
		existing["acceptance_responses"] = existing_responses
		if existing_responses and str(existing.get("acceptance_status", "")).strip().lower() not in {"cancelada", "finalizada"}:
			existing["acceptance_status"] = "aceptada"

	return [_normalize_visit_record(visit, include_display=include_display) for visit in grouped.values()]


def _visit_week_dir(date_iso: str) -> Path:
	"""Returns the weekly folder path (semana_YYYY-MM-DD, Monday-based) for a visit date."""
	try:
		d = datetime.strptime(date_iso, "%Y-%m-%d").date()
	except ValueError:
		d = date.today()
	monday = d - timedelta(days=d.weekday())
	return VISITS_DIR / f"semana_{monday.strftime('%Y-%m-%d')}"


def _quarter_dir(quarter: str, year: int) -> Path:
	"""Returns the folder path for a given quarter+year (e.g. T1_2026)."""
	return TRIMESTRAL_DIR / f"{quarter}_{year}"


def _read_all_visits() -> list[dict[str, Any]]:
	"""Read all visits from all weekly folders in VISITS_DIR."""
	if not VISITS_DIR.exists():
		return []
	visits: list[dict[str, Any]] = []
	for folder in sorted(VISITS_DIR.iterdir()):
		if folder.is_dir():
			visits.extend(_read_json(folder / "visitas.json", []))
	return _merge_visit_records(visits)


def _read_all_quarterly_scores() -> list[dict[str, Any]]:
	"""Read all quarterly scores from all quarter folders in TRIMESTRAL_DIR."""
	if not TRIMESTRAL_DIR.exists():
		return []
	scores: list[dict[str, Any]] = []
	for folder in sorted(TRIMESTRAL_DIR.iterdir()):
		if folder.is_dir():
			scores.extend(_read_json(folder / "trimestral.json", []))
	return scores


def _default_state() -> dict[str, Any]:
	return {
		"evaluations": {},
		"document_counters": {"criterio_evaluacion_tecnica": 0},
		"criteria_documents": [],
	}


@lru_cache(maxsize=4)
def _load_module(module_path: str):
	path = Path(module_path)

	# Ensure the app root is on sys.path so modules loaded from
	# sub-directories (e.g. "Documentos PDF.py/") can resolve
	# top-level imports like ``runtime_paths`` in frozen builds.
	import sys
	root = str(resource_path("."))
	if root not in sys.path:
		sys.path.insert(0, root)

	spec = importlib.util.spec_from_file_location(path.stem, path)
	if spec is None or spec.loader is None:
		raise ImportError(f"No se pudo cargar el modulo: {path}")

	module = importlib.util.module_from_spec(spec)
	spec.loader.exec_module(module)
	return module


class CalibrationController:
	def delete_client_agreement_pdf(self, client_name: str, pdf_filename: str) -> bool:
		"""Delete a specific agreement PDF for a client. Returns True if deleted."""
		target_dir = self._agreement_client_dir(str(client_name))
		pdf_path = target_dir / pdf_filename
		if pdf_path.exists() and pdf_path.is_file():
			pdf_path.unlink()
			return True
		return False

	def delete_criteria_document(self, resolution_number: str) -> bool:
		"""Delete a criteria document by its resolution number from app_state and filesystem."""
		documents_log = self.app_state.get("criteria_documents", [])
		idx = next((i for i, doc in enumerate(documents_log)
					if str(doc.get("resolution_number")) == str(resolution_number)), None)
		if idx is None:
			return False
		doc = documents_log[idx]
		# Remove file from filesystem
		output_path = doc.get("output_path")
		if output_path:
			try:
				path = Path(output_path)
				if path.exists() and path.is_file():
					path.unlink()
			except Exception:
				pass
		# Remove from log and save
		del documents_log[idx]
		self.app_state["criteria_documents"] = documents_log
		_write_json(STATE_FILE, self.app_state)
		self.reload()
		return True
	def __init__(self, root_dir: Path | None = None) -> None:
		_seed_runtime_data_dir()
		self.root_dir = root_dir or ROOT_DIR
		self.current_user: dict[str, Any] | None = None
		self.raw_records: list[dict[str, Any]] = []
		self.norms_catalog: list[dict[str, Any]] = []
		self.users_catalog: list[dict[str, Any]] = []
		self.clients_catalog: list[dict[str, Any]] = []
		self.app_state: dict[str, Any] = _default_state()
		self._consolidation_done: bool = False
		self._visits_normalized: bool = False
		self._history_dir_index: dict[str, Path] = {}
		self._record_index: dict[str, dict[str, Any]] = {}
		self._record_index_normalized: dict[str, dict[str, Any]] = {}
		self._history_cache: dict[str, list[dict[str, Any]]] = {}
		self._visits_cache: list[dict[str, Any]] | None = None
		self._quarterly_scores_cache: list[dict[str, Any]] | None = None
		self._catalog_norms_cache: list[dict[str, str]] | None = None
		self._dashboard_people_cache: list[str] | None = None
		self._assignable_inspectors_cache: list[str] | None = None
		self._client_names_cache: list[str] | None = None
		self._norm_card_metrics_cache: list[dict[str, Any]] | None = None
		self._latest_evaluation_cache: dict[str, dict[str, Any]] = {}
		self._latest_evaluation_by_norm_cache: dict[tuple[str, str], dict[str, Any]] = {}
		self._principal_rows_cache: dict[tuple[str, str], list[dict[str, Any]]] = {}
		self._overview_metrics_cache: dict[str, Any] | None = None
		self._executive_profile_cache: dict[str, dict[str, Any]] = {}
		self.reload()

	def reload(self) -> None:
		self._reset_runtime_caches()
		self.raw_records = _read_json(BD_FILE, [])
		self.norms_catalog = _read_json(NORMS_FILE, [])
		self.clients_catalog = _read_json(CLIENTS_FILE, [])

		users_payload = _read_json(USERS_FILE, {"users": []})
		self.users_catalog = users_payload.get("users", [])

		self.app_state = _read_json(STATE_FILE, _default_state())
		self.app_state.setdefault("evaluations", {})
		self.app_state.setdefault("document_counters", {"criterio_evaluacion_tecnica": 0})
		self.app_state.setdefault("criteria_documents", [])
		self.app_state.setdefault("vacations", [])
		self.app_state.setdefault("workshops", [])
		HISTORY_DIR.mkdir(parents=True, exist_ok=True)
		VISITS_DIR.mkdir(parents=True, exist_ok=True)
		TRIMESTRAL_DIR.mkdir(parents=True, exist_ok=True)
		if not NORMS_REPORT_FILE.exists():
			_write_json(NORMS_REPORT_FILE, [])
		self._migrate_legacy_state()
		if not self._visits_normalized:
			self._normalize_visit_storage()
			self._visits_normalized = True
		if not self._consolidation_done:
			self._consolidate_history_storage()
			self._consolidation_done = True
		self._rebuild_runtime_indexes()

	def _reset_runtime_caches(self) -> None:
		self._history_dir_index = {}
		self._record_index = {}
		self._record_index_normalized = {}
		self._history_cache = {}
		self._visits_cache = None
		self._quarterly_scores_cache = None
		self._catalog_norms_cache = None
		self._dashboard_people_cache = None
		self._assignable_inspectors_cache = None
		self._client_names_cache = None
		self._norm_card_metrics_cache = None
		self._latest_evaluation_cache = {}
		self._latest_evaluation_by_norm_cache = {}
		self._principal_rows_cache = {}
		self._overview_metrics_cache = None
		self._executive_profile_cache = {}

	def _rebuild_runtime_indexes(self) -> None:
		for record in self.raw_records:
			name = str(record.get("NOMBRE", "")).strip()
			if not name:
				continue
			self._record_index[name] = record
			normalized_name = _normalize_person_name(name)
			if normalized_name and normalized_name not in self._record_index_normalized:
				self._record_index_normalized[normalized_name] = record

		if HISTORY_DIR.exists():
			for child in HISTORY_DIR.iterdir():
				if not child.is_dir():
					continue
				identity = _folder_identity(child.name)
				if identity and identity not in self._history_dir_index:
					self._history_dir_index[identity] = child

		for payload in self.app_state.get("evaluations", {}).values():
			if not isinstance(payload, dict):
				continue
			inspector_name = str(
				payload.get("inspector_name")
				or payload.get("inspector_supervised")
				or ""
			).strip()
			if not inspector_name:
				continue
			self._cache_latest_evaluation_entry(inspector_name, payload)

	def _cache_latest_evaluation_entry(self, inspector_name: str, payload: dict[str, Any]) -> None:
		clean_name = str(inspector_name or "").strip()
		if not clean_name or not isinstance(payload, dict):
			return

		current_latest = self._latest_evaluation_cache.get(clean_name)
		payload_saved_at = str(payload.get("saved_at", ""))
		if current_latest is None or payload_saved_at >= str(current_latest.get("saved_at", "")):
			self._latest_evaluation_cache[clean_name] = payload

		norm_key = _normalize_norm_key(payload.get("selected_norm"))
		current_by_norm = self._latest_evaluation_by_norm_cache.get((clean_name, norm_key))
		if current_by_norm is None or payload_saved_at >= str(current_by_norm.get("saved_at", "")):
			self._latest_evaluation_by_norm_cache[(clean_name, norm_key)] = payload

	def _get_all_visits_cached(self) -> list[dict[str, Any]]:
		if self._visits_cache is None:
			visits = [_normalize_visit_record(visit, include_display=True) for visit in _read_all_visits()]
			visits.sort(
				key=lambda item: (
					item.get("visit_date", ""),
					item.get("assignment_time", ""),
					item.get("updated_at", ""),
				),
				reverse=True,
			)
			self._visits_cache = visits
		return self._visits_cache

	def _get_all_quarterly_scores_cached(self) -> list[dict[str, Any]]:
		if self._quarterly_scores_cache is None:
			scores = _read_all_quarterly_scores()
			scores.sort(
				key=lambda item: (
					int(item.get("year", 0)),
					str(item.get("quarter", "")),
					str(item.get("updated_at", "")),
				),
				reverse=True,
			)
			self._quarterly_scores_cache = scores
		return self._quarterly_scores_cache

	def _migrate_legacy_state(self) -> None:
		"""One-time migration: moves visits and quarterly_scores from app_state.json
		to separate weekly/quarterly JSON files."""
		dirty = False

		old_visits = self.app_state.pop("visits", None)
		if old_visits:
			for visit in old_visits:
				date_iso = (
					_normalize_visit_date(str(visit.get("visit_date", "")))
					or date.today().strftime("%Y-%m-%d")
				)
				week_dir = _visit_week_dir(date_iso)
				existing = _read_json(week_dir / "visitas.json", [])
				vid = visit.get("id")
				if vid and not any(v.get("id") == vid for v in existing):
					existing.append(visit)
					_write_json(week_dir / "visitas.json", existing)
			dirty = True

		old_scores = self.app_state.pop("quarterly_scores", None)
		if old_scores:
			for score in old_scores:
				q = str(score.get("quarter", "T1")).upper()
				y = int(score.get("year", date.today().year))
				q_dir = _quarter_dir(q, y)
				existing = _read_json(q_dir / "trimestral.json", [])
				sid = score.get("id")
				if sid and not any(s.get("id") == sid for s in existing):
					existing.append(score)
					_write_json(q_dir / "trimestral.json", existing)
			dirty = True

		if dirty:
			_write_json(STATE_FILE, self.app_state)

	def _normalize_visit_storage(self) -> None:
		if not VISITS_DIR.exists():
			return

		for week_dir in VISITS_DIR.iterdir():
			if not week_dir.is_dir():
				continue
			visits_path = week_dir / "visitas.json"
			raw_visits = _read_json(visits_path, [])
			normalized_visits = _merge_visit_records(raw_visits)
			if normalized_visits != raw_visits:
				_write_json(visits_path, normalized_visits)

	def _redistribute_misplaced_history(self) -> None:
		"""Scan all historico.json files and move entries to the correct inspector folder."""
		if not HISTORY_DIR.exists():
			return
		moved: dict[str, list[dict[str, Any]]] = {}
		for child in list(HISTORY_DIR.iterdir()):
			if not child.is_dir():
				continue
			history_file = child / "historico.json"
			if not history_file.exists():
				continue
			entries = _read_json(history_file, [])
			if not isinstance(entries, list) or not entries:
				continue
			folder_identity = _folder_identity(child.name)
			keep: list[dict[str, Any]] = []
			for entry in entries:
				if not isinstance(entry, dict):
					continue
				entry_name = str(entry.get("inspector_name", "")).strip()
				entry_identity = _normalize_person_name(entry_name) if entry_name else ""
				if not entry_identity or entry_identity == folder_identity:
					keep.append(entry)
				else:
					canonical = self._resolve_canonical_person_name(entry_name) or entry_name
					target_folder = _safe_folder_name(canonical)
					moved.setdefault(target_folder, []).append(entry)
			if len(keep) != len(entries):
				_write_json(history_file, keep)
		for target_folder, entries in moved.items():
			target_dir = HISTORY_DIR / target_folder
			target_dir.mkdir(parents=True, exist_ok=True)
			target_file = target_dir / "historico.json"
			existing = _read_json(target_file, [])
			if not isinstance(existing, list):
				existing = []
			seen: set[str] = set()
			merged: list[dict[str, Any]] = []
			for item in [*existing, *entries]:
				try:
					key = json.dumps(item, ensure_ascii=False, sort_keys=True)
				except TypeError:
					key = str(item)
				if key not in seen:
					seen.add(key)
					merged.append(item)
			_write_json(target_file, merged)

	def _consolidate_history_storage(self) -> None:
		if not HISTORY_DIR.exists():
			return

		self._redistribute_misplaced_history()

		# Build identity → folders map in one pass (name-based, no JSON reads)
		existing_dirs = [child for child in HISTORY_DIR.iterdir() if child.is_dir()]
		identity_map: dict[str, list[Path]] = {}
		for child in existing_dirs:
			ident = _folder_identity(child.name)
			if ident:
				identity_map.setdefault(ident, []).append(child)

		candidate_names: set[str] = {
			str(record.get("NOMBRE", "")).strip()
			for record in self.raw_records
			if str(record.get("NOMBRE", "")).strip()
		}
		candidate_names.update(
			str(user.get("name", "")).strip()
			for user in self.users_catalog
			if str(user.get("name", "")).strip()
		)
		candidate_names.update(
			child.name.replace("_", " ").strip()
			for child in existing_dirs
			if child.name.strip()
		)

		for name in sorted(candidate_names):
			if not name:
				continue
			preferred = HISTORY_DIR / _safe_folder_name(name)
			ident = _folder_identity(name)
			folders_for_ident = identity_map.get(ident, [])
			# Skip if preferred folder already exists and is the only match — nothing to merge
			if (
				preferred.exists()
				and len(folders_for_ident) == 1
				and folders_for_ident[0].resolve() == preferred.resolve()
			):
				continue
			self._history_dir(name)

	def authenticate(self, username: str, password: str) -> dict[str, Any] | None:
		normalized_user = username.strip().lower()
		normalized_password = password.strip()

		for user in self.users_catalog:
			if user.get("username", "").strip().lower() != normalized_user:
				continue
			if user.get("password", "").strip() != normalized_password:
				continue

			self.current_user = dict(user)
			self.current_user["role"] = _normalize_role_name(self.current_user.get("role"))
			return self.current_user

		return None

	def logout(self) -> None:
		self.current_user = None

	def is_admin(self, user: dict[str, Any] | None = None) -> bool:
		return self.has_full_access(user)

	def _role_name(self, user: dict[str, Any] | None = None) -> str:
		candidate = user or self.current_user or {}
		return _normalize_role_name(candidate.get("role"))

	def has_full_access(self, user: dict[str, Any] | None = None) -> bool:
		return self._role_name(user) in {
			"admin",
			"gerente",
			"sub gerente",
			"coordinador operativo",
			"coordinadora en fiabilidad",
		}

	def is_executive_role(self, user: dict[str, Any] | None = None) -> bool:
		return self._role_name(user) in {"ejecutivo tecnico", "especialidades"}
	
	## --- Roles y control de vistas.
	def available_sections(self, user: dict[str, Any] | None = None) -> list[str]:
		if self.has_full_access(user):
			return ["Supervisión", "Criterios", "Dashboard", "Calendario", "Trimestral", "Configuraciones"]
		role_name = self._role_name(user)
		if role_name == "talento humano":
			return ["Supervisión", "Dashboard", "Calendario"]
		if role_name == "supervisor":
			return ["Supervisión","Criterios", "Dashboard", "Trimestral", "Calendario"]
		if self.is_executive_role(user):
			return ["Calendario", "Criterios", "Trimestral", "Supervisión"]
		return ["Calendario", "Trimestral"]

	def _resolve_canonical_person_name(self, value: str | None) -> str:
		clean_name = str(value or "").strip()
		if not clean_name:
			return ""

		normalized_target = _normalize_person_name(clean_name)
		if not normalized_target:
			return clean_name

		candidate_names: list[str] = []
		candidate_names.extend(
			str(user.get("name", "")).strip()
			for user in self.users_catalog
			if str(user.get("name", "")).strip()
		)
		candidate_names.extend(
			str(record.get("NOMBRE", "")).strip()
			for record in self.raw_records
			if str(record.get("NOMBRE", "")).strip()
		)

		target_tokens = normalized_target.split()
		best_match = clean_name
		best_token_count = len(target_tokens)
		for candidate in candidate_names:
			normalized_candidate = _normalize_person_name(candidate)
			candidate_tokens = normalized_candidate.split()
			if len(candidate_tokens) <= best_token_count:
				continue
			if candidate_tokens[: len(target_tokens)] == target_tokens:
				best_match = candidate
				best_token_count = len(candidate_tokens)

		if best_match != clean_name:
			return best_match

		for candidate in candidate_names:
			if _normalize_person_name(candidate) == normalized_target:
				return candidate

		return best_match

	def get_record(self, inspector_name: str) -> dict[str, Any] | None:
		target_name = str(inspector_name or "").strip()
		if not target_name:
			return None

		record = self._record_index.get(target_name)
		if record is not None:
			return record

		normalized_target = _normalize_person_name(target_name)
		if not normalized_target:
			return None

		record = self._record_index_normalized.get(normalized_target)
		if record is not None:
			return record

		target_tokens = normalized_target.split()
		best_match: dict[str, Any] | None = None
		best_token_count = 0
		for candidate_identity, candidate_record in self._record_index_normalized.items():
			candidate_tokens = candidate_identity.split()
			if not candidate_tokens:
				continue
			prefix_match = candidate_tokens[: len(target_tokens)] == target_tokens
			reverse_prefix_match = target_tokens[: len(candidate_tokens)] == candidate_tokens
			if not prefix_match and not reverse_prefix_match:
				continue
			if len(candidate_tokens) > best_token_count:
				best_match = candidate_record
				best_token_count = len(candidate_tokens)

		return best_match

	def get_catalog_norms(self) -> list[dict[str, str]]:
		if self._catalog_norms_cache is not None:
			return [dict(item) for item in self._catalog_norms_cache]

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

		self._catalog_norms_cache = sorted(catalog, key=lambda item: self._norm_sort_key(item["token"]))
		return [dict(item) for item in self._catalog_norms_cache]

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
		if self._dashboard_people_cache is not None:
			return list(self._dashboard_people_cache)

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
		self._dashboard_people_cache = sorted(names)
		return list(self._dashboard_people_cache)

	def get_assignable_inspectors(self) -> list[str]:
		if self._assignable_inspectors_cache is not None:
			return list(self._assignable_inspectors_cache)

		# Roles a los que se les asignan visitas
		# Incluir ejecutivos técnicos, especialistas y supervisores
		assignable_roles = {"ejecutivo tecnico", "especialista", "supervisor"}
		assignables = sorted(
			str(user.get("name", "")).strip()
			for user in self.users_catalog
			if _normalize_role_name(user.get("role")) in assignable_roles
			and str(user.get("name", "")).strip()
		)
		self._assignable_inspectors_cache = assignables or self.get_dashboard_people()
		return list(self._assignable_inspectors_cache)

	def get_busy_executives(self, date_iso: str, exclude_visit_id: str | None = None) -> set[str]:
		"""Returns names of executives who already have a visit assigned on date_iso."""
		busy: set[str] = set()
		for visit in self._get_all_visits_cached():
			if exclude_visit_id and visit.get("id") == exclude_visit_id:
				continue
			if _normalize_visit_date(str(visit.get("visit_date", ""))) == date_iso:
				busy.update(visit.get("inspectors", []))
		return busy

	def get_available_executives(self, date_iso: str, exclude_visit_id: str | None = None) -> list[str]:
		"""Returns executives not yet assigned to any visit on date_iso."""
		busy = self.get_busy_executives(date_iso, exclude_visit_id)
		return [e for e in self.get_assignable_inspectors() if e not in busy]

	def get_client_names(self) -> list[str]:
		if self._client_names_cache is not None:
			return list(self._client_names_cache)

		names = {
			str(client.get("CLIENTE", "")).strip()
			for client in self.clients_catalog
			if str(client.get("CLIENTE", "")).strip()
		}
		self._client_names_cache = sorted(names)
		return list(self._client_names_cache)

	def get_client_addresses(self, client_name: str) -> list[dict[str, str]]:
		options: list[dict[str, str]] = []
		for client in self.clients_catalog:
			if str(client.get("CLIENTE", "")).strip() != client_name:
				continue

			for index, address in enumerate(client.get("DIRECCIONES", []), start=1):
				warehouse = str(address.get("ALMACEN", "")).strip()
				street = str(address.get("CALLE Y NO", "")).strip()
				colony = str(address.get("COLONIA O POBLACION", "")).strip()
				municipality = str(address.get("MUNICIPIO O ALCADIA", "")).strip()
				state = str(address.get("CIUDAD O ESTADO", "")).strip()
				postal_code = str(address.get("CP", "")).strip()
				service = str(address.get("SERVICIO", "")).strip()
				location_text = ", ".join(
					part for part in [street, colony, municipality, state, postal_code] if part
				)
				address_text = " | ".join(part for part in [warehouse, location_text] if part)
				label = f"Sede {index}: {warehouse or municipality or state or 'Sin ubicacion'} | {service or 'Sin servicio'}"
				options.append(
					{
						"label": label,
						"address": address_text,
						"location": location_text,
						"warehouse": warehouse,
						"service": service or "Sin servicio",
					}
				)

		return options

	def get_client_warehouse_for_address(self, client_name: str, address_text: str) -> str:
		client_name = str(client_name).strip()
		address_text = str(address_text).strip()
		if not client_name or not address_text:
			return ""

		for client in self.clients_catalog:
			if str(client.get("CLIENTE", "")).strip() != client_name:
				continue

			for address in client.get("DIRECCIONES", []):
				if not isinstance(address, dict):
					continue

				warehouse = str(address.get("ALMACEN", "")).strip()
				location_text = ", ".join(
					part for part in [
						str(address.get("CALLE Y NO", "")).strip(),
						str(address.get("COLONIA O POBLACION", "")).strip(),
						str(address.get("MUNICIPIO O ALCADIA", "")).strip(),
						str(address.get("CIUDAD O ESTADO", "")).strip(),
						str(address.get("CP", "")).strip(),
					] if part
				)
				full_text = " | ".join(part for part in [warehouse, location_text] if part)

				if address_text == full_text or address_text == location_text or (location_text and location_text in address_text):
					return warehouse

		return ""

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
		clean_name = str(inspector_name or "").strip()
		if not clean_name:
			return []

		cached_history = self._history_cache.get(clean_name)
		if cached_history is None:
			history_path = self._history_file(clean_name)
			history = _read_json(history_path, [])
			if not isinstance(history, list):
				history = []
			target_identity = _normalize_person_name(clean_name)
			filtered: list[dict[str, Any]] = []
			for item in history:
				if not isinstance(item, dict):
					continue
				entry_name = str(item.get("inspector_name", "")).strip()
				if entry_name and _normalize_person_name(entry_name) != target_identity:
					continue
				filtered.append(item)

			# Also read individual JSON files from FORMATOS DE SUPERVISION
			formatos_dir = self._history_dir(clean_name) / "FORMATOS DE SUPERVISION"
			if formatos_dir.exists() and formatos_dir.is_dir():
				existing_ids = {item.get("saved_at", "") for item in filtered if isinstance(item, dict)}
				for json_file in formatos_dir.glob("*.json"):
					try:
						entry = _read_json(json_file, None)
						if not isinstance(entry, dict):
							continue
						entry_name = str(entry.get("inspector_name", "")).strip()
						if entry_name and _normalize_person_name(entry_name) != target_identity:
							continue
						entry_saved = entry.get("saved_at", "")
						entry_norm = entry.get("selected_norm", "")
						dup_key = (entry_saved, entry_norm)
						if entry_saved and any(
							(item.get("saved_at", ""), item.get("selected_norm", "")) == dup_key
							for item in filtered
						):
							continue
						entry["archivo_path"] = str(json_file)
						filtered.append(entry)
					except Exception:
						continue

			cached_history = sorted(filtered, key=lambda item: item.get("saved_at", ""))
			self._history_cache[clean_name] = cached_history
			if cached_history:
				self._cache_latest_evaluation_entry(clean_name, cached_history[-1])

		return [dict(item) for item in cached_history if isinstance(item, dict)]

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
				status = "Feedback"
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
				"actions_text": "Supervisar",
			}

			searchable_blob = f"{name} {row['norms_text']} {status}".lower()
			if lowered_query and lowered_query not in searchable_blob:
				continue
			if status_filter == "Completos" and not row["form_completed"]:
				continue
			if status_filter == "Pendientes" and row["form_completed"]:
				continue
			if status_filter == "Feedback" and row["status"] != "Feedback":
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
		if self._norm_card_metrics_cache is not None:
			return [dict(item) for item in self._norm_card_metrics_cache]

		catalog = self.get_catalog_norms()
		counts = {item["token"]: 0 for item in catalog}
		for record in self.raw_records:
			for token in self.get_accredited_norms(record):
				counts[token] = counts.get(token, 0) + 1

		self._norm_card_metrics_cache = [
			{
				"token": item["token"],
				"label": item["nom"],
				"description": item["nombre"],
				"count": counts.get(item["token"], 0),
			}
			for item in catalog
		]
		return [dict(item) for item in self._norm_card_metrics_cache]

	def get_executive_profile(self, inspector_name: str) -> dict[str, Any]:
		clean_name = str(inspector_name or "").strip()
		if not clean_name:
			return {}

		cached_profile = self._executive_profile_cache.get(clean_name)
		if cached_profile is not None:
			return {
				**cached_profile,
				"accredited_norms": list(cached_profile.get("accredited_norms", [])),
				"history": [dict(point) for point in cached_profile.get("history", [])],
				"recent_visits": [dict(visit) for visit in cached_profile.get("recent_visits", [])],
			}

		history = self.get_history(clean_name)
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

		profile = {
			"name": clean_name,
			"accredited_norms": self.get_accredited_norms(clean_name),
			"history": chart_points,
			"latest_score": latest_score,
			"average_score": average_score,
			"focus_required": focus_required,
			"latest_status": history[-1].get("status", "Sin seguimiento") if history else "Sin seguimiento",
			"recent_visits": self.get_recent_visits(clean_name),
		}
		self._executive_profile_cache[clean_name] = profile
		return {
			**profile,
			"accredited_norms": list(profile.get("accredited_norms", [])),
			"history": [dict(point) for point in profile.get("history", [])],
			"recent_visits": [dict(visit) for visit in profile.get("recent_visits", [])],
		}

	def save_principal_record(
		self,
		name: str,
		accredited_norms: list[str],
		original_name: str | None = None,
	) -> dict[str, Any]:
		clean_name = name.strip()
		if not clean_name:
			raise ValueError("El nombre del ejecutivo tecnico es obligatorio.")

		if original_name:
			record = self.get_record(original_name)
			if record is None:
				raise ValueError("No se encontro el registro a editar.")
		else:
			record = None

		duplicate = self.get_record(clean_name)
		if duplicate and clean_name != original_name:
			raise ValueError("Ya existe un ejecutivo tecnico con ese nombre.")

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
		inspector_supervised = str(payload.get("inspector_supervised", inspector_name)).strip() or inspector_name
		protocol_answers = _normalize_supervision_answers(payload.get("protocol_answers"))
		process_answers = _normalize_supervision_answers(payload.get("process_answers"))
		technical_normative_rows = _normalize_technical_normative_rows(payload.get("technical_normative_rows"))
		image_folder = str(payload.get("image_folder", "")).strip()
		score_breakdown = payload.get("score_breakdown", {})
		if not isinstance(score_breakdown, dict):
			score_breakdown = {}
		score_by_norm = payload.get("score_by_norm", {})
		if not isinstance(score_by_norm, dict):
			score_by_norm = {}
		soft_skills_score = _coerce_score(payload.get("soft_skills_score"))
		technical_skills_score = _coerce_score(payload.get("technical_skills_score"))
		soft_skills_breakdown = payload.get("soft_skills_breakdown", {})
		if not isinstance(soft_skills_breakdown, dict):
			soft_skills_breakdown = {}
		technical_skills_breakdown = payload.get("technical_skills_breakdown", {})
		if not isinstance(technical_skills_breakdown, dict):
			technical_skills_breakdown = {}
		normalized_score_by_norm: dict[str, float] = {}
		for norm_name, raw_score in score_by_norm.items():
			clean_norm_name = str(norm_name).strip()
			numeric_score = _coerce_score(raw_score)
			if not clean_norm_name or numeric_score is None:
				continue
			normalized_score_by_norm[clean_norm_name] = numeric_score
		score = _coerce_score(payload.get("score"))

		if not clean_date:
			raise ValueError("La fecha de seguimiento es obligatoria.")
		if not clean_client:
			raise ValueError("Debes seleccionar cliente o almacen.")
		if score is None:
			raise ValueError("El puntaje debe ser numerico.")
		if score < 0 or score > 100:
			raise ValueError("El puntaje debe estar entre 0 y 100.")
		if soft_skills_score is not None and (soft_skills_score < 0 or soft_skills_score > 100):
			raise ValueError("La calificación de habilidades blandas debe estar entre 0 y 100.")
		if technical_skills_score is not None and (technical_skills_score < 0 or technical_skills_score > 100):
			raise ValueError("La calificación de habilidades técnicas debe estar entre 0 y 100.")

		evaluation = {
			"inspector_name": inspector_name,
			"inspector_supervised": inspector_supervised,
			"selected_norm": clean_norm or "Sin norma",
			"client": clean_client,
			"visit_date": clean_date,
			"score": score,
			"soft_skills_score": soft_skills_score,
			"technical_skills_score": technical_skills_score,
			"status": clean_status,
			"observations": clean_observations,
			"corrective_actions": clean_actions,
			"evaluator": evaluator or (self.current_user or {}).get("name", "Sin evaluador"),
			"protocol_answers": protocol_answers,
			"process_answers": process_answers,
			"technical_normative_rows": technical_normative_rows,
			"image_folder": image_folder,
			"score_breakdown": score_breakdown,
			"soft_skills_breakdown": soft_skills_breakdown,
			"technical_skills_breakdown": technical_skills_breakdown,
			"score_by_norm": normalized_score_by_norm,
			"form_structure": "supervision_v2",
			"saved_at": _timestamp(),
			"form_completed": True,
		}

		# Guardar cada formulario en un archivo JSON individual
		from uuid import uuid4
		import json
		from pathlib import Path

		# Carpeta: data/historico/<EJECUTIVO>/FORMATOS DE SUPERVISION
		from pathlib import Path
		import shutil
		base_folder = self._history_dir(inspector_name)
		formatos_folder = base_folder / "FORMATOS DE SUPERVISION"
		formatos_folder.mkdir(parents=True, exist_ok=True)
		ts = datetime.now().strftime("%Y%m%d_%H%M%S")
		unique_id = uuid4().hex[:8]
		norm_slug = str(clean_norm).replace(" ", "_").replace("/", "-")
		base_filename = f"Supervision_{ts}_{norm_slug}_{unique_id}"
		json_path = formatos_folder / f"{base_filename}.json"
		pdf_path = formatos_folder / f"{base_filename}.pdf"
		# Guardar JSON
		with open(json_path, "w", encoding="utf-8") as f:
			json.dump(evaluation, f, ensure_ascii=False, indent=2)

		# Si el PDF ya fue generado en otro flujo, moverlo aquí si es necesario
		# Si no, el flujo de generación de PDF debe usar pdf_path como destino
		evaluation["pdf_path"] = str(pdf_path)
		evaluation["json_path"] = str(json_path)

		# Actualizar app_state con el último formulario
		self.app_state["evaluations"][inspector_name] = evaluation
		_write_json(STATE_FILE, self.app_state)
		self.reload()
		return evaluation

	def list_visits(
		self,
		current_user: dict[str, Any] | None = None,
		name: str | None = None,
	) -> list[dict[str, Any]]:
		visits = self._get_all_visits_cached()
		if name:
			visits = [visit for visit in visits if name in visit.get("inspectors", [])]

		candidate = current_user or self.current_user
		if candidate and self.is_executive_role(candidate):
			visits = [visit for visit in visits if candidate.get("name") in visit.get("inspectors", [])]

		return [dict(item) for item in visits]

	def list_trimestral_scores(
		self,
		inspector_name: str | None = None,
		norm: str | None = None,
		year: int | None = None,
		quarter: str | None = None,
		current_user: dict[str, Any] | None = None,
		include_unsent: bool = False,
	) -> list[dict[str, Any]]:
		scores = self._get_all_quarterly_scores_cached()
		candidate = current_user or self.current_user
		is_executive_view = bool(candidate and self.is_executive_role(candidate))
		if not inspector_name and is_executive_view:
			inspector_name = str(candidate.get("name", "")).strip() or None
		if inspector_name:
			target = str(inspector_name).strip()
			target_identity = _normalize_person_name(target)
			if target_identity:
				scores = [
					item
					for item in scores
					if _normalize_person_name(str(item.get("inspector", ""))) == target_identity
				]
			else:
				scores = [item for item in scores if str(item.get("inspector", "")).strip() == target]
		if is_executive_view and not include_unsent:
			scores = [item for item in scores if str(item.get("sent_at", "")).strip()]
		if norm:
			norm_value = _extract_norm_token(str(norm)) or str(norm).strip().upper()
			scores = [
				item
				for item in scores
				if (_extract_norm_token(str(item.get("norm", ""))) or str(item.get("norm", "")).strip().upper()) == norm_value
			]
		if year is not None:
			year_value = int(year)
			scores = [item for item in scores if int(item.get("year", 0)) == year_value]
		if quarter:
			quarter_value = str(quarter).strip().upper()
			scores = [item for item in scores if str(item.get("quarter", "")).strip().upper() == quarter_value]

		enriched_scores: list[dict[str, Any]] = []
		for item in scores:
			if not isinstance(item, dict):
				continue
			row = dict(item)
			medal_payload = self.get_trimestral_medal(row.get("score"))
			row["medal"] = medal_payload["key"]
			row["medal_title"] = medal_payload["title"]
			row["medal_label"] = medal_payload["label"]
			row["medal_message"] = medal_payload["message"]
			enriched_scores.append(row)

		return enriched_scores

	def get_trimestral_medals_summary(
		self,
		inspector_name: str | None = None,
		current_user: dict[str, Any] | None = None,
		include_unsent: bool = True,
	) -> dict[str, Any]:
		scores = self.list_trimestral_scores(
			inspector_name=inspector_name,
			current_user=current_user,
			include_unsent=include_unsent,
		)

		## --- MEDALLAS TRIMESTRALES: ORO >=100, PLATA >=90, BRONCE >=80 --- ##
		# Si no se especifica inspector_name, sumar medallas por ejecutivo y periodo (para admin/coordinación)
		counts = {"ORO": 0, "PLATA": 0, "BRONCE": 0}
		medals_by_period = {}
		if inspector_name:
			# Lógica individual (igual que antes)
			period_map = {}
			for row in scores:
				year = str(row.get("year", "")).strip()
				quarter = str(row.get("quarter", "")).strip().upper()
				if not year or not quarter:
					continue
				key = f"{year}-{quarter}"
				period_map.setdefault(key, []).append(row)
			for period, period_rows in period_map.items():
				califs = [float(item.get("score", 0)) for item in period_rows if item.get("score") is not None]
				if not califs:
					continue
				promedio = sum(califs) / len(califs)
				if promedio >= 100:
					counts["ORO"] += 1
					medals_by_period[period] = "ORO"
				elif promedio >= 90:
					counts["PLATA"] += 1
					medals_by_period[period] = "PLATA"
				elif promedio >= 80:
					counts["BRONCE"] += 1
					medals_by_period[period] = "BRONCE"
				else:
					medals_by_period[period] = "SIN_MEDALLA"
		else:
			# Lógica global: agrupar por inspector y periodo
			inspector_period_map = {}
			for row in scores:
				inspector = str(row.get("inspector", "")).strip()
				year = str(row.get("year", "")).strip()
				quarter = str(row.get("quarter", "")).strip().upper()
				if not inspector or not year or not quarter:
					continue
				key = f"{inspector}|{year}-{quarter}"
				inspector_period_map.setdefault(key, []).append(row)
			for key, period_rows in inspector_period_map.items():
				# key = "inspector|YYYY-Tx"
				period = key.split("|", 1)[-1]
				califs = [float(item.get("score", 0)) for item in period_rows if item.get("score") is not None]
				if not califs:
					continue
				promedio = sum(califs) / len(califs)
				if promedio >= 100:
					counts["ORO"] += 1
					medals_by_period[key] = "ORO"
				elif promedio >= 90:
					counts["PLATA"] += 1
					medals_by_period[key] = "PLATA"
				elif promedio >= 80:
					counts["BRONCE"] += 1
					medals_by_period[key] = "BRONCE"
				else:
					medals_by_period[key] = "SIN_MEDALLA"

		latest = None
		if scores:
			latest = max(
				scores,
				key=lambda item: (
					int(item.get("year", 0) or 0),
					str(item.get("quarter", "")),
					str(item.get("updated_at", "")),
				),
			)
		latest_medal = self.get_trimestral_medal((latest or {}).get("score"))

		return {
			"counts": counts,
			"total": sum(counts.values()),
			"latest_medal": latest_medal,
			"scores_count": len(scores),
			"medals_by_period": medals_by_period,
		}

	@staticmethod
	def get_trimestral_medal(score_value: Any) -> dict[str, str]:
		score = _coerce_score(score_value)
		if score is None:
			return {
				"key": "",
				"title": "Sin medalla",
				"label": "Sin medalla",
				"message": "Sigue mejorando para alcanzar una medalla.",
			}

		for minimum, key, title, message in TRIMESTRAL_MEDAL_RULES:
			if score >= minimum:
				return {
					"key": key,
					"title": title,
					"label": f"{title} {key}",
					"message": message,
				}

		return {
			"key": "",
			"title": "Sin medalla",
			"label": "Sin medalla",
			"message": "Sigue mejorando para alcanzar una medalla.",
		}

	def _boleta_path_for_score(self, score_row: dict[str, Any]) -> Path | None:
		inspector_name = str(score_row.get("inspector", "")).strip()
		quarter = str(score_row.get("quarter", "")).strip().upper()
		try:
			year = int(score_row.get("year", 0))
		except (TypeError, ValueError):
			year = 0

		if not inspector_name or not year or quarter not in {"T1", "T2", "T3", "T4"}:
			return None

		boleta_dir = self._history_dir(inspector_name) / "boletas" / str(year)
		boleta_dir.mkdir(parents=True, exist_ok=True)
		return boleta_dir / f"{quarter}_boleta.json"

	def _upsert_boleta_record(self, score_row: dict[str, Any]) -> None:
		boleta_path = self._boleta_path_for_score(score_row)
		if boleta_path is None:
			return

		score_id = str(score_row.get("id", "")).strip()
		if not score_id:
			return

		boleta_records = _read_json(boleta_path, [])
		if not isinstance(boleta_records, list):
			boleta_records = []
		boleta_records = [row for row in boleta_records if str(row.get("id", "")).strip() != score_id]
		boleta_records.append(dict(score_row))
		_write_json(boleta_path, boleta_records)

	def _remove_boleta_record(self, score_row: dict[str, Any]) -> None:
		boleta_path = self._boleta_path_for_score(score_row)
		if boleta_path is None or not boleta_path.exists():
			return

		score_id = str(score_row.get("id", "")).strip()
		if not score_id:
			return

		boleta_records = _read_json(boleta_path, [])
		if not isinstance(boleta_records, list):
			return
		filtered = [row for row in boleta_records if str(row.get("id", "")).strip() != score_id]
		if filtered != boleta_records:
			_write_json(boleta_path, filtered)

	def send_trimestral_scores(self, inspector_name: str, score_ids: list[str] | None = None) -> int:
		target_identity = _normalize_person_name(inspector_name)
		if not target_identity or not TRIMESTRAL_DIR.exists():
			return 0

		target_ids = {str(item).strip() for item in (score_ids or []) if str(item).strip()}
		sent_count = 0

		for folder in TRIMESTRAL_DIR.iterdir():
			if not folder.is_dir():
				continue

			scores_path = folder / "trimestral.json"
			scores = _read_json(scores_path, [])
			changed = False
			for row in scores:
				if not isinstance(row, dict):
					continue
				score_id = str(row.get("id", "")).strip()
				if target_ids and score_id not in target_ids:
					continue
				if _normalize_person_name(str(row.get("inspector", ""))) != target_identity:
					continue
				if str(row.get("sent_at", "")).strip():
					continue

				row["sent_at"] = _timestamp()
				row["updated_at"] = _timestamp()
				changed = True
				sent_count += 1
				self._upsert_boleta_record(row)

			if changed:
				_write_json(scores_path, scores)

		if sent_count:
			self.reload()

		return sent_count

	def save_trimestral_score(self, payload: dict[str, Any], score_id: str | None = None) -> dict[str, Any]:
		inspector = str(payload.get("inspector", "")).strip()
		norm = _extract_norm_token(str(payload.get("norm", ""))) or str(payload.get("norm", "")).strip().upper()
		quarter = str(payload.get("quarter", "")).strip().upper()
		year_value = str(payload.get("year", "")).strip()
		notes = str(payload.get("notes", "")).strip()
		score = _coerce_score(payload.get("score"))

		if not inspector:
			raise ValueError("Debes seleccionar un ejecutivo tecnico.")
		if not norm:
			raise ValueError("Debes seleccionar la norma de la calificacion trimestral.")
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

		# Find and remove existing score from whichever quarter folder it lives in
		existing = None
		if score_id and TRIMESTRAL_DIR.exists():
			for _folder in TRIMESTRAL_DIR.iterdir():
				if not _folder.is_dir():
					continue
				_fscores = _read_json(_folder / "trimestral.json", [])
				_found = next((s for s in _fscores if s.get("id") == score_id), None)
				if _found:
					existing = _found
					_fscores = [s for s in _fscores if s.get("id") != score_id]
					_write_json(_folder / "trimestral.json", _fscores)
					break

		# Deduplicate: if no explicit score_id, check for same inspector+norm+quarter+year
		if existing is None:
			_insp_norm = _normalize_person_name(inspector)
			_q_path = _quarter_dir(quarter, year) / "trimestral.json"
			if _q_path.exists():
				_q_check = _read_json(_q_path, [])
				_dup = next(
					(
						s for s in _q_check
						if _normalize_person_name(str(s.get("inspector", ""))) == _insp_norm
						and (_extract_norm_token(str(s.get("norm", ""))) or str(s.get("norm", "")).strip().upper()) == norm
						and str(s.get("quarter", "")).strip().upper() == quarter
						and str(s.get("year", "")) == str(year)
					),
					None,
				)
				if _dup is not None:
					existing = _dup
					_write_json(_q_path, [s for s in _q_check if s.get("id") != _dup.get("id")])

		# Lock: prevent modifications to already-sent scores
		if existing is not None and str(existing.get("sent_at", "")).strip():
			raise ValueError("Esta calificacion ya fue enviada y no puede ser modificada.")

		record = dict(existing or {})
		record["id"] = record.get("id") or uuid4().hex
		record["inspector"] = inspector
		record["norm"] = norm
		record["quarter"] = quarter
		record["year"] = year
		record["score"] = score
		record["notes"] = notes
		record["evaluator"] = (self.current_user or {}).get("name", "Sistema")
		record["confirmed_at"] = None
		record["confirmed_by"] = ""
		previous_sent_at = str(record.get("sent_at", "")).strip()
		record["sent_at"] = previous_sent_at if previous_sent_at else ""
		record["boleta_status"] = "CRITICO" if score < 90 else "CALIFICADO"
		medal_payload = self.get_trimestral_medal(score)
		record["medal"] = medal_payload["key"]
		record["medal_title"] = medal_payload["title"]
		record["medal_label"] = medal_payload["label"]
		record["medal_message"] = medal_payload["message"]
		record["updated_at"] = _timestamp()

		q_dir = _quarter_dir(quarter, year)
		q_scores = _read_json(q_dir / "trimestral.json", [])
		q_scores.append(record)
		_write_json(q_dir / "trimestral.json", q_scores)
		self._upsert_boleta_record(record)

		self.reload()
		return record

	def confirm_trimestral_scores(self, inspector_name: str, score_ids: list[str]) -> int:
		target_identity = _normalize_person_name(inspector_name)
		pending_ids = {str(score_id).strip() for score_id in score_ids if str(score_id).strip()}
		if not pending_ids:
			return 0

		viewer_name = str((self.current_user or {}).get("name", "")).strip() or inspector_name
		confirmed = 0
		if not TRIMESTRAL_DIR.exists():
			return 0

		for folder in TRIMESTRAL_DIR.iterdir():
			if not folder.is_dir():
				continue

			scores_path = folder / "trimestral.json"
			scores = _read_json(scores_path, [])
			changed = False
			for row in scores:
				if not isinstance(row, dict):
					continue
				score_id = str(row.get("id", "")).strip()
				if score_id not in pending_ids:
					continue
				if target_identity and _normalize_person_name(str(row.get("inspector", ""))) != target_identity:
					continue
				if str(row.get("confirmed_at") or "").strip():
					continue

				row["confirmed_at"] = _timestamp()
				row["confirmed_by"] = viewer_name
				row["updated_at"] = _timestamp()
				self._upsert_boleta_record(row)
				changed = True
				confirmed += 1

			if changed:
				_write_json(scores_path, scores)

		if confirmed:
			self.reload()

		return confirmed

	def delete_trimestral_score(self, score_id: str) -> None:
		if not TRIMESTRAL_DIR.exists():
			return
		for _folder in TRIMESTRAL_DIR.iterdir():
			if not _folder.is_dir():
				continue
			_scores = _read_json(_folder / "trimestral.json", [])
			_existing = next((s for s in _scores if s.get("id") == score_id), None)
			if _existing:
				_scores = [s for s in _scores if s.get("id") != score_id]
				_write_json(_folder / "trimestral.json", _scores)
				self._remove_boleta_record(_existing)
				self.reload()
				return

	def save_visit(self, payload: dict[str, Any], visit_id: str | None = None) -> dict[str, Any]:
		inspectors = _normalize_visit_inspectors(
			payload.get("inspectors"),
			str(payload.get("inspector", "")),
		)
		client = str(payload.get("client", "")).strip()
		address = str(payload.get("address", "")).strip()
		service = str(payload.get("service", "")).strip()
		visit_date = _normalize_visit_date(str(payload.get("visit_date", "")))
		assignment_time = _normalize_visit_time(str(payload.get("assignment_time", "")))
		departure_time = _normalize_visit_time(str(payload.get("departure_time", "")))
		status = str(payload.get("status", "")).strip() or "Programada"
		if status == "En ruta":
			status = "Programada"
		notes = str(payload.get("notes", "")).strip()

		if not inspectors:
			raise ValueError("Debes seleccionar un ejecutivo tecnico.")
		if not client:
			raise ValueError("Debes seleccionar un cliente.")
		if not address:
			raise ValueError("Debes seleccionar una direccion.")
		if not visit_date:
			raise ValueError("La fecha de visita es obligatoria y debe tener formato YYYY-MM-DD.")
		if not assignment_time:
			raise ValueError("Debes capturar la hora de asignacion al almacen con formato HH:MM.")
		if not departure_time:
			raise ValueError("Debes capturar la hora de salida con formato HH:MM.")
		if departure_time <= assignment_time:
			raise ValueError("La hora de salida debe ser posterior a la hora de asignacion al almacen.")

		# Find existing visit across all week folders
		existing = None
		existing_week_dir: Path | None = None
		if visit_id and VISITS_DIR.exists():
			for _wfolder in VISITS_DIR.iterdir():
				if not _wfolder.is_dir():
					continue
				_wvisits = _read_json(_wfolder / "visitas.json", [])
				_found = next((v for v in _wvisits if v.get("id") == visit_id), None)
				if _found:
					existing = _found
					existing_week_dir = _wfolder
					break

		today_iso = date.today().strftime("%Y-%m-%d")
		existing_date = _normalize_visit_date(str(existing.get("visit_date", ""))) if existing else ""
		if visit_date < today_iso and (existing is None or existing_date != visit_date):
			raise ValueError("No puedes agendar visitas en fechas anteriores al dia de hoy.")

		for scheduled_visit in self.list_visits():
			if visit_id and scheduled_visit.get("id") == visit_id:
				continue
			if _normalize_visit_date(str(scheduled_visit.get("visit_date", ""))) != visit_date:
				continue
			overlap = sorted(set(inspectors) & set(scheduled_visit.get("inspectors", [])))
			if overlap:
				names_text = ", ".join(overlap)
				raise ValueError(
					f"Los siguientes ejecutivos tecnicos ya tienen una visita asignada para {visit_date}: {names_text}"
				)

		existing_inspectors = _normalize_visit_inspectors(
			existing.get("inspectors") if existing else [],
			str(existing.get("inspector", "")) if existing else "",
		)

		visit = dict(existing or {})
		visit["id"] = visit.get("id") or uuid4().hex
		visit["group_id"] = str(visit.get("group_id", "")).strip() or visit["id"]
		visit["inspectors"] = inspectors
		visit["inspector"] = inspectors[0]
		visit["client"] = client
		visit["address"] = address
		visit["service"] = service or "Sin servicio"
		visit["visit_date"] = visit_date
		visit["assignment_time"] = assignment_time
		visit["departure_time"] = departure_time
		visit["status"] = status
		visit["notes"] = notes

		raw_status = str(visit.get("acceptance_status", "asignada")).strip().lower() or "asignada"
		existing_responses = _normalize_visit_acceptance_responses(visit.get("acceptance_responses"))
		responses_by_identity = {
			_normalize_person_name(name): payload
			for name, payload in existing_responses.items()
			if _normalize_person_name(name)
		}
		filtered_responses: dict[str, dict[str, str]] = {}
		for inspector_name in inspectors:
			identity = _normalize_person_name(inspector_name)
			payload = responses_by_identity.get(identity)
			if payload is None:
				continue
			filtered_responses[inspector_name] = dict(payload)
		visit["acceptance_responses"] = filtered_responses

		if raw_status in {"cancelada", "finalizada"}:
			visit["acceptance_status"] = raw_status
		elif filtered_responses:
			visit["acceptance_status"] = "aceptada"
		elif raw_status == "reasignada":
			visit["acceptance_status"] = "reasignada"
		else:
			visit["acceptance_status"] = "asignada"

		visit["accepted_by"] = ", ".join(filtered_responses.keys()) if filtered_responses else ""
		visit["assigned_by"] = (self.current_user or {}).get("name", "Sistema")
		visit["updated_at"] = _timestamp()

		# Remove from old week folder (handles date change between weeks)
		if existing_week_dir is not None:
			_old_list = _read_json(existing_week_dir / "visitas.json", [])
			_old_list = [v for v in _old_list if v.get("id") != visit["id"]]
			_write_json(existing_week_dir / "visitas.json", _old_list)

		# Write to (possibly new) week folder
		new_week_dir = _visit_week_dir(visit_date)
		new_list = _read_json(new_week_dir / "visitas.json", [])
		new_list.append(visit)
		_write_json(new_week_dir / "visitas.json", _merge_visit_records(new_list))

		affected_inspectors = set(existing_inspectors) | set(inspectors)
		for inspector_name in affected_inspectors:
			self._sync_visit_history(inspector_name)
		self.reload()
		return visit

	def delete_visit(self, visit_id: str) -> None:
		if not VISITS_DIR.exists():
			return
		for _wfolder in VISITS_DIR.iterdir():
			if not _wfolder.is_dir():
				continue
			_wvisits = _read_json(_wfolder / "visitas.json", [])
			_existing = next((v for v in _wvisits if v.get("id") == visit_id), None)
			if _existing:
				affected_inspectors = _normalize_visit_inspectors(
					_existing.get("inspectors"),
					str(_existing.get("inspector", "")),
				)
				_wvisits = [v for v in _wvisits if v.get("id") != visit_id]
				_write_json(_wfolder / "visitas.json", _wvisits)
				for inspector_name in affected_inspectors:
					self._sync_visit_history(inspector_name)
				self.reload()
				return

	def accept_visit(self, visit_id: str) -> None:
		"""Mark a visit as accepted by the current technical executive with timestamp."""
		if not VISITS_DIR.exists():
			return

		viewer = self.current_user or {}
		viewer_name = str(viewer.get("name", "")).strip()
		if not self.is_executive_role(viewer):
			raise ValueError("Solo un ejecutivo tecnico puede confirmar la visita.")
		if not viewer_name:
			raise ValueError("No se pudo identificar al ejecutivo que confirma.")
		viewer_identity = _normalize_person_name(viewer_name)

		for _wfolder in VISITS_DIR.iterdir():
			if not _wfolder.is_dir():
				continue
			_wvisits = _read_json(_wfolder / "visitas.json", [])
			_existing = next((v for v in _wvisits if v.get("id") == visit_id), None)
			if _existing:
				status = str(_existing.get("acceptance_status", "asignada")).strip().lower() or "asignada"
				if status == "cancelada":
					raise ValueError("La visita esta cancelada y no puede confirmarse.")
				if status == "finalizada":
					raise ValueError("La visita ya fue finalizada y no admite nuevas confirmaciones.")

				inspectors = _normalize_visit_inspectors(
					_existing.get("inspectors"),
					str(_existing.get("inspector", "")),
				)
				inspector_by_identity = {
					_normalize_person_name(name): name
					for name in inspectors
					if _normalize_person_name(name)
				}
				canonical_viewer = inspector_by_identity.get(viewer_identity)
				if not canonical_viewer:
					raise ValueError("No tienes permiso para confirmar esta visita.")

				responses = _normalize_visit_acceptance_responses(_existing.get("acceptance_responses"))
				response_identities = {
					_normalize_person_name(name)
					for name in responses
					if _normalize_person_name(name)
				}
				if viewer_identity in response_identities:
					raise ValueError("Ya confirmaste esta visita.")

				confirmed_at = _timestamp()
				responses[canonical_viewer] = {
					"confirmed_at": confirmed_at,
					"confirmed_by": canonical_viewer,
				}

				_existing["acceptance_responses"] = responses
				_existing["acceptance_status"] = "aceptada"
				_existing["accepted_by"] = ", ".join(sorted(responses.keys(), key=str.casefold))
				_existing["updated_at"] = confirmed_at
				_write_json(_wfolder / "visitas.json", _wvisits)
				self.reload()
				return

		raise ValueError("No se encontro la visita seleccionada.")

	def cancel_visit(self, visit_id: str, reason: str = "") -> None:
		"""Cancel a visit (admin only)."""
		if not VISITS_DIR.exists():
			return
		for _wfolder in VISITS_DIR.iterdir():
			if not _wfolder.is_dir():
				continue
			_wvisits = _read_json(_wfolder / "visitas.json", [])
			_existing = next((v for v in _wvisits if v.get("id") == visit_id), None)
			if _existing:
				_existing["acceptance_status"] = "cancelada"
				_existing["cancellation_reason"] = str(reason or "").strip()
				_existing["cancelled_by"] = (self.current_user or {}).get("name", "Sistema")
				_existing["updated_at"] = _timestamp()
				_write_json(_wfolder / "visitas.json", _wvisits)
				self.reload()
				return

	def mark_visit_reasignada(self, visit_id: str, new_date: str) -> None:
		"""Mark a visit as reasignada (grey) when moving it to another date."""
		if not VISITS_DIR.exists():
			return
		for _wfolder in VISITS_DIR.iterdir():
			if not _wfolder.is_dir():
				continue
			_wvisits = _read_json(_wfolder / "visitas.json", [])
			_existing = next((v for v in _wvisits if v.get("id") == visit_id), None)
			if _existing:
				_existing["acceptance_status"] = "reasignada"
				_existing["cancellation_reason"] = f"Reasignada al {new_date}"
				_existing["reassigned_to_date"] = new_date
				_existing["reassigned_by"] = (self.current_user or {}).get("name", "Sistema")
				_existing["updated_at"] = _timestamp()
				_write_json(_wfolder / "visitas.json", _wvisits)
				self.reload()
				return

	def reassign_visit(self, visit_id: str, new_inspectors: list[str]) -> None:
		"""Reassign a visit to different technical executives."""
		if not new_inspectors:
			raise ValueError("Debes seleccionar al menos un ejecutivo tecnico.")
		if not VISITS_DIR.exists():
			return
		for _wfolder in VISITS_DIR.iterdir():
			if not _wfolder.is_dir():
				continue
			_wvisits = _read_json(_wfolder / "visitas.json", [])
			_existing = next((v for v in _wvisits if v.get("id") == visit_id), None)
			if _existing:
				old_inspectors = _normalize_visit_inspectors(_existing.get("inspectors"), "")
				_existing["inspectors"] = new_inspectors
				_existing["inspector"] = new_inspectors[0]
				_existing["acceptance_status"] = "reasignada"
				_existing["acceptance_responses"] = {}
				_existing["accepted_by"] = ""
				_existing["reassigned_from"] = old_inspectors
				_existing["reassigned_by"] = (self.current_user or {}).get("name", "Sistema")
				_existing["updated_at"] = _timestamp()
				_write_json(_wfolder / "visitas.json", _wvisits)
				affected = set(old_inspectors) | set(new_inspectors)
				for inspector_name in affected:
					self._sync_visit_history(inspector_name)
				self.reload()
				return

	def _find_visit_storage(self, visit_id: str) -> tuple[Path, list[dict[str, Any]], dict[str, Any]] | None:
		if not VISITS_DIR.exists():
			return None
		for week_dir in VISITS_DIR.iterdir():
			if not week_dir.is_dir():
				continue
			visits_path = week_dir / "visitas.json"
			visits = _read_json(visits_path, [])
			if not isinstance(visits, list):
				continue
			for row in visits:
				if not isinstance(row, dict):
					continue
				if str(row.get("id", "")).strip() == str(visit_id).strip():
					return visits_path, visits, row
		return None

	def get_visit_available_norms(self, visit_id: str, inspector_name: str | None = None) -> list[str]:
		found = self._find_visit_storage(visit_id)
		if found is None:
			return []
		_visit_path, _visits, visit = found

		all_norms: set[str] = set()
		visit_inspectors = _normalize_visit_inspectors(visit.get("inspectors"), str(visit.get("inspector", "")))
		if not visit_inspectors and inspector_name:
			visit_inspectors = [str(inspector_name).strip()]

		for name in visit_inspectors:
			for token in self.get_accredited_norms(name):
				all_norms.add(token)
		return sorted(all_norms, key=self._norm_sort_key)

	def get_visit_reported_norms(self, visit_id: str, inspector_name: str) -> list[str]:
		reports = _read_json(NORMS_REPORT_FILE, [])
		if not isinstance(reports, list):
			return []

		target_visit = str(visit_id).strip()
		target_identity = _normalize_person_name(inspector_name)
		for row in reports:
			if not isinstance(row, dict):
				continue
			if str(row.get("visit_id", "")).strip() != target_visit:
				continue
			if _normalize_person_name(str(row.get("inspector", ""))) != target_identity:
				continue
			raw_norms = row.get("norms", [])
			if not isinstance(raw_norms, list):
				return []
			cleaned = [_normalize_norm_key(item) for item in raw_norms if str(item).strip()]
			return sorted(set(cleaned), key=self._norm_sort_key)

		return []

	def save_visit_norm_report(self, visit_id: str, norms_applied: list[str], inspector_name: str | None = None) -> dict[str, Any]:
		found = self._find_visit_storage(visit_id)
		if found is None:
			raise ValueError("No se encontro la visita seleccionada.")

		visits_path, visits, visit = found
		visit_inspectors = _normalize_visit_inspectors(visit.get("inspectors"), str(visit.get("inspector", "")))

		viewer_name = str((self.current_user or {}).get("name", "")).strip()
		target_name = str(inspector_name or viewer_name).strip()
		if not target_name:
			target_name = visit_inspectors[0] if visit_inspectors else ""

		if not target_name:
			raise ValueError("No se pudo identificar al ejecutivo tecnico del reporte.")

		if self.current_user and self.is_executive_role(self.current_user):
			if target_name != viewer_name:
				raise ValueError("Solo puedes reportar normas para tus propias visitas.")
			if target_name not in visit_inspectors:
				raise ValueError("No tienes permiso para reportar normas en esta visita.")

		clean_norms = sorted(
			{_normalize_norm_key(item) for item in norms_applied if str(item).strip()},
			key=self._norm_sort_key,
		)
		if not clean_norms:
			raise ValueError("Selecciona al menos una norma aplicada durante la visita.")

		reports = _read_json(NORMS_REPORT_FILE, [])
		if not isinstance(reports, list):
			reports = []

		target_visit = str(visit_id).strip()
		target_identity = _normalize_person_name(target_name)
		existing_index = next(
			(
				index
				for index, item in enumerate(reports)
				if isinstance(item, dict)
				and str(item.get("visit_id", "")).strip() == target_visit
				and _normalize_person_name(str(item.get("inspector", ""))) == target_identity
			),
			None,
		)

		visit_date = _normalize_visit_date(str(visit.get("visit_date", "")))
		month_key = visit_date[:7] if visit_date else datetime.now().strftime("%Y-%m")
		year_value = int(month_key.split("-")[0]) if "-" in month_key else datetime.now().year

		record = dict(reports[existing_index]) if existing_index is not None else {}
		record["id"] = str(record.get("id", "")).strip() or uuid4().hex
		record["visit_id"] = target_visit
		record["visit_date"] = visit_date
		record["month"] = month_key
		record["year"] = year_value
		record["inspector"] = target_name
		record["client"] = str(visit.get("client", "")).strip()
		record["address"] = str(visit.get("address", "")).strip()
		record["service"] = str(visit.get("service", "")).strip()
		record["norms"] = clean_norms
		record["norm_count"] = len(clean_norms)
		record["updated_at"] = _timestamp()

		if existing_index is None:
			reports.append(record)
		else:
			reports[existing_index] = record
		_write_json(NORMS_REPORT_FILE, reports)

		reported_norms = visit.get("reported_norms", {})
		if not isinstance(reported_norms, dict):
			reported_norms = {}
		reported_norms[target_name] = clean_norms
		visit["reported_norms"] = reported_norms
		visit["acceptance_status"] = "finalizada"
		visit["updated_at"] = _timestamp()
		_write_json(visits_path, visits)

		self.reload()
		return record

	def mark_visit_finalized(self, visit_id: str, finalized_at: str) -> None:
		"""Marca la visita como Finalizada, registra la hora y actualiza el reporte de normas."""
		found = self._find_visit_storage(visit_id)
		if found is None:
			return
		visits_path, visits, visit = found
		visit["status"] = "Finalizada"
		visit["acceptance_status"] = "finalizada"
		visit["finalized_at"] = finalized_at
		visit["updated_at"] = _timestamp()
		_write_json(visits_path, visits)

		# Actualizar también el reporte de normas con la hora de finalización
		reports = _read_json(NORMS_REPORT_FILE, [])
		if isinstance(reports, list):
			target_visit = str(visit_id).strip()
			for item in reports:
				if isinstance(item, dict) and str(item.get("visit_id", "")).strip() == target_visit:
					item["finalized_at"] = finalized_at
			_write_json(NORMS_REPORT_FILE, reports)

		self.reload()

	def list_norm_visit_reports(self, month: str | None = None) -> list[dict[str, Any]]:
		reports = _read_json(NORMS_REPORT_FILE, [])
		if not isinstance(reports, list):
			reports = []

		if month:
			month_key = str(month).strip()
			reports = [item for item in reports if str(item.get("month", "")).strip() == month_key]

		reports.sort(
			key=lambda item: (
				str(item.get("month", "")),
				str(item.get("visit_date", "")),
				str(item.get("updated_at", "")),
			),
			reverse=True,
		)
		return [dict(item) for item in reports if isinstance(item, dict)]

	def get_norm_report_months(self) -> list[str]:
		reports = self.list_norm_visit_reports()
		months = sorted({str(item.get("month", "")).strip() for item in reports if str(item.get("month", "")).strip()})
		return months

	def get_monthly_norm_demand(self, month: str | None = None) -> list[dict[str, Any]]:
		target_month = str(month).strip() if month else ""
		rows = self.list_norm_visit_reports(month=target_month or None)
		counts: dict[str, int] = {}
		for row in rows:
			for norm in row.get("norms", []):
				norm_key = _normalize_norm_key(norm)
				counts[norm_key] = counts.get(norm_key, 0) + 1

		result = [{"norm": norm, "count": count} for norm, count in counts.items()]
		result.sort(key=lambda item: (item["count"], item["norm"]), reverse=True)
		return result

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
		role = _normalize_role_name(payload.get("role")) or "ejecutivo tecnico"

		if not all([name, username, password]):
			raise ValueError("Nombre, usuario y contrasena son obligatorios.")
		if role not in {
			"admin",
			"gerente",
			"sub gerente",
			"coordinador operativo",
			"coordinadora en fiabilidad",
			"talento humano",
			"supervisor",
			"ejecutivo tecnico",
			"especialidades",
		}:
			raise ValueError("El rol no es valido para este sistema.")

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

	def save_client(self, payload: dict[str, Any], original_client: str | None = None) -> dict[str, Any]:
		client_name = str(payload.get("CLIENTE", "")).strip()
		rfc = str(payload.get("RFC", "")).strip()
		warehouse = str(payload.get("ALMACEN", "")).strip()
		street = str(payload.get("CALLE Y NO", "")).strip()
		colony = str(payload.get("COLONIA O POBLACION", "")).strip()
		municipality = str(payload.get("MUNICIPIO O ALCADIA", "")).strip()
		state = str(payload.get("CIUDAD O ESTADO", "")).strip()
		postal_code = str(payload.get("CP", "")).strip()
		service = str(payload.get("SERVICIO", "")).strip()

		if not client_name:
			raise ValueError("El nombre del cliente es obligatorio.")

		normalized_client = client_name.casefold()
		duplicate = next(
			(
				item
				for item in self.clients_catalog
				if str(item.get("CLIENTE", "")).strip().casefold() == normalized_client
			),
			None,
		)
		if duplicate and str(duplicate.get("CLIENTE", "")).strip() != str(original_client or "").strip():
			raise ValueError("Ya existe un cliente con ese nombre.")

		if original_client:
			original = next(
				(item for item in self.clients_catalog if str(item.get("CLIENTE", "")).strip() == str(original_client).strip()),
				None,
			)
			if original is None:
				raise ValueError("No se encontro el cliente a editar.")
		else:
			original = None

		existing_addresses = []
		if original and isinstance(original.get("DIRECCIONES"), list):
			existing_addresses = [
				dict(address)
				for address in original.get("DIRECCIONES", [])
				if isinstance(address, dict)
			]

		has_address_data = any([warehouse, street, colony, municipality, state, postal_code, service])
		if has_address_data:
			primary_address = {
				"ALMACEN": warehouse,
				"CALLE Y NO": street,
				"COLONIA O POBLACION": colony,
				"MUNICIPIO O ALCADIA": municipality,
				"CIUDAD O ESTADO": state,
				"CP": postal_code,
				"SERVICIO": service or "DICTAMEN",
			}
			if existing_addresses:
				existing_addresses[0] = primary_address
			else:
				existing_addresses = [primary_address]

		updated = dict(original or {})
		updated["RFC"] = rfc
		updated["CLIENTE"] = client_name
		updated["DIRECCIONES"] = existing_addresses

		if original is None:
			self.clients_catalog.append(updated)
		else:
			index = self.clients_catalog.index(original)
			self.clients_catalog[index] = updated

		_write_json(CLIENTS_FILE, self.clients_catalog)
		self.reload()
		return updated

	def delete_client(self, client_name: str) -> None:
		original = next(
			(item for item in self.clients_catalog if str(item.get("CLIENTE", "")).strip() == str(client_name).strip()),
			None,
		)
		if original is None:
			return

		self.clients_catalog.remove(original)
		_write_json(CLIENTS_FILE, self.clients_catalog)
		self.reload()

	def save_client_address(self, client_name: str, address: dict[str, Any], address_index: int | None = None) -> None:
		client = next(
			(item for item in self.clients_catalog if str(item.get("CLIENTE", "")).strip() == str(client_name).strip()),
			None,
		)
		if client is None:
			raise ValueError("No se encontro el cliente.")
		addresses = list(client.get("DIRECCIONES") or [])
		addr: dict[str, Any] = {
			"ALMACEN": str(address.get("ALMACEN", "")).strip(),
			"CALLE Y NO": str(address.get("CALLE Y NO", "")).strip(),
			"COLONIA O POBLACION": str(address.get("COLONIA O POBLACION", "")).strip(),
			"MUNICIPIO O ALCADIA": str(address.get("MUNICIPIO O ALCADIA", "")).strip(),
			"CIUDAD O ESTADO": str(address.get("CIUDAD O ESTADO", "")).strip(),
			"CP": str(address.get("CP", "")).strip(),
			"SERVICIO": str(address.get("SERVICIO", "DICTAMEN")).strip() or "DICTAMEN",
		}
		if address_index is not None and 0 <= address_index < len(addresses):
			addresses[address_index] = addr
		else:
			addresses.append(addr)
		client["DIRECCIONES"] = addresses
		_write_json(CLIENTS_FILE, self.clients_catalog)
		self.reload()

	def delete_client_address(self, client_name: str, address_index: int) -> None:
		client = next(
			(item for item in self.clients_catalog if str(item.get("CLIENTE", "")).strip() == str(client_name).strip()),
			None,
		)
		if client is None:
			return
		addresses = list(client.get("DIRECCIONES") or [])
		if 0 <= address_index < len(addresses):
			addresses.pop(address_index)
			client["DIRECCIONES"] = addresses
			_write_json(CLIENTS_FILE, self.clients_catalog)
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
			module = _load_module(str(DOCUMENT_MODULE_DIR / "FormatoSupervision.py"))
			builder = getattr(module, "build_formato_supervision_pdf")
		elif document_kind == "criterio":
			module = _load_module(str(DOCUMENT_MODULE_DIR / "CriterioEvaluacionTecnica.py"))
			builder = getattr(module, "build_criterio_evaluacion_pdf")
		else:
			raise ValueError("Tipo de documento no soportado.")

		builder(output_path, payload)
		return output_path

	def preview_criterio_resolution_number(self) -> str:
		counters = self.app_state.setdefault("document_counters", {"criterio_evaluacion_tecnica": 0})
		current_value = int(counters.get("criterio_evaluacion_tecnica", 0) or 0)
		return _format_criterio_resolution_number(current_value + 1)

	def _peek_criterio_resolution_number(self) -> tuple[int, str]:
		counters = self.app_state.setdefault("document_counters", {"criterio_evaluacion_tecnica": 0})
		next_value = int(counters.get("criterio_evaluacion_tecnica", 0) or 0) + 1
		return next_value, _format_criterio_resolution_number(next_value)

	def generate_criterio_document(self, destination: str | Path, payload: dict[str, Any]) -> Path:
		output_path = Path(destination)
		output_path.parent.mkdir(parents=True, exist_ok=True)

		resolved_payload = dict(payload)
		next_value, resolution_number = self._peek_criterio_resolution_number()
		resolved_payload["resolution_number"] = resolution_number
		resolved_payload["visit_date"] = str(resolved_payload.get("visit_date", "")).strip() or date.today().strftime("%Y-%m-%d")
		resolved_payload["generated_at"] = _timestamp()

		module = _load_module(str(DOCUMENT_MODULE_DIR / "CriterioEvaluacionTecnica.py"))
		builder = getattr(module, "build_criterio_evaluacion_pdf")
		builder(output_path, resolved_payload)

		client_name = str(resolved_payload.get("client", "")).strip() or "sin_cliente"
		client_folder = _safe_folder_name(client_name)
		timestamp_folder = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
		archive_path = CRITERIA_ARCHIVE_DIR / client_folder / timestamp_folder / "CriterioEvaluacion.pdf"
		archive_path.parent.mkdir(parents=True, exist_ok=True)
		shutil.copy2(output_path, archive_path)

		counters = self.app_state.setdefault("document_counters", {"criterio_evaluacion_tecnica": 0})
		counters["criterio_evaluacion_tecnica"] = next_value
		_write_json(STATE_FILE, self.app_state)
		documents_log = self.app_state.setdefault("criteria_documents", [])
		documents_log.append(
			{
				"resolution_number": resolution_number,
				"client": str(resolved_payload.get("client", "")).strip(),
				"selected_norm": str(resolved_payload.get("selected_norm", "")).strip(),
				"executive_name": str(resolved_payload.get("executive_name", "")).strip(),
				"evaluated_product": str(resolved_payload.get("evaluated_product", "")).strip(),
				"generated_at": resolved_payload["generated_at"],
				"output_path": str(archive_path),
				"user_output_path": str(output_path),
			}
		)
		_write_json(STATE_FILE, self.app_state)
		return output_path

	def get_criteria_history(self, client_name: str | None = None) -> list[dict[str, Any]]:
		"""Retorna el historial de PDFs de criterios generados, opcionalmente filtrados por cliente."""
		documents_log = self.app_state.get("criteria_documents", [])
		if not client_name:
			return documents_log
		client_clean = str(client_name or "").strip().lower()
		return [doc for doc in documents_log if str(doc.get("client", "")).strip().lower() == client_clean]

	def _agreement_client_dir(self, client_name: str) -> Path:
		client_folder = _safe_folder_name(str(client_name or "").strip() or "sin_cliente")
		path = AGREEMENTS_ARCHIVE_DIR / client_folder
		path.mkdir(parents=True, exist_ok=True)
		return path

	def save_client_agreement_pdf(self, client_name: str, source_path: str | Path) -> Path:
		clean_client = str(client_name or "").strip()
		if not clean_client:
			raise ValueError("Debes seleccionar un cliente para cargar acuerdos.")

		source = Path(source_path)
		if not source.exists() or not source.is_file():
			raise ValueError(f"El archivo de acuerdos '{source_path}' no existe o no es válido.")

		# Guardar con el nombre original, si ya existe, agregar sufijo con timestamp
		target_dir = self._agreement_client_dir(clean_client)
		base_name = source.name
		target_path = target_dir / base_name
		if target_path.exists():
			timestamp = datetime.now().strftime("_%Y%m%d_%H%M%S")
			target_path = target_dir / f"{source.stem}{timestamp}{source.suffix}"
		shutil.copy2(source, target_path)
		return target_path

	def get_client_agreements(self, client_name: str | None = None) -> list[dict[str, Any]]:
		if not client_name:
			return []

		target_dir = self._agreement_client_dir(str(client_name))
		agreements: list[dict[str, Any]] = []
		for pdf_path in sorted(target_dir.glob("*.pdf"), key=lambda item: item.stat().st_mtime, reverse=True):
			agreements.append(
				{
					"client": str(client_name).strip(),
					"title": pdf_path.name,
					"output_path": str(pdf_path),
					"generated_at": datetime.fromtimestamp(pdf_path.stat().st_mtime).strftime("%Y-%m-%d %H:%M"),
				}
			)
		return agreements

	def _history_dir(self, inspector_name: str) -> Path:
		requested_name = str(inspector_name or "").strip()
		canonical_name = self._resolve_canonical_person_name(requested_name) or requested_name
		preferred_path = HISTORY_DIR / _safe_folder_name(canonical_name)
		alias_paths: list[Path] = []
		if canonical_name and requested_name and _normalize_person_name(canonical_name) != _normalize_person_name(requested_name):
			for alias_path in {
				HISTORY_DIR / _safe_folder_name(requested_name),
				HISTORY_DIR / _safe_slug(requested_name),
			}:
				alias_paths.append(alias_path)

		target_identity = _folder_identity(canonical_name)
		cached_path = self._history_dir_index.get(target_identity) if target_identity else None
		has_pending_alias = any(
			alias_path.exists()
			and cached_path is not None
			and alias_path.resolve() != cached_path.resolve()
			for alias_path in alias_paths
		)
		if cached_path is not None and cached_path.exists() and not has_pending_alias:
			return cached_path

		preferred_path.mkdir(parents=True, exist_ok=True)
		if target_identity and HISTORY_DIR.exists():
			candidates: list[Path] = []
			for child in HISTORY_DIR.iterdir():
				if not child.is_dir():
					continue
				child_identity = _folder_identity(child.name)
				if child_identity == target_identity:
					candidates.append(child)
					continue
				if not self._consolidation_done and _history_folder_matches_identity(child, target_identity):
					candidates.append(child)

			legacy_path = HISTORY_DIR / _safe_slug(canonical_name)
			if legacy_path.exists() and legacy_path.is_dir() and legacy_path not in candidates:
				candidates.append(legacy_path)

			for alias_path in alias_paths:
				if alias_path.exists() and alias_path.is_dir() and alias_path not in candidates:
					candidates.append(alias_path)

			for source in candidates:
				if source.resolve() == preferred_path.resolve():
					continue
				_merge_history_directories(preferred_path, source)

		if target_identity:
			self._history_dir_index[target_identity] = preferred_path
		return preferred_path

	def _history_file(self, inspector_name: str) -> Path:
		return self._history_dir(inspector_name) / "historico.json"

	# ─── Vacaciones ─────────────────────────────────────────────────────────

	def list_vacations(self) -> list[dict[str, Any]]:
		return list(self.app_state.get("vacations", []))

	def save_vacation(self, executive: str, start_date: str, end_date: str) -> dict[str, Any]:
		if not executive or not start_date or not end_date:
			raise ValueError("Ejecutivo, fecha inicio y fecha fin son obligatorios.")
		if end_date < start_date:
			raise ValueError("La fecha fin no puede ser anterior a la fecha inicio.")
		entry = {
			"id": datetime.now().strftime("%Y%m%d%H%M%S%f"),
			"executive": executive.strip(),
			"start_date": start_date,
			"end_date": end_date,
			"created_at": datetime.now().isoformat(),
		}
		vacations = self.app_state.setdefault("vacations", [])
		vacations.append(entry)
		_write_json(STATE_FILE, self.app_state)
		return entry

	def delete_vacation(self, vacation_id: str) -> bool:
		vacations = self.app_state.get("vacations", [])
		idx = next((i for i, v in enumerate(vacations) if v.get("id") == vacation_id), None)
		if idx is None:
			return False
		del vacations[idx]
		_write_json(STATE_FILE, self.app_state)
		return True

	def get_vacations_for_date(self, iso_date: str) -> list[dict[str, Any]]:
		result = []
		for v in self.app_state.get("vacations", []):
			if v.get("start_date", "") <= iso_date <= v.get("end_date", ""):
				result.append(v)
		return result

	# ─── Talleres ───────────────────────────────────────────────────────────

	def list_workshops(self) -> list[dict[str, Any]]:
		return list(self.app_state.get("workshops", []))

	def save_workshop(self, title: str, workshop_date: str, description: str = "", executives=None, place: str = "", start_time: str = "", end_time: str = "", **kwargs) -> dict[str, Any]:
		if not title or not workshop_date:
			raise ValueError("Titulo y fecha del taller son obligatorios.")
		if executives is None:
			executives = "ALL"
		entry = {
			"id": datetime.now().strftime("%Y%m%d%H%M%S%f"),
			"title": title.strip(),
			"date": workshop_date,
			"description": description.strip(),
			"created_at": datetime.now().isoformat(),
			"executives": executives,
			"place": place.strip() if place else "",
			"start_time": start_time.strip() if start_time else "",
			"end_time": end_time.strip() if end_time else "",
		}
		# Permitir campos extra como 'type'
		entry.update(kwargs)
		workshops = self.app_state.setdefault("workshops", [])
		workshops.append(entry)
		_write_json(STATE_FILE, self.app_state)
		self._notify_all_users_about_workshop(entry)
		return entry

	def _notify_all_users_about_workshop(self, workshop_entry):
		# Cargar usuarios desde el archivo USERS_FILE
		users_data = _read_json(USERS_FILE, {"users": []})
		users = users_data.get("users", [])
		notifications = self.app_state.setdefault("notifications", [])
		# Determinar a quién notificar
		if workshop_entry.get("executives") == "ALL":
			notify_usernames = [user.get("username") for user in users if _normalize_role_name(user.get("role")) in {"ejecutivo tecnico", "especialidades"}]
		else:
			exec_names = set(workshop_entry.get("executives", []))
			notify_usernames = [user.get("username") for user in users if user.get("name") in exec_names]
		for username in notify_usernames:
			notifications.append({
				"username": username,
				"message": f"Nuevo taller: {workshop_entry['title']} el {workshop_entry['date']}",
				"created_at": datetime.now().isoformat(),
				"type": "workshop"
			})
		_write_json(STATE_FILE, self.app_state)

	def delete_workshop(self, workshop_id: str) -> bool:
		workshops = self.app_state.get("workshops", [])
		idx = next((i for i, w in enumerate(workshops) if w.get("id") == workshop_id), None)
		if idx is None:
			return False
		del workshops[idx]
		_write_json(STATE_FILE, self.app_state)
		return True

	def get_workshops_for_date(self, iso_date: str) -> list[dict[str, Any]]:
		return [w for w in self.app_state.get("workshops", []) if w.get("date") == iso_date]

	def _visits_file(self, inspector_name: str) -> Path:
		return self._history_dir(inspector_name) / "visitas.json"

	def _sync_visit_history(self, inspector_name: str) -> None:
		_write_json(self._visits_file(inspector_name), self.list_visits(name=inspector_name))




def _controller_evaluation_key(inspector_name: str, norm_token: str | None = None) -> str:
	normalized_norm = _normalize_norm_key(norm_token)
	return f"{inspector_name}::{normalized_norm}"


def _controller_get_norm_display_name(self, norm_value: str | None) -> str:
	clean_value = str(norm_value or "").strip()
	if not clean_value:
		return "Sin norma"

	token = _extract_norm_token(clean_value)
	if token:
		for item in self.get_catalog_norms():
			if str(item.get("token", "")).strip() != token:
				continue
			nom_value = " ".join(str(item.get("nom", token)).split())
			return nom_value

	return " ".join(clean_value.split())


def _controller_get_norm_score_history(
	self,
	inspector_name: str,
	norm_token: str | None = None,
) -> list[dict[str, Any]]:
	target_norm_key = _normalize_norm_key(norm_token) if norm_token else ""
	history = self.get_history(inspector_name)
	rows: list[dict[str, Any]] = []

	for entry in history:
		visit_date = str(entry.get("visit_date", "")).strip()
		saved_at = str(entry.get("saved_at", "")).strip()
		status = str(entry.get("status", "")).strip() or "Sin estatus"
		evaluator = str(entry.get("evaluator", "")).strip() or "Sin supervisor"
		soft_skills_score = _coerce_score(entry.get("soft_skills_score"))
		technical_skills_score = _coerce_score(entry.get("technical_skills_score"))

		raw_scores = entry.get("score_by_norm", {})
		score_rows: list[tuple[str, float]] = []
		if isinstance(raw_scores, dict):
			for raw_norm, raw_score in raw_scores.items():
				norm_name = str(raw_norm).strip()
				score = _coerce_score(raw_score)
				if not norm_name or score is None:
					continue
				score_rows.append((norm_name, score))

		if not score_rows:
			norm_name = str(entry.get("selected_norm", "")).strip() or "Sin norma"
			score = _coerce_score(entry.get("score"))
			if score is not None:
				score_rows.append((norm_name, score))

		for raw_norm_name, score in score_rows:
			norm_key = _normalize_norm_key(raw_norm_name)
			if target_norm_key and norm_key != target_norm_key:
				continue

			rows.append(
				{
					"norm": self.get_norm_display_name(raw_norm_name),
					"score": score,
					"soft_skills_score": soft_skills_score,
					"technical_skills_score": technical_skills_score,
					"status": status,
					"evaluator": evaluator,
					"visit_date": visit_date,
					"saved_at": saved_at,
					"archivo_path": entry.get("archivo_path") or entry.get("json_path") or "",
					"enviado": entry.get("enviado", False),
					"confirmado": entry.get("confirmado", False),
				}
			)

	rows.sort(
		key=lambda item: (
			str(item.get("visit_date", "")),
			str(item.get("saved_at", "")),
		),
		reverse=True,
	)
	return rows


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

	if VISITS_DIR.exists():
		for week_dir in VISITS_DIR.iterdir():
			if not week_dir.is_dir():
				continue
			visits_path = week_dir / "visitas.json"
			raw_visits = _read_json(visits_path, [])
			changed = False
			for visit in raw_visits:
				inspectors = _normalize_visit_inspectors(visit.get("inspectors"), str(visit.get("inspector", "")))
				renamed = [new_name if name == original_name else name for name in inspectors]
				if renamed != inspectors:
					visit["inspectors"] = renamed
					visit["inspector"] = renamed[0] if renamed else ""
					changed = True
			if changed:
				_write_json(visits_path, _merge_visit_records(raw_visits))

	original_dir = self._history_dir(original_name)
	target_dir = HISTORY_DIR / _safe_folder_name(new_name)
	target_dir.mkdir(parents=True, exist_ok=True)
	if original_dir.exists() and original_dir.resolve() != target_dir.resolve():
		_merge_history_directories(target_dir, original_dir)

	self._history_dir(new_name)

	_write_json(STATE_FILE, self.app_state)
	self._sync_visit_history(new_name)


def _controller_norm_sort_key(token: str) -> tuple[int, str]:
	match = re.search(r"(\d{3})", token)
	if match:
		return int(match.group(1)), token
	return 999, token


def _controller_get_latest_evaluation(self, inspector_name: str, norm_token: str | None = None) -> dict[str, Any]:
	clean_name = str(inspector_name or "").strip()
	if not clean_name:
		return {}

	if norm_token:
		normalized_norm = _normalize_norm_key(norm_token)
		latest = self._latest_evaluation_by_norm_cache.get((clean_name, normalized_norm))
		if latest:
			return dict(latest)

		history = self.get_history(clean_name)
		for entry in reversed(history):
			current_norm = _normalize_norm_key(str(entry.get("selected_norm", "")).strip())
			if current_norm == normalized_norm:
				self._cache_latest_evaluation_entry(clean_name, entry)
				return dict(entry)
		return {}

	latest = self._latest_evaluation_cache.get(clean_name)
	if latest:
		return dict(latest)

	history = self.get_history(clean_name)
	if history:
		self._cache_latest_evaluation_entry(clean_name, history[-1])
		return dict(history[-1])
	return {}


def _controller_has_completed_form(self, inspector_name: str, norm_token: str | None = None) -> bool:
	latest = self.get_latest_evaluation(inspector_name, norm_token)
	return bool(latest.get("form_completed"))


def _controller_get_principal_rows(self, search_text: str = "", status_filter: str = "Todos") -> list[dict[str, Any]]:
	lowered_query = search_text.strip().lower()
	clean_status_filter = str(status_filter or "Todos").strip() or "Todos"
	cache_key = (lowered_query, clean_status_filter)
	cached_rows = self._principal_rows_cache.get(cache_key)
	if cached_rows is not None:
		return [dict(item) for item in cached_rows]

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
			status = "Feedback"
		else:
			status = "Estable"

		actions_text = "Supervisar"

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
		if clean_status_filter == "Completos" and not row["form_completed"]:
			continue
		if clean_status_filter == "Pendientes" and row["form_completed"]:
			continue
		if clean_status_filter == "Feedback" and row["status"] != "Feedback":
			continue

		rows.append(row)

	rows = sorted(rows, key=lambda item: item["name"].lower())
	self._principal_rows_cache[cache_key] = rows
	return [dict(item) for item in rows]


def _controller_get_overview_metrics(self) -> dict[str, Any]:
	if self._overview_metrics_cache is not None:
		return dict(self._overview_metrics_cache)

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
	self._overview_metrics_cache = {
		"inspectors": len(inspector_names),
		"completed_forms": completed,
		"average_score": round(mean(scores), 1) if scores else None,
		"alerts": alerts,
	}
	return dict(self._overview_metrics_cache)


def _controller_save_evaluation(self, inspector_name: str, payload: dict[str, Any]) -> dict[str, Any]:
	clean_norm = str(payload.get("selected_norm", "")).strip()
	clean_client = str(payload.get("client", "")).strip()
	clean_date = str(payload.get("visit_date", "")).strip()
	clean_status = str(payload.get("status", "")).strip() or "En seguimiento"
	clean_observations = str(payload.get("observations", "")).strip()
	clean_actions = str(payload.get("corrective_actions", "")).strip()
	evaluator = str(payload.get("evaluator", "")).strip()
	inspector_supervised = str(payload.get("inspector_supervised", inspector_name)).strip() or inspector_name
	protocol_answers = _normalize_supervision_answers(payload.get("protocol_answers"))
	process_answers = _normalize_supervision_answers(payload.get("process_answers"))
	technical_normative_rows = _normalize_technical_normative_rows(payload.get("technical_normative_rows"))
	image_folder = str(payload.get("image_folder", "")).strip()
	score_breakdown = payload.get("score_breakdown", {})
	if not isinstance(score_breakdown, dict):
		score_breakdown = {}
	score_by_norm = payload.get("score_by_norm", {})
	if not isinstance(score_by_norm, dict):
		score_by_norm = {}
	soft_skills_score = _coerce_score(payload.get("soft_skills_score"))
	technical_skills_score = _coerce_score(payload.get("technical_skills_score"))
	soft_skills_breakdown = payload.get("soft_skills_breakdown", {})
	if not isinstance(soft_skills_breakdown, dict):
		soft_skills_breakdown = {}
	technical_skills_breakdown = payload.get("technical_skills_breakdown", {})
	if not isinstance(technical_skills_breakdown, dict):
		technical_skills_breakdown = {}
	normalized_score_by_norm: dict[str, float] = {}
	for norm_name, raw_score in score_by_norm.items():
		clean_norm_name = str(norm_name).strip()
		numeric_score = _coerce_score(raw_score)
		if not clean_norm_name or numeric_score is None:
			continue
		normalized_score_by_norm[clean_norm_name] = numeric_score
	score = _coerce_score(payload.get("score"))

	if not clean_date:
		raise ValueError("La fecha de seguimiento es obligatoria.")
	if not clean_client:
		raise ValueError("Debes seleccionar cliente o almacen.")
	if score is None:
		raise ValueError("El puntaje debe ser numerico.")
	if score < 0 or score > 100:
		raise ValueError("El puntaje debe estar entre 0 y 100.")
	if soft_skills_score is not None and (soft_skills_score < 0 or soft_skills_score > 100):
		raise ValueError("La calificación de habilidades blandas debe estar entre 0 y 100.")
	if technical_skills_score is not None and (technical_skills_score < 0 or technical_skills_score > 100):
		raise ValueError("La calificación de habilidades técnicas debe estar entre 0 y 100.")

	evaluation = {
		"inspector_name": inspector_name,
		"inspector_supervised": inspector_supervised,
		"selected_norm": clean_norm or "Sin norma",
		"client": clean_client,
		"visit_date": clean_date,
		"score": score,
		"soft_skills_score": soft_skills_score,
		"technical_skills_score": technical_skills_score,
		"status": clean_status,
		"observations": clean_observations,
		"corrective_actions": clean_actions,
		"evaluator": evaluator or (self.current_user or {}).get("name", "Sin evaluador"),
		"protocol_answers": protocol_answers,
		"process_answers": process_answers,
		"technical_normative_rows": technical_normative_rows,
		"image_folder": image_folder,
		"score_breakdown": score_breakdown,
		"soft_skills_breakdown": soft_skills_breakdown,
		"technical_skills_breakdown": technical_skills_breakdown,
		"score_by_norm": normalized_score_by_norm,
		"form_structure": "supervision_v2",
		"saved_at": _timestamp(),
		"form_completed": True,
	}

	base_folder = self._history_dir(inspector_name)
	formatos_folder = base_folder / "FORMATOS DE SUPERVISION"
	formatos_folder.mkdir(parents=True, exist_ok=True)
	ts = datetime.now().strftime("%Y%m%d_%H%M%S")
	unique_id = uuid4().hex[:8]
	norm_slug = str(clean_norm).replace(" ", "_").replace("/", "-")
	base_filename = f"Supervision_{ts}_{norm_slug}_{unique_id}"
	json_path = formatos_folder / f"{base_filename}.json"
	pdf_path = formatos_folder / f"{base_filename}.pdf"
	with open(json_path, "w", encoding="utf-8") as f:
		json.dump(evaluation, f, ensure_ascii=False, indent=2)

	evaluation["pdf_path"] = str(pdf_path)
	evaluation["json_path"] = str(json_path)

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
	norm_slug = _safe_slug(_normalize_norm_key(norm_token).lower())
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
	payload["selected_norm"] = self.get_norm_display_name(payload.get("selected_norm") or norm_token)
	raw_score_by_norm = payload.get("score_by_norm", {})
	if isinstance(raw_score_by_norm, dict):
		normalized_score_by_norm: dict[str, float] = {}
		for raw_norm, raw_score in raw_score_by_norm.items():
			numeric_score = _coerce_score(raw_score)
			if numeric_score is None:
				continue
			normalized_score_by_norm[self.get_norm_display_name(str(raw_norm))] = numeric_score
		if normalized_score_by_norm:
			payload["score_by_norm"] = normalized_score_by_norm
	payload["accredited_norms"] = self.get_accredited_norms(inspector_name)

	if document_kind == "formato":
		module = _load_module(str(DOCUMENT_MODULE_DIR / "FormatoSupervision.py"))
		builder = getattr(module, "build_formato_supervision_pdf")
	elif document_kind == "criterio":
		module = _load_module(str(DOCUMENT_MODULE_DIR / "CriterioEvaluacionTecnica.py"))
		builder = getattr(module, "build_criterio_evaluacion_pdf")
	else:
		raise ValueError("Tipo de documento no soportado.")

	builder(output_path, payload)
	return output_path


def _controller_get_default_trimestral_report_path(self) -> Path:
	reports_dir = DATA_DIR / "reportes"
	reports_dir.mkdir(parents=True, exist_ok=True)

	current_user = self.current_user or {}
	viewer_name = str(current_user.get("name", "reporte_trimestral")).strip() or "reporte_trimestral"
	viewer_role = _normalize_role_name(current_user.get("role"))
	scope_slug = (
		"global"
		if viewer_role in {"admin", "gerente", "sub gerente", "coordinador operativo", "coordinadora en fiabilidad"}
		else _safe_slug(viewer_name).lower()
	)
	timestamp = datetime.now().strftime("%Y%m%d_%H%M")
	filename = f"reporte_trimestral_{scope_slug}_{timestamp}.pdf"
	return reports_dir / filename


def _controller_generate_trimestral_dashboard_report(
	self,
	destination: str | Path,
	payload: dict[str, Any],
) -> Path:
	output_path = Path(destination)
	output_path.parent.mkdir(parents=True, exist_ok=True)

	module = _load_module(str(DOCUMENT_MODULE_DIR / "ReporteTrimestral.py"))
	builder = getattr(module, "build_trimestral_dashboard_pdf")
	builder(output_path, payload)
	return output_path


CalibrationController._evaluation_key = staticmethod(_controller_evaluation_key)
CalibrationController.get_norm_display_name = _controller_get_norm_display_name
CalibrationController.get_norm_score_history = _controller_get_norm_score_history
CalibrationController._rename_related_history = _controller_rename_related_history
CalibrationController._norm_sort_key = staticmethod(_controller_norm_sort_key)
CalibrationController.get_latest_evaluation = _controller_get_latest_evaluation
CalibrationController.has_completed_form = _controller_has_completed_form
CalibrationController.get_principal_rows = _controller_get_principal_rows
CalibrationController.get_overview_metrics = _controller_get_overview_metrics
CalibrationController.save_evaluation = _controller_save_evaluation
CalibrationController.get_default_document_path = _controller_get_default_document_path
CalibrationController.generate_document = _controller_generate_document
CalibrationController.get_default_trimestral_report_path = _controller_get_default_trimestral_report_path
CalibrationController.generate_trimestral_dashboard_report = _controller_generate_trimestral_dashboard_report
