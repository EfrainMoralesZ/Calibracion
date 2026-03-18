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


ROOT_DIR = Path(__file__).resolve().parent
DATA_DIR = ROOT_DIR / "data"
BD_FILE = DATA_DIR / "BD-Calibracion.json"
CLIENTS_FILE = DATA_DIR / "Clientes.json"
NORMS_FILE = DATA_DIR / "Normas.json"
USERS_FILE = DATA_DIR / "Usuarios.json"
STATE_FILE = DATA_DIR / "app_state.json"
HISTORY_DIR = DATA_DIR / "historico"
VISITS_DIR = DATA_DIR / "visitas"
TRIMESTRAL_DIR = DATA_DIR / "trimestral"
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


def _safe_folder_name(value: str) -> str:
	text = str(value or "").strip()
	if not text:
		return "sin_nombre"

	text = re.sub(r"[<>:\"/\\|?*]", "", text)
	text = re.sub(r"\s+", "_", text)
	text = text.strip("._")
	return text or "sin_nombre"


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
	return {"evaluations": {}}


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
		HISTORY_DIR.mkdir(parents=True, exist_ok=True)
		VISITS_DIR.mkdir(parents=True, exist_ok=True)
		TRIMESTRAL_DIR.mkdir(parents=True, exist_ok=True)
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

	def _consolidate_history_storage(self) -> None:
		if not HISTORY_DIR.exists():
			return

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
		return ["Calendario", "Trimestral"]

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

		return self._record_index_normalized.get(normalized_target)

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

		executives = sorted(
			str(user.get("name", "")).strip()
			for user in self.users_catalog
			if user.get("role") == "ejecutivo" and str(user.get("name", "")).strip()
		)
		self._assignable_inspectors_cache = executives or self.get_dashboard_people()
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
		clean_name = str(inspector_name or "").strip()
		if not clean_name:
			return []

		cached_history = self._history_cache.get(clean_name)
		if cached_history is None:
			history_path = self._history_file(clean_name)
			history = _read_json(history_path, [])
			if not isinstance(history, list):
				history = []
			cached_history = sorted(history, key=lambda item: item.get("saved_at", ""))
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
				"actions_text": "Formulario",
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
			"score_breakdown": score_breakdown,
			"soft_skills_breakdown": soft_skills_breakdown,
			"technical_skills_breakdown": technical_skills_breakdown,
			"score_by_norm": normalized_score_by_norm,
			"form_structure": "supervision_v2",
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
		visits = self._get_all_visits_cached()
		if name:
			visits = [visit for visit in visits if name in visit.get("inspectors", [])]

		candidate = current_user or self.current_user
		if candidate and candidate.get("role") == "ejecutivo":
			visits = [visit for visit in visits if candidate.get("name") in visit.get("inspectors", [])]

		return [dict(item) for item in visits]

	def list_trimestral_scores(
		self,
		inspector_name: str | None = None,
		year: int | None = None,
		quarter: str | None = None,
		current_user: dict[str, Any] | None = None,
	) -> list[dict[str, Any]]:
		scores = self._get_all_quarterly_scores_cached()
		candidate = current_user or self.current_user
		if not inspector_name and candidate and candidate.get("role") == "ejecutivo":
			inspector_name = str(candidate.get("name", "")).strip() or None
		if inspector_name:
			target = str(inspector_name).strip()
			scores = [item for item in scores if str(item.get("inspector", "")).strip() == target]
		if year is not None:
			year_value = int(year)
			scores = [item for item in scores if int(item.get("year", 0)) == year_value]
		if quarter:
			quarter_value = str(quarter).strip().upper()
			scores = [item for item in scores if str(item.get("quarter", "")).strip().upper() == quarter_value]

		return [dict(item) for item in scores]

	def save_trimestral_score(self, payload: dict[str, Any], score_id: str | None = None) -> dict[str, Any]:
		inspector = str(payload.get("inspector", "")).strip()
		quarter = str(payload.get("quarter", "")).strip().upper()
		year_value = str(payload.get("year", "")).strip()
		notes = str(payload.get("notes", "")).strip()
		score = _coerce_score(payload.get("score"))

		if not inspector:
			raise ValueError("Debes seleccionar un ejecutivo tecnico.")
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

		record = dict(existing or {})
		record["id"] = record.get("id") or uuid4().hex
		record["inspector"] = inspector
		record["quarter"] = quarter
		record["year"] = year
		record["score"] = score
		record["notes"] = notes
		record["evaluator"] = (self.current_user or {}).get("name", "Sistema")
		record["updated_at"] = _timestamp()

		q_dir = _quarter_dir(quarter, year)
		q_scores = _read_json(q_dir / "trimestral.json", [])
		q_scores.append(record)
		_write_json(q_dir / "trimestral.json", q_scores)
		self.reload()
		return record

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
		visit["acceptance_status"] = visit.get("acceptance_status", "asignada")
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
		"""Mark a visit as accepted by the assigned technical executive."""
		if not VISITS_DIR.exists():
			return
		for _wfolder in VISITS_DIR.iterdir():
			if not _wfolder.is_dir():
				continue
			_wvisits = _read_json(_wfolder / "visitas.json", [])
			_existing = next((v for v in _wvisits if v.get("id") == visit_id), None)
			if _existing:
				_existing["acceptance_status"] = "aceptada"
				_existing["updated_at"] = _timestamp()
				_write_json(_wfolder / "visitas.json", _wvisits)
				self.reload()
				return

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
				_existing["acceptance_status"] = "asignada"
				_existing["reassigned_from"] = old_inspectors
				_existing["reassigned_by"] = (self.current_user or {}).get("name", "Sistema")
				_existing["updated_at"] = _timestamp()
				_write_json(_wfolder / "visitas.json", _wvisits)
				affected = set(old_inspectors) | set(new_inspectors)
				for inspector_name in affected:
					self._sync_visit_history(inspector_name)
				self.reload()
				return

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
			module = _load_module(str(DOCUMENT_MODULE_DIR / "FormatoSupervision.py"))
			builder = getattr(module, "build_formato_supervision_pdf")
		elif document_kind == "criterio":
			module = _load_module(str(DOCUMENT_MODULE_DIR / "CriterioEvaluacionTecnica.py"))
			builder = getattr(module, "build_criterio_evaluacion_pdf")
		else:
			raise ValueError("Tipo de documento no soportado.")

		builder(output_path, payload)
		return output_path

	def _history_dir(self, inspector_name: str) -> Path:
		preferred_path = HISTORY_DIR / _safe_folder_name(inspector_name)

		target_identity = _folder_identity(inspector_name)
		cached_path = self._history_dir_index.get(target_identity) if target_identity else None
		if cached_path is not None and cached_path.exists():
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

			legacy_path = HISTORY_DIR / _safe_slug(inspector_name)
			if legacy_path.exists() and legacy_path.is_dir() and legacy_path not in candidates:
				candidates.append(legacy_path)

			for source in candidates:
				if source.resolve() == preferred_path.resolve():
					continue
				_merge_history_directories(preferred_path, source)

		if target_identity:
			self._history_dir_index[target_identity] = preferred_path
		return preferred_path

	def _history_file(self, inspector_name: str) -> Path:
		return self._history_dir(inspector_name) / "historico.json"

	def _visits_file(self, inspector_name: str) -> Path:
		return self._history_dir(inspector_name) / "visitas.json"

	def _write_history(self, inspector_name: str, history: list[dict[str, Any]]) -> None:
		_write_json(self._history_file(inspector_name), history)

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
			status = "En enfoque"
		else:
			status = "Estable"

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
		if clean_status_filter == "Completos" and not row["form_completed"]:
			continue
		if clean_status_filter == "Pendientes" and row["form_completed"]:
			continue
		if clean_status_filter == "En enfoque" and row["status"] != "En enfoque":
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
		"score_breakdown": score_breakdown,
		"soft_skills_breakdown": soft_skills_breakdown,
		"technical_skills_breakdown": technical_skills_breakdown,
		"score_by_norm": normalized_score_by_norm,
		"form_structure": "supervision_v2",
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
