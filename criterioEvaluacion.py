from __future__ import annotations

import os
import re
import threading
from datetime import datetime
from tkinter import filedialog, messagebox

import customtkinter as ctk

from ui_shared import FONTS, STYLE, _position_toplevel, _safe_focus


PROTOCOL_QUESTIONS = [
	"El ejecutivo cumple con su horario establecido?",
	"Se cuenta con cobertura completa durante toda la jornada?",
	"Porta chaleco limpio y planchado?",
	"Se encuentra rasurado y con buena presentacion personal?",
	"Cumple con el protocolo de vestimenta ejecutiva de V&C?",
	"Lleva las botas limpias y en buen estado?",
	"Viste la camisa asignada por V&C?",
	"Saluda al cliente principal en campo de manera cordial?",
	"Se despide del cliente principal en campo de manera cordial?",
	"El colaborador escucha las necesidades del cliente?",
	"Usa un lenguaje claro y sin ambiguedades?",
	"Personaliza la comunicacion de acuerdo con el perfil del cliente?",
	"Responde con rapidez y mantiene informado al cliente?",
	"Muestra comprension hacia la situacion del cliente?",
	"Se anticipa a necesidades sin esperar preguntas?",
	"Es cordial, formal y respetuoso en todo momento?",
	"Cierra confirmando acuerdos con el cliente?",
	"Sabe clasificar entre los diferentes campos de aplicacion de las NOMs?",
	"Tiene conocimiento tecnico para evaluar criterios normativos?",
]

PROCESS_QUESTIONS = [
	"Realiza muestreo y revision de mercancias?",
	"Se escanean todas las cajas para identificar articulos?",
	"Valida la correcta colocacion de etiquetas?",
	"Genera completa y precisa la base de etiquetado NOM?",
	"Valida que la informacion entregada este actualizada?",
	"Solicita y archiva documentacion requerida por embarque?",
	"Toma fotografias claras y suficientes como evidencia?",
	"Carga documentacion a la nube con reglas correctas?",
	"Identifica tecnico liberador y tipo de liberacion en correo?",
]

TECHNICAL_RESULT_OPTIONS = ["C", "NC"]


class CriteriaNormDialog(ctk.CTkToplevel):
	def __init__(self, master, inspector_name: str, norms: list[str], on_select) -> None:
		super().__init__(master)
		self.on_select = on_select
		self.norms = norms or ["Sin norma"]
		self.norm_var = ctk.StringVar(value=self.norms[0])

		self.title(f"Seleccionar criterio - {inspector_name}")
		self.geometry("560x270")
		self.resizable(False, False)
		self.configure(fg_color=STYLE["fondo"])
		self.transient(master)
		self.grab_set()

		wrapper = ctk.CTkFrame(self, fg_color=STYLE["surface"], corner_radius=20)
		wrapper.pack(fill="both", expand=True, padx=18, pady=18)
		wrapper.grid_columnconfigure(0, weight=1)

		ctk.CTkLabel(
			wrapper,
			text="Formulario de Criterios por Cliente",
			font=FONTS["subtitle"],
			text_color=STYLE["texto_oscuro"],
		).grid(row=0, column=0, padx=16, pady=(18, 8), sticky="w")

		ctk.CTkLabel(
			wrapper,
			text="Selecciona la NOM acreditada para abrir el formulario y generar el PDF.",
			font=FONTS["small"],
			text_color="#6D7480",
			justify="left",
		).grid(row=1, column=0, padx=16, pady=(0, 12), sticky="w")

		ctk.CTkComboBox(
			wrapper,
			variable=self.norm_var,
			values=self.norms,
			height=38,
			fg_color="#FFFFFF",
			border_color="#94A3B8",
			button_color=STYLE["primario"],
			dropdown_hover_color=STYLE["primario"],
			state="readonly",
		).grid(row=2, column=0, padx=16, pady=(0, 16), sticky="ew")

		actions = ctk.CTkFrame(wrapper, fg_color="transparent")
		actions.grid(row=3, column=0, padx=16, pady=(0, 16), sticky="ew")
		actions.grid_columnconfigure(0, weight=1)
		actions.grid_columnconfigure(1, weight=1)

		ctk.CTkButton(
			actions,
			text="Cancelar",
			fg_color=STYLE["fondo"],
			text_color=STYLE["texto_oscuro"],
			hover_color="#E9ECEF",
			command=self.destroy,
		).grid(row=0, column=0, padx=(0, 8), sticky="ew")

		ctk.CTkButton(
			actions,
			text="Abrir formulario",
			fg_color=STYLE["secundario"],
			hover_color="#1D1D1D",
			command=self._submit,
		).grid(row=0, column=1, padx=(8, 0), sticky="ew")

		_position_toplevel(self, master, 560, 270)

	def _submit(self) -> None:
		self.on_select(self.norm_var.get().strip() or "Sin norma")
		self.destroy()


class CriteriaEvaluationDialog(ctk.CTkToplevel):
	def __init__(self, master, controller, inspector_name: str, can_edit: bool, on_saved, initial_norm: str) -> None:
		super().__init__(master)
		self.controller = controller
		self.inspector_name = inspector_name
		self.can_edit = can_edit
		self.on_saved = on_saved
		self.initial_norm = initial_norm

		self.norm_var = ctk.StringVar(value=initial_norm)
		self.client_var = ctk.StringVar(value="")
		self.date_var = ctk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
		self.supervisor_var = ctk.StringVar(value=str((controller.current_user or {}).get("name", "")).strip())
		self.image_folder_var = ctk.StringVar(value="")
		self.form_status_var = ctk.StringVar(value="")

		self.protocol_result_vars: list[ctk.StringVar] = []
		self.protocol_obs_vars: list[ctk.StringVar] = []
		self.process_result_vars: list[ctk.StringVar] = []
		self.process_obs_vars: list[ctk.StringVar] = []
		self.technical_rows: list[dict[str, object]] = []

		self.download_formato_button: ctk.CTkButton | None = None
		self.download_criterio_button: ctk.CTkButton | None = None
		self._document_generation_in_progress = False
		self._document_worker: threading.Thread | None = None

		self.title(f"Criterios por cliente - {inspector_name}")
		self.geometry("980x700")
		self.configure(fg_color=STYLE["fondo"])
		self.transient(master)
		self.grab_set()
		self.protocol("WM_DELETE_WINDOW", self._handle_close_request)

		self._build_ui()
		self._sync_download_state()
		_position_toplevel(self, master, 980, 700)

	def _build_ui(self) -> None:
		wrapper = ctk.CTkFrame(self, fg_color=STYLE["surface"], corner_radius=24)
		wrapper.pack(fill="both", expand=True, padx=18, pady=18)
		wrapper.grid_columnconfigure(0, weight=1)
		wrapper.grid_rowconfigure(1, weight=1)

		ctk.CTkLabel(
			wrapper,
			text=f"Criterios de evaluacion por cliente - {self.inspector_name}",
			font=FONTS["subtitle"],
			text_color=STYLE["texto_oscuro"],
		).grid(row=0, column=0, padx=18, pady=(16, 8), sticky="w")

		tabs = ctk.CTkTabview(wrapper, fg_color=STYLE["fondo"], corner_radius=20)
		tabs.grid(row=1, column=0, padx=18, pady=(0, 12), sticky="nsew")
		tabs.add("Informacion")
		tabs.add("Protocolo")
		tabs.add("Procesos")
		tabs.add("Tecnico")

		self._build_info_tab(tabs.tab("Informacion"))
		self._build_answers_tab(tabs.tab("Protocolo"), PROTOCOL_QUESTIONS, self.protocol_result_vars, self.protocol_obs_vars)
		self._build_answers_tab(tabs.tab("Procesos"), PROCESS_QUESTIONS, self.process_result_vars, self.process_obs_vars)
		self._build_technical_tab(tabs.tab("Tecnico"))

		actions = ctk.CTkFrame(wrapper, fg_color="transparent")
		actions.grid(row=2, column=0, padx=18, pady=(0, 14), sticky="ew")
		actions.grid_columnconfigure(0, weight=1)
		actions.grid_columnconfigure(1, weight=1)
		actions.grid_columnconfigure(2, weight=1)

		self.download_formato_button = ctk.CTkButton(
			actions,
			text="PDF Formato supervisión",
			fg_color=STYLE["primario"],
			text_color=STYLE["texto_oscuro"],
			hover_color="#D8C220",
			command=lambda: self._download_document("formato"),
		)
		self.download_formato_button.grid(row=0, column=0, padx=(0, 8), sticky="ew")

		self.download_criterio_button = ctk.CTkButton(
			actions,
			text="PDF Criterios por cliente",
			fg_color=STYLE["secundario"],
			text_color=STYLE["texto_claro"],
			hover_color="#1D1D1D",
			command=lambda: self._download_document("criterio"),
		)
		self.download_criterio_button.grid(row=0, column=1, padx=8, sticky="ew")

		ctk.CTkButton(
			actions,
			text="Cerrar",
			fg_color=STYLE["fondo"],
			text_color=STYLE["texto_oscuro"],
			hover_color="#E9ECEF",
			command=self._handle_close_request,
		).grid(row=0, column=2, padx=(8, 0), sticky="ew")

	def _build_info_tab(self, parent: ctk.CTkFrame) -> None:
		parent.grid_columnconfigure(0, weight=1)
		clients = self.controller.get_client_names() or ["Sin cliente"]
		if clients and not self.client_var.get().strip():
			self.client_var.set(clients[0])

		body = ctk.CTkFrame(parent, fg_color="transparent")
		body.pack(fill="both", expand=True, padx=12, pady=12)
		body.grid_columnconfigure(0, weight=1)

		self._add_field(
			body,
			0,
			"NOM evaluada",
			ctk.CTkEntry(body, textvariable=self.norm_var, height=38, border_color="#94A3B8", state="readonly"),
		)
		client_combo = ctk.CTkComboBox(
			body,
			variable=self.client_var,
			values=clients,
			height=38,
			fg_color="#FFFFFF",
			border_color="#94A3B8",
			button_color=STYLE["primario"],
			dropdown_hover_color=STYLE["primario"],
			state="readonly",
		)
		self._add_field(body, 1, "Cliente", client_combo)
		self._add_field(
			body,
			2,
			"Fecha",
			ctk.CTkEntry(body, textvariable=self.date_var, height=38, border_color="#94A3B8", state="readonly"),
		)
		self._add_field(
			body,
			3,
			"Ejecutivo supervisado",
			ctk.CTkEntry(body, height=38, border_color="#94A3B8", state="readonly", textvariable=ctk.StringVar(value=self.inspector_name)),
		)
		self._add_field(
			body,
			4,
			"Supervisor",
			ctk.CTkEntry(body, textvariable=self.supervisor_var, height=38, border_color="#94A3B8"),
		)

	def _build_answers_tab(self, parent: ctk.CTkFrame, questions: list[str], result_vars: list[ctk.StringVar], obs_vars: list[ctk.StringVar]) -> None:
		parent.grid_columnconfigure(0, weight=1)
		parent.grid_rowconfigure(0, weight=1)

		scroll = ctk.CTkScrollableFrame(parent, fg_color=STYLE["fondo"], corner_radius=0)
		scroll.grid(row=0, column=0, padx=12, pady=12, sticky="nsew")
		scroll.grid_columnconfigure(0, weight=1)

		for idx, question in enumerate(questions):
			result_var = ctk.StringVar(value="")
			obs_var = ctk.StringVar(value="")
			result_vars.append(result_var)
			obs_vars.append(obs_var)

			card = ctk.CTkFrame(scroll, fg_color="#FFFFFF", border_width=1, border_color="#E3E6EA", corner_radius=10)
			card.grid(row=idx, column=0, sticky="ew", pady=(0, 10))
			card.grid_columnconfigure(0, weight=1)

			ctk.CTkLabel(
				card,
				text=f"{idx + 1}. {question}",
				font=FONTS["small_bold"],
				text_color=STYLE["texto_oscuro"],
				justify="left",
				wraplength=860,
			).grid(row=0, column=0, padx=10, pady=(8, 6), sticky="w")

			controls = ctk.CTkFrame(card, fg_color="transparent")
			controls.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="ew")
			controls.grid_columnconfigure(0, weight=1)
			controls.grid_columnconfigure(1, weight=1)
			controls.grid_columnconfigure(2, weight=1)
			controls.grid_columnconfigure(3, weight=3)

			for col, (label, value) in enumerate([
				("Conforme", "conforme"),
				("No conforme", "no_conforme"),
				("No aplica", "no_aplica"),
			]):
				cell = ctk.CTkFrame(controls, fg_color="#F5F7FA", corner_radius=8)
				cell.grid(row=0, column=col, padx=(0 if col == 0 else 6, 6), sticky="nsew")
				ctk.CTkLabel(cell, text=label, font=FONTS["small"], text_color=STYLE["texto_oscuro"]).pack(pady=(6, 2))
				ctk.CTkRadioButton(cell, text="", variable=result_var, value=value, width=20).pack(pady=(0, 6))

			obs_cell = ctk.CTkFrame(controls, fg_color="#F5F7FA", corner_radius=8)
			obs_cell.grid(row=0, column=3, padx=(6, 0), sticky="nsew")
			ctk.CTkLabel(obs_cell, text="Observaciones", font=FONTS["small"], text_color=STYLE["texto_oscuro"]).pack(anchor="w", padx=8, pady=(6, 2))
			ctk.CTkEntry(obs_cell, textvariable=obs_var, height=32, border_color="#94A3B8").pack(fill="x", padx=8, pady=(0, 6))

	def _build_technical_tab(self, parent: ctk.CTkFrame) -> None:
		parent.grid_columnconfigure(0, weight=1)
		parent.grid_rowconfigure(3, weight=1)

		evidence = ctk.CTkFrame(parent, fg_color="transparent")
		evidence.grid(row=0, column=0, padx=12, pady=(12, 8), sticky="ew")
		evidence.grid_columnconfigure(1, weight=1)

		ctk.CTkButton(
			evidence,
			text="Cargar carpeta de imagenes",
			fg_color=STYLE["fondo"],
			text_color=STYLE["texto_oscuro"],
			border_width=1,
			border_color="#D5D8DC",
			hover_color="#E9ECEF",
			command=self._select_images_folder,
		).grid(row=0, column=0, padx=(0, 10), sticky="w")

		ctk.CTkLabel(
			evidence,
			textvariable=self.image_folder_var,
			font=FONTS["small"],
			text_color="#6D7480",
			anchor="w",
			justify="left",
		).grid(row=0, column=1, sticky="ew")

		header = ctk.CTkFrame(parent, fg_color=STYLE["primario"], corner_radius=10)
		header.grid(row=1, column=0, padx=12, pady=(0, 8), sticky="ew")
		header.grid_columnconfigure(0, weight=3)
		header.grid_columnconfigure(1, weight=2)
		header.grid_columnconfigure(2, weight=1)
		header.grid_columnconfigure(3, weight=3)
		header.grid_columnconfigure(4, weight=1)

		ctk.CTkLabel(header, text="SKU/ITEM/CODIGO/UPC", font=FONTS["small_bold"], text_color=STYLE["texto_oscuro"]).grid(row=0, column=0, padx=8, pady=8, sticky="w")
		ctk.CTkLabel(header, text="NOM aplicable", font=FONTS["small_bold"], text_color=STYLE["texto_oscuro"]).grid(row=0, column=1, padx=8, pady=8, sticky="w")
		ctk.CTkLabel(header, text="C/NC", font=FONTS["small_bold"], text_color=STYLE["texto_oscuro"]).grid(row=0, column=2, padx=8, pady=8)
		ctk.CTkLabel(header, text="Observaciones", font=FONTS["small_bold"], text_color=STYLE["texto_oscuro"]).grid(row=0, column=3, padx=8, pady=8, sticky="w")
		ctk.CTkButton(header, text="+ Fila", width=70, fg_color=STYLE["secundario"], hover_color="#1D1D1D", command=self._add_technical_row).grid(row=0, column=4, padx=8, pady=6)

		self.tech_container = ctk.CTkScrollableFrame(parent, fg_color=STYLE["fondo"], corner_radius=0)
		self.tech_container.grid(row=3, column=0, padx=12, pady=(0, 8), sticky="nsew")
		self.tech_container.grid_columnconfigure(0, weight=1)

		ctk.CTkLabel(parent, textvariable=self.form_status_var, font=FONTS["small"], text_color="#6D7480").grid(row=4, column=0, padx=12, pady=(0, 6), sticky="w")
		self._add_technical_row()

	def _add_technical_row(self) -> None:
		frame = ctk.CTkFrame(self.tech_container, fg_color="#FFFFFF", border_width=1, border_color="#E3E6EA", corner_radius=10)
		frame.grid_columnconfigure(0, weight=3)
		frame.grid_columnconfigure(1, weight=2)
		frame.grid_columnconfigure(2, weight=1)
		frame.grid_columnconfigure(3, weight=3)
		frame.grid_columnconfigure(4, weight=1)

		sku_var = ctk.StringVar(value="")
		norm_var = ctk.StringVar(value=self.norm_var.get().strip() or "Sin norma")
		result_var = ctk.StringVar(value="")
		obs_var = ctk.StringVar(value="")

		ctk.CTkEntry(frame, textvariable=sku_var, height=34, border_color="#94A3B8").grid(row=0, column=0, padx=8, pady=8, sticky="ew")
		ctk.CTkEntry(frame, textvariable=norm_var, height=34, border_color="#94A3B8").grid(row=0, column=1, padx=8, pady=8, sticky="ew")
		ctk.CTkComboBox(
			frame,
			variable=result_var,
			values=TECHNICAL_RESULT_OPTIONS,
			width=90,
			height=34,
			fg_color="#FFFFFF",
			border_color="#94A3B8",
			button_color=STYLE["primario"],
			dropdown_hover_color=STYLE["primario"],
			state="readonly",
		).grid(row=0, column=2, padx=8, pady=8)
		ctk.CTkEntry(frame, textvariable=obs_var, height=34, border_color="#94A3B8").grid(row=0, column=3, padx=8, pady=8, sticky="ew")

		remove_btn = ctk.CTkButton(
			frame,
			text="-",
			width=34,
			fg_color="#F7E0DE",
			text_color=STYLE["peligro"],
			hover_color="#F2C9C5",
			command=lambda row_frame=frame: self._remove_technical_row(row_frame),
		)
		remove_btn.grid(row=0, column=4, padx=(4, 8), pady=8)

		self.technical_rows.append(
			{
				"frame": frame,
				"sku_var": sku_var,
				"norm_var": norm_var,
				"result_var": result_var,
				"obs_var": obs_var,
			}
		)
		self._repack_technical_rows()

	def _remove_technical_row(self, row_frame) -> None:
		if len(self.technical_rows) == 1:
			self.technical_rows[0]["sku_var"].set("")
			self.technical_rows[0]["result_var"].set("")
			self.technical_rows[0]["obs_var"].set("")
			return
		self.technical_rows = [row for row in self.technical_rows if row.get("frame") != row_frame]
		row_frame.destroy()
		self._repack_technical_rows()

	def _repack_technical_rows(self) -> None:
		for idx, row in enumerate(self.technical_rows):
			row["frame"].grid(row=idx, column=0, pady=(0, 8), sticky="ew")

	def _select_images_folder(self) -> None:
		selected = filedialog.askdirectory(parent=self, title="Selecciona la carpeta de imagenes")
		if selected:
			self.image_folder_var.set(selected)

	def _add_field(self, parent, row: int, label: str, widget) -> None:
		ctk.CTkLabel(parent, text=label, font=FONTS["label"], text_color=STYLE["texto_oscuro"]).grid(row=row * 2, column=0, padx=8, pady=(12 if row == 0 else 8, 4), sticky="w")
		widget.grid(row=row * 2 + 1, column=0, padx=8, sticky="ew")

	@staticmethod
	def _collect_answer_rows(questions: list[str], result_vars: list[ctk.StringVar], obs_vars: list[ctk.StringVar]) -> list[dict[str, str]]:
		rows: list[dict[str, str]] = []
		for idx, question in enumerate(questions):
			rows.append(
				{
					"activity": question,
					"result": result_vars[idx].get().strip().lower(),
					"observations": obs_vars[idx].get().strip(),
				}
			)
		return rows

	def _collect_technical_rows(self) -> list[dict[str, str]]:
		rows: list[dict[str, str]] = []
		for idx, row in enumerate(self.technical_rows, start=1):
			sku = str(row["sku_var"].get()).strip()
			norm = str(row["norm_var"].get()).strip()
			c_nc = str(row["result_var"].get()).strip().upper()
			observations = str(row["obs_var"].get()).strip()
			if not any([sku, norm, c_nc, observations]):
				continue
			if not sku:
				raise ValueError(f"Completa SKU/ITEM/CODIGO/UPC en fila tecnica {idx}.")
			if not norm:
				raise ValueError(f"Completa NOM aplicable en fila tecnica {idx}.")
			if c_nc not in {"C", "NC"}:
				raise ValueError(f"Selecciona C o NC en fila tecnica {idx}.")
			rows.append(
				{
					"sku": sku,
					"applicable_norm": norm,
					"c_nc": c_nc,
					"result": "conforme" if c_nc == "C" else "no_conforme",
					"observations": observations,
				}
			)
		return rows

	@staticmethod
	def _section_score(answers: list[dict[str, str]]) -> tuple[float, dict[str, int]]:
		conforme = 0
		no_conforme = 0
		no_aplica = 0
		for answer in answers:
			result = str(answer.get("result", "")).strip().lower()
			if result == "conforme":
				conforme += 1
			elif result == "no_conforme":
				no_conforme += 1
			elif result == "no_aplica":
				no_aplica += 1
		aplicables = conforme + no_conforme
		score = round((conforme / aplicables) * 100.0, 1) if aplicables > 0 else 0.0
		return score, {
			"conforme": conforme,
			"no_conforme": no_conforme,
			"no_aplica": no_aplica,
			"aplicables": aplicables,
		}

	def _calculate_scores(
		self,
		selected_norm: str,
		protocol_answers: list[dict[str, str]],
		process_answers: list[dict[str, str]],
		technical_rows: list[dict[str, str]],
	) -> tuple[float, str, dict[str, int], dict[str, float]]:
		norm_stats: dict[str, dict[str, int]] = {}

		def ensure_norm(norm_name: str) -> dict[str, int]:
			clean = norm_name.strip() or "Sin norma"
			if clean not in norm_stats:
				norm_stats[clean] = {"conforme": 0, "no_conforme": 0, "no_aplica": 0}
			return norm_stats[clean]

		for answer in [*protocol_answers, *process_answers]:
			bucket = ensure_norm(selected_norm)
			result = str(answer.get("result", "")).strip().lower()
			if result == "conforme":
				bucket["conforme"] += 1
			elif result == "no_conforme":
				bucket["no_conforme"] += 1
			elif result == "no_aplica":
				bucket["no_aplica"] += 1

		for row in technical_rows:
			bucket = ensure_norm(str(row.get("applicable_norm", "")))
			result = str(row.get("result", "")).strip().lower()
			if result == "conforme":
				bucket["conforme"] += 1
			elif result == "no_conforme":
				bucket["no_conforme"] += 1

		total_conforme = 0
		total_no_conforme = 0
		total_no_aplica = 0
		score_by_norm: dict[str, float] = {}
		for norm_name, stats in norm_stats.items():
			conforme = int(stats.get("conforme", 0))
			no_conforme = int(stats.get("no_conforme", 0))
			no_aplica = int(stats.get("no_aplica", 0))
			aplicables = conforme + no_conforme
			score_by_norm[norm_name] = round((conforme / aplicables) * 100.0, 1) if aplicables > 0 else 0.0
			total_conforme += conforme
			total_no_conforme += no_conforme
			total_no_aplica += no_aplica

		final_score = round(sum(score_by_norm.values()) / len(score_by_norm), 1) if score_by_norm else 0.0
		status = "Estable" if final_score >= 90 else ("En seguimiento" if final_score >= 70 else "Critico")
		breakdown = {
			"conforme": total_conforme,
			"no_conforme": total_no_conforme,
			"no_aplica": total_no_aplica,
			"aplicables": total_conforme + total_no_conforme,
		}
		return final_score, status, breakdown, score_by_norm

	def _build_evaluation_payload(self) -> dict[str, object] | None:
		if not self.client_var.get().strip():
			messagebox.showerror("Criterios", "Debes seleccionar un cliente.", parent=self)
			return None
		if not self.supervisor_var.get().strip():
			messagebox.showerror("Criterios", "Debes capturar el nombre del supervisor.", parent=self)
			return None

		protocol_answers = self._collect_answer_rows(PROTOCOL_QUESTIONS, self.protocol_result_vars, self.protocol_obs_vars)
		process_answers = self._collect_answer_rows(PROCESS_QUESTIONS, self.process_result_vars, self.process_obs_vars)

		missing_protocol = [str(i + 1) for i, row in enumerate(protocol_answers) if str(row.get("result", "")).strip() not in {"conforme", "no_conforme", "no_aplica"}]
		if missing_protocol:
			messagebox.showerror("Criterios", "Completa Protocolo. Faltan preguntas: " + ", ".join(missing_protocol), parent=self)
			return None
		missing_process = [str(i + 1) for i, row in enumerate(process_answers) if str(row.get("result", "")).strip() not in {"conforme", "no_conforme", "no_aplica"}]
		if missing_process:
			messagebox.showerror("Criterios", "Completa Procesos. Faltan preguntas: " + ", ".join(missing_process), parent=self)
			return None

		try:
			technical_rows = self._collect_technical_rows()
		except ValueError as error:
			messagebox.showerror("Criterios", str(error), parent=self)
			return None

		if not technical_rows:
			messagebox.showerror("Criterios", "Debes capturar al menos una fila tecnica.", parent=self)
			return None

		selected_norm = self.norm_var.get().strip() or "Sin norma"
		score, status, score_breakdown, score_by_norm = self._calculate_scores(
			selected_norm,
			protocol_answers,
			process_answers,
			technical_rows,
		)
		soft_score, soft_breakdown = self._section_score(protocol_answers)
		technical_score, technical_breakdown = self._section_score(process_answers)

		return {
			"selected_norm": selected_norm,
			"client": self.client_var.get().strip(),
			"visit_date": self.date_var.get().strip(),
			"score": f"{score:.1f}",
			"status": status,
			"observations": "",
			"corrective_actions": "",
			"evaluator": self.supervisor_var.get().strip(),
			"inspector_supervised": self.inspector_name,
			"protocol_answers": protocol_answers,
			"process_answers": process_answers,
			"technical_normative_rows": technical_rows,
			"image_folder": self.image_folder_var.get().strip(),
			"score_breakdown": score_breakdown,
			"score_by_norm": score_by_norm,
			"soft_skills_score": soft_score,
			"technical_skills_score": technical_score,
			"soft_skills_breakdown": soft_breakdown,
			"technical_skills_breakdown": technical_breakdown,
		}

	def _persist_evaluation(self, payload: dict[str, object]) -> bool:
		try:
			self.controller.save_evaluation(self.inspector_name, payload)
		except ValueError as error:
			messagebox.showerror("Criterios", str(error), parent=self)
			return False
		self.on_saved()
		return True

	def _sync_download_state(self) -> None:
		can_use = self.can_edit and not self._document_generation_in_progress
		if self.download_formato_button is not None:
			self.download_formato_button.configure(state="normal" if can_use else "disabled")
		if self.download_criterio_button is not None:
			self.download_criterio_button.configure(state="normal" if can_use else "disabled")
		self.form_status_var.set(
			"Completa formulario para generar PDF con criterios por cliente y evidencias."
			if can_use
			else "Solo lectura para tu rol o hay una generacion en curso."
		)

	def _set_document_busy(self, busy: bool, status_message: str | None = None) -> None:
		self._document_generation_in_progress = busy
		if status_message is not None:
			self.form_status_var.set(status_message)
		self._sync_download_state()

	def _run_document_generation(self, kind: str, destination: str, selected_norm: str) -> None:
		try:
			output = self.controller.generate_document(self.inspector_name, kind, destination, selected_norm)
		except Exception as error:
			self.after(0, lambda error=error: self._finish_document_generation(error=error))
			return
		self.after(0, lambda output=output: self._finish_document_generation(output=output))

	def _finish_document_generation(self, output=None, error: Exception | None = None) -> None:
		self._document_worker = None
		self._set_document_busy(False)

		if error is not None:
			messagebox.showerror("Documentos", str(error), parent=self)
			return
		if output is None:
			return

		if hasattr(os, "startfile"):
			try:
				os.startfile(output)
			except OSError:
				pass
		messagebox.showinfo("Documentos", f"Archivo generado en:\n{output}", parent=self)

	def _download_document(self, kind: str) -> None:
		if not self.can_edit or self._document_generation_in_progress:
			return

		payload = self._build_evaluation_payload()
		if payload is None:
			return

		selected_norm = self.norm_var.get().strip() or self.initial_norm
		default_path = self.controller.get_default_document_path(self.inspector_name, kind, selected_norm)
		destination = filedialog.asksaveasfilename(
			parent=self,
			title="Guardar documento",
			defaultextension=".pdf",
			initialdir=str(default_path.parent),
			initialfile=default_path.name,
			filetypes=[("PDF", "*.pdf")],
		)
		if not destination:
			return

		if not self._persist_evaluation(payload):
			return

		mode_text = "Formato de supervision" if kind == "formato" else "Criterios por cliente"
		self._set_document_busy(True, f"Generando PDF de {mode_text}. Espera un momento...")
		self._document_worker = threading.Thread(
			target=self._run_document_generation,
			args=(kind, destination, selected_norm),
			daemon=True,
		)
		self._document_worker.start()

	def _handle_close_request(self) -> None:
		if self._document_generation_in_progress:
			return
		self.destroy()


class CriteriaEvaluationView(ctk.CTkFrame):
	def __init__(self, master, controller, can_edit: bool, on_change) -> None:
		super().__init__(master, fg_color=STYLE["fondo"])
		self.controller = controller
		self.can_edit = can_edit
		self.on_change = on_change
		self.search_var = ctk.StringVar(value="")
		self.results_var = ctk.StringVar(value="0 ejecutivos visibles")
		self.cards_frame: ctk.CTkScrollableFrame | None = None

		self.grid_columnconfigure(0, weight=1)
		self.grid_rowconfigure(1, weight=1)

		self._build_ui()
		self.search_var.trace_add("write", lambda *_args: self.refresh())
		self.refresh()

	def _build_ui(self) -> None:
		toolbar = ctk.CTkFrame(self, fg_color="#FFFFFF", corner_radius=20, border_width=1, border_color="#E3E6EA")
		toolbar.grid(row=0, column=0, padx=16, pady=(16, 10), sticky="ew")
		toolbar.grid_columnconfigure(1, weight=1)

		ctk.CTkLabel(toolbar, text="Buscar ejecutivo", font=FONTS["small_bold"], text_color=STYLE["texto_oscuro"]).grid(row=0, column=0, padx=(14, 10), pady=(12, 4), sticky="w")
		entry = ctk.CTkEntry(toolbar, textvariable=self.search_var, height=36, border_color="#94A3B8")
		entry.grid(row=0, column=1, padx=(0, 10), pady=(12, 4), sticky="ew")

		ctk.CTkButton(
			toolbar,
			text="Limpiar",
			width=90,
			fg_color=STYLE["fondo"],
			text_color=STYLE["texto_oscuro"],
			hover_color="#E9ECEF",
			command=lambda: self.search_var.set(""),
		).grid(row=0, column=2, padx=(0, 14), pady=(12, 4))

		ctk.CTkLabel(toolbar, textvariable=self.results_var, font=FONTS["small"], text_color="#6D7480").grid(row=1, column=0, columnspan=3, padx=14, pady=(0, 12), sticky="w")

		self.cards_frame = ctk.CTkScrollableFrame(self, fg_color=STYLE["fondo"], corner_radius=20)
		self.cards_frame.grid(row=1, column=0, padx=16, pady=(0, 14), sticky="nsew")
		for col in range(4):
			self.cards_frame.grid_columnconfigure(col, weight=1, uniform="criteria_cards")

		_safe_focus(entry)

	def refresh(self) -> None:
		if self.cards_frame is None:
			return
		for child in self.cards_frame.winfo_children():
			child.destroy()

		rows = self.controller.get_principal_rows(self.search_var.get(), "Todos")
		self.results_var.set(f"{len(rows)} ejecutivos visibles para criterios por cliente.")

		if not rows:
			empty = ctk.CTkFrame(self.cards_frame, fg_color="#FFFFFF", corner_radius=18, border_width=1, border_color="#E3E6EA")
			empty.grid(row=0, column=0, columnspan=4, padx=8, pady=8, sticky="ew")
			ctk.CTkLabel(empty, text="No hay ejecutivos para mostrar.", font=FONTS["label_bold"], text_color=STYLE["texto_oscuro"]).grid(row=0, column=0, padx=14, pady=(14, 6), sticky="w")
			ctk.CTkLabel(empty, text="Ajusta la busqueda para abrir un formulario de criterios.", font=FONTS["small"], text_color="#6D7480").grid(row=1, column=0, padx=14, pady=(0, 14), sticky="w")
			return

		for index, row in enumerate(rows):
			self._build_card(index, row)

	def _build_card(self, index: int, row: dict) -> None:
		if self.cards_frame is None:
			return

		card = ctk.CTkFrame(self.cards_frame, fg_color="#FFFFFF", corner_radius=18, border_width=1, border_color="#E3E6EA")
		card.grid(row=index // 4, column=index % 4, padx=8, pady=8, sticky="nsew")
		card.grid_columnconfigure(0, weight=1)

		ctk.CTkLabel(card, text=self._truncate_text(str(row.get("name", "")), 34), font=FONTS["label_bold"], text_color=STYLE["texto_oscuro"], wraplength=220, justify="left").grid(row=0, column=0, padx=12, pady=(12, 6), sticky="w")
		ctk.CTkLabel(card, text=f"Normas: {row.get('norm_count', 0)}", font=FONTS["small"], text_color="#6D7480").grid(row=1, column=0, padx=12, pady=(0, 2), sticky="w")
		ctk.CTkLabel(card, text=f"Estado: {row.get('status', '--')}", font=FONTS["small"], text_color="#6D7480").grid(row=2, column=0, padx=12, pady=(0, 8), sticky="w")

		ctk.CTkButton(
			card,
			text="Abrir Criterios",
			fg_color=STYLE["primario"],
			text_color=STYLE["texto_oscuro"],
			hover_color="#D8C220",
			command=lambda name=str(row.get("name", "")).strip(): self._open_criteria(name),
		).grid(row=3, column=0, padx=12, pady=(0, 12), sticky="ew")

	def _open_criteria(self, inspector_name: str) -> None:
		if not inspector_name:
			return
		norms = self.controller.get_accredited_norms(inspector_name) or ["Sin norma"]

		def _open_form(norm_token: str) -> None:
			CriteriaEvaluationDialog(
				self,
				self.controller,
				inspector_name,
				self.can_edit,
				self._handle_change,
				initial_norm=norm_token,
			)

		CriteriaNormDialog(self, inspector_name, norms, _open_form)

	def _handle_change(self) -> None:
		self.refresh()
		self.on_change()

	@staticmethod
	def _truncate_text(text: str, limit: int) -> str:
		clean = re.sub(r"\s+", " ", text or "").strip()
		if len(clean) <= limit:
			return clean
		return clean[: limit - 3].rstrip() + "..."
