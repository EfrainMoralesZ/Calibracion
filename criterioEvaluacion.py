from __future__ import annotations

import os
import re
import threading
from datetime import datetime
from pathlib import Path
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


class CriteriaInspectorDialog(ctk.CTkToplevel):
	def __init__(self, master, inspectors: list[str], on_select) -> None:
		super().__init__(master)
		self.on_select = on_select
		self.inspectors = inspectors or ["Sin ejecutivo"]
		self.inspector_var = ctk.StringVar(value=self.inspectors[0])

		self.title("Seleccionar ejecutivo")
		self.geometry("560x260")
		self.resizable(False, False)
		self.configure(fg_color=STYLE["fondo"])
		self.transient(master)
		self.grab_set()

		wrapper = ctk.CTkFrame(self, fg_color=STYLE["surface"], corner_radius=20)
		wrapper.pack(fill="both", expand=True, padx=18, pady=18)
		wrapper.grid_columnconfigure(0, weight=1)

		ctk.CTkLabel(
			wrapper,
			text="Selecciona ejecutivo para abrir criterios",
			font=FONTS["subtitle"],
			text_color=STYLE["texto_oscuro"],
		).grid(row=0, column=0, padx=16, pady=(18, 8), sticky="w")

		ctk.CTkComboBox(
			wrapper,
			variable=self.inspector_var,
			values=self.inspectors,
			height=38,
			fg_color="#FFFFFF",
			border_color="#94A3B8",
			button_color=STYLE["primario"],
			dropdown_hover_color=STYLE["primario"],
			state="readonly",
		).grid(row=1, column=0, padx=16, pady=(0, 16), sticky="ew")

		actions = ctk.CTkFrame(wrapper, fg_color="transparent")
		actions.grid(row=2, column=0, padx=16, pady=(0, 16), sticky="ew")
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
			text="Continuar",
			fg_color=STYLE["secundario"],
			hover_color="#1D1D1D",
			command=self._submit,
		).grid(row=0, column=1, padx=(8, 0), sticky="ew")

		_position_toplevel(self, master, 560, 260)

	def _submit(self) -> None:
		self.on_select(self.inspector_var.get().strip())
		self.destroy()


class CriteriaEvaluationDialog(ctk.CTkToplevel):
	def __init__(
		self,
		master,
		controller,
		inspector_name: str | None,
		can_edit: bool,
		on_saved,
		initial_norm: str | None = None,
		initial_client: str | None = None,
	) -> None:
		super().__init__(master)
		self.controller = controller
		self.inspector_name = str(inspector_name or "").strip()
		self.can_edit = can_edit
		self.on_saved = on_saved
		self.initial_norm = str(initial_norm or "").strip()
		self.initial_client = str(initial_client or "").strip()

		current_user = controller.current_user or {}
		self._inspector_locked = controller.is_executive_role(current_user)
		if self._inspector_locked:
			current_name = str(current_user.get("name", "")).strip()
			base_name = current_name or self.inspector_name
			self.inspector_options = [base_name] if base_name else []
		else:
			self.inspector_options = controller.get_assignable_inspectors() or []
			if self.inspector_name and self.inspector_name not in self.inspector_options:
				self.inspector_options.insert(0, self.inspector_name)

		default_inspector = self.inspector_name
		if not default_inspector and self.inspector_options:
			default_inspector = self.inspector_options[0]

		self.inspector_var = ctk.StringVar(value=default_inspector)
		self.norm_options: list[str] = []
		self.norm_var = ctk.StringVar(value=self.initial_norm or "Sin norma")
		self.client_var = ctk.StringVar(value="")
		self.date_var = ctk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
		self.supervisor_var = ctk.StringVar(value=str((controller.current_user or {}).get("name", "")).strip())
		self.executive_name_var = ctk.StringVar(
			value=str((controller.current_user or {}).get("name", "")).strip() or default_inspector
		)
		self.resolution_number_var = ctk.StringVar(value=controller.preview_criterio_resolution_number())
		self.product_var = ctk.StringVar(value="")
		self.evidence_summary_var = ctk.StringVar(value="Sin imagenes cargadas.")
		self.image_folder_var = ctk.StringVar(value="")
		self.form_status_var = ctk.StringVar(value="")
		self.header_title_var = ctk.StringVar(value="")

		self.protocol_result_vars: list[ctk.StringVar] = []
		self.protocol_obs_vars: list[ctk.StringVar] = []
		self.process_result_vars: list[ctk.StringVar] = []
		self.process_obs_vars: list[ctk.StringVar] = []
		self.technical_rows: list[dict[str, object]] = []
		self.evidence_files: list[str] = []
		self.comment_box: ctk.CTkTextbox | None = None
		self.resolution_box: ctk.CTkTextbox | None = None

		self.download_criterio_button: ctk.CTkButton | None = None
		self.norm_combo: ctk.CTkComboBox | None = None
		self._document_generation_in_progress = False
		self._document_worker: threading.Thread | None = None

		self._refresh_norms_for_inspector(self.initial_norm)
		self._update_header_title()

		self.title("Criterios por cliente")
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
			textvariable=self.header_title_var,
			font=FONTS["subtitle"],
			text_color=STYLE["texto_oscuro"],
		).grid(row=0, column=0, padx=18, pady=(16, 8), sticky="w")

		tabs = ctk.CTkTabview(wrapper, fg_color=STYLE["fondo"], corner_radius=20)
		tabs.grid(row=1, column=0, padx=18, pady=(0, 12), sticky="nsew")
		tabs.add("Informacion")
		tabs.add("Comentario")
		tabs.add("Evidencia")
		tabs.add("Resolucion")

		self._build_info_tab(tabs.tab("Informacion"))
		self._build_comment_tab(tabs.tab("Comentario"))
		self._build_evidence_tab(tabs.tab("Evidencia"))
		self._build_resolution_tab(tabs.tab("Resolucion"))

		actions = ctk.CTkFrame(wrapper, fg_color="transparent")
		actions.grid(row=2, column=0, padx=18, pady=(0, 14), sticky="ew")
		actions.grid_columnconfigure(0, weight=1)
		actions.grid_columnconfigure(1, weight=1)

		self.download_criterio_button = ctk.CTkButton(
			actions,
			text="PDF Criterios de evaluación",
			fg_color=STYLE["secundario"],
			text_color=STYLE["texto_claro"],
			hover_color="#1D1D1D",
			command=lambda: self._download_document("criterio"),
		)
		self.download_criterio_button.grid(row=0, column=0, padx=(0, 8), sticky="ew")

		ctk.CTkButton(
			actions,
			text="Cerrar",
			fg_color=STYLE["fondo"],
			text_color=STYLE["texto_oscuro"],
			hover_color="#E9ECEF",
			command=self._handle_close_request,
		).grid(row=0, column=1, padx=(8, 0), sticky="ew")

	def _build_info_tab(self, parent: ctk.CTkFrame) -> None:
		parent.grid_columnconfigure(0, weight=1)
		clients = self.controller.get_client_names() or ["Sin cliente"]
		if self.initial_client and self.initial_client in clients:
			self.client_var.set(self.initial_client)
		elif clients and not self.client_var.get().strip():
			self.client_var.set(clients[0])

		body = ctk.CTkScrollableFrame(parent, fg_color="transparent")
		body.pack(fill="both", expand=True, padx=12, pady=12)
		body.grid_columnconfigure(0, weight=1)
		body.grid_columnconfigure(1, weight=1)

		self.norm_combo = ctk.CTkComboBox(
			body,
			variable=self.norm_var,
			values=self.norm_options or ["Sin norma"],
			height=38,
			fg_color="#FFFFFF",
			border_color="#94A3B8",
			button_color=STYLE["primario"],
			dropdown_hover_color=STYLE["primario"],
			state="readonly",
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

		# Campos en dos columnas: (label, widget, columna)
		fields = [
			("Numero de resolucion",    ctk.CTkEntry(body, textvariable=self.resolution_number_var, height=38, border_color="#94A3B8", state="readonly"), 0),
			("Ejecutivo en sesion",     ctk.CTkEntry(body, textvariable=self.executive_name_var,    height=38, border_color="#94A3B8", state="readonly"), 1),
			("Norma aplicable",         self.norm_combo,                                                                                                  0),
			("Cliente",                 client_combo,                                                                                                     1),
			("Fecha",                   ctk.CTkEntry(body, textvariable=self.date_var,               height=38, border_color="#94A3B8", state="readonly"), 0),
			("Producto evaluado",       ctk.CTkEntry(body, textvariable=self.product_var,            height=38, border_color="#94A3B8"),                   1),
		]

		row_counters = [0, 0]
		for label, widget, col in fields:
			lbl_row = row_counters[col]
			px = (8, 6) if col == 0 else (6, 8)
			ctk.CTkLabel(body, text=label, font=FONTS["label"], text_color=STYLE["texto_oscuro"]).grid(
				row=lbl_row, column=col, padx=px, pady=(12 if lbl_row == 0 else 8, 4), sticky="w"
			)
			widget.grid(row=lbl_row + 1, column=col, padx=px, sticky="ew")
			row_counters[col] += 2

	def _build_comment_tab(self, parent: ctk.CTkFrame) -> None:
		parent.grid_columnconfigure(0, weight=1)
		parent.grid_rowconfigure(1, weight=1)

		ctk.CTkLabel(
			parent,
			text="Captura el comentario que se colocara en la primera columna del formato.",
			font=FONTS["small"],
			text_color="#6D7480",
			justify="left",
		).grid(row=0, column=0, padx=12, pady=(12, 8), sticky="w")

		self.comment_box = ctk.CTkTextbox(parent, corner_radius=16, border_width=1, border_color="#D5D8DC")
		self.comment_box.grid(row=1, column=0, padx=12, pady=(0, 12), sticky="nsew")

	def _build_evidence_tab(self, parent: ctk.CTkFrame) -> None:
		parent.grid_columnconfigure(0, weight=1)
		parent.grid_rowconfigure(2, weight=1)

		toolbar = ctk.CTkFrame(parent, fg_color="transparent")
		toolbar.grid(row=0, column=0, padx=12, pady=(12, 8), sticky="ew")
		toolbar.grid_columnconfigure(1, weight=1)

		ctk.CTkButton(
			toolbar,
			text="Cargar imagenes",
			fg_color=STYLE["primario"],
			text_color=STYLE["texto_oscuro"],
			hover_color="#D8C220",
			command=self._select_evidence_files,
		).grid(row=0, column=0, padx=(0, 10), sticky="w")

		ctk.CTkLabel(
			toolbar,
			textvariable=self.evidence_summary_var,
			font=FONTS["small"],
			text_color="#6D7480",
			justify="left",
			anchor="w",
		).grid(row=0, column=1, sticky="ew")

		ctk.CTkLabel(
			parent,
			text="Las imagenes cargadas se insertaran en la columna EVIDENCIA del PDF con tamano uniforme.",
			font=FONTS["small"],
			text_color="#6D7480",
			justify="left",
		).grid(row=1, column=0, padx=12, pady=(0, 8), sticky="w")

		preview = ctk.CTkTextbox(parent, corner_radius=16, border_width=1, border_color="#D5D8DC")
		preview.grid(row=2, column=0, padx=12, pady=(0, 12), sticky="nsew")
		preview.insert("1.0", "No hay imagenes seleccionadas.")
		preview.configure(state="disabled")
		self.evidence_preview_box = preview

	def _build_resolution_tab(self, parent: ctk.CTkFrame) -> None:
		parent.grid_columnconfigure(0, weight=1)
		parent.grid_rowconfigure(1, weight=1)

		ctk.CTkLabel(
			parent,
			text="Captura la resolucion que se colocara en la tercera columna del formato.",
			font=FONTS["small"],
			text_color="#6D7480",
			justify="left",
		).grid(row=0, column=0, padx=12, pady=(12, 8), sticky="w")

		self.resolution_box = ctk.CTkTextbox(parent, corner_radius=16, border_width=1, border_color="#D5D8DC")
		self.resolution_box.grid(row=1, column=0, padx=12, pady=(0, 12), sticky="nsew")

	def _get_selected_inspector(self) -> str:
		return self.inspector_var.get().strip() or self.inspector_name

	def _update_header_title(self) -> None:
		inspector = self._get_selected_inspector()
		if inspector:
			self.header_title_var.set(f"Criterios de evaluacion por cliente - {inspector}")
			self.title(f"Criterios por cliente - {inspector}")
		else:
			self.header_title_var.set("Criterios de evaluacion por cliente")
			self.title("Criterios por cliente")

	def _refresh_norms_for_inspector(self, preferred_norm: str | None = None) -> None:
		catalog_norms = [self.controller.get_norm_display_name(item.get("nom") or item.get("token")) for item in self.controller.get_catalog_norms()]
		self.norm_options = catalog_norms or ["Sin norma"]
		if self.norm_combo is not None:
			self.norm_combo.configure(values=self.norm_options)

		wanted = str(preferred_norm or "").strip()
		if not wanted or wanted not in self.norm_options:
			wanted = self.norm_options[0]
		self.norm_var.set(wanted)

	def _on_inspector_change(self) -> None:
		self._refresh_norms_for_inspector()
		self._update_header_title()

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

	def _select_evidence_files(self) -> None:
		selected = filedialog.askopenfilenames(
			parent=self,
			title="Selecciona imagenes de evidencia",
			filetypes=[("Imagenes", "*.png *.jpg *.jpeg *.webp")],
		)
		if not selected:
			return
		self.evidence_files = [str(path) for path in selected]
		self._refresh_evidence_preview()

	def _refresh_evidence_preview(self) -> None:
		total = len(self.evidence_files)
		if total == 0:
			self.evidence_summary_var.set("Sin imagenes cargadas.")
			preview_text = "No hay imagenes seleccionadas."
		else:
			self.evidence_summary_var.set(f"{total} imagen(es) cargadas para evidencia.")
			visible_names = [os.path.basename(path) for path in self.evidence_files[:8]]
			preview_text = "\n".join(visible_names)
			if total > len(visible_names):
				preview_text += f"\n... y {total - len(visible_names)} imagen(es) mas"
		if hasattr(self, "evidence_preview_box") and self.evidence_preview_box is not None:
			self.evidence_preview_box.configure(state="normal")
			self.evidence_preview_box.delete("1.0", "end")
			self.evidence_preview_box.insert("1.0", preview_text)
			self.evidence_preview_box.configure(state="disabled")

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
		inspector_name = self._get_selected_inspector()
		if not inspector_name:
			messagebox.showerror("Criterios", "Debes seleccionar un ejecutivo supervisado.", parent=self)
			return None
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
			"inspector_supervised": inspector_name,
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

	def _read_textbox(self, textbox: ctk.CTkTextbox | None) -> str:
		if textbox is None:
			return ""
		return textbox.get("1.0", "end").strip()

	def _build_criterio_payload(self) -> dict[str, object] | None:
		client_name = self.client_var.get().strip()
		selected_norm = self.norm_var.get().strip() or "Sin norma"
		executive_name = self.executive_name_var.get().strip() or self._get_selected_inspector()
		product_name = self.product_var.get().strip()
		comment_text = self._read_textbox(self.comment_box)
		resolution_text = self._read_textbox(self.resolution_box)

		if not client_name:
			messagebox.showerror("Criterios", "Debes seleccionar un cliente.", parent=self)
			return None
		if not executive_name:
			messagebox.showerror("Criterios", "No se pudo identificar al ejecutivo en sesion.", parent=self)
			return None
		if not product_name:
			messagebox.showerror("Criterios", "Debes capturar el producto evaluado.", parent=self)
			return None
		if not comment_text:
			messagebox.showerror("Criterios", "Debes capturar el comentario.", parent=self)
			return None
		if not resolution_text:
			messagebox.showerror("Criterios", "Debes capturar la resolucion.", parent=self)
			return None

		return {
			"resolution_number": self.resolution_number_var.get().strip(),
			"visit_date": self.date_var.get().strip(),
			"client": client_name,
			"executive_name": executive_name,
			"selected_norm": selected_norm,
			"evaluated_product": product_name,
			"comment": comment_text,
			"resolution_text": resolution_text,
			"evidence_files": list(self.evidence_files),
			"reviewed_by": self.supervisor_var.get().strip(),
			"inspector_supervised": self._get_selected_inspector(),
		}

	def _reset_form(self) -> None:
		"""Limpia todos los campos del formulario de criterios."""
		self.product_var.set("")
		if self.comment_box is not None:
			self.comment_box.delete("1.0", "end")
		if self.resolution_box is not None:
			self.resolution_box.delete("1.0", "end")
		self.evidence_files = []
		self._refresh_evidence_preview()

	def _open_history(self) -> None:
		"""Abre ventana con el historial de PDFs de criterios para este cliente."""
		client_name = self.client_var.get().strip()
		if not client_name:
			messagebox.showinfo("Historial", "Selecciona un cliente para ver el historial.", parent=self)
			return
		
		history = self.controller.get_criteria_history(client_name)
		if not history:
			messagebox.showinfo("Historial", f"No hay criterios generados para {client_name}.", parent=self)
			return
		
		history_window = ctk.CTkToplevel(self)
		history_window.title(f"Historial de criterios - {client_name}")
		history_window.geometry("900x500")
		history_window.configure(fg_color=STYLE["fondo"])
		history_window.transient(self)
		_position_toplevel(history_window, self, 900, 500)
		
		wrapper = ctk.CTkFrame(history_window, fg_color=STYLE["surface"], corner_radius=20)
		wrapper.pack(fill="both", expand=True, padx=18, pady=18)
		wrapper.grid_columnconfigure(0, weight=1)
		wrapper.grid_rowconfigure(1, weight=1)
		
		ctk.CTkLabel(
			wrapper,
			text=f"Historial de criterios - {client_name}",
			font=FONTS["subtitle"],
			text_color=STYLE["texto_oscuro"],
		).grid(row=0, column=0, padx=18, pady=(16, 12), sticky="w")
		
		scroll_frame = ctk.CTkScrollableFrame(wrapper, fg_color=STYLE["fondo"], corner_radius=0)
		scroll_frame.grid(row=1, column=0, padx=18, pady=(0, 12), sticky="nsew")
		scroll_frame.grid_columnconfigure(0, weight=1)
		
		for idx, doc in enumerate(sorted(history, key=lambda x: x.get("generated_at", ""), reverse=True)):
			card = ctk.CTkFrame(scroll_frame, fg_color="#FFFFFF", border_width=1, border_color="#E3E6EA", corner_radius=12)
			card.grid(row=idx, column=0, sticky="ew", pady=(0, 10))
			card.grid_columnconfigure(0, weight=1)
			card.grid_columnconfigure(1, weight=0)
			
			info_text = (
				f"Res: {doc.get('resolution_number', '-')} | "
				f"Norma: {doc.get('selected_norm', '-')} | "
				f"Producto: {doc.get('evaluated_product', '-')} | "
				f"Ejecutivo: {doc.get('executive_name', '-')} | "
				f"Fecha: {doc.get('generated_at', '-')}"
			)
			ctk.CTkLabel(
				card,
				text=info_text,
				font=FONTS["small"],
				text_color=STYLE["texto_oscuro"],
				justify="left",
				anchor="w",
			).grid(row=0, column=0, padx=12, pady=10, sticky="ew")
			
			output_path = doc.get("output_path", "")
			if output_path and Path(output_path).exists():
				ctk.CTkButton(
					card,
					text="Abrir",
					width=80,
					fg_color=STYLE["primario"],
					text_color=STYLE["texto_oscuro"],
					hover_color="#D8C220",
					command=lambda path=output_path: self._open_pdf(path),
				).grid(row=0, column=1, padx=12, pady=10)
			
		ctk.CTkButton(
			wrapper,
			text="Cerrar",
			fg_color=STYLE["fondo"],
			text_color=STYLE["texto_oscuro"],
			hover_color="#E9ECEF",
			command=history_window.destroy,
		).grid(row=2, column=0, padx=18, pady=(0, 12), sticky="ew")
	
	def _open_pdf(self, path: str) -> None:
		"""Abre un archivo PDF."""
		pdf_path = Path(path)
		if not pdf_path.exists():
			messagebox.showerror("Historial", f"El archivo no existe:\n{path}", parent=self)
			return
		if hasattr(os, "startfile"):
			try:
				os.startfile(pdf_path)
			except OSError as e:
				messagebox.showerror("Historial", f"No se pudo abrir el archivo:\n{e}", parent=self)
	
	def _persist_evaluation(self, payload: dict[str, object]) -> bool:
		inspector_name = self._get_selected_inspector()
		try:
			self.controller.save_evaluation(inspector_name, payload)
		except ValueError as error:
			messagebox.showerror("Criterios", str(error), parent=self)
			return False
		self.on_saved()
		return True

	def _sync_download_state(self) -> None:
		can_use = self.can_edit and not self._document_generation_in_progress
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

	def _run_document_generation(
		self,
		inspector_name: str,
		kind: str,
		destination: str,
		selected_norm: str,
		payload_override: dict[str, object] | None = None,
	) -> None:
		try:
			if kind == "criterio" and payload_override is not None:
				output = self.controller.generate_criterio_document(destination, payload_override)
			else:
				output = self.controller.generate_document(inspector_name, kind, destination, selected_norm)
		except Exception as error:
			self.after(0, lambda error=error, kind=kind: self._finish_document_generation(error=error, kind=kind))
			return
		self.after(0, lambda output=output, kind=kind: self._finish_document_generation(output=output, kind=kind))

	def _finish_document_generation(self, output=None, error: Exception | None = None, kind: str | None = None) -> None:
		self._document_worker = None
		self._set_document_busy(False)
		if kind == "criterio":
			self.resolution_number_var.set(self.controller.preview_criterio_resolution_number())

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
		
		if kind == "criterio":
			self._reset_form()

	def _download_document(self, kind: str) -> None:
		if not self.can_edit or self._document_generation_in_progress:
			return

		if kind == "criterio":
			payload = self._build_criterio_payload()
			if payload is None:
				return
			inspector_name = str(payload.get("executive_name", "")).strip() or self._get_selected_inspector()
			selected_norm = str(payload.get("selected_norm", "")).strip() or self.initial_norm or "Sin norma"
		else:
			payload = self._build_evaluation_payload()
			if payload is None:
				return
			inspector_name = self._get_selected_inspector()
			if not inspector_name:
				messagebox.showerror("Criterios", "Debes seleccionar un ejecutivo supervisado.", parent=self)
				return
			selected_norm = self.norm_var.get().strip() or self.initial_norm or "Sin norma"

		default_path = self.controller.get_default_document_path(inspector_name, kind, selected_norm)
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

		if kind != "criterio" and not self._persist_evaluation(payload):
			return

		mode_text = "Formato de supervision" if kind == "formato" else "Criterios por cliente"
		self._set_document_busy(True, f"Generando PDF de {mode_text}. Espera un momento...")
		self._document_worker = threading.Thread(
			target=self._run_document_generation,
			args=(inspector_name, kind, destination, selected_norm, payload if kind == "criterio" else None),
			daemon=True,
		)
		self._document_worker.start()

	def _handle_close_request(self) -> None:
		if self._document_generation_in_progress:
			return
		self.destroy()


class CriteriaEvaluationView(ctk.CTkFrame):
	PAGE_SIZE = 20
	CARD_COLUMNS = 1

	def __init__(self, master, controller, can_edit: bool, on_change) -> None:
		super().__init__(master, fg_color=STYLE["fondo"])
		self.controller = controller
		self.can_edit = can_edit
		self.on_change = on_change
		self.search_var = ctk.StringVar(value="")
		self.results_var = ctk.StringVar(value="0 clientes visibles")
		self.cards_frame: ctk.CTkScrollableFrame | None = None
		self.pager_frame: ctk.CTkFrame | None = None
		self.current_page = 0
		self._filtered_clients: list[str] = []

		self.grid_columnconfigure(0, weight=1)
		self.grid_rowconfigure(1, weight=1)

		self._build_ui()
		self.search_var.trace_add("write", lambda *_args: self._on_search_change())
		self.refresh()

	def _build_ui(self) -> None:
		toolbar = ctk.CTkFrame(self, fg_color="#FFFFFF", corner_radius=20, border_width=1, border_color="#E3E6EA")
		toolbar.grid(row=0, column=0, padx=16, pady=(16, 10), sticky="ew")
		toolbar.grid_columnconfigure(1, weight=1)

		ctk.CTkLabel(toolbar, text="Buscar cliente", font=FONTS["small_bold"], text_color=STYLE["texto_oscuro"]).grid(row=0, column=0, padx=(14, 10), pady=(12, 4), sticky="w")
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
		self.cards_frame.grid_columnconfigure(0, weight=1)

		self.pager_frame = ctk.CTkFrame(self, fg_color="transparent")
		self.pager_frame.grid(row=2, column=0, padx=16, pady=(0, 12), sticky="ew")

		_safe_focus(entry)

	def _on_search_change(self) -> None:
		self.current_page = 0
		self.refresh()

	def refresh(self) -> None:
		if self.cards_frame is None:
			return
		for child in self.cards_frame.winfo_children():
			child.destroy()

		search_text = self.search_var.get().strip().casefold()
		clients = self.controller.get_client_names() or []
		if search_text:
			clients = [name for name in clients if search_text in str(name).casefold()]
		self._filtered_clients = clients

		total = len(clients)
		total_pages = max(1, (total + self.PAGE_SIZE - 1) // self.PAGE_SIZE)
		self.current_page = max(0, min(self.current_page, total_pages - 1))

		start = self.current_page * self.PAGE_SIZE
		page_clients = clients[start : start + self.PAGE_SIZE]
		if total > 0:
			self.results_var.set(
				f"{total} clientes visibles para criterios por cliente. "
				f"Pagina {self.current_page + 1} de {total_pages}."
			)
		else:
			self.results_var.set("0 clientes visibles para criterios por cliente.")

		if not page_clients:
			empty = ctk.CTkFrame(self.cards_frame, fg_color="#FFFFFF", corner_radius=18, border_width=1, border_color="#E3E6EA")
			empty.grid(row=0, column=0, columnspan=self.CARD_COLUMNS, padx=8, pady=8, sticky="ew")
			ctk.CTkLabel(empty, text="No hay clientes para mostrar.", font=FONTS["label_bold"], text_color=STYLE["texto_oscuro"]).grid(row=0, column=0, padx=14, pady=(14, 6), sticky="w")
			ctk.CTkLabel(empty, text="Ajusta la busqueda para abrir un formulario de criterios.", font=FONTS["small"], text_color="#6D7480").grid(row=1, column=0, padx=14, pady=(0, 14), sticky="w")
			self._rebuild_pager(total)
			return

		for index, client_name in enumerate(page_clients):
			self._build_card(index, client_name)

		self._rebuild_pager(total)

	def _build_card(self, index: int, client_name: str) -> None:
		if self.cards_frame is None:
			return

		has_agreements = bool(self.controller.get_client_agreements(client_name))
		border_color = "#F0C040" if has_agreements else "#E3E6EA"

		card = ctk.CTkFrame(self.cards_frame, fg_color="#FFFFFF", corner_radius=14, border_width=1, border_color=border_color)
		card.grid(row=index, column=0, padx=6, pady=5, sticky="ew")
		card.grid_columnconfigure(0, weight=1)
		card.grid_columnconfigure(1, weight=0)

		name_row = ctk.CTkFrame(card, fg_color="transparent")
		name_row.grid(row=0, column=0, padx=(12, 8), pady=(10, 2) if has_agreements else 10, sticky="w")

		ctk.CTkLabel(
			name_row,
			text=client_name,
			font=FONTS["small_bold"],
			text_color=STYLE["texto_oscuro"],
			wraplength=540,
			justify="left",
		).pack(side="left")

		if has_agreements:
			badge = ctk.CTkFrame(name_row, fg_color="#FFF3CD", corner_radius=6)
			badge.pack(side="left", padx=(10, 0))
			ctk.CTkLabel(
				badge,
				text="📋 Con acuerdos",
				font=FONTS["small"],
				text_color="#856404",
			).pack(padx=8, pady=2)

		actions = ctk.CTkFrame(card, fg_color="transparent")
		actions.grid(row=0, column=1, padx=(8, 10), pady=8, sticky="e")
		ctk.CTkButton(
			actions,
			text="Formulario",
			fg_color=STYLE["primario"],
			text_color=STYLE["texto_oscuro"],
			hover_color="#D8C220",
			width=120,
			height=30,
			command=lambda client=client_name: self._open_criteria(client),
		).pack(side="left", padx=(0, 8))

		ctk.CTkButton(
			actions,
			text="Historial",
			fg_color=STYLE["secundario"],
			text_color=STYLE["texto_claro"],
			hover_color="#1D1D1D",
			width=120,
			height=30,
			command=lambda client=client_name: self._open_history_for_client(client),
		).pack(side="left", padx=(0, 8))

		ctk.CTkButton(
			actions,
			text="Acuerdos",
			fg_color=STYLE["fondo"],
			text_color=STYLE["texto_oscuro"],
			hover_color="#E9ECEF",
			width=120,
			height=30,
			command=lambda client=client_name: self._upload_agreement_for_client(client),
		).pack(side="left")

	def _open_criteria(self, client_name: str) -> None:
		if not client_name:
			return

		current_user = self.controller.current_user or {}
		selected_inspector: str | None = None
		if self.controller.is_executive_role(current_user):
			selected_inspector = str(current_user.get("name", "")).strip()
			if not selected_inspector:
				messagebox.showerror("Criterios", "No se pudo identificar al ejecutivo actual.", parent=self)
				return
		elif not (self.controller.get_assignable_inspectors() or []):
			messagebox.showinfo("Criterios", "No hay ejecutivos disponibles para abrir criterios.", parent=self)
			return

		CriteriaEvaluationDialog(
			self,
			self.controller,
			selected_inspector,
			self.can_edit,
			self._handle_change,
			initial_client=client_name,
		)

	def _upload_agreement_for_client(self, client_name: str) -> None:
		if not client_name:
			return

		source_path = filedialog.askopenfilename(
			title=f"Selecciona minuta de acuerdos - {client_name}",
			filetypes=[("Archivos PDF", "*.pdf")],
		)
		if not source_path:
			return

		try:
			saved_path = self.controller.save_client_agreement_pdf(client_name, source_path)
		except ValueError as error:
			messagebox.showerror("Acuerdos", str(error), parent=self)
			return

		messagebox.showinfo(
			"Acuerdos",
			f"Se guardo la minuta en:\n{saved_path}",
			parent=self,
		)

	def _open_history_for_client(self, client_name: str) -> None:
		if not client_name:
			return

		history = self.controller.get_criteria_history(client_name)
		agreements = self.controller.get_client_agreements(client_name)
		if not history and not agreements:
			messagebox.showinfo("Historial", f"No hay criterios ni acuerdos para {client_name}.", parent=self)
			return

		history_window = ctk.CTkToplevel(self)
		history_window.title(f"Historial de criterios - {client_name}")
		history_window.geometry("980x560")
		history_window.configure(fg_color=STYLE["fondo"])
		history_window.transient(self)
		_position_toplevel(history_window, self, 980, 560)

		wrapper = ctk.CTkFrame(history_window, fg_color=STYLE["surface"], corner_radius=20)
		wrapper.pack(fill="both", expand=True, padx=18, pady=18)
		wrapper.grid_columnconfigure(0, weight=1)
		wrapper.grid_rowconfigure(2, weight=1)

		ctk.CTkLabel(
			wrapper,
			text=f"Historial de criterios - {client_name}",
			font=FONTS["subtitle"],
			text_color=STYLE["texto_oscuro"],
		).grid(row=0, column=0, padx=18, pady=(16, 12), sticky="w")

		norm_values = sorted(
			{
				str(item.get("selected_norm", "")).strip() or "Sin norma"
				for item in history
			}
		)
		norm_filter_var = ctk.StringVar(value="Todas")

		filters = ctk.CTkFrame(wrapper, fg_color="transparent")
		filters.grid(row=1, column=0, padx=18, pady=(0, 10), sticky="ew")
		filters.grid_columnconfigure(1, weight=1)

		ctk.CTkLabel(
			filters,
			text="Filtrar por norma",
			font=FONTS["small_bold"],
			text_color=STYLE["texto_oscuro"],
		).grid(row=0, column=0, padx=(0, 10), sticky="w")

		norm_combo = ctk.CTkComboBox(
			filters,
			variable=norm_filter_var,
			values=["Todas", *norm_values],
			height=34,
			fg_color="#FFFFFF",
			border_color="#94A3B8",
			button_color=STYLE["primario"],
			dropdown_hover_color=STYLE["primario"],
			state="readonly",
		)
		norm_combo.grid(row=0, column=1, sticky="w")

		scroll_frame = ctk.CTkScrollableFrame(wrapper, fg_color=STYLE["fondo"], corner_radius=0)
		scroll_frame.grid(row=2, column=0, padx=18, pady=(0, 12), sticky="nsew")
		scroll_frame.grid_columnconfigure(0, weight=1)

		def _render_rows() -> None:
			for child in scroll_frame.winfo_children():
				child.destroy()

			row_idx = 0
			if agreements:
				ctk.CTkLabel(
					scroll_frame,
					text="Acuerdos del cliente",
					font=FONTS["small_bold"],
					text_color=STYLE["texto_oscuro"],
				).grid(row=row_idx, column=0, padx=6, pady=(0, 8), sticky="w")
				row_idx += 1

				for agreement in agreements:
					agreement_row = ctk.CTkFrame(
						scroll_frame,
						fg_color="#FFFFFF",
						border_width=1,
						border_color="#E3E6EA",
						corner_radius=12,
					)
					agreement_row.grid(row=row_idx, column=0, sticky="ew", pady=(0, 10))
					agreement_row.grid_columnconfigure(0, weight=1)
					agreement_row.grid_columnconfigure(1, weight=0)

					text = f"{agreement.get('title', 'Acuerdo')} | Fecha: {agreement.get('generated_at', '-') }"
					ctk.CTkLabel(
						agreement_row,
						text=text,
						font=FONTS["small"],
						text_color=STYLE["texto_oscuro"],
						justify="left",
						anchor="w",
					).grid(row=0, column=0, padx=12, pady=10, sticky="ew")

					agreement_path = str(agreement.get("output_path", "")).strip()
					if agreement_path and Path(agreement_path).exists():
						ctk.CTkButton(
							agreement_row,
							text="Abrir acuerdo",
							width=110,
							fg_color=STYLE["primario"],
							text_color=STYLE["texto_oscuro"],
							hover_color="#D8C220",
							command=lambda path=agreement_path: self._open_pdf_from_history(path),
						).grid(row=0, column=1, padx=12, pady=10)
					row_idx += 1

			if agreements and history:
				ctk.CTkLabel(
					scroll_frame,
					text="Criterios generados",
					font=FONTS["small_bold"],
					text_color=STYLE["texto_oscuro"],
				).grid(row=row_idx, column=0, padx=6, pady=(4, 8), sticky="w")
				row_idx += 1

			selected_norm = norm_filter_var.get().strip()
			filtered_history = sorted(history, key=lambda x: x.get("generated_at", ""), reverse=True)
			if selected_norm and selected_norm != "Todas":
				filtered_history = [
					item
					for item in filtered_history
					if (str(item.get("selected_norm", "")).strip() or "Sin norma") == selected_norm
				]

			if not filtered_history:
				ctk.CTkLabel(
					scroll_frame,
					text="No hay criterios para la norma seleccionada.",
					font=FONTS["small"],
					text_color="#6D7480",
				).grid(row=row_idx, column=0, padx=6, pady=(2, 8), sticky="w")
				return

			for doc in filtered_history:
				row = ctk.CTkFrame(scroll_frame, fg_color="#FFFFFF", border_width=1, border_color="#E3E6EA", corner_radius=12)
				row.grid(row=row_idx, column=0, sticky="ew", pady=(0, 10))
				row.grid_columnconfigure(0, weight=1)
				row.grid_columnconfigure(1, weight=0)

				info_text = (
					f"Res: {doc.get('resolution_number', '-')} | "
					f"Norma: {doc.get('selected_norm', '-')} | "
					f"Producto: {doc.get('evaluated_product', '-')} | "
					f"Ejecutivo: {doc.get('executive_name', '-')} | "
					f"Fecha: {doc.get('generated_at', '-')}"
				)
				ctk.CTkLabel(
					row,
					text=info_text,
					font=FONTS["small"],
					text_color=STYLE["texto_oscuro"],
					justify="left",
					anchor="w",
				).grid(row=0, column=0, padx=12, pady=10, sticky="ew")

				output_path = str(doc.get("output_path", "")).strip()
				if output_path and Path(output_path).exists():
					ctk.CTkButton(
						row,
						text="Abrir",
						width=80,
						fg_color=STYLE["primario"],
						text_color=STYLE["texto_oscuro"],
						hover_color="#D8C220",
						command=lambda path=output_path: self._open_pdf_from_history(path),
					).grid(row=0, column=1, padx=12, pady=10)
				row_idx += 1

		norm_combo.configure(command=lambda _value: _render_rows())
		_render_rows()

		ctk.CTkButton(
			wrapper,
			text="Cerrar",
			fg_color=STYLE["fondo"],
			text_color=STYLE["texto_oscuro"],
			hover_color="#E9ECEF",
			command=history_window.destroy,
		).grid(row=3, column=0, padx=18, pady=(0, 12), sticky="ew")

	def _open_pdf_from_history(self, path: str) -> None:
		pdf_path = Path(path)
		if not pdf_path.exists():
			messagebox.showerror("Historial", f"El archivo no existe:\n{path}", parent=self)
			return
		if hasattr(os, "startfile"):
			try:
				os.startfile(pdf_path)
			except OSError as error:
				messagebox.showerror("Historial", f"No se pudo abrir el archivo:\n{error}", parent=self)

	def _rebuild_pager(self, total: int) -> None:
		if self.pager_frame is None:
			return

		for child in self.pager_frame.winfo_children():
			child.destroy()

		if total == 0:
			return

		total_pages = max(1, (total + self.PAGE_SIZE - 1) // self.PAGE_SIZE)
		if total_pages <= 1:
			return

		inner = ctk.CTkFrame(self.pager_frame, fg_color="transparent")
		inner.pack(anchor="center")

		ctk.CTkButton(
			inner,
			text="← Anterior",
			width=110,
			height=32,
			fg_color=STYLE["fondo"],
			text_color=STYLE["texto_oscuro"],
			hover_color="#E9ECEF",
			state="normal" if self.current_page > 0 else "disabled",
			command=lambda: self._go_page(-1),
		).pack(side="left", padx=(0, 8))

		ctk.CTkLabel(
			inner,
			text=f"Pagina {self.current_page + 1} de {total_pages}  —  {total} clientes",
			font=FONTS["small"],
			text_color="#6D7480",
		).pack(side="left", padx=12)

		ctk.CTkButton(
			inner,
			text="Siguiente →",
			width=110,
			height=32,
			fg_color=STYLE["fondo"],
			text_color=STYLE["texto_oscuro"],
			hover_color="#E9ECEF",
			state="normal" if self.current_page < total_pages - 1 else "disabled",
			command=lambda: self._go_page(1),
		).pack(side="left", padx=(8, 0))

	def _go_page(self, delta: int) -> None:
		total = len(self._filtered_clients)
		total_pages = max(1, (total + self.PAGE_SIZE - 1) // self.PAGE_SIZE)
		new_page = max(0, min(total_pages - 1, self.current_page + delta))
		if new_page == self.current_page:
			return
		self.current_page = new_page
		self.refresh()

	def _handle_change(self) -> None:
		self.refresh()
		self.on_change()

