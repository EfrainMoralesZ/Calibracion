from __future__ import annotations

import re
import unicodedata
from datetime import datetime
from pathlib import Path
from statistics import mean
from tkinter import filedialog, messagebox, ttk
import tkinter as tk

import customtkinter as ctk
from PIL import Image

from runtime_paths import resource_path, writable_path


class TrimestralView(ctk.CTkFrame):
	def __init__(self, master, controller, style: dict, fonts: dict, can_edit: bool) -> None:
		super().__init__(master, fg_color=style["fondo"])
		self.controller = controller
		self.style = style
		self.fonts = fonts
		self.can_edit = can_edit

		self.selected_score_id: str | None = None
		self.row_cache: dict[str, dict] = {}
		self._history_signature: tuple | None = None
		self._cards_signature: tuple | None = None
		self.cards_page_size = 9
		self.current_cards_page = 0
		self.cards_medal_filter_var = ctk.StringVar(value="Todas")

		# Para capturas múltiples
		self.temp_scores = []  # Lista temporal de capturas antes de guardar
		self.temp_table = None  # Widget de tabla temporal
		self.add_button = None  # Botón para agregar a la tabla temporal
		self.save_all_button = None  # Botón para guardar todas las capturas
		self.inspector_var = ctk.StringVar()
		self.norm_var = ctk.StringVar()
		self.year_var = ctk.StringVar(value=str(datetime.now().year))
		self.quarter_var = ctk.StringVar(value="T1")
		self.score_var = ctk.StringVar()

		self.tree: ttk.Treeview | None = None
		self.notes_box: ctk.CTkTextbox | None = None
		self.score_entry: ctk.CTkEntry | None = None
		self.norm_selector: ctk.CTkComboBox | None = None
		self.quarter_selector_capture: ctk.CTkComboBox | None = None
		self.cards_frame: ctk.CTkScrollableFrame | None = None
		self.capture_dialog: ctk.CTkToplevel | None = None
		self.capture_title_label: ctk.CTkLabel | None = None
		self.capture_inspector_value_label: ctk.CTkLabel | None = None
		self.capture_delete_button: ctk.CTkButton | None = None
		self.capture_history_box: ctk.CTkTextbox | None = None
		self.cards_pager_frame: ctk.CTkFrame | None = None
		self.history_dashboard_frame: ctk.CTkScrollableFrame | None = None
		self.history_dashboard_summary_label: ctk.CTkLabel | None = None
		self.cards_medal_filter_combo: ctk.CTkComboBox | None = None
		self.medal_images = self._load_medal_images(size=(64, 64))
		self.medal_images_small = self._load_medal_images(size=(24, 24))
		self.alert_images = self._load_alert_images()
		self.tabview: ctk.CTkTabview | None = None
		self.capture_tab_name = "Captura trimestral"
		self.history_tab_name = "Historial trimestral"

		self.grid_columnconfigure(0, weight=1)
		self.grid_rowconfigure(1, weight=1)
		self._build_ui()

	def _build_ui(self) -> None:
		header = ctk.CTkFrame(self, fg_color="transparent")
		header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
		header.grid_columnconfigure(0, weight=1)

		ctk.CTkLabel(
			header,
			text="Trimestral",
			font=self.fonts["subtitle"],
			text_color=self.style["texto_oscuro"],
		).grid(row=0, column=0, sticky="w")
		subtitle_text = (
			"Asigna calificaciones trimestrales por norma, consulta demanda de normas y analiza frecuencias de uso por pestana."
			if self.can_edit
			else "Consulta tus calificaciones trimestrales por norma en cards y analiza demanda de uso en historial."
		)
		ctk.CTkLabel(
			header,
			text=subtitle_text,
			font=self.fonts["small"],
			text_color="#6D7480",
		).grid(row=1, column=0, sticky="w", pady=(4, 0))

		if self.can_edit:
			self.tabview = ctk.CTkTabview(
				self,
				fg_color=self.style["surface"],
				segmented_button_selected_color=self.style["secundario"],
				segmented_button_selected_hover_color="#1D1D1D",
			)
			self.tabview.grid(row=1, column=0, sticky="nsew")
			self.tabview.add(self.capture_tab_name)
			self.tabview.add(self.history_tab_name)

			self._build_capture_tab(self.tabview.tab(self.capture_tab_name))
			self._build_history_tab(self.tabview.tab(self.history_tab_name))
		else:
			executive_panel = ctk.CTkFrame(self, fg_color="transparent")
			executive_panel.grid(row=1, column=0, sticky="nsew")
			executive_panel.grid_columnconfigure(0, weight=1)
			executive_panel.grid_rowconfigure(0, weight=1)
			self._build_capture_tab(executive_panel)

	def _load_medal_images(self, size: tuple[int, int] = (46, 46)) -> dict[str, ctk.CTkImage | None]:
		files = {
			"ORO": "medalla_oro.png",
			"PLATA": "medalla_plata.png",
			"BRONCE": "medalla_bronce.png",
		}
		images: dict[str, ctk.CTkImage | None] = {}
		for key, file_name in files.items():
			image_path: Path | None = None
			for candidate in [resource_path("img", file_name), writable_path("img", file_name)]:
				if candidate.exists():
					image_path = candidate
					break
			if image_path is None:
				images[key] = None
				continue

			try:
				source = Image.open(image_path)
			except OSError:
				images[key] = None
				continue

			images[key] = ctk.CTkImage(light_image=source, dark_image=source, size=size)
		return images

	def _load_alert_images(self) -> dict[str, ctk.CTkImage | None]:
		image_path: Path | None = None
		for candidate in [resource_path("img", "alerta.png"), writable_path("img", "alerta.png")]:
			if candidate.exists():
				image_path = candidate
				break
		if image_path is None:
			return {"small": None, "large": None}

		try:
			source = Image.open(image_path)
		except OSError:
			return {"small": None, "large": None}

		return {
			"small": ctk.CTkImage(light_image=source, dark_image=source, size=(20, 20)),
			"large": ctk.CTkImage(light_image=source, dark_image=source, size=(26, 26)),
		}

	def _build_capture_tab(self, tab) -> None:
		tab.grid_columnconfigure(0, weight=1)
		tab.grid_rowconfigure(0, weight=1)

		cards_wrapper = ctk.CTkFrame(tab, fg_color=self.style["surface"], corner_radius=22)
		cards_wrapper.grid(row=0, column=0, sticky="nsew", pady=12)
		cards_wrapper.grid_columnconfigure(0, weight=1)
		cards_row_for_scroll = 3 if self.can_edit else 2
		cards_wrapper.grid_rowconfigure(cards_row_for_scroll, weight=1)

		cards_title = "Inspectores y normas acreditadas" if self.can_edit else "Tu card trimestral"
		cards_hint = (
			"Asigna calificaciones trimestrales por norma, consulta demanda de uso de normas y confirma capturas."
			if self.can_edit
			else "Aqui aparece tu card trimestral con acceso a estadisticas de demanda de normas y confirmacion de visto."
		)
		ctk.CTkLabel(
			cards_wrapper,
			text=cards_title,
			font=self.fonts["label_bold"],
			text_color=self.style["texto_oscuro"],
		).grid(row=0, column=0, padx=18, pady=(14, 4), sticky="w")
		ctk.CTkLabel(
			cards_wrapper,
			text=cards_hint,
			font=self.fonts["small"],
			text_color="#6D7480",
		).grid(row=1, column=0, padx=18, pady=(0, 8), sticky="w")

		if self.can_edit:
			filter_row = ctk.CTkFrame(cards_wrapper, fg_color="transparent")
			filter_row.grid(row=2, column=0, padx=18, pady=(0, 8), sticky="ew")
			filter_row.grid_columnconfigure(3, weight=1)
			ctk.CTkLabel(
				filter_row,
				text="Filtro por medallas",
				font=self.fonts["small_bold"],
				text_color="#6D7480",
			).grid(row=0, column=0, padx=(0, 8), sticky="w")
			self.cards_medal_filter_combo = ctk.CTkComboBox(
				filter_row,
				variable=self.cards_medal_filter_var,
				values=["Todas", "Oro", "Plata", "Bronce", "Sin medalla"],
				width=170,
				height=34,
				fg_color="#FFFFFF",
				border_color="#D5D8DC",
				button_color=self.style["primario"],
				dropdown_hover_color=self.style["primario"],
				command=lambda _value: self._render_inspector_cards(),
			)
			self.cards_medal_filter_combo.grid(row=0, column=1, sticky="w")
			ctk.CTkLabel(
				filter_row,
				text="Rangos: 80 Bronce | 90 Plata | 100 Oro",
				font=self.fonts["small"],
				text_color="#6D7480",
			).grid(row=0, column=2, padx=(12, 0), sticky="w")

		self.cards_frame = ctk.CTkScrollableFrame(cards_wrapper, fg_color="transparent")
		self.cards_frame.grid(row=cards_row_for_scroll, column=0, padx=10, pady=(0, 10), sticky="nsew")
		for card_col in range(3):
			self.cards_frame.grid_columnconfigure(card_col, weight=1, uniform="trimestral_cards")

		self.cards_pager_frame = ctk.CTkFrame(cards_wrapper, fg_color="transparent")
		self.cards_pager_frame.grid(row=cards_row_for_scroll + 1, column=0, padx=18, pady=(0, 14), sticky="ew")

	def _build_history_tab(self, tab) -> None:
		tab.grid_columnconfigure(0, weight=1)
		tab.grid_rowconfigure(0, weight=1)

		table_panel = ctk.CTkFrame(tab, fg_color=self.style["surface"], corner_radius=22)
		table_panel.grid(row=0, column=0, sticky="nsew")
		table_panel.grid_columnconfigure(0, weight=1)
		table_panel.grid_rowconfigure(2, weight=1)

		dashboard_panel = ctk.CTkFrame(table_panel, fg_color="#FFFFFF", corner_radius=16, border_width=1, border_color="#E3E6EA")
		dashboard_panel.grid(row=2, column=0, padx=18, pady=(0, 18), sticky="nsew")
		dashboard_panel.grid_columnconfigure(0, weight=1)
		dashboard_panel.grid_rowconfigure(3, weight=1)

		dashboard_title = "Demanda de normas por ejecutivo" if self.can_edit else "Demanda de tus normas acreditadas"
		ctk.CTkLabel(
			dashboard_panel,
			text=dashboard_title,
			font=self.fonts["label_bold"],
			text_color=self.style["texto_oscuro"],
		).grid(row=0, column=0, padx=12, pady=(10, 2), sticky="w")

		default_summary = "Sin reportes de normas aplicadas para mostrar."
		self.history_dashboard_summary_label = ctk.CTkLabel(
			dashboard_panel,
			text=default_summary,
			font=self.fonts["small"],
			text_color="#6D7480",
			justify="left",
		)
		self.history_dashboard_summary_label.grid(row=1, column=0, padx=12, pady=(0, 6), sticky="w")

		legend_row = ctk.CTkFrame(dashboard_panel, fg_color="transparent")
		legend_row.grid(row=2, column=0, padx=12, pady=(0, 8), sticky="ew")
		legend_row.grid_columnconfigure(2, weight=1)
		ctk.CTkLabel(
			legend_row,
			text="Alta demanda (más usos)",
			font=self.fonts["small"],
			text_color="#282828",
			fg_color=self.style["primario"],
			corner_radius=999,
			height=26,
		).grid(row=0, column=0, padx=(0, 8), sticky="w")
		ctk.CTkLabel(
			legend_row,
			text="Baja demanda (menos usos)",
			font=self.fonts["small"],
			text_color="#6D7480",
			fg_color="#E8EAEB",
			corner_radius=999,
			height=26,
		).grid(row=0, column=1, sticky="w")
		ctk.CTkButton(
			legend_row,
			text="Exportar PDF",
			width=128,
			height=34,
			fg_color=self.style["primario"],
			text_color=self.style["texto_oscuro"],
			hover_color="#D9C31E",
			command=self._export_history_dashboard_pdf,
		).grid(row=0, column=3, sticky="e")

		self.history_dashboard_frame = ctk.CTkScrollableFrame(
			dashboard_panel,
			fg_color="#FFFFFF",
			corner_radius=12,
		)
		self.history_dashboard_frame.grid(row=3, column=0, padx=12, pady=(0, 12), sticky="nsew")
		self.history_dashboard_frame.grid_columnconfigure(0, weight=1)

		self.tree = None
		self.row_cache = {}
		self._history_signature = None

	def _field(self, parent, index: int, label: str, widget) -> None:
		base_row = (index - 1) * 2 + 1
		ctk.CTkLabel(parent, text=label, font=self.fonts["label"], text_color=self.style["texto_oscuro"]).grid(
			row=base_row,
			column=0,
			padx=18,
			pady=(0 if index == 1 else 8, 6),
			sticky="w",
		)
		widget.grid(row=base_row + 1, column=0, padx=18, sticky="ew")

	def _build_capture_dialog(self) -> None:
		if not self.can_edit:
			return

		self._close_capture_dialog(reset_state=False)

		self.capture_dialog = ctk.CTkToplevel(self)
		self.capture_dialog.title("Captura trimestral")
		self.capture_dialog.geometry("1100x700")
		self.capture_dialog.minsize(900, 600)
		self.capture_dialog.configure(fg_color=self.style["fondo"])
		self.capture_dialog.transient(self.winfo_toplevel())
		self.capture_dialog.grab_set()
		self.capture_dialog.protocol("WM_DELETE_WINDOW", self._close_capture_dialog)

		# Outer wrapper: scrollable content on top, fixed action bar at bottom
		outer = ctk.CTkFrame(self.capture_dialog, fg_color=self.style["fondo"])
		outer.pack(fill="both", expand=True, padx=18, pady=18)
		outer.grid_rowconfigure(0, weight=1)
		outer.grid_rowconfigure(1, weight=0)
		outer.grid_columnconfigure(0, weight=1)

		# Scrollable form area en dos columnas
		form_scroll = ctk.CTkScrollableFrame(outer, fg_color=self.style["surface"], corner_radius=18)
		form_scroll.grid(row=0, column=0, sticky="nsew", pady=(0, 10))
		form_scroll.grid_columnconfigure(0, weight=1, minsize=350)
		form_scroll.grid_columnconfigure(1, weight=1, minsize=350)
		form_panel_left = ctk.CTkFrame(form_scroll, fg_color="transparent")
		form_panel_left.grid(row=0, column=0, sticky="nsew", padx=(18, 9), pady=10)
		form_panel_right = ctk.CTkFrame(form_scroll, fg_color="transparent")
		form_panel_right.grid(row=0, column=1, sticky="nsew", padx=(9, 18), pady=10)

		# Título
		self.capture_title_label = ctk.CTkLabel(
			form_panel_left,
			text="Nueva captura trimestral",
			font=self.fonts["label_bold"],
			text_color=self.style["texto_oscuro"],
		)
		self.capture_title_label.grid(row=0, column=0, padx=0, pady=(0, 8), sticky="w")

		# Columna izquierda: datos principales
		self.capture_inspector_value_label = ctk.CTkLabel(
			form_panel_left,
			textvariable=self.inspector_var,
			font=self.fonts["label"],
			text_color=self.style["texto_oscuro"],
			fg_color="#FFFFFF",
			corner_radius=10,
			height=38,
			anchor="w",
		)
		self._field(form_panel_left, 1, "Ejecutivo Tecnico", self.capture_inspector_value_label)

		self.norm_selector = ctk.CTkComboBox(
			form_panel_left,
			variable=self.norm_var,
			values=["SIN_NORMA"],
			height=38,
			fg_color="#FFFFFF",
			border_color="#D5D8DC",
			button_color=self.style["primario"],
			dropdown_hover_color=self.style["primario"],
			command=lambda _value: self._refresh_capture_history_preview(),
		)
		self._field(form_panel_left, 2, "Norma", self.norm_selector)

		# Menú de año (últimos 5 años y el actual)
		from datetime import datetime
		current_year = datetime.now().year
		year_options = [str(y) for y in range(current_year, current_year - 6, -1)]
		self._field(form_panel_left, 3, "Año", ctk.CTkComboBox(
			form_panel_left,
			variable=self.year_var,
			values=year_options,
			height=38,
			fg_color="#FFFFFF",
			border_color="#D5D8DC",
			button_color=self.style["primario"],
			dropdown_hover_color=self.style["primario"],
			command=lambda _v: self._on_capture_period_change(),
		))
		self.quarter_selector_capture = ctk.CTkComboBox(
			form_panel_left,
			variable=self.quarter_var,
			values=["T1", "T2", "T3", "T4"],
			height=38,
			fg_color="#FFFFFF",
			border_color="#D5D8DC",
			button_color=self.style["primario"],
			dropdown_hover_color=self.style["primario"],
			command=lambda _v: self._on_capture_period_change(),
		)
		self._field(form_panel_left, 4, "Trimestre", self.quarter_selector_capture)

		# Columna derecha: calificación y notas

		# Calificación con aclaración de porcentaje
		self.score_entry = ctk.CTkEntry(form_panel_right, textvariable=self.score_var, height=38, border_color="#D5D8DC")
		self._field(form_panel_right, 1, "Calificación (%) [1-100]", self.score_entry)

		ctk.CTkLabel(
			form_panel_right,
			text="📝 Notas / Observaciones",
			font=self.fonts["label_bold"],
			text_color=self.style["texto_oscuro"],
		).grid(row=3, column=0, padx=0, pady=(10, 6), sticky="w")
		self.notes_box = ctk.CTkTextbox(
			form_panel_right,
			height=140,
			corner_radius=12,
			border_width=2,
			border_color="#D5D8DC",
			fg_color="#FFFFFF",
			text_color=self.style["texto_oscuro"],
			font=self.fonts["label"],
		)
		self.notes_box.grid(row=4, column=0, padx=0, pady=(0, 4), sticky="ew")
		ctk.CTkLabel(
			form_panel_right,
			text="Estas notas seran visibles para el ejecutivo evaluado.",
			font=self.fonts["small"],
			text_color="#6D7480",
		).grid(row=5, column=0, padx=0, pady=(0, 4), sticky="w")

		# Tabla temporal para capturas múltiples (abajo, ocupa ambas columnas)
		self.temp_table = ttk.Treeview(form_scroll, columns=("norma", "año", "trimestre", "calificacion", "estado"), show="headings", height=4)
		for col, width in zip(["norma", "año", "trimestre", "calificacion", "estado"], [100, 60, 80, 100, 120]):
			self.temp_table.heading(col, text=col.capitalize())
			self.temp_table.column(col, width=width)
		self.temp_table.grid(row=1, column=0, columnspan=2, padx=18, pady=(0, 8), sticky="ew")

		# Botones debajo de la tabla
		button_row = ctk.CTkFrame(form_scroll, fg_color="transparent")
		button_row.grid(row=2, column=0, columnspan=2, padx=18, pady=(0, 8), sticky="ew")
		button_row.grid_columnconfigure(0, weight=1)
		button_row.grid_columnconfigure(1, weight=1)
		self.add_button = ctk.CTkButton(
			button_row,
			text="Agregar a lista",
			fg_color=self.style["primario"],
			text_color=self.style["texto_oscuro"],
			hover_color="#D8C220",
			command=self._add_temp_score,
		)
		self.add_button.grid(row=0, column=0, padx=(0, 8), sticky="ew")
		self.save_all_button = ctk.CTkButton(
			button_row,
			text="Guardar todo",
			fg_color=self.style["secundario"],
			hover_color="#1D1D1D",
			command=self._save_all_temp_scores,
		)
		self.save_all_button.grid(row=0, column=1, padx=(8, 0), sticky="ew")
		# Botón eliminar (opcional: elimina la fila seleccionada de la tabla)
		self.delete_button = ctk.CTkButton(
			button_row,
			text="Eliminar",
			fg_color=self.style["peligro"],
			hover_color="#B43C31",
			command=self._delete_selected_temp_score,
		)
		self.delete_button.grid(row=0, column=2, padx=(8, 0), sticky="ew")

	def _delete_selected_temp_score(self):
		selected = self.temp_table.selection()
		if not selected:
			messagebox.showinfo("Trimestral", "Selecciona una fila para eliminar.")
			return
		idx = self.temp_table.index(selected[0])
		del self.temp_scores[idx]
		self._refresh_temp_table()

	       # (El formulario y la tabla ya están correctamente definidos arriba, no se requiere duplicar ni redefinir aquí)
		# (Botones de acción ya están definidos bajo la tabla temporal)

	def _add_temp_score(self):
		# Obtiene los datos del formulario
		norm = self.norm_var.get()
		year = self.year_var.get()
		quarter = self.quarter_var.get()
		score = self.score_var.get()
		inspector = self.inspector_var.get().strip()
		notes = self.notes_box.get("1.0", "end").strip() if self.notes_box is not None else ""
		# Validación básica
		if not norm or not year or not quarter or not score or not inspector:
			messagebox.showerror("Trimestral", "Completa todos los campos antes de agregar.")
			return
		# Validar porcentaje
		try:
			score_val = float(score)
			if not (1 <= score_val <= 100):
				raise ValueError
		except Exception:
			messagebox.showerror("Trimestral", "La calificación debe ser un número entre 1 y 100.")
			return
		# No permitir duplicados de norma+trimestre+año ni si ya está calificado/enviado
		inspector_scores = self.controller.list_trimestral_scores(inspector_name=inspector, include_unsent=True)
		for item in inspector_scores:
			if item.get("norm") == norm and str(item.get("year")) == year and str(item.get("quarter")) == quarter:
				if item.get("sent_at"):
					messagebox.showwarning("Trimestral", f"Ya existe una calificación enviada para {norm} {quarter} {year}. No se puede modificar.")
					return
				else:
					messagebox.showwarning("Trimestral", f"Ya existe una calificación pendiente para {norm} {quarter} {year}.")
					return
		for item in self.temp_scores:
			if item["norm"] == norm and item["year"] == year and item["quarter"] == quarter:
				messagebox.showwarning("Trimestral", "Ya capturaste esta norma para ese trimestre/año.")
				return
		# Agrega a la lista temporal
		self.temp_scores.append({
			"inspector": inspector,
			"norm": norm,
			"year": year,
			"quarter": quarter,
			"score": score,
			"notes": notes,
			"estado": "Pendiente"
		})
		self._refresh_temp_table()
		self.score_var.set("")
		if self.notes_box is not None:
			try:
				if self.notes_box.winfo_exists():
					self.notes_box.delete("1.0", "end")
			except tk.TclError:
				pass

	def _refresh_temp_table(self):
		# Limpia la tabla y la repuebla
		for row in self.temp_table.get_children():
			self.temp_table.delete(row)
		for item in self.temp_scores:
			self.temp_table.insert("", "end", values=(item["norm"], item["year"], item["quarter"], item["score"], item["estado"]))

	def _save_all_temp_scores(self):
		if not self.temp_scores:
			messagebox.showinfo("Trimestral", "No hay capturas pendientes para guardar.")
			return
		errores = 0
		for item in self.temp_scores:
			try:
				self.controller.save_trimestral_score(item, None)
				item["estado"] = "Asignada"
			except Exception as e:
				errores += 1
				item["estado"] = f"Error: {e}"
		self._refresh_temp_table()
		if errores == 0:
			messagebox.showinfo("Trimestral", "Todas las calificaciones fueron guardadas correctamente.")
			self.temp_scores.clear()
			self._refresh_temp_table()
		else:
			messagebox.showwarning("Trimestral", f"{errores} calificaciones no se pudieron guardar.")

		self.refresh()

	# Bloqueo de edición tras envío (en la lógica de edición y guardado real ya existe, pero aquí puedes agregar validación visual si lo deseas)

		self._sync_capture_norm_selector()
		self._refresh_capture_history_preview()
		self._update_capture_title()
		self._sync_capture_delete_state()
		if self.score_entry is not None:
			self.capture_dialog.after(80, self.score_entry.focus_set)

	def _capture_norm_values(self, inspector_name: str) -> list[str]:
		clean_name = str(inspector_name or "").strip()
		norm_values: list[str] = []

		if clean_name:
			norm_values.extend(self.controller.get_accredited_norms(clean_name))
			for item in self.controller.list_trimestral_scores(inspector_name=clean_name):
				norm_values.append(self._norm_key(item))

		if not norm_values:
			norm_values.extend(item["token"] for item in self.controller.get_catalog_norms())

		unique: list[str] = []
		seen: set[str] = set()
		for value in norm_values:
			token = self._norm_key(value)
			if token in seen:
				continue
			unique.append(token)
			seen.add(token)

		# Excluir normas ya calificadas para el trimestre/año seleccionado
		selected_year = self.year_var.get().strip()
		selected_quarter = self.quarter_var.get().strip().upper()
		if clean_name and selected_year and selected_quarter:
			existing_scores = self.controller.list_trimestral_scores(
				inspector_name=clean_name, include_unsent=True,
			)
			graded_norms: set[str] = set()
			for item in existing_scores:
				if str(item.get("year", "")).strip() == selected_year and str(item.get("quarter", "")).strip().upper() == selected_quarter:
					graded_norms.add(self._norm_key(item))
			# Tambien excluir las ya agregadas en la tabla temporal
			for item in self.temp_scores:
				if str(item.get("year", "")).strip() == selected_year and str(item.get("quarter", "")).strip().upper() == selected_quarter:
					graded_norms.add(self._norm_key(item))
			unique = [token for token in unique if token not in graded_norms]

		return unique or ["SIN_NORMA"]

	def _sync_capture_norm_selector(self) -> None:
		if self.norm_selector is None:
			return
		try:
			if not self.norm_selector.winfo_exists():
				return
		except tk.TclError:
			return

		norm_values = self._capture_norm_values(self.inspector_var.get())
		self.norm_selector.configure(values=norm_values)
		current_norm = self._norm_key(self.norm_var.get())
		if current_norm not in norm_values:
			self.norm_var.set(norm_values[0])
		else:
			self.norm_var.set(current_norm)

	def _on_capture_period_change(self) -> None:
		"""Muestra alerta si el trimestre seleccionado ya fue enviado y re-sincroniza normas."""
		inspector_name = self.inspector_var.get().strip()
		if inspector_name:
			existing = self.controller.list_trimestral_scores(
				inspector_name=inspector_name, include_unsent=True,
			)
			selected_year = self.year_var.get().strip()
			selected_quarter = self.quarter_var.get().strip().upper()
			# Verificar si el trimestre seleccionado ya fue enviado
			if selected_quarter in ("T1", "T2", "T3", "T4") and selected_year:
				quarter_scores = [
					item for item in existing
					if str(item.get("year", "")).strip() == selected_year
					and str(item.get("quarter", "")).strip().upper() == selected_quarter
				]
				if quarter_scores and all(str(item.get("sent_at", "")).strip() for item in quarter_scores):
					messagebox.showinfo(
						"Trimestral",
						f"El trimestre {selected_quarter} {selected_year} ya fue calificado y enviado.",
						parent=self,
					)
		self._sync_capture_norm_selector()
		self._refresh_capture_history_preview()

	def _refresh_capture_history_preview(self) -> None:
		if self.capture_history_box is None:
			return
		try:
			if not self.capture_history_box.winfo_exists():
				return
		except tk.TclError:
			return

		inspector_name = self.inspector_var.get().strip()
		norm_value = self._norm_key(self.norm_var.get())
		if not inspector_name:
			lines = ["Selecciona un inspector para mostrar historial."]
		else:
			scores = self.controller.list_trimestral_scores(inspector_name=inspector_name, norm=norm_value)
			if not scores:
				lines = ["Sin calificaciones previas para esta norma."]
			else:
				lines = [f"Norma: {norm_value}", ""]
				for item in scores[:8]:
					score_value = self._coerce_score(item.get("score"))
					score_text = f"{score_value:.1f}%" if score_value is not None else "--"
					medal = self._score_medal(item)
					medal_text = medal["title"] if medal["key"] else "Sin medalla"
					period = f"{item.get('quarter', '--')} {item.get('year', '--')}"
					updated = str(item.get("updated_at", "--"))
					sent_at = str(item.get("sent_at", "")).strip()
					confirmed_at = str(item.get("confirmed_at", "")).strip()
					if not sent_at:
						state = "Asignada (sin enviar)"
					elif confirmed_at and confirmed_at.lower() != "none":
						state = "Confirmado"
					else:
						state = "Pendiente"
					lines.append(f"- {period} | {score_text} | {medal_text} | {updated} | {state}")

		self.capture_history_box.configure(state="normal")
		self.capture_history_box.delete("1.0", "end")
		self.capture_history_box.insert("1.0", "\n".join(lines))
		self.capture_history_box.configure(state="disabled")

	def _open_capture_dialog(self, inspector_name: str | None = None, score: dict | None = None) -> None:
		if not self.can_edit:
			return

		# Block editing if the score was already sent
		if score is not None and str(score.get("sent_at", "")).strip():
			messagebox.showinfo(
				"Trimestral",
				"Esta calificacion ya fue enviada y no puede ser modificada.",
				parent=self,
			)
			return

		if score is None:
			self.selected_score_id = None
			self.inspector_var.set(str(inspector_name or "").strip())
			self.year_var.set(str(datetime.now().year))
			self.quarter_var.set("T1")
			self.score_var.set("")
			available_norms = self._capture_norm_values(self.inspector_var.get())
			self.norm_var.set(available_norms[0] if available_norms else "SIN_NORMA")
			notes = ""
		else:
			self.selected_score_id = str(score.get("id", "")).strip() or None
			self.inspector_var.set(str(score.get("inspector", inspector_name or "")).strip())
			self.norm_var.set(self._norm_key(score))
			self.year_var.set(str(score.get("year", datetime.now().year)))
			self.quarter_var.set(str(score.get("quarter", "T1")) or "T1")
			raw_score = self._coerce_score(score.get("score"))
			self.score_var.set("" if raw_score is None else f"{raw_score:.1f}")
			notes = str(score.get("notes", ""))

		if not self.inspector_var.get().strip():
			messagebox.showerror("Trimestral", "No se detecto un ejecutivo tecnico para la captura.", parent=self)
			return

		self._build_capture_dialog()
		if self.notes_box is not None:
			self.notes_box.delete("1.0", "end")
			if notes:
				self.notes_box.insert("1.0", notes)
		self._sync_capture_norm_selector()
		self._refresh_capture_history_preview()
		self._update_capture_title()
		self._sync_capture_delete_state()

	def _close_capture_dialog(self, reset_state: bool = True) -> None:
		dialog = self.capture_dialog
		if reset_state:
			self.clear_form(full_reset=True)

		self.capture_dialog = None
		self.capture_title_label = None
		self.capture_inspector_value_label = None
		self.capture_delete_button = None
		self.capture_history_box = None
		self.norm_selector = None
		self.quarter_selector_capture = None
		self.notes_box = None
		self.score_entry = None

		if dialog is None:
			return

		try:
			if dialog.winfo_exists():
				dialog.grab_release()
				dialog.destroy()
		except tk.TclError:
			return

	def _capture_message_parent(self):
		if self.capture_dialog is not None:
			try:
				if self.capture_dialog.winfo_exists():
					return self.capture_dialog
			except tk.TclError:
				pass
		return self

	def _update_capture_title(self) -> None:
		if self.capture_title_label is None:
			return

		base_text = "Editar captura trimestral" if self.selected_score_id else "Nueva captura trimestral"
		inspector_name = self.inspector_var.get().strip()
		if inspector_name:
			base_text = f"{base_text} | {inspector_name}"

		self.capture_title_label.configure(text=base_text)
		if self.capture_dialog is not None:
			try:
				if self.capture_dialog.winfo_exists():
					self.capture_dialog.title(base_text)
			except tk.TclError:
				return

	def _sync_capture_delete_state(self) -> None:
		if self.capture_delete_button is None:
			return
		self.capture_delete_button.configure(state="normal" if self.selected_score_id else "disabled")

	def _go_cards_page(self, delta: int, total_cards: int) -> None:
		total_pages = max(1, (total_cards + self.cards_page_size - 1) // self.cards_page_size)
		new_page = max(0, min(total_pages - 1, self.current_cards_page + delta))
		if new_page == self.current_cards_page:
			return
		self.current_cards_page = new_page
		self._render_inspector_cards()
		try:
			if self.cards_frame is not None:
				self.cards_frame._parent_canvas.yview_moveto(0)
		except Exception:
			pass

	def _rebuild_cards_pager(self, total_cards: int) -> None:
		if self.cards_pager_frame is None:
			return

		for child in self.cards_pager_frame.winfo_children():
			child.destroy()

		if total_cards == 0:
			return

		total_pages = max(1, (total_cards + self.cards_page_size - 1) // self.cards_page_size)
		if total_pages <= 1:
			return

		inner = ctk.CTkFrame(self.cards_pager_frame, fg_color="transparent")
		inner.pack(anchor="center")

		ctk.CTkButton(
			inner,
			text="\u2190  Anterior",
			width=110,
			height=34,
			fg_color=self.style["fondo"],
			text_color=self.style["texto_oscuro"],
			hover_color="#E9ECEF",
			state="normal" if self.current_cards_page > 0 else "disabled",
			command=lambda: self._go_cards_page(-1, total_cards),
		).pack(side="left", padx=(0, 8))

		ctk.CTkLabel(
			inner,
			text=f"Pagina {self.current_cards_page + 1} de {total_pages}  \u2014  {total_cards} registros",
			font=self.fonts["small"],
			text_color="#6D7480",
		).pack(side="left", padx=12)

		ctk.CTkButton(
			inner,
			text="Siguiente  \u2192",
			width=110,
			height=34,
			fg_color=self.style["fondo"],
			text_color=self.style["texto_oscuro"],
			hover_color="#E9ECEF",
			state="normal" if self.current_cards_page < total_pages - 1 else "disabled",
			command=lambda: self._go_cards_page(1, total_cards),
		).pack(side="left", padx=(8, 0))

	def refresh(self) -> None:
		self._render_inspector_cards()
		self._refresh_history_table()
		self._render_history_dashboard()
		self._sync_capture_norm_selector()
		self._refresh_capture_history_preview()
		self._update_capture_title()
		self._sync_capture_delete_state()

	def _build_personalized_messages(self) -> tuple[str, str]:
		current_name = str((self.controller.current_user or {}).get("name", "")).strip() or "Equipo"
		hour = datetime.now().hour
		if hour < 12:
			saludo = "Buenos dias"
		elif hour < 19:
			saludo = "Buenas tardes"
		else:
			saludo = "Buenas noches"

		greeting_text = f"{saludo}, {current_name}."
		if not self.controller.current_user or not self.controller.is_executive_role(self.controller.current_user):
			return greeting_text, "Monitorea resultados trimestrales y reconoce el desempeno destacado del equipo."

		scores = self.controller.list_trimestral_scores(inspector_name=current_name, include_unsent=True)
		totals = self._accumulated_medals(scores)
		if totals["ORO"] > 0:
			message = (
				f"{current_name}, llevas {totals['ORO']} medalla(s) Oro. "
				"¡Excelente desempeño!"
			)
		elif totals["PLATA"] > 0:
			message = (
				f"{current_name}, llevas {totals['PLATA']} medalla(s) Plata. "
				"¡Óptimo desempeño!"
			)
		elif totals["BRONCE"] > 0:
			message = (
				f"{current_name}, ya obtuviste {totals['BRONCE']} medalla(s) Bronce. "
				"Vas bien: favor de reforzar para subir al siguiente nivel."
			)
		else:
			message = f"{current_name}, cada trimestre cuenta: apunta a 80 para tu primera medalla Bronce."

		return greeting_text, message

	def _score_medal(self, scores_or_value) -> dict[str, str]:
		# Ahora espera una lista de scores para calcular el promedio general
		if isinstance(scores_or_value, list):
			# Lista de dicts (todas las normas del inspector)
			califs = []
			for item in scores_or_value:
				try:
					val = float(item.get("score", 0))
					califs.append(val)
				except Exception:
					pass
			if not califs:
				key = ""
			else:
				promedio = sum(califs) / len(califs)
				if promedio >= 100:
					key = "ORO"
				elif promedio >= 90:
					key = "PLATA"
				elif promedio >= 80:
					key = "BRONCE"
				else:
					key = ""
		else:
			# Para compatibilidad, si recibe un valor individual
			try:
				score_value = float(scores_or_value)
			except Exception:
				score_value = 0
			if score_value >= 100:
				key = "ORO"
			elif score_value >= 90:
				key = "PLATA"
			elif score_value >= 80:
				key = "BRONCE"
			else:
				key = ""
		color = "#6D7480"
		if key == "ORO":
			color = "#B98500"
		elif key == "PLATA":
			color = "#4F5D73"
		elif key == "BRONCE":
			color = "#8C4B20"
		label = {"ORO": "Oro", "PLATA": "Plata", "BRONCE": "Bronce"}.get(key, "Sin medalla")
		title = label
		message = {
			"ORO": "¡Excelente desempeño!",
			"PLATA": "¡Óptimo desempeño!",
			"BRONCE": "Vas bien: favor de reforzar para subir al siguiente nivel."
		}.get(key, "Sin medalla")
		return {
			"key": key,
			"label": label,
			"title": title,
			"message": message,
			"color": color,
		}

	def _accumulated_medals(self, scores: list[dict]) -> dict[str, int]:
		# Solo se asigna UNA medalla basada en el promedio general
		totals = {"ORO": 0, "PLATA": 0, "BRONCE": 0}
		if not scores:
			return totals
		# Calcular el promedio general de las calificaciones
		califs = []
		for item in scores:
			try:
				val = float(item.get("score", 0))
				califs.append(val)
			except Exception:
				pass
		if not califs:
			return totals
		promedio = sum(califs) / len(califs)
		# Asignar medalla según el promedio
		if promedio >= 100:
			totals["ORO"] = 1
		elif promedio >= 90:
			totals["PLATA"] = 1
		elif promedio >= 80:
			totals["BRONCE"] = 1
		return totals

	def _medal_summary_text(self, scores: list[dict]) -> str:
		totals = self._accumulated_medals(scores)
		parts = []
		if totals["ORO"]:
			parts.append(f"Oro x{totals['ORO']}")
		if totals["PLATA"]:
			parts.append(f"Plata x{totals['PLATA']}")
		if totals["BRONCE"]:
			parts.append(f"Bronce x{totals['BRONCE']}")
		return " | ".join(parts) if parts else "Sin medallas acumuladas"

	def _match_medal_filter(self, scores: list[dict]) -> bool:
		selected = str(self.cards_medal_filter_var.get() or "Todas").strip().lower()
		if selected == "todas":
			return True
		totals = self._accumulated_medals(scores)
		if selected == "oro":
			return totals["ORO"] > 0
		if selected == "plata":
			return totals["PLATA"] > 0
		if selected == "bronce":
			return totals["BRONCE"] > 0
		if selected == "sin medalla":
			return sum(totals.values()) == 0
		return True

	def _refresh_history_table(self) -> None:
		if self.tree is None:
			return

		target_inspector = None if self.can_edit else str((self.controller.current_user or {}).get("name", "")).strip()
		row_cache: dict[str, dict] = {}
		rows_to_render: list[tuple[str, tuple[str, str, str, str, str, str]]] = []
		for score in self.controller.list_trimestral_scores(
			inspector_name=target_inspector or None,
			include_unsent=self.can_edit,
		):
			score_id = score.get("id")
			if not score_id:
				continue
			row_cache[score_id] = score
			raw_score = self._coerce_score(score.get("score"))
			score_text = f"{raw_score:.1f}%" if raw_score is not None else "--"
			sent_at = str(score.get("sent_at", "")).strip()
			confirmed_raw = score.get("confirmed_at")
			confirmed_at = "" if confirmed_raw in (None, "") else str(confirmed_raw).strip()
			if sent_at:
				confirmed_text = f"Si ({confirmed_at[:16]})" if confirmed_at and confirmed_at.lower() != "none" else "Pendiente"
			else:
				confirmed_text = "No enviado"
			rows_to_render.append(
				(
					score_id,
					(
						str(score.get("inspector", "--")),
						self._norm_key(score),
						f"{score.get('quarter', '--')} {score.get('year', '--')}",
						score_text,
						confirmed_text,
						str(score.get("updated_at", "--")),
					),
				),
			)

		signature = tuple((score_id, *values) for score_id, values in rows_to_render)
		self.row_cache = row_cache
		if signature == self._history_signature:
			return
		self._history_signature = signature

		selected_score_id = self.selected_score_id if self.selected_score_id in row_cache else None
		for item in self.tree.get_children():
			self.tree.delete(item)

		for score_id, values in rows_to_render:
			self.tree.insert("", "end", iid=score_id, values=values)

		if selected_score_id:
			self.tree.selection_set(selected_score_id)
			self.tree.see(selected_score_id)

	def _collect_history_dashboard_data(self) -> dict[str, object]:
		bars: list[dict[str, object]] = []
		empty_message = "Sin reportes de normas aplicadas para graficar."
		metrics: dict[str, int] = {}
		summary_text = "Sin reportes de normas aplicadas para mostrar."

		reports = self.controller.list_norm_visit_reports()

		if self.can_edit:
			aggregate: dict[tuple[str, str], int] = {}
			for row in reports:
				inspector = str(row.get("inspector", "")).strip()
				raw_norms = row.get("norms", [])
				if not inspector or not isinstance(raw_norms, list):
					continue
				for raw_norm in raw_norms:
					norm = self._norm_key(raw_norm)
					if not norm or norm == "SIN_NORMA":
						continue
					aggregate[(inspector, norm)] = aggregate.get((inspector, norm), 0) + 1

			for (inspector, norm), usage_count in aggregate.items():
				bars.append(
					{
						"inspector": inspector,
						"norm": self._norm_full(norm),
						"usage_count": usage_count,
						"status": "Demanda media",
					}
				)

			if bars:
				highest_usage = max(int(item.get("usage_count", 0) or 0) for item in bars)
				lowest_usage = min(int(item.get("usage_count", 0) or 0) for item in bars)
				for item in bars:
					usage_count = int(item.get("usage_count", 0) or 0)
					if usage_count == highest_usage:
						item["status"] = "Mayor demanda"
					elif usage_count == lowest_usage:
						item["status"] = "Menor demanda"
					else:
						item["status"] = "Demanda media"
			else:
				highest_usage = 0
				lowest_usage = 0

			metrics = {
				"total_combinations": len(bars),
				"total_uses": sum(int(item.get("usage_count", 0) or 0) for item in bars),
				"highest_usage": highest_usage,
				"lowest_usage": lowest_usage,
				"norms_used": len({str(item.get("norm", "")).strip() for item in bars if str(item.get("norm", "")).strip()}),
				"executivos": len({str(item.get("inspector", "")).strip() for item in bars if str(item.get("inspector", "")).strip()}),
				"reports_count": len(reports),
			}
			summary_text = (
				f"Reportes de visita: {metrics['reports_count']} | Combinaciones ejecutivo-norma: {metrics['total_combinations']} | "
				f"Usos totales: {metrics['total_uses']} | Mayor demanda: {metrics['highest_usage']} usos | "
				f"Menor demanda: {metrics['lowest_usage']} usos | Ordenadas por demanda (mayor a menor)."
			)
			report_title = "Reporte de demanda de normas por ejecutivo"
			scope_label = "Global"
		else:
			inspector_name = str((self.controller.current_user or {}).get("name", "")).strip()
			target_identity = self._normalize_identity(inspector_name)
			own_reports = [
				item
				for item in reports
				if self._normalize_identity(str(item.get("inspector", "")).strip()) == target_identity
			]

			aggregate: dict[str, int] = {}
			for row in own_reports:
				raw_norms = row.get("norms", [])
				if not isinstance(raw_norms, list):
					continue
				for raw_norm in raw_norms:
					norm = self._norm_key(raw_norm)
					if not norm or norm == "SIN_NORMA":
						continue
					aggregate[norm] = aggregate.get(norm, 0) + 1

			for norm, usage_count in aggregate.items():
				bars.append(
					{
						"inspector": inspector_name,
						"norm": self._norm_full(norm),
						"usage_count": usage_count,
						"status": "Demanda media",
					}
				)

			if not bars:
				empty_message = "Todavia no has reportado normas aplicadas en tus visitas."
				highest_usage = 0
				lowest_usage = 0
			else:
				highest_usage = max(int(item.get("usage_count", 0) or 0) for item in bars)
				lowest_usage = min(int(item.get("usage_count", 0) or 0) for item in bars)
				for item in bars:
					usage_count = int(item.get("usage_count", 0) or 0)
					if usage_count == highest_usage:
						item["status"] = "Mayor demanda"
					elif usage_count == lowest_usage:
						item["status"] = "Menor demanda"
					else:
						item["status"] = "Demanda media"

			metrics = {
				"total_norms": len(bars),
				"total_uses": sum(int(item.get("usage_count", 0) or 0) for item in bars),
				"highest_usage": highest_usage,
				"lowest_usage": lowest_usage,
				"reports_count": len(own_reports),
			}
			summary_text = (
				f"Reportes de visita: {metrics['reports_count']} | Normas reportadas: {metrics['total_norms']} | "
				f"Usos totales: {metrics['total_uses']} | Mayor demanda: {metrics['highest_usage']} usos | "
				f"Menor demanda: {metrics['lowest_usage']} usos."
			)
			report_title = f"Reporte de demanda de normas de {inspector_name}" if inspector_name else "Reporte de demanda de normas"
			scope_label = inspector_name or "Ejecutivo"

		bars.sort(
			key=lambda item: (
				-int(item.get("usage_count", 0) or 0),
				str(item.get("inspector", "")),
				str(item.get("norm", "")),
			)
		)
		return {
			"bars": bars,
			"empty_message": empty_message,
			"summary_text": summary_text,
			"metrics": metrics,
			"report_title": report_title,
			"scope_label": scope_label,
		}

	def _render_history_dashboard(self) -> None:
		if self.history_dashboard_frame is None:
			return

		container = self.history_dashboard_frame
		for child in container.winfo_children():
			child.destroy()

		dashboard_data = self._collect_history_dashboard_data()
		bars = list(dashboard_data.get("bars", []))
		empty_message = str(dashboard_data.get("empty_message", "Sin datos para mostrar."))
		summary_text = str(dashboard_data.get("summary_text", "Sin reportes de normas aplicadas para mostrar."))
		if self.history_dashboard_summary_label is not None:
			self.history_dashboard_summary_label.configure(text=summary_text)

		if not bars:
			ctk.CTkLabel(
				container,
				text=empty_message,
				font=self.fonts["label"],
				text_color="#6D7480",
			).grid(row=0, column=0, padx=12, pady=24, sticky="n")
			return

		header_row = ctk.CTkFrame(container, fg_color="#F3F5F7", corner_radius=12)
		header_row.grid(row=0, column=0, padx=4, pady=(0, 8), sticky="ew")
		header_row.grid_columnconfigure(0, minsize=340 if self.can_edit else 230)
		header_row.grid_columnconfigure(1, weight=1)
		header_row.grid_columnconfigure(2, minsize=96)
		ctk.CTkLabel(
			header_row,
			text="Ejecutivo y norma" if self.can_edit else "Norma reportada",
			font=self.fonts["small_bold"],
			text_color="#6D7480",
		).grid(row=0, column=0, padx=(14, 10), pady=8, sticky="w")
		ctk.CTkLabel(
			header_row,
			text="Frecuencia de uso (demanda)",
			font=self.fonts["small_bold"],
			text_color="#6D7480",
		).grid(row=0, column=1, padx=(0, 10), pady=8, sticky="w")
		ctk.CTkLabel(
			header_row,
			text="Usos",
			font=self.fonts["small_bold"],
			text_color="#6D7480",
		).grid(row=0, column=2, padx=(0, 14), pady=8, sticky="e")

		label_width = 340 if self.can_edit else 230
		max_usage = max([item["usage_count"] for item in bars], default=1)
		
		for index, item in enumerate(bars, start=1):
			usage_count = int(item.get("usage_count", 0) or 0)
			status_text = str(item.get("status", ""))
			if status_text == "Mayor demanda":
				fill_color = self.style["exito"]
			elif status_text == "Menor demanda":
				fill_color = self.style["advertencia"]
			else:
				fill_color = self.style["primario"]

			row_card = ctk.CTkFrame(
				container,
				fg_color="#FFFFFF",
				corner_radius=14,
				border_width=1,
				border_color="#E9ECEF",
			)
			row_card.grid(row=index, column=0, padx=4, pady=(0, 8), sticky="ew")
			row_card.grid_columnconfigure(0, minsize=label_width)
			row_card.grid_columnconfigure(1, weight=1)
			row_card.grid_columnconfigure(2, minsize=96)

			meta_frame = ctk.CTkFrame(row_card, fg_color="transparent")
			meta_frame.grid(row=0, column=0, padx=(12, 10), pady=10, sticky="nsew")
			norm_label = self._norm_display(str(item.get("norm", "")))

			if self.can_edit:
				title_text = str(item.get("inspector", "--"))
				subtitle_text = f"Norma: {norm_label}"
			else:
				title_text = norm_label
				subtitle_text = "Frecuencia reportada en visitas finalizadas"

			ctk.CTkLabel(
				meta_frame,
				text=title_text,
				font=self.fonts["label_bold"],
				text_color=self.style["texto_oscuro"],
				justify="left",
				anchor="w",
				wraplength=label_width - 26,
			).pack(anchor="w")
			ctk.CTkLabel(
				meta_frame,
				text=subtitle_text,
				font=self.fonts["small"],
				text_color="#6D7480",
				justify="left",
				anchor="w",
				wraplength=label_width - 26,
			).pack(anchor="w", pady=(2, 0))

			bar_frame = ctk.CTkFrame(row_card, fg_color="transparent")
			bar_frame.grid(row=0, column=1, padx=(0, 10), pady=10, sticky="ew")
			bar_frame.grid_columnconfigure(0, weight=1)

			progress = ctk.CTkProgressBar(
				bar_frame,
				height=18,
				corner_radius=9,
				fg_color="#EEF1F4",
				progress_color=fill_color,
			)
			progress.grid(row=0, column=0, sticky="ew")
			progress.set(max(0.0, min(usage_count / max(max_usage, 1), 1.0)))

			usage_text = "uso" if usage_count == 1 else "usos"
			ctk.CTkLabel(
				bar_frame,
				text=f"{usage_count} {usage_text} reportados | {status_text}",
				font=self.fonts["small"],
				text_color="#6D7480",
				anchor="w",
				justify="left",
			).grid(row=1, column=0, pady=(5, 0), sticky="w")

			ctk.CTkLabel(
				row_card,
				text=f"{usage_count}",
				font=self.fonts["label_bold"],
				text_color=fill_color,
			).grid(row=0, column=2, padx=(0, 14), pady=10, sticky="e")

	def _export_history_dashboard_pdf(self) -> None:
		dashboard_data = self._collect_history_dashboard_data()
		bars = list(dashboard_data.get("bars", []))
		if not bars:
			messagebox.showinfo("Trimestral", "No hay reportes de normas para exportar.", parent=self)
			return

		default_path = self.controller.get_default_trimestral_report_path()
		destination = filedialog.asksaveasfilename(
			parent=self,
			title="Guardar reporte trimestral",
			defaultextension=".pdf",
			initialdir=str(default_path.parent),
			initialfile=default_path.name,
			filetypes=[("PDF", "*.pdf")],
		)
		if not destination:
			return

		payload = {
			"can_edit": self.can_edit,
			"report_title": dashboard_data.get("report_title"),
			"scope_label": dashboard_data.get("scope_label"),
			"summary_text": dashboard_data.get("summary_text"),
			"metrics": dashboard_data.get("metrics"),
			"bars": bars,
			"viewer_name": str((self.controller.current_user or {}).get("name", "")).strip() or "Sistema",
			"exported_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
		}

		try:
			output = self.controller.generate_trimestral_dashboard_report(destination, payload)
		except ValueError as error:
			messagebox.showerror("Trimestral", str(error), parent=self)
			return
		except Exception as error:
			messagebox.showerror("Trimestral", f"No se pudo generar el PDF.\n{error}", parent=self)
			return

		messagebox.showinfo("Trimestral", f"Reporte generado en:\n{output}", parent=self)

	def _render_inspector_cards(self) -> None:
		if self.cards_frame is None:
			return

		rows = self.controller.get_principal_rows()
		current_name = str((self.controller.current_user or {}).get("name", "")).strip()
		if not self.can_edit:
			current_identity = self._normalize_identity(current_name)
			rows = [
				row
				for row in rows
				if self._normalize_identity(str(row.get("name", "")).strip()) == current_identity
			]
			if not rows and current_name:
				profile = self.controller.get_executive_profile(current_name)
				accredited_norms = list(profile.get("accredited_norms", [])) if isinstance(profile, dict) else []
				norms_text = ", ".join(accredited_norms) if accredited_norms else "Sin acreditaciones"
				rows = [{"name": current_name, "norms_text": norms_text}]

		score_rows = self.controller.list_trimestral_scores(
			inspector_name=None if self.can_edit else current_name,
			include_unsent=self.can_edit,
		)
		scores_by_inspector: dict[str, list[dict]] = {}
		for score in score_rows:
			inspector_name = str(score.get("inspector", "")).strip()
			inspector_identity = self._normalize_identity(inspector_name)
			if not inspector_identity:
				continue
			scores_by_inspector.setdefault(inspector_identity, []).append(score)

		for values in scores_by_inspector.values():
			values.sort(
				key=lambda item: (
					int(item.get("year", 0)),
					self._quarter_sort_key(item.get("quarter", "")),
					str(item.get("updated_at", "")),
				),
				reverse=True,
			)

		card_models: list[dict[str, str]] = []
		for row in rows:
			inspector_name = str(row.get("name", "")).strip()
			inspector_identity = self._normalize_identity(inspector_name)
			norms_text = str(row.get("norms_text", "Sin acreditaciones")).strip() or "Sin acreditaciones"
			assigned_scores = scores_by_inspector.get(inspector_identity, [])
			if not assigned_scores and not self.can_edit and current_name:
				assigned_scores = scores_by_inspector.get(self._normalize_identity(current_name), [])
			if not self._match_medal_filter(assigned_scores):
				continue
			summary = self._format_assigned_scores(assigned_scores)
			medals = self._accumulated_medals(assigned_scores)
			norm_count = len({self._norm_key(item) for item in assigned_scores})
			card_models.append(
				{
					"inspector_name": inspector_name or current_name,
					"norms_text": norms_text,
					"skills_text": f"Normas calificadas: {norm_count}",
					"pending_text": summary["pending_text"],
					"califications_hint": summary["califications_hint"],
					"medal_summary": self._medal_summary_text(assigned_scores),
					"medal_oro": str(medals["ORO"]),
					"medal_plata": str(medals["PLATA"]),
					"medal_bronce": str(medals["BRONCE"]),
					"send_ready": "1" if summary["send_ready"] else "0",
					"confirm_ready": "1" if summary["confirm_ready"] else "0",
					"check_text": summary["check_text"],
					"check_ok": "1" if summary["check_ok"] else "0",
					"workshop_alert": summary.get("workshop_alert", ""),
				}
			)

		signature = tuple(
			(
				model["inspector_name"],
				model["norms_text"],
				model["skills_text"],
				model["pending_text"],
				model["califications_hint"],
				model["workshop_alert"],
				model["medal_summary"],
				model["medal_oro"],
				model["medal_plata"],
				model["medal_bronce"],
				model["send_ready"],
				model["confirm_ready"],
				model["check_text"],
				model["check_ok"],
				str(self.cards_medal_filter_var.get() or ""),
			)
			for model in card_models
		)
		total_pages = max(1, (len(card_models) + self.cards_page_size - 1) // self.cards_page_size)
		self.current_cards_page = min(self.current_cards_page, total_pages - 1)
		render_signature = (self.current_cards_page, signature)
		if render_signature == self._cards_signature and self.cards_frame.winfo_children():
			self._rebuild_cards_pager(len(card_models))
			return
		self._cards_signature = render_signature

		for child in self.cards_frame.winfo_children():
			child.destroy()

		if not card_models:
			ctk.CTkLabel(
				self.cards_frame,
				text="Sin resultados para el filtro de medallas seleccionado.",
				font=self.fonts["small"],
				text_color="#6D7480",
			).grid(row=0, column=0, padx=10, pady=8, sticky="w")
			self._rebuild_cards_pager(0)
			return

		start_index = self.current_cards_page * self.cards_page_size
		page_models = card_models[start_index:start_index + self.cards_page_size]

		# DEBUG: Mostrar en la UI los inspectores y modelos de card generados
		debug_info = "\n".join([
			f"{i+1}. {m['inspector_name']} | send_ready={m['send_ready']} | confirm_ready={m['confirm_ready']} | pending_text={m['pending_text']} | medals: ORO={m['medal_oro']} PLATA={m['medal_plata']} BRONCE={m['medal_bronce']}"
			for i, m in enumerate(page_models)
		])
		ctk.CTkLabel(
			self.cards_frame,
			text=f"[DEBUG] Inspectores cards generados en página actual:\n{debug_info}",
			font=self.fonts["small"],
			text_color="#B22222",
			anchor="w",
			justify="left",
		).grid(row=99, column=0, columnspan=3, padx=10, pady=2, sticky="w")

		for index, model in enumerate(page_models):
			inspector_name = model["inspector_name"]
			card_host = ctk.CTkFrame(self.cards_frame, fg_color="transparent")
			card_host.grid(row=index // 3, column=index % 3, padx=6, pady=6, sticky="nsew")
			card_host.grid_columnconfigure(0, weight=1)
			card_host.grid_rowconfigure(0, weight=1)

			card = ctk.CTkFrame(
				card_host,
				fg_color="#FFFFFF",
				corner_radius=20,
				border_width=1,
				border_color="#E3E6EA",
				height=320,
			)
			card.grid(row=0, column=0, sticky="nsew")
			card.grid_propagate(False)
			card.grid_columnconfigure(0, weight=1)
			card.grid_columnconfigure(1, minsize=178)

			ctk.CTkLabel(
				card,
				text=inspector_name,
				font=self.fonts["label_bold"],
				text_color=self.style["texto_oscuro"],
				anchor="w",
				justify="left",
				wraplength=200,
			).grid(row=0, column=0, padx=12, pady=(12, 4), sticky="ew")

			ctk.CTkLabel(
				card,
				text=model["skills_text"],
				font=self.fonts["small_bold"],
				text_color="#6D7480",
			).grid(row=1, column=0, padx=12, pady=(0, 6), sticky="w")

			ctk.CTkLabel(
				card,
				text=model["pending_text"],
				font=self.fonts["small_bold"],
				text_color="#6D7480",
			).grid(row=2, column=0, padx=12, pady=(0, 4), sticky="w")

			hint_frame = ctk.CTkFrame(card, fg_color="transparent")
			hint_frame.grid(row=3, column=0, padx=12, pady=(0, 4), sticky="w")
			if model["califications_hint"]:
				ctk.CTkLabel(
					hint_frame,
					text=model["califications_hint"],
					font=self.fonts["small"],
					text_color="#6D7480",
					justify="left",
					wraplength=220,
				).pack(anchor="w")
			if model["workshop_alert"] == "1":
				alert_row = ctk.CTkFrame(hint_frame, fg_color="transparent")
				alert_row.pack(anchor="w", pady=(2, 0))
				alert_img = self.alert_images.get("small")
				if alert_img is not None:
					ctk.CTkLabel(alert_row, text="", image=alert_img).pack(side="left", padx=(0, 4))
				ctk.CTkLabel(
					alert_row,
					text="Requiere tomar taller",
					font=self.fonts["small_bold"],
					text_color="#E67E22",
				).pack(side="left")

			ctk.CTkLabel(
				card,
				text=model["check_text"],
				font=self.fonts["small_bold"],
				text_color=self.style["exito"] if model["check_ok"] == "1" else "#6D7480",
			).grid(row=4, column=0, padx=12, pady=(0, 8), sticky="w")

			medal_panel = ctk.CTkFrame(
				card,
				fg_color="transparent",
				corner_radius=0,
				border_width=0,
				width=250,
				height=154,
			)
			medal_panel.grid(row=0, column=1, rowspan=4, padx=(0, 8), pady=(8, 4), sticky="new")
			medal_panel.grid_propagate(False)
			ctk.CTkLabel(
				medal_panel,
				text="🏆  Medallas",
				font=self.fonts["label_bold"],
				text_color="#8A6A17",
			).pack(padx=8, pady=(8, 2), anchor="center")

			achieved = []
			for key, count_key, color, emoji in [
				("ORO", "medal_oro", "#B98500", "🥇"),
				("PLATA", "medal_plata", "#4F5D73", "🥈"),
				("BRONCE", "medal_bronce", "#8C4B20", "🥉"),
			]:
				try:
					count_value = int(model[count_key])
				except (TypeError, ValueError):
					count_value = 0
				if count_value > 0:
					achieved.append((key, count_value, color, emoji))

			if not achieved:
				ctk.CTkLabel(
					medal_panel,
					text="Sin medallas",
					font=self.fonts["small"],
					text_color="#9AA1AB",
				).pack(padx=8, pady=(26, 0))
			else:
				row_frame = ctk.CTkFrame(medal_panel, fg_color="transparent")
				row_frame.pack(padx=6, pady=(8, 6), fill="x")
				for key, count_value, color, emoji in achieved:
					pill = ctk.CTkFrame(row_frame, fg_color="transparent", corner_radius=0, border_width=0)
					pill.pack(side="left", padx=8)
					image_ref = self.medal_images.get(key)
					if image_ref is not None:
						ctk.CTkLabel(pill, text="", image=image_ref).pack(side="left", padx=(0, 4), pady=0)
					else:
						ctk.CTkLabel(pill, text=emoji, font=("Segoe UI Emoji", 34), text_color=color).pack(side="left", padx=(0, 4), pady=0)
					ctk.CTkLabel(
						pill,
						text=f"x{count_value}",
						font=("Inter", 17, "bold"),
						text_color=color,
					).pack(side="left", padx=(0, 2), pady=(12, 0))

			actions_row = ctk.CTkFrame(card, fg_color="transparent")
			actions_row.grid(row=5, column=0, columnspan=2, padx=12, pady=(2, 10), sticky="ew")
			actions_row.grid_columnconfigure(0, weight=1)
			if self.can_edit:
				actions_row.grid_columnconfigure(1, weight=1)
				actions_row.grid_columnconfigure(2, weight=1)
			else:
				actions_row.grid_columnconfigure(1, weight=1)

			ctk.CTkButton(
				actions_row,
				text="Ver calificaciones",
				fg_color=self.style["primario"],
				text_color=self.style["texto_oscuro"],
				hover_color="#D8C220",
				command=lambda name=inspector_name: self._open_preview_popup(name),
			).grid(row=0, column=0, padx=(0, 4 if self.can_edit else 3), sticky="ew")

			if self.can_edit:
				ctk.CTkButton(
					actions_row,
					text="Captura",
					fg_color=self.style["fondo"],
					text_color=self.style["texto_oscuro"],
					hover_color="#E9ECEF",
					state="normal",
					command=lambda name=inspector_name: self._open_capture_for_inspector(name),
				).grid(row=0, column=1, padx=4, sticky="ew")
				ctk.CTkButton(
					actions_row,
					text="Enviar",
					fg_color=self.style["secundario"],
					text_color=self.style["texto_claro"],
					hover_color="#1D1D1D",
					state="normal",
					command=lambda name=inspector_name: self._send_scores_for_inspector(name),
				).grid(row=0, column=2, padx=(4, 0), sticky="ew")
			else:
				ctk.CTkButton(
					actions_row,
					text="Confirmar visto",
					fg_color=self.style["exito"],
					hover_color="#0B7A4D",
					state="normal" if model["confirm_ready"] == "1" else "disabled",
					command=lambda name=inspector_name: self._confirm_scores_for_inspector(name),
				).grid(row=0, column=1, padx=(3, 0), sticky="ew")

		self._rebuild_cards_pager(len(card_models))

	def _format_assigned_scores(self, scores: list[dict]) -> dict[str, str]:
		if not scores:
			return {
				"pending_text": "Estado: Sin calificaciones",
				"califications_hint": "Registra calificaciones trimestrales.",
				"send_ready": "",
				"confirm_ready": "",
				"check_text": "Check ejecutivo: --",
				"check_ok": "",
			}

		unsent_rows = [item for item in scores if not str(item.get("sent_at", "")).strip()]
		sent_rows = [item for item in scores if str(item.get("sent_at", "")).strip()]
		pending_confirm = [item for item in sent_rows if not str(item.get("confirmed_at") or "").strip()]
		confirmed_rows = [item for item in sent_rows if str(item.get("confirmed_at") or "").strip()]
		medals = self._accumulated_medals(scores)
		critical_count = 0
		workshop_count = 0
		latest_score_value = None
		for item in scores:
			score_value = self._coerce_score(item.get("score"))
			if latest_score_value is None and score_value is not None:
				latest_score_value = score_value
			if score_value is not None and score_value < 90:
				critical_count += 1
			if score_value is not None and score_value <= 80:
				workshop_count += 1

		if unsent_rows:
			state_text = "Estado: Calificaciones asignadas"
			hint_text = f"Envio: {len(sent_rows)} enviadas | {len(unsent_rows)} pendientes"
		else:
			state_text = "Estado: Calificado"
			hint_text = ""

		workshop_alert = ""
		if latest_score_value is not None and latest_score_value <= 80:
			workshop_alert = "1"
		elif workshop_count > 0:
			workshop_alert = "1"

		if sent_rows:
			if pending_confirm:
				check_text = f"Check ejecutivo: {len(confirmed_rows)}/{len(sent_rows)} confirmadas"
				check_ok = ""
			else:
				check_text = "Check ejecutivo: ✓ Confirmado"
				check_ok = "1"
		else:
			check_text = "Check ejecutivo: --"
			check_ok = ""

		return {
			"pending_text": state_text,
			"califications_hint": hint_text,
			"workshop_alert": workshop_alert,
			"send_ready": "1" if len(unsent_rows) > 0 else "",
			"confirm_ready": "1" if len(pending_confirm) > 0 else "",
			"check_text": check_text,
			"check_ok": check_ok,
		}

	def _send_scores_for_inspector(self, inspector_name: str) -> None:
		if not self.can_edit:
			return

		scores = self.controller.list_trimestral_scores(inspector_name=inspector_name, include_unsent=True)
		unsent_ids = [
			str(item.get("id", "")).strip()
			for item in scores
			if str(item.get("id", "")).strip() and not str(item.get("sent_at", "")).strip()
		]
		if not unsent_ids:
			messagebox.showinfo("Trimestral", "No hay calificaciones pendientes por enviar para este ejecutivo.", parent=self)
			return

		if not messagebox.askyesno(
			"Trimestral",
			f"Deseas enviar {len(unsent_ids)} calificaciones al ejecutivo tecnico?",
			parent=self,
		):
			return

		sent_count = self.controller.send_trimestral_scores(inspector_name, unsent_ids)
		if sent_count <= 0:
			messagebox.showwarning("Trimestral", "No se pudieron enviar las calificaciones seleccionadas.", parent=self)
			return

		messagebox.showinfo("Trimestral", f"Se enviaron {sent_count} calificaciones.", parent=self)
		self.refresh()

	def _confirm_scores_for_inspector(self, inspector_name: str) -> None:
		if self.can_edit:
			return

		scores = self.controller.list_trimestral_scores(inspector_name=inspector_name, include_unsent=False)
		pending_ids = [
			str(item.get("id", "")).strip()
			for item in scores
			if str(item.get("id", "")).strip()
			and str(item.get("sent_at", "")).strip()
			and not str(item.get("confirmed_at") or "").strip()
		]
		if not pending_ids:
			messagebox.showinfo("Trimestral", "No hay calificaciones pendientes por confirmar.", parent=self)
			return

		if not messagebox.askyesno("Trimestral", "Confirmas que ya revisaste tus calificaciones?", parent=self):
			return

		confirmed_count = self.controller.confirm_trimestral_scores(inspector_name, pending_ids)
		if confirmed_count <= 0:
			messagebox.showwarning("Trimestral", "No se pudieron confirmar las calificaciones.", parent=self)
			return

		messagebox.showinfo("Trimestral", f"Se confirmaron {confirmed_count} calificaciones.", parent=self)
		self.refresh()

	def _open_preview_popup(self, inspector_name: str) -> None:
		scores = self.controller.list_trimestral_scores(inspector_name=inspector_name, include_unsent=True)

		dialog = ctk.CTkToplevel(self)
		dialog.title(f"Preview calificaciones — {inspector_name}")
		dialog.geometry("950x550")
		dialog.minsize(950,550)
		dialog.configure(fg_color=self.style["fondo"])
		dialog.transient(self.winfo_toplevel())
		dialog.grab_set()

		wrapper = ctk.CTkFrame(dialog, fg_color=self.style["surface"], corner_radius=20)
		wrapper.pack(fill="both", expand=True, padx=18, pady=18)
		wrapper.grid_columnconfigure(0, weight=1)
		wrapper.grid_rowconfigure(4, weight=1)

		ctk.CTkLabel(
			wrapper,
			text=f"Calificaciones asignadas — {inspector_name}",
			font=self.fonts["label_bold"],
			text_color=self.style["texto_oscuro"],
		).grid(row=0, column=0, padx=18, pady=(14, 10), sticky="w")

		filters = ctk.CTkFrame(wrapper, fg_color="transparent")
		filters.grid(row=1, column=0, padx=18, pady=(0, 10), sticky="ew")
		filters.grid_columnconfigure(9, weight=1)

		valid_years: list[int] = []
		for score_row in scores:
			try:
				year_value = int(score_row.get("year", 0))
			except (TypeError, ValueError):
				continue
			if year_value > 0 and year_value not in valid_years:
				valid_years.append(year_value)
		valid_years.sort(reverse=True)

		year_var = ctk.StringVar(value="Todos")
		quarter_var = ctk.StringVar(value="Todos")
		# Quitar filtro de medalla
		norm_var = ctk.StringVar(value="Todas")

		norm_tokens = sorted({self._norm_key(item) for item in scores if self._norm_key(item)})
		norm_options = ["Todas", *norm_tokens]

		ctk.CTkLabel(
			filters,
			text="Año",
			font=self.fonts["small_bold"],
			text_color="#6D7480",
		).grid(row=0, column=0, padx=(0, 8), sticky="w")
		year_selector = ctk.CTkComboBox(
			filters,
			variable=year_var,
			values=["Todos", *[str(value) for value in valid_years]],
			width=120,
			height=34,
			fg_color="#FFFFFF",
			border_color="#D5D8DC",
			button_color=self.style["primario"],
			dropdown_hover_color=self.style["primario"],
		)
		year_selector.grid(row=0, column=1, padx=(0, 16), sticky="w")

		ctk.CTkLabel(
			filters,
			text="Trimestre",
			font=self.fonts["small_bold"],
			text_color="#6D7480",
		).grid(row=0, column=2, padx=(0, 8), sticky="w")
		quarter_selector = ctk.CTkComboBox(
			filters,
			variable=quarter_var,
			values=["Todos"],
			width=120,
			height=34,
			fg_color="#FFFFFF",
			border_color="#D5D8DC",
			button_color=self.style["primario"],
			dropdown_hover_color=self.style["primario"],
		)
		quarter_selector.grid(row=0, column=3, padx=(0, 16), sticky="w")

		# Eliminado filtro de medalla

		ctk.CTkLabel(
			filters,
			text="Norma",
			font=self.fonts["small_bold"],
			text_color="#6D7480",
		).grid(row=0, column=6, padx=(0, 8), sticky="w")
		norm_selector = ctk.CTkComboBox(
			filters,
			variable=norm_var,
			values=norm_options,
			width=180,
			height=34,
			fg_color="#FFFFFF",
			border_color="#D5D8DC",
			button_color=self.style["primario"],
			dropdown_hover_color=self.style["primario"],
		)
		norm_selector.grid(row=0, column=7, padx=(0, 12), sticky="w")

		ctk.CTkLabel(
			filters,
			text=(
				"Vista admin: se muestran solo calificaciones criticas (<90%) filtradas por año, trimestre y medalla."
				if self.can_edit
				else "Consulta las calificaciones pasadas filtrando por año, trimestre y medalla."
			),
			font=self.fonts["small"],
			text_color="#6D7480",
		).grid(row=0, column=8, sticky="w")

		summary_label = ctk.CTkLabel(
			wrapper,
			text="",
			font=self.fonts["small"],
			text_color="#6D7480",
			justify="left",
		)
		summary_label.grid(row=2, column=0, padx=18, pady=(0, 8), sticky="w")

		curve_canvas = tk.Canvas(
			wrapper,
			height=170,
			bg=self.style["surface"],
			highlightthickness=0,
		)
		curve_canvas.grid(row=3, column=0, padx=18, pady=(0, 8), sticky="ew")

		results_frame = ctk.CTkScrollableFrame(wrapper, fg_color="#FFFFFF", corner_radius=16)
		results_frame.grid(row=4, column=0, padx=18, pady=(0, 10), sticky="nsew")
		results_frame.grid_columnconfigure(0, weight=1)

		def _score_state(score_row: dict) -> tuple[str, str]:
			raw_score = self._coerce_score(score_row.get("score"))
			if self.can_edit and raw_score is not None and raw_score < 90:
				return "Critico (<90%)", self.style["advertencia"]
			sent_at = str(score_row.get("sent_at", "")).strip()
			confirmed_at = str(score_row.get("confirmed_at") or "").strip()
			if not sent_at:
				return "Sin enviar", "#6D7480"
			if confirmed_at and confirmed_at.lower() != "none":
				return "Confirmado", self.style["exito"]
			return "Enviada", self.style["advertencia"]

		def _medal_message(medal_key: str) -> str:
			if medal_key == "ORO":
				return "¡Felicidades, excelente desempeño!"
			if medal_key == "PLATA":
				return "¡Felicidades, óptimo desempeño!"
			if medal_key == "BRONCE":
				return "Sigue esforzándote"
			return "--"

		def _available_quarters() -> list[str]:
			selected_year = year_var.get().strip()
			quarters: set[str] = set()
			for score_row in scores:
				if selected_year != "Todos" and str(score_row.get("year", "")).strip() != selected_year:
					continue
				quarter_value = str(score_row.get("quarter", "")).strip().upper()
				if quarter_value in {"T1", "T2", "T3", "T4"}:
					quarters.add(quarter_value)
			ordered = sorted(quarters, key=self._quarter_sort_key)
			return ["Todos", *ordered]

		def _filtered_scores() -> list[dict]:
			selected_year = year_var.get().strip()
			selected_quarter = quarter_var.get().strip().upper()
			selected_norm = norm_var.get().strip()
			filtered = list(scores)
			if selected_year != "Todos":
				filtered = [item for item in filtered if str(item.get("year", "")).strip() == selected_year]
			if selected_quarter != "TODOS":
				filtered = [item for item in filtered if str(item.get("quarter", "")).strip().upper() == selected_quarter]
			if selected_norm != "Todas":
				filtered = [item for item in filtered if self._norm_key(item) == selected_norm]
			filtered.sort(
				key=lambda item: (
					int(item.get("year", 0)),
					self._quarter_sort_key(item.get("quarter", "")),
					self._norm_key(item),
					str(item.get("updated_at", "")),
				),
				reverse=True,
			)
			return filtered

		def _render_preview(_value=None) -> None:
			# Calcular promedio general
			filtered_scores = _filtered_scores()
			score_values = [self._coerce_score(item.get("score")) for item in filtered_scores if self._coerce_score(item.get("score")) is not None]
			avg_score = sum(score_values) / len(score_values) if score_values else None
			for child in results_frame.winfo_children():
				child.destroy()

			selected_year = year_var.get().strip()
			selected_quarter = quarter_var.get().strip().upper()
			# selected_medal = medal_var.get().strip()  # Eliminado

			period_label = selected_year if selected_year != "Todos" else "todos los años"
			quarter_label = selected_quarter if selected_quarter != "TODOS" else "todos los trimestres"
			norm_label = norm_var.get().strip()

			curve_rows = sorted(
				filtered_scores,
				key=lambda item: (
					int(item.get("year", 0) or 0),
					self._quarter_sort_key(item.get("quarter", "")),
					str(item.get("updated_at", "")),
				),
			)
			curve_history = []
			for item in curve_rows:
				score_value = self._coerce_score(item.get("score"))
				if score_value is None:
					continue
				curve_history.append(
					{
						"label": f"{item.get('quarter', '--')} {item.get('year', '--')}",
						"score": score_value,
					}
				)
			self._draw_curve_on_canvas(curve_canvas, curve_history, "Sin datos para dibujar curva de dispersion.")
			if not filtered_scores:
				summary_label.configure(
					text=(
						f"Sin calificaciones criticas (<90%) para {period_label}, {quarter_label}, norma {norm_label}."
						if self.can_edit
						else f"Sin registros para {period_label}, {quarter_label}, norma {norm_label}."
					)
				)
				ctk.CTkLabel(
					results_frame,
					text=(
						"No hay calificaciones criticas por debajo de 90% para el filtro seleccionado."
						if self.can_edit
						else "No hay calificaciones históricas para el filtro seleccionado."
					),
					font=self.fonts["label"],
					text_color="#6D7480",
				).grid(row=0, column=0, padx=12, pady=24, sticky="n")
				return

			periods = {
				(int(item.get("year", 0)), str(item.get("quarter", "")).strip().upper())
				for item in filtered_scores
			}
			norms = {self._norm_key(item) for item in filtered_scores}
			promedio_text = f" | Promedio general: {avg_score:.1f}%" if avg_score is not None else ""
			summary_label.configure(
				text=(
					f"Mostrando {len(filtered_scores)} calificaciones criticas (<90%) en {len(periods)} periodos | "
					f"Normas: {len(norms)} | Año: {period_label} | Trimestre: {quarter_label} | Norma: {norm_label}.{promedio_text}"
					if self.can_edit
					else f"Mostrando {len(filtered_scores)} calificaciones en {len(periods)} periodos | "
					f"Normas: {len(norms)} | Año: {period_label} | Trimestre: {quarter_label} | Norma: {norm_label}.{promedio_text}"
				)
			)

			grouped: dict[tuple[int, str], list[dict]] = {}
			for score_row in filtered_scores:
				try:
					year_value = int(score_row.get("year", 0))
				except (TypeError, ValueError):
					year_value = 0
				quarter_value = str(score_row.get("quarter", "")).strip().upper() or "--"
				grouped.setdefault((year_value, quarter_value), []).append(score_row)

			ordered_periods = sorted(grouped, key=lambda item: (item[0], self._quarter_sort_key(item[1])), reverse=True)
			for index, period_key in enumerate(ordered_periods):
				year_value, quarter_value = period_key
				period_rows = sorted(grouped[period_key], key=lambda item: self._norm_key(item))

				card = ctk.CTkFrame(
					results_frame,
					fg_color="#FFFFFF",
					corner_radius=14,
					border_width=1,
					border_color="#E9ECEF",
				)
				card.grid(row=index, column=0, padx=4, pady=(0, 10), sticky="ew")
				card.grid_columnconfigure(0, weight=1)

				header = ctk.CTkFrame(card, fg_color="#F3F5F7", corner_radius=12)
				header.grid(row=0, column=0, padx=10, pady=(10, 8), sticky="ew")
				header.grid_columnconfigure(0, weight=1)
				ctk.CTkLabel(
					header,
					text=f"{quarter_value} {year_value}",
					font=self.fonts["label_bold"],
					text_color=self.style["texto_oscuro"],
				).grid(row=0, column=0, padx=12, pady=8, sticky="w")
				ctk.CTkLabel(
					header,
					text=f"Registros: {len(period_rows)}",
					font=self.fonts["small_bold"],
					text_color="#6D7480",
				).grid(row=0, column=1, padx=12, pady=8, sticky="e")

				columns = ctk.CTkFrame(card, fg_color="transparent")
				columns.grid(row=1, column=0, padx=12, pady=(0, 4), sticky="ew")
				columns.grid_columnconfigure(0, weight=1)
				columns.grid_columnconfigure(1, minsize=100)
				columns.grid_columnconfigure(2, minsize=130)
				columns.grid_columnconfigure(3, minsize=140)
				columns.grid_columnconfigure(4, minsize=120)
				columns.grid_columnconfigure(5, minsize=230)
				for col_index, label_text in enumerate(["Norma", "Calificacion", "Estado", "Medalla", "Actualizado", "Mensaje"]):
					anchor = "w" if col_index == 0 else "center"
					sticky = "w" if col_index in (0, 5) else "ew"
					ctk.CTkLabel(
						columns,
						text=label_text,
						font=self.fonts["small_bold"],
						text_color="#6D7480",
						anchor=anchor,
					).grid(row=0, column=col_index, pady=(0, 2), sticky=sticky)

				for row_index, score_row in enumerate(period_rows, start=2):
					row_frame = ctk.CTkFrame(card, fg_color="transparent")
					row_frame.grid(row=row_index, column=0, padx=12, pady=(0, 6), sticky="ew")
					row_frame.grid_columnconfigure(0, weight=1)
					row_frame.grid_columnconfigure(1, minsize=100)
					row_frame.grid_columnconfigure(2, minsize=130)
					row_frame.grid_columnconfigure(3, minsize=140)
					row_frame.grid_columnconfigure(4, minsize=120)
					row_frame.grid_columnconfigure(5, minsize=230)

					norm_text = self._norm_display(self._norm_key(score_row))
					raw_score = self._coerce_score(score_row.get("score"))
					score_text = f"{raw_score:.1f}%" if raw_score is not None else "--"
					state_text, state_color = _score_state(score_row)
					medal = self._score_medal(score_row)
					medal_text = medal["title"] if medal["key"] else "Sin medalla"
					medal_message = _medal_message(medal.get("key", ""))
					updated_text = str(score_row.get("updated_at", "")).strip() or "--"
					notes_text = str(score_row.get("notes", "")).strip()

					ctk.CTkLabel(
						row_frame,
						text=norm_text,
						font=self.fonts["small"],
						text_color=self.style["texto_oscuro"],
						justify="left",
						anchor="w",
						wraplength=320,
					).grid(row=0, column=0, sticky="w")
					ctk.CTkLabel(
						row_frame,
						text=score_text,
						font=self.fonts["small_bold"],
						text_color=self.style["texto_oscuro"],
					).grid(row=0, column=1, sticky="ew")
					ctk.CTkLabel(
						row_frame,
						text=state_text,
						font=self.fonts["small_bold"],
						text_color=state_color,
					).grid(row=0, column=2, sticky="ew")
					medal_cell = ctk.CTkFrame(row_frame, fg_color="transparent")
					medal_cell.grid(row=0, column=3, sticky="ew")
					medal_img = self.medal_images_small.get(medal["key"]) if medal["key"] else None
					if medal_img is not None:
						ctk.CTkLabel(medal_cell, text="", image=medal_img).pack(side="left", padx=(16, 4))
						ctk.CTkLabel(
							medal_cell,
							text=medal_text,
							font=self.fonts["small_bold"],
							text_color=medal["color"],
						).pack(side="left")
					else:
						ctk.CTkLabel(
							medal_cell,
							text=medal_text,
							font=self.fonts["small_bold"],
							text_color=medal["color"],
						).pack(side="left", padx=16)
					ctk.CTkLabel(
						row_frame,
						text=updated_text[:16],
						font=self.fonts["small"],
						text_color="#6D7480",
					).grid(row=0, column=4, sticky="ew")
					ctk.CTkLabel(
						row_frame,
						text=medal_message,
						font=self.fonts["small"],
						text_color="#6D7480" if medal_message == "--" else self.style["texto_oscuro"],
						anchor="w",
						justify="left",
						wraplength=220,
					).grid(row=0, column=5, sticky="w")

					if notes_text:
						notes_frame = ctk.CTkFrame(card, fg_color="#FFF8E1", corner_radius=8)
						notes_frame.grid(row=row_index + len(period_rows) + 10, column=0, padx=20, pady=(0, 6), sticky="ew")
						ctk.CTkLabel(
							notes_frame,
							text=f"📝 Notas: {notes_text}",
							font=self.fonts["small"],
							text_color="#5D4E00",
							anchor="w",
							justify="left",
							wraplength=800,
						).pack(padx=10, pady=6, anchor="w")

		def _on_year_change(_value=None) -> None:
			quarter_values = _available_quarters()
			quarter_selector.configure(values=quarter_values)
			if quarter_var.get() not in quarter_values:
				quarter_var.set("Todos")
			_render_preview()

		year_selector.configure(command=_on_year_change)
		quarter_selector.configure(command=_render_preview)
		# medal_selector ya no existe, eliminado
		norm_selector.configure(command=_render_preview)
		curve_canvas.bind("<Configure>", _render_preview)
		_on_year_change()

		ctk.CTkButton(
			wrapper,
			text="Cerrar",
			fg_color=self.style["fondo"],
			text_color=self.style["texto_oscuro"],
			hover_color="#E9ECEF",
			command=dialog.destroy,
		).grid(row=5, column=0, padx=18, pady=(0, 14), sticky="e")

	def _open_capture_for_inspector(self, inspector_name: str) -> None:
		if not self.can_edit:
			return
		self._open_capture_dialog(inspector_name=inspector_name)

	def _open_inspector_detail(self, inspector_name: str) -> None:
		profile = self.controller.get_executive_profile(inspector_name)
		accredited_norms = list(profile.get("accredited_norms", [])) if isinstance(profile, dict) else []
		scores_for_inspector = self.controller.list_trimestral_scores(inspector_name=inspector_name)
		norm_values = self._capture_norm_values(inspector_name)
		for token in accredited_norms:
			norm_token = self._norm_key(token)
			if norm_token not in norm_values:
				norm_values.append(norm_token)

		if scores_for_inspector:
			default_norm = self._norm_key(scores_for_inspector[0])
			if default_norm not in norm_values:
				norm_values.insert(0, default_norm)
		else:
			default_norm = norm_values[0] if norm_values else "SIN_NORMA"

		dialog = ctk.CTkToplevel(self)
		dialog.title(f"Detalle trimestral - {inspector_name}")
		dialog.geometry("1120x680")
		dialog.minsize(980, 560)
		dialog.configure(fg_color=self.style["fondo"])
		dialog.transient(self.winfo_toplevel())
		dialog.grab_set()

		wrapper = ctk.CTkFrame(dialog, fg_color=self.style["surface"], corner_radius=20)
		wrapper.pack(fill="both", expand=True, padx=18, pady=18)
		wrapper.grid_columnconfigure(0, weight=3)
		wrapper.grid_columnconfigure(1, weight=2)
		wrapper.grid_rowconfigure(1, weight=1)

		ctk.CTkLabel(
			wrapper,
			text=f"Trimestral de {inspector_name}",
			font=self.fonts["label_bold"],
			text_color=self.style["texto_oscuro"],
		).grid(row=0, column=0, padx=(18, 10), pady=(16, 10), sticky="w")

		norm_var = ctk.StringVar(value=default_norm)
		norm_selector = ctk.CTkComboBox(
			wrapper,
			variable=norm_var,
			values=norm_values,
			height=36,
			fg_color="#FFFFFF",
			border_color="#D5D8DC",
			button_color=self.style["primario"],
			dropdown_hover_color=self.style["primario"],
		)
		norm_selector.grid(row=0, column=1, padx=(10, 18), pady=(16, 10), sticky="ew")

		curve_canvas = tk.Canvas(
			wrapper,
			bg=self.style["surface"],
			highlightthickness=0,
			height=360,
		)
		curve_canvas.grid(row=1, column=0, padx=(18, 10), pady=(0, 16), sticky="nsew")

		info_scroll = ctk.CTkScrollableFrame(wrapper, fg_color="#F8F9FA", corner_radius=16)
		info_scroll.grid(row=1, column=1, padx=(10, 18), pady=(0, 16), sticky="nsew")
		info_scroll.grid_columnconfigure(0, weight=1)

		detail_scores = list(scores_for_inspector)
		confirm_button: ctk.CTkButton | None = None

		def _scores_for_selected_norm() -> list[dict]:
			target_norm = self._norm_key(norm_var.get())
			filtered = [item for item in detail_scores if self._norm_key(item) == target_norm]
			filtered.sort(
				key=lambda item: (
					int(item.get("year", 0)),
					self._quarter_sort_key(item.get("quarter", "")),
					str(item.get("updated_at", "")),
				),
				reverse=True,
			)
			return filtered

		def _is_conf(row: dict) -> bool:
			v = str(row.get("confirmed_at") or "").strip().lower()
			return bool(v) and v != "none"

		def _render_detail(_value=None) -> None:
			# clear previous content
			for _w in info_scroll.winfo_children():
				_w.destroy()

			norm_token = self._norm_key(norm_var.get())
			norm_display = self._norm_display(norm_token)
			selected_rows = _scores_for_selected_norm()

			# -- update chart --
			curve_history = []
			for row in reversed(selected_rows):
				score_value = self._coerce_score(row.get("score"))
				if score_value is None:
					continue
				label = f"{row.get('quarter', '--')} {row.get('year', '--')}"
				curve_history.append({"label": label, "score": score_value})
			self._draw_curve_on_canvas(curve_canvas, curve_history, f"Sin curva disponible para {norm_token}.")

			# ── Norm header ──────────────────────────────────────────
			ctk.CTkLabel(
				info_scroll,
				text=norm_token,
				font=("Arial", 20, "bold"),
				text_color=self.style["texto_oscuro"],
			).pack(anchor="w", padx=14, pady=(12, 0))
			ctk.CTkLabel(
				info_scroll,
				text=norm_display,
				font=("Arial", 10),
				text_color="#6D7480",
				wraplength=280,
				justify="left",
			).pack(anchor="w", padx=14, pady=(0, 8))

			ctk.CTkFrame(info_scroll, height=1, fg_color="#E0E3E8").pack(fill="x", padx=14, pady=(0, 10))

			if selected_rows:
				numeric_scores = [
					v for v in (self._coerce_score(r.get("score")) for r in selected_rows)
					if v is not None
				]
				latest = selected_rows[0]
				latest_score = self._coerce_score(latest.get("score"))
				latest_medal = self._score_medal(latest)
				latest_period = f"{latest.get('quarter', '--')} {latest.get('year', '--')}"
				is_latest_confirmed = _is_conf(latest)
				is_latest_sent = bool(str(latest.get("sent_at") or "").strip())

				# ── Big score ────────────────────────────────────────
				score_color = "#1E9E5F" if (latest_score or 0) >= 90 else "#C0392B"
				score_display = f"{latest_score:.1f}%" if latest_score is not None else "--"
				score_row = ctk.CTkFrame(info_scroll, fg_color="transparent")
				score_row.pack(anchor="w", padx=14, pady=(0, 2))
				if latest_score is not None and latest_score <= 80:
					alert_large = self.alert_images.get("large")
					if alert_large is not None:
						ctk.CTkLabel(score_row, text="", image=alert_large).pack(side="left", padx=(0, 6))
					else:
						ctk.CTkLabel(score_row, text="!", font=("Arial", 18, "bold"), text_color="#C0392B").pack(side="left", padx=(0, 6))
				ctk.CTkLabel(
					score_row,
					text=score_display,
					font=("Arial", 36, "bold"),
					text_color=score_color,
				).pack(side="left")
				ctk.CTkLabel(
					info_scroll,
					text=f"Periodo: {latest_period}",
					font=("Arial", 11),
					text_color="#6D7480",
				).pack(anchor="w", padx=14, pady=(0, 6))

				medal_row = ctk.CTkFrame(info_scroll, fg_color="transparent")
				medal_row.pack(anchor="w", padx=14, pady=(0, 6))
				medal_key = str(latest_medal.get("key", "")).strip().upper()
				medal_image = self.medal_images.get(medal_key)
				if medal_image is not None:
					ctk.CTkLabel(
						medal_row,
						text="",
						image=medal_image,
					).pack(side="left", padx=(0, 6))
				ctk.CTkLabel(
					medal_row,
					text=f"Medalla: {latest_medal['label'] if latest_medal['key'] else 'Sin medalla'}",
					font=("Arial", 11, "bold"),
					text_color=latest_medal["color"],
				).pack(side="left")

				if latest_score is not None and latest_score <= 80:
					ctk.CTkLabel(
						info_scroll,
						text="Recomendacion: asistir a taller para reforzar conocimientos y aumentar su nivel.",
						font=("Arial", 11, "bold"),
						text_color="#C0392B",
						wraplength=290,
						justify="left",
					).pack(anchor="w", padx=14, pady=(0, 8))

				# ── Status badge ─────────────────────────────────────
				if not is_latest_sent:
					badge_text, badge_bg, badge_fg = "Sin enviar", "#F0F0F0", "#6D7480"
				elif is_latest_confirmed:
					badge_text, badge_bg, badge_fg = "✓ Confirmado", "#D4EDDA", "#1E9E5F"
				else:
					badge_text, badge_bg, badge_fg = "⏳ Pendiente de confirmar", "#FFF3CD", "#856404"
				ctk.CTkLabel(
					info_scroll,
					text=badge_text,
					fg_color=badge_bg,
					text_color=badge_fg,
					corner_radius=8,
					padx=10,
					font=("Arial", 11, "bold"),
				).pack(anchor="w", padx=14, pady=(0, 12))

				# ── Stats grid ───────────────────────────────────────
				avg_score = round(mean(numeric_scores), 1) if numeric_scores else None
				best_score = max(numeric_scores) if numeric_scores else None
				worst_score = min(numeric_scores) if numeric_scores else None

				stats_frame = ctk.CTkFrame(info_scroll, fg_color="#FFFFFF", corner_radius=10)
				stats_frame.pack(fill="x", padx=14, pady=(0, 10))
				for _col in range(3):
					stats_frame.grid_columnconfigure(_col, weight=1)
				for _ci, (_lbl, _val) in enumerate([
					("Promedio", f"{avg_score:.1f}%" if avg_score is not None else "--"),
					("Mejor", f"{best_score:.1f}%" if best_score is not None else "--"),
					("Mas baja", f"{worst_score:.1f}%" if worst_score is not None else "--"),
				]):
					_cell = ctk.CTkFrame(stats_frame, fg_color="transparent")
					_cell.grid(row=0, column=_ci, padx=8, pady=8)
					ctk.CTkLabel(_cell, text=_lbl, font=("Arial", 9), text_color="#6D7480").pack()
					value_row = ctk.CTkFrame(_cell, fg_color="transparent")
					value_row.pack()
					if _lbl == "Mas baja" and worst_score is not None and worst_score <= 80:
						alert_small = self.alert_images.get("small")
						if alert_small is not None:
							ctk.CTkLabel(value_row, text="", image=alert_small).pack(side="left", padx=(0, 4))
						else:
							ctk.CTkLabel(value_row, text="!", font=("Arial", 10, "bold"), text_color="#C0392B").pack(side="left", padx=(0, 4))
					ctk.CTkLabel(value_row, text=_val, font=("Arial", 13, "bold"), text_color=self.style["texto_oscuro"]).pack(side="left")
			else:
				ctk.CTkLabel(
					info_scroll,
					text="Sin calificaciones para esta norma.",
					font=("Arial", 12),
					text_color="#6D7480",
				).pack(anchor="w", padx=14, pady=12)

			# ── Calificaciones del año ────────────────────────────────
			current_year = datetime.now().year
			current_year_rows = [r for r in selected_rows if str(r.get("year", "")) == str(current_year)]

			ctk.CTkFrame(info_scroll, height=1, fg_color="#E0E3E8").pack(fill="x", padx=14, pady=(4, 8))
			ctk.CTkLabel(
				info_scroll,
				text=f"Calificaciones {current_year}",
				font=("Arial", 12, "bold"),
				text_color=self.style["texto_oscuro"],
			).pack(anchor="w", padx=14, pady=(0, 6))

			if current_year_rows:
				for row in current_year_rows:
					sv = self._coerce_score(row.get("score"))
					score_txt = f"{sv:.1f}%" if sv is not None else "--"
					medal = self._score_medal(row)
					medal_txt = medal["title"] if medal["key"] else "Sin medalla"
					period = f"{row.get('quarter', '--')} {row.get('year', '--')}"
					sentv = str(row.get("sent_at") or "").strip()
					if not sentv:
						st, st_color = "Sin enviar", "#6D7480"
					elif _is_conf(row):
						st, st_color = "✓ Confirmado", "#1E9E5F"
					else:
						st, st_color = "Pendiente", "#856404"
					row_f = ctk.CTkFrame(info_scroll, fg_color="transparent")
					row_f.pack(fill="x", padx=14, pady=2)
					row_f.grid_columnconfigure(0, weight=1)
					left_row = ctk.CTkFrame(row_f, fg_color="transparent")
					left_row.grid(row=0, column=0, sticky="w")
					if sv is not None and sv <= 80:
						alert_small = self.alert_images.get("small")
						if alert_small is not None:
							ctk.CTkLabel(left_row, text="", image=alert_small).pack(side="left", padx=(0, 4))
						else:
							ctk.CTkLabel(left_row, text="!", font=("Arial", 10, "bold"), text_color="#C0392B").pack(side="left", padx=(0, 4))
					ctk.CTkLabel(left_row, text=f"{period}  {score_txt}  |  {medal_txt}", font=("Arial", 11), text_color=self.style["texto_oscuro"]).pack(side="left")
					ctk.CTkLabel(row_f, text=st, font=("Arial", 10, "bold"), text_color=st_color).grid(row=0, column=1, sticky="e")
			else:
				ctk.CTkLabel(info_scroll, text=f"Sin calificaciones para {current_year}.", font=("Arial", 11), text_color="#6D7480").pack(anchor="w", padx=14)

			# ── Historial ────────────────────────────────────────────
			if len(selected_rows) > len(current_year_rows):
				ctk.CTkFrame(info_scroll, height=1, fg_color="#E0E3E8").pack(fill="x", padx=14, pady=(10, 6))
				ctk.CTkLabel(
					info_scroll, text="Historial anterior",
					font=("Arial", 11, "bold"), text_color="#6D7480",
				).pack(anchor="w", padx=14, pady=(0, 4))
				for row in selected_rows:
					if str(row.get("year", "")) == str(current_year):
						continue
					sv = self._coerce_score(row.get("score"))
					score_txt = f"{sv:.1f}%" if sv is not None else "--"
					medal = self._score_medal(row)
					medal_txt = medal["title"] if medal["key"] else "Sin medalla"
					period = f"{row.get('quarter', '--')} {row.get('year', '--')}"
					sentv = str(row.get("sent_at") or "").strip()
					if not sentv:
						st, st_color = "Sin enviar", "#6D7480"
					elif _is_conf(row):
						st, st_color = "✓ Confirmado", "#1E9E5F"
					else:
						st, st_color = "Pendiente", "#856404"
					row_f = ctk.CTkFrame(info_scroll, fg_color="transparent")
					row_f.pack(fill="x", padx=14, pady=2)
					row_f.grid_columnconfigure(0, weight=1)
					left_row = ctk.CTkFrame(row_f, fg_color="transparent")
					left_row.grid(row=0, column=0, sticky="w")
					if sv is not None and sv <= 80:
						alert_small = self.alert_images.get("small")
						if alert_small is not None:
							ctk.CTkLabel(left_row, text="", image=alert_small).pack(side="left", padx=(0, 4))
						else:
							ctk.CTkLabel(left_row, text="!", font=("Arial", 10, "bold"), text_color="#C0392B").pack(side="left", padx=(0, 4))
					ctk.CTkLabel(left_row, text=f"{period}  {score_txt}  |  {medal_txt}", font=("Arial", 11), text_color="#6D7480").pack(side="left")
					ctk.CTkLabel(row_f, text=st, font=("Arial", 10), text_color=st_color).grid(row=0, column=1, sticky="e")

			# ── Update confirm button state ───────────────────────────
			if confirm_button is not None:
				has_pending = any(
					str(r.get("sent_at") or "").strip() and not _is_conf(r)
					for r in selected_rows
				)
				confirm_button.configure(state="normal" if has_pending else "disabled")

		def _confirm_selected_norm() -> None:
			nonlocal detail_scores
			selected_rows = _scores_for_selected_norm()
			pending_ids = [
				str(item.get("id", "")).strip()
				for item in selected_rows
				if str(item.get("id", "")).strip() and not str(item.get("confirmed_at", "")).strip()
			]
			if not pending_ids:
				messagebox.showinfo("Trimestral", "No hay calificaciones pendientes por confirmar para esta norma.", parent=dialog)
				return

			if not messagebox.askyesno("Trimestral", "Confirmas que ya revisaste estas calificaciones?", parent=dialog):
				return

			confirmed_count = self.controller.confirm_trimestral_scores(inspector_name, pending_ids)
			if confirmed_count <= 0:
				messagebox.showinfo("Trimestral", "No se pudieron confirmar las calificaciones seleccionadas.", parent=dialog)
				return

			messagebox.showinfo("Trimestral", f"Se confirmaron {confirmed_count} calificaciones.", parent=dialog)
			detail_scores = self.controller.list_trimestral_scores(inspector_name=inspector_name)
			self.refresh()
			_render_detail()

		norm_selector.configure(command=_render_detail)
		curve_canvas.bind("<Configure>", _render_detail)

		actions = ctk.CTkFrame(wrapper, fg_color="transparent")
		actions.grid(row=2, column=1, padx=(10, 18), pady=(0, 14), sticky="e")
		if self.can_edit:
			ctk.CTkButton(
				actions,
				text="Cerrar",
				fg_color=self.style["fondo"],
				text_color=self.style["texto_oscuro"],
				hover_color="#E9ECEF",
				command=dialog.destroy,
			).grid(row=0, column=0, sticky="e")
		else:
			actions.grid_columnconfigure(0, weight=1)
			actions.grid_columnconfigure(1, weight=1)
			confirm_button = ctk.CTkButton(
				actions,
				text="Confirmado",
				fg_color=self.style["exito"],
				hover_color="#0B7A4D",
				command=_confirm_selected_norm,
			)
			confirm_button.grid(row=0, column=0, padx=(0, 6), sticky="ew")
			ctk.CTkButton(
				actions,
				text="Cerrar",
				fg_color=self.style["fondo"],
				text_color=self.style["texto_oscuro"],
				hover_color="#E9ECEF",
				command=dialog.destroy,
			).grid(row=0, column=1, sticky="ew")

		dialog.after(80, _render_detail)

	@staticmethod
	def _normalize_identity(value: str | None) -> str:
		text = str(value or "").strip().lower()
		if not text:
			return ""
		normalized = unicodedata.normalize("NFKD", text)
		without_accents = "".join(ch for ch in normalized if not unicodedata.combining(ch))
		without_symbols = re.sub(r"[^a-z0-9]+", " ", without_accents)
		return re.sub(r"\s+", " ", without_symbols).strip()

	@staticmethod
	def _norm_key(value: dict | str | None) -> str:
		if isinstance(value, dict):
			raw_value = value.get("norm", "")
		else:
			raw_value = value
		text = str(raw_value or "").strip().upper()
		if not text:
			return "SIN_NORMA"
		match = re.search(r"NOM-\d{3}", text)
		if match:
			return match.group(0)
		return text

	def _norm_display(self, norm_token: str) -> str:
		normalized = self._norm_key(norm_token)
		for item in self.controller.get_catalog_norms():
			token = self._norm_key(item.get("token", ""))
			if token != normalized:
				continue
			nombre = str(item.get("nombre", "")).strip()
			return f"{token} | {nombre}" if nombre else token
		return normalized

	def _norm_full(self, norm_token: str) -> str:
		"""Return the full NOM identifier (e.g. NOM-004-SE-2021) from a short key."""
		normalized = self._norm_key(norm_token)
		for item in self.controller.get_catalog_norms():
			if self._norm_key(item.get("token", "")) == normalized:
				full_nom = str(item.get("nom", "")).strip()
				if full_nom:
					return full_nom
		return normalized

	@staticmethod
	def _quarter_sort_key(value: object) -> int:
		quarter = str(value or "").strip().upper()
		return {"T1": 1, "T2": 2, "T3": 3, "T4": 4}.get(quarter, 9)

	@staticmethod
	def _normalize_label(value) -> str:
		text = str(value or "").strip()
		if not text:
			return "--"
		if len(text) >= 10:
			return text[:10]
		return text

	@staticmethod
	def _coerce_score(value) -> float | None:
		if value in (None, ""):
			return None
		try:
			return round(float(value), 2)
		except (TypeError, ValueError):
			return None

	def _draw_curve_on_canvas(self, canvas: tk.Canvas | None, history: list[dict], empty_message: str) -> None:
		if canvas is None:
			return

		canvas.delete("all")
		width = max(canvas.winfo_width(), 320)
		height = max(canvas.winfo_height(), 180)

		left = 42
		right = width - 20
		top = 18
		bottom = height - 30

		for axis_score in [0, 50, 90, 100]:
			y = bottom - ((axis_score / 100) * (bottom - top))
			axis_color = self.style["advertencia"] if axis_score == 90 else "#D5D8DC"
			canvas.create_line(left, y, right, y, fill=axis_color, dash=(4, 4) if axis_score != 100 else ())
			canvas.create_text(left - 18, y, text=str(axis_score), fill="#6D7480", font=("Arial", 9))

		if not history:
			canvas.create_text(width / 2, height / 2, text=empty_message, fill="#6D7480", font=("Arial", 12))
			return

		usable_width = max(30, right - left)
		step = usable_width / max(len(history) - 1, 1)
		points = []
		for index, item in enumerate(history):
			x = left + index * step
			score = max(0, min(100, float(item.get("score", 0))))
			y = bottom - ((score / 100) * (bottom - top))
			label = self._normalize_label(item.get("label", ""))
			points.append((x, y, label, score, index))

		for point_index in range(len(points) - 1):
			canvas.create_line(
				points[point_index][0],
				points[point_index][1],
				points[point_index + 1][0],
				points[point_index + 1][1],
				fill=self.style["secundario"],
				width=3,
				smooth=True,
			)

		label_step = max(1, len(points) // 7)
		for x, y, label, score, point_index in points:
			fill = self.style["exito"] if score >= 90 else self.style["advertencia"]
			canvas.create_oval(x - 5, y - 5, x + 5, y + 5, fill=fill, outline="")
			canvas.create_text(x, y - 12, text=f"{score:.0f}", fill=self.style["texto_oscuro"], font=("Arial", 9, "bold"))
			if point_index == 0 or point_index == len(points) - 1 or point_index % label_step == 0:
				canvas.create_text(x, bottom + 14, text=label, fill="#6D7480", font=("Arial", 8))

	def clear_form(self, full_reset: bool = False) -> None:
		inspector_name = "" if full_reset else self.inspector_var.get().strip()
		norm_value = "" if full_reset else self._norm_key(self.norm_var.get())

		self.selected_score_id = None
		self.inspector_var.set(inspector_name)
		self.norm_var.set(norm_value)
		self.year_var.set(str(datetime.now().year))
		self.quarter_var.set("T1")
		self.score_var.set("")
		if self.notes_box is not None:
			try:
				if self.notes_box.winfo_exists():
					self.notes_box.delete("1.0", "end")
			except tk.TclError:
				pass
		self._sync_capture_norm_selector()
		self._refresh_capture_history_preview()
		self._update_capture_title()
		self._sync_capture_delete_state()

	def save_score(self) -> None:
		if not self.can_edit:
			return

		parent = self._capture_message_parent()
		inspector_name = self.inspector_var.get().strip()
		if not inspector_name:
			messagebox.showerror("Trimestral", "No hay ejecutivo tecnico seleccionado para guardar.", parent=parent)
			return

		notes = self.notes_box.get("1.0", "end").strip() if self.notes_box is not None else ""
		try:
			self.controller.save_trimestral_score(
				{
					"inspector": inspector_name,
					"norm": self.norm_var.get(),
					"year": self.year_var.get(),
					"quarter": self.quarter_var.get(),
					"score": self.score_var.get(),
					"notes": notes,
				},
				self.selected_score_id,
			)
		except ValueError as error:
			messagebox.showerror("Trimestral", str(error), parent=parent)
			return

		self.score_var.set("")
		if self.notes_box is not None:
			try:
				if self.notes_box.winfo_exists():
					self.notes_box.delete("1.0", "end")
			except tk.TclError:
				pass
		self.selected_score_id = None
		self._sync_capture_delete_state()
		self._refresh_capture_history_preview()
		self.refresh()
		messagebox.showinfo("Trimestral", "Calificacion trimestral guardada correctamente.", parent=parent)

	def delete_score(self) -> None:
		if not self.can_edit or not self.selected_score_id:
			return
		parent = self._capture_message_parent()
		if not messagebox.askyesno("Trimestral", "Deseas eliminar la calificacion seleccionada?", parent=parent):
			return

		self.controller.delete_trimestral_score(self.selected_score_id)
		self.clear_form(full_reset=True)
		self.refresh()
		self._close_capture_dialog(reset_state=False)

	def _on_tree_select(self, _event=None) -> None:
		if self.tree is None:
			return
		selected = self.tree.selection()
		if not selected:
			return
		score = self.row_cache.get(selected[0])
		if score is None:
			return

		self.selected_score_id = selected[0]
		if not self.can_edit:
			return
		self._open_capture_dialog(score=score)
