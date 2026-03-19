from __future__ import annotations

import re
from datetime import datetime
from statistics import mean
from tkinter import messagebox, ttk
import tkinter as tk

import customtkinter as ctk


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

		self.inspector_var = ctk.StringVar()
		self.norm_var = ctk.StringVar()
		self.year_var = ctk.StringVar(value=str(datetime.now().year))
		self.quarter_var = ctk.StringVar(value="T1")
		self.score_var = ctk.StringVar()

		self.tree: ttk.Treeview | None = None
		self.notes_box: ctk.CTkTextbox | None = None
		self.score_entry: ctk.CTkEntry | None = None
		self.norm_selector: ctk.CTkComboBox | None = None
		self.cards_frame: ctk.CTkScrollableFrame | None = None
		self.capture_dialog: ctk.CTkToplevel | None = None
		self.capture_title_label: ctk.CTkLabel | None = None
		self.capture_inspector_value_label: ctk.CTkLabel | None = None
		self.capture_delete_button: ctk.CTkButton | None = None
		self.capture_history_box: ctk.CTkTextbox | None = None
		self.cards_pager_frame: ctk.CTkFrame | None = None
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
			"Asigna calificaciones trimestrales por norma, revisa cards y consulta historial por pestana."
			if self.can_edit
			else "Consulta tus calificaciones trimestrales por norma en cards y en historial de solo lectura."
		)
		ctk.CTkLabel(
			header,
			text=subtitle_text,
			font=self.fonts["small"],
			text_color="#6D7480",
		).grid(row=1, column=0, sticky="w", pady=(4, 0))

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

	def _build_capture_tab(self, tab) -> None:
		tab.grid_columnconfigure(0, weight=1)
		tab.grid_rowconfigure(0, weight=1)

		cards_wrapper = ctk.CTkFrame(tab, fg_color=self.style["surface"], corner_radius=22)
		cards_wrapper.grid(row=0, column=0, sticky="nsew", pady=12)
		cards_wrapper.grid_columnconfigure(0, weight=1)
		cards_wrapper.grid_rowconfigure(2, weight=1)

		cards_title = "Inspectores y normas acreditadas" if self.can_edit else "Tu card trimestral"
		cards_hint = (
			"Usa Ver calificaciones para abrir el historial por norma/anio y Captura trimestral para enviar la calificacion."
			" El circulo muestra la ultima calificacion pendiente (ej. T1 9) hasta su confirmacion."
			if self.can_edit
			else "Aqui aparece tu card trimestral por norma y el estado de confirmacion de lectura."
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

		self.cards_frame = ctk.CTkScrollableFrame(cards_wrapper, fg_color="transparent")
		self.cards_frame.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="nsew")
		for card_col in range(3):
			self.cards_frame.grid_columnconfigure(card_col, weight=1, uniform="trimestral_cards")

		self.cards_pager_frame = ctk.CTkFrame(cards_wrapper, fg_color="transparent")
		self.cards_pager_frame.grid(row=3, column=0, padx=18, pady=(0, 14), sticky="ew")

	def _build_history_tab(self, tab) -> None:
		tab.grid_columnconfigure(0, weight=1)
		tab.grid_rowconfigure(0, weight=1)

		table_panel = ctk.CTkFrame(tab, fg_color=self.style["surface"], corner_radius=22)
		table_panel.grid(row=0, column=0, sticky="nsew")
		table_panel.grid_columnconfigure(0, weight=1)
		table_panel.grid_rowconfigure(2, weight=1)

		title_text = "Historial trimestral global" if self.can_edit else "Historial trimestral asignado"
		hint_text = (
			"En esta pestana se muestran las calificaciones trimestrales por norma y su confirmacion."
			if self.can_edit
			else "En esta pestana aparece tu historial trimestral por norma."
		)
		ctk.CTkLabel(
			table_panel,
			text=title_text,
			font=self.fonts["label_bold"],
			text_color=self.style["texto_oscuro"],
		).grid(row=0, column=0, padx=18, pady=(14, 4), sticky="w")
		ctk.CTkLabel(
			table_panel,
			text=hint_text,
			font=self.fonts["small"],
			text_color="#6D7480",
		).grid(row=1, column=0, padx=18, pady=(0, 8), sticky="w")

		container = ctk.CTkFrame(table_panel, fg_color="transparent")
		container.grid(row=2, column=0, padx=18, pady=(0, 18), sticky="nsew")
		container.grid_columnconfigure(0, weight=1)
		container.grid_rowconfigure(0, weight=1)

		columns = ("inspector", "norm", "periodo", "calificacion", "confirmado", "updated")
		self.tree = ttk.Treeview(container, columns=columns, show="headings", height=18)
		self.tree.grid(row=0, column=0, sticky="nsew")
		scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
		scrollbar.grid(row=0, column=1, sticky="ns")
		self.tree.configure(yscrollcommand=scrollbar.set)

		self.tree.heading("inspector", text="Ejecutivo Tecnico")
		self.tree.heading("norm", text="Norma")
		self.tree.heading("periodo", text="Periodo")
		self.tree.heading("calificacion", text="Calificacion")
		self.tree.heading("confirmado", text="Confirmado")
		self.tree.heading("updated", text="Ultima actualizacion")
		self.tree.column("inspector", width=230, anchor="w")
		self.tree.column("norm", width=110, anchor="center")
		self.tree.column("periodo", width=110, anchor="center")
		self.tree.column("calificacion", width=110, anchor="center")
		self.tree.column("confirmado", width=130, anchor="center")
		self.tree.column("updated", width=170, anchor="w")
		self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

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
		self.capture_dialog.geometry("620x700")
		self.capture_dialog.minsize(580, 640)
		self.capture_dialog.configure(fg_color=self.style["fondo"])
		self.capture_dialog.transient(self.winfo_toplevel())
		self.capture_dialog.grab_set()
		self.capture_dialog.protocol("WM_DELETE_WINDOW", self._close_capture_dialog)

		form_panel = ctk.CTkFrame(self.capture_dialog, fg_color=self.style["surface"], corner_radius=22)
		form_panel.pack(fill="both", expand=True, padx=18, pady=18)
		form_panel.grid_columnconfigure(0, weight=1)

		self.capture_title_label = ctk.CTkLabel(
			form_panel,
			text="Nueva captura trimestral",
			font=self.fonts["label_bold"],
			text_color=self.style["texto_oscuro"],
		)
		self.capture_title_label.grid(row=0, column=0, padx=18, pady=(14, 8), sticky="w")

		self.capture_inspector_value_label = ctk.CTkLabel(
			form_panel,
			textvariable=self.inspector_var,
			font=self.fonts["label"],
			text_color=self.style["texto_oscuro"],
			fg_color="#FFFFFF",
			corner_radius=10,
			height=38,
			anchor="w",
		)
		self._field(form_panel, 1, "Ejecutivo Tecnico", self.capture_inspector_value_label)

		self.norm_selector = ctk.CTkComboBox(
			form_panel,
			variable=self.norm_var,
			values=["SIN_NORMA"],
			height=38,
			fg_color="#FFFFFF",
			border_color="#D5D8DC",
			button_color=self.style["primario"],
			dropdown_hover_color=self.style["primario"],
			command=lambda _value: self._refresh_capture_history_preview(),
		)
		self._field(form_panel, 2, "Norma", self.norm_selector)

		self._field(form_panel, 3, "Anio", ctk.CTkEntry(form_panel, textvariable=self.year_var, height=38, border_color="#D5D8DC"))
		self._field(
			form_panel,
			4,
			"Trimestre",
			ctk.CTkComboBox(
				form_panel,
				variable=self.quarter_var,
				values=["T1", "T2", "T3", "T4"],
				height=38,
				fg_color="#FFFFFF",
				border_color="#D5D8DC",
				button_color=self.style["primario"],
				dropdown_hover_color=self.style["primario"],
			),
		)
		self.score_entry = ctk.CTkEntry(form_panel, textvariable=self.score_var, height=38, border_color="#D5D8DC")
		self._field(form_panel, 5, "Calificacion", self.score_entry)

		ctk.CTkLabel(
			form_panel,
			text="Notas",
			font=self.fonts["label"],
			text_color=self.style["texto_oscuro"],
		).grid(row=11, column=0, padx=18, pady=(10, 6), sticky="w")
		self.notes_box = ctk.CTkTextbox(form_panel, height=110, corner_radius=16)
		self.notes_box.grid(row=12, column=0, padx=18, sticky="ew")

		ctk.CTkLabel(
			form_panel,
			text="Historial reciente de la norma",
			font=self.fonts["label"],
			text_color=self.style["texto_oscuro"],
		).grid(row=13, column=0, padx=18, pady=(10, 6), sticky="w")
		self.capture_history_box = ctk.CTkTextbox(form_panel, height=120, corner_radius=16)
		self.capture_history_box.grid(row=14, column=0, padx=18, sticky="ew")
		self.capture_history_box.configure(state="disabled")

		actions = ctk.CTkFrame(form_panel, fg_color="transparent")
		actions.grid(row=15, column=0, padx=18, pady=(14, 16), sticky="ew")
		actions.grid_columnconfigure(0, weight=1)
		actions.grid_columnconfigure(1, weight=1)
		actions.grid_columnconfigure(2, weight=1)
		actions.grid_columnconfigure(3, weight=1)
		ctk.CTkButton(
			actions,
			text="Limpiar",
			fg_color=self.style["fondo"],
			text_color=self.style["texto_oscuro"],
			hover_color="#E9ECEF",
			command=lambda: self.clear_form(full_reset=False),
		).grid(row=0, column=0, padx=(0, 6), sticky="ew")
		ctk.CTkButton(
			actions,
			text="Guardar",
			fg_color=self.style["secundario"],
			hover_color="#1D1D1D",
			command=self.save_score,
		).grid(row=0, column=1, padx=3, sticky="ew")
		self.capture_delete_button = ctk.CTkButton(
			actions,
			text="Eliminar",
			fg_color=self.style["peligro"],
			hover_color="#B43C31",
			command=self.delete_score,
		)
		self.capture_delete_button.grid(row=0, column=2, padx=3, sticky="ew")
		ctk.CTkButton(
			actions,
			text="Cerrar",
			fg_color=self.style["fondo"],
			text_color=self.style["texto_oscuro"],
			hover_color="#E9ECEF",
			command=self._close_capture_dialog,
		).grid(row=0, column=3, padx=(6, 0), sticky="ew")

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
					period = f"{item.get('quarter', '--')} {item.get('year', '--')}"
					updated = str(item.get("updated_at", "--"))
					state = "Confirmado" if str(item.get("confirmed_at", "")).strip() else "Pendiente"
					lines.append(f"- {period} | {score_text} | {updated} | {state}")

		self.capture_history_box.configure(state="normal")
		self.capture_history_box.delete("1.0", "end")
		self.capture_history_box.insert("1.0", "\n".join(lines))
		self.capture_history_box.configure(state="disabled")

	def _open_capture_dialog(self, inspector_name: str | None = None, score: dict | None = None) -> None:
		if not self.can_edit:
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
		self._sync_capture_norm_selector()
		self._refresh_capture_history_preview()
		self._update_capture_title()
		self._sync_capture_delete_state()

	def _refresh_history_table(self) -> None:
		if self.tree is None:
			return

		target_inspector = None if self.can_edit else str((self.controller.current_user or {}).get("name", "")).strip()
		row_cache: dict[str, dict] = {}
		rows_to_render: list[tuple[str, tuple[str, str, str, str, str, str]]] = []
		for score in self.controller.list_trimestral_scores(inspector_name=target_inspector or None):
			score_id = score.get("id")
			if not score_id:
				continue
			row_cache[score_id] = score
			raw_score = self._coerce_score(score.get("score"))
			score_text = f"{raw_score:.1f}%" if raw_score is not None else "--"
			confirmed_at = str(score.get("confirmed_at", "")).strip()
			confirmed_text = f"Si ({confirmed_at[:16]})" if confirmed_at else "Pendiente"
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

	def _render_inspector_cards(self) -> None:
		if self.cards_frame is None:
			return

		rows = self.controller.get_principal_rows()
		current_name = str((self.controller.current_user or {}).get("name", "")).strip()
		if not self.can_edit:
			rows = [row for row in rows if str(row.get("name", "")).strip() == current_name]

		score_rows = self.controller.list_trimestral_scores(inspector_name=None if self.can_edit else current_name)
		scores_by_inspector: dict[str, list[dict]] = {}
		for score in score_rows:
			inspector_name = str(score.get("inspector", "")).strip()
			if not inspector_name:
				continue
			scores_by_inspector.setdefault(inspector_name, []).append(score)

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
			norms_text = str(row.get("norms_text", "Sin acreditaciones")).strip() or "Sin acreditaciones"
			assigned_scores = scores_by_inspector.get(inspector_name, [])
			summary = self._format_assigned_scores(assigned_scores)
			norm_count = len({self._norm_key(item) for item in assigned_scores})
			card_models.append(
				{
					"inspector_name": inspector_name,
					"norms_text": norms_text,
					"skills_text": f"Normas calificadas: {norm_count}",
					"pending_text": summary["pending_text"],
					"califications_hint": summary["califications_hint"],
					"pending_badge": summary["pending_badge"],
				}
			)

		signature = tuple(
			(
				model["inspector_name"],
				model["norms_text"],
				model["skills_text"],
				model["pending_text"],
				model["califications_hint"],
				model["pending_badge"],
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
				text="Sin inspectores disponibles para mostrar en Trimestral.",
				font=self.fonts["small"],
				text_color="#6D7480",
			).grid(row=0, column=0, padx=10, pady=8, sticky="w")
			self._rebuild_cards_pager(0)
			return

		start_index = self.current_cards_page * self.cards_page_size
		page_models = card_models[start_index:start_index + self.cards_page_size]

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
				height=260,
			)
			card.grid(row=0, column=0, sticky="nsew")
			card.grid_propagate(False)
			card.grid_columnconfigure(0, weight=1)

			if model["pending_badge"]:
				ctk.CTkLabel(
					card_host,
					text=model["pending_badge"],
					font=("Arial", 10, "bold"),
					text_color=self.style["texto_oscuro"],
					fg_color=self.style["primario"],
					corner_radius=28,
					width=56,
					height=56,
					justify="center",
				).place(relx=1.0, x=-8, y=-8, anchor="ne")

			ctk.CTkLabel(
				card,
				text=inspector_name,
				font=self.fonts["label_bold"],
				text_color=self.style["texto_oscuro"],
				anchor="w",
				justify="left",
				wraplength=240,
			).grid(row=0, column=0, padx=12, pady=(12, 4), sticky="ew")

			ctk.CTkLabel(
				card,
				text="Normas acreditadas",
				font=self.fonts["small_bold"],
				text_color="#6D7480",
			).grid(row=1, column=0, padx=12, sticky="w")

			ctk.CTkLabel(
				card,
				text=model["norms_text"],
				font=self.fonts["small"],
				text_color=self.style["texto_oscuro"],
				justify="left",
				wraplength=240,
			).grid(row=2, column=0, padx=12, pady=(0, 6), sticky="ew")

			ctk.CTkLabel(
				card,
				text=model["skills_text"],
				font=self.fonts["small_bold"],
				text_color="#6D7480",
			).grid(row=3, column=0, padx=12, pady=(0, 8), sticky="w")

			ctk.CTkLabel(
				card,
				text=model["pending_text"],
				font=self.fonts["small_bold"],
				text_color="#6D7480",
			).grid(row=4, column=0, padx=12, pady=(0, 4), sticky="w")

			ctk.CTkLabel(
				card,
				text=model["califications_hint"],
				font=self.fonts["small"],
				text_color="#6D7480",
				justify="left",
				wraplength=250,
			).grid(row=5, column=0, padx=12, pady=(0, 6), sticky="w")

			actions_row = ctk.CTkFrame(card, fg_color="transparent")
			actions_row.grid(row=6, column=0, padx=12, pady=(0, 12), sticky="ew")
			actions_row.grid_columnconfigure(0, weight=1)
			if self.can_edit:
				actions_row.grid_columnconfigure(1, weight=1)
				actions_row.grid_columnconfigure(2, weight=1)

			ctk.CTkButton(
				actions_row,
				text="Ver calificaciones",
				fg_color=self.style["primario"],
				text_color=self.style["texto_oscuro"],
				hover_color="#D8C220",
				command=lambda name=inspector_name: self._open_inspector_detail(name),
			).grid(row=0, column=0, padx=(0, 4 if self.can_edit else 0), sticky="ew")

			if self.can_edit:
				ctk.CTkButton(
					actions_row,
					text="Preview",
					fg_color=self.style["surface"],
					text_color=self.style["texto_oscuro"],
					hover_color="#E9ECEF",
					border_width=1,
					border_color="#D5D8DC",
					command=lambda name=inspector_name: self._open_preview_popup(name),
				).grid(row=0, column=1, padx=4, sticky="ew")
				ctk.CTkButton(
					actions_row,
					text="Captura",
					fg_color=self.style["fondo"],
					text_color=self.style["texto_oscuro"],
					hover_color="#E9ECEF",
					command=lambda name=inspector_name: self._open_capture_for_inspector(name),
				).grid(row=0, column=2, padx=(4, 0), sticky="ew")

		self._rebuild_cards_pager(len(card_models))

	def _format_assigned_scores(self, scores: list[dict]) -> dict[str, str]:
		if not scores:
			return {
				"pending_text": "Pendientes de enterado: 0 | Confirmadas: 0",
				"califications_hint": "Ver calificaciones muestra el historial por anio y trimestre.",
				"pending_badge": "",
			}

		pending_rows = [item for item in scores if not str(item.get("confirmed_at", "")).strip()]
		confirmed_rows = [item for item in scores if str(item.get("confirmed_at", "")).strip()]
		pending_count = len(pending_rows)
		confirmed_count = len(confirmed_rows)

		badge_text = self._pending_badge_text(pending_rows[0]) if pending_rows else ""
		if pending_rows:
			hint_text = (
				"Calificacion pendiente visible en el circulo."
				" Al confirmar de enterado se mueve a Ver calificaciones."
			)
		else:
			hint_text = "Sin pendientes. Revisa Ver calificaciones para el historico anual/trimestral."

		return {
			"pending_text": f"Pendientes de enterado: {pending_count} | Confirmadas: {confirmed_count}",
			"califications_hint": hint_text,
			"pending_badge": badge_text,
		}

	def _pending_badge_text(self, score_row: dict) -> str:
		quarter = str(score_row.get("quarter", "--")).strip().upper() or "--"
		raw_score = self._coerce_score(score_row.get("score"))
		if raw_score is None:
			score_text = "--"
		elif float(raw_score).is_integer():
			score_text = str(int(raw_score))
		else:
			score_text = f"{raw_score:.1f}"
		return f"{quarter}\n{score_text}"

	def _open_preview_popup(self, inspector_name: str) -> None:
		scores = self.controller.list_trimestral_scores(inspector_name=inspector_name)

		dialog = ctk.CTkToplevel(self)
		dialog.title(f"Preview calificaciones — {inspector_name}")
		dialog.geometry("620x520")
		dialog.minsize(520, 420)
		dialog.configure(fg_color=self.style["fondo"])
		dialog.transient(self.winfo_toplevel())
		dialog.grab_set()

		wrapper = ctk.CTkFrame(dialog, fg_color=self.style["surface"], corner_radius=20)
		wrapper.pack(fill="both", expand=True, padx=18, pady=18)
		wrapper.grid_columnconfigure(0, weight=1)
		wrapper.grid_rowconfigure(1, weight=1)

		ctk.CTkLabel(
			wrapper,
			text=f"Calificaciones asignadas — {inspector_name}",
			font=self.fonts["label_bold"],
			text_color=self.style["texto_oscuro"],
		).grid(row=0, column=0, padx=18, pady=(14, 10), sticky="w")

		info_box = ctk.CTkTextbox(wrapper, corner_radius=16)
		info_box.grid(row=1, column=0, padx=18, pady=(0, 10), sticky="nsew")

		lines: list[str] = []
		if not scores:
			lines = ["Sin calificaciones asignadas para este ejecutivo tecnico."]
		else:
			scores_by_norm: dict[str, list[dict]] = {}
			for score in scores:
				norm_token = self._norm_key(score)
				scores_by_norm.setdefault(norm_token, []).append(score)

			for norm_token in sorted(scores_by_norm):
				norm_scores = sorted(
					scores_by_norm[norm_token],
					key=lambda s: (int(s.get("year", 0)), self._quarter_sort_key(s.get("quarter", ""))),
				)
				lines.append(f"Norma: {norm_token}")
				for score_row in norm_scores:
					raw_score = self._coerce_score(score_row.get("score"))
					score_text = f"{raw_score:.1f}%" if raw_score is not None else "--"
					period = f"{score_row.get('quarter', '--')} {score_row.get('year', '--')}"
					state = "Confirmado" if str(score_row.get("confirmed_at", "")).strip() else "Pendiente"
					lines.append(f"  - {period} | {score_text} | {state}")
				lines.append("")

		info_box.configure(state="normal")
		info_box.delete("1.0", "end")
		info_box.insert("1.0", "\n".join(lines).rstrip())
		info_box.configure(state="disabled")

		ctk.CTkButton(
			wrapper,
			text="Cerrar",
			fg_color=self.style["fondo"],
			text_color=self.style["texto_oscuro"],
			hover_color="#E9ECEF",
			command=dialog.destroy,
		).grid(row=2, column=0, padx=18, pady=(0, 14), sticky="e")

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

		info_box = ctk.CTkTextbox(wrapper, corner_radius=16)
		info_box.grid(row=1, column=1, padx=(10, 18), pady=(0, 16), sticky="nsew")

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

		def _render_detail(_value=None) -> None:
			norm_token = self._norm_key(norm_var.get())
			norm_display = self._norm_display(norm_token)
			selected_rows = _scores_for_selected_norm()

			curve_history = []
			for row in reversed(selected_rows):
				score_value = self._coerce_score(row.get("score"))
				if score_value is None:
					continue
				label = f"{row.get('quarter', '--')} {row.get('year', '--')}"
				curve_history.append({"label": label, "score": score_value})

			self._draw_curve_on_canvas(curve_canvas, curve_history, f"Sin curva disponible para {norm_token}.")

			norms_line = ", ".join(accredited_norms) if accredited_norms else "Sin acreditaciones"
			lines = [
				"Resumen trimestral por norma",
				"",
				f"Ejecutivo Tecnico: {inspector_name}",
				f"Norma seleccionada: {norm_display}",
				f"Normas acreditadas: {norms_line}",
				"",
			]

			if selected_rows:
				numeric_scores = [
					value
					for value in (self._coerce_score(item.get("score")) for item in selected_rows)
					if value is not None
				]
				latest = selected_rows[0]
				latest_score = self._coerce_score(latest.get("score"))
				latest_period = f"{latest.get('quarter', '--')} {latest.get('year', '--')}"
				avg_score = round(mean(numeric_scores), 1) if numeric_scores else None
				best_score = max(numeric_scores) if numeric_scores else None
				worst_score = min(numeric_scores) if numeric_scores else None

				lines.append(
					f"Ultima calificacion: {latest_score:.1f}% ({latest_period})"
					if latest_score is not None
					else f"Ultima calificacion: -- ({latest_period})"
				)
				lines.append(f"Promedio: {avg_score:.1f}%" if avg_score is not None else "Promedio: --")
				lines.append(f"Mejor calificacion: {best_score:.1f}%" if best_score is not None else "Mejor calificacion: --")
				lines.append(f"Calificacion mas baja: {worst_score:.1f}%" if worst_score is not None else "Calificacion mas baja: --")
			else:
				lines.append("Sin calificaciones trimestrales registradas para esta norma.")

			current_year = datetime.now().year
			current_year_rows: list[dict] = []
			for item in selected_rows:
				try:
					if int(item.get("year", 0)) == current_year:
						current_year_rows.append(item)
				except (TypeError, ValueError):
					continue

			lines.extend(["", f"Calificaciones {current_year}:"])
			if current_year_rows:
				for row in current_year_rows:
					score_value = self._coerce_score(row.get("score"))
					score_text = f"{score_value:.1f}%" if score_value is not None else "--"
					period = f"{row.get('quarter', '--')} {row.get('year', '--')}"
					confirmed_at = str(row.get("confirmed_at", "")).strip()
					state = "Confirmado" if confirmed_at else "Pendiente de confirmacion"
					lines.append(f"- {period} | {score_text} | {state}")
			else:
				lines.append(f"- Sin calificaciones para {current_year}.")

			lines.extend(["", "Historial acumulado por anio/trimestre:"])
			if selected_rows:
				for row in selected_rows:
					score_value = self._coerce_score(row.get("score"))
					score_text = f"{score_value:.1f}%" if score_value is not None else "--"
					period = f"{row.get('quarter', '--')} {row.get('year', '--')}"
					date_text = str(row.get("updated_at", "--"))
					confirmed_at = str(row.get("confirmed_at", "")).strip()
					state = "Confirmado" if confirmed_at else "Pendiente de confirmacion"
					lines.append(f"- {period} | {score_text} | Fecha: {date_text} | {state}")
			else:
				lines.append("- Sin historial disponible.")

			info_box.configure(state="normal")
			info_box.delete("1.0", "end")
			info_box.insert("1.0", "\n".join(lines))
			info_box.configure(state="disabled")

			if confirm_button is not None:
				has_pending = any(not str(item.get("confirmed_at", "")).strip() for item in selected_rows)
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
