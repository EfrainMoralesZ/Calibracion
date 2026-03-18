from __future__ import annotations

from datetime import datetime
from tkinter import messagebox, ttk

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

		self.inspector_var = ctk.StringVar()
		self.year_var = ctk.StringVar(value=str(datetime.now().year))
		self.quarter_var = ctk.StringVar(value="T1")
		self.score_var = ctk.StringVar()

		self.tree: ttk.Treeview | None = None
		self.inspector_selector: ctk.CTkComboBox | None = None
		self.notes_box: ctk.CTkTextbox | None = None

		self.grid_columnconfigure(0, weight=3)
		self.grid_columnconfigure(1, weight=2)
		self.grid_rowconfigure(1, weight=1)
		self._build_ui()

	def _build_ui(self) -> None:
		header = ctk.CTkFrame(self, fg_color="transparent")
		header.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 14))
		header.grid_columnconfigure(0, weight=1)

		ctk.CTkLabel(
			header,
			text="Trimestral",
			font=self.fonts["subtitle"],
			text_color=self.style["texto_oscuro"],
		).grid(row=0, column=0, sticky="w")
		ctk.CTkLabel(
			header,
			text="Asigna calificaciones trimestrales a ejecutivos y conserva el historial de evaluacion.",
			font=self.fonts["small"],
			text_color="#6D7480",
		).grid(row=1, column=0, sticky="w", pady=(4, 0))

		table_panel = ctk.CTkFrame(self, fg_color=self.style["surface"], corner_radius=22)
		table_panel.grid(row=1, column=0, sticky="nsew", padx=(0, 12))
		table_panel.grid_columnconfigure(0, weight=1)
		table_panel.grid_rowconfigure(1, weight=1)

		ctk.CTkLabel(
			table_panel,
			text="Historial trimestral",
			font=self.fonts["label_bold"],
			text_color=self.style["texto_oscuro"],
		).grid(row=0, column=0, padx=18, pady=(14, 8), sticky="w")

		container = ctk.CTkFrame(table_panel, fg_color="transparent")
		container.grid(row=1, column=0, padx=18, pady=(0, 18), sticky="nsew")
		container.grid_columnconfigure(0, weight=1)
		container.grid_rowconfigure(0, weight=1)

		columns = ("inspector", "periodo", "calificacion", "evaluator", "updated")
		self.tree = ttk.Treeview(container, columns=columns, show="headings", height=16)
		self.tree.grid(row=0, column=0, sticky="nsew")
		scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
		scrollbar.grid(row=0, column=1, sticky="ns")
		self.tree.configure(yscrollcommand=scrollbar.set)

		self.tree.heading("inspector", text="Ejecutivo Tecnico")
		self.tree.heading("periodo", text="Periodo")
		self.tree.heading("calificacion", text="Calificacion")
		self.tree.heading("evaluator", text="Evaluador")
		self.tree.heading("updated", text="Ultima actualizacion")
		self.tree.column("inspector", width=250, anchor="w")
		self.tree.column("periodo", width=120, anchor="center")
		self.tree.column("calificacion", width=120, anchor="center")
		self.tree.column("evaluator", width=180, anchor="w")
		self.tree.column("updated", width=170, anchor="w")
		self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

		form_panel = ctk.CTkFrame(self, fg_color=self.style["surface"], corner_radius=22)
		form_panel.grid(row=1, column=1, sticky="nsew")
		form_panel.grid_columnconfigure(0, weight=1)

		ctk.CTkLabel(
			form_panel,
			text="Captura trimestral",
			font=self.fonts["label_bold"],
			text_color=self.style["texto_oscuro"],
		).grid(row=0, column=0, padx=18, pady=(14, 8), sticky="w")

		if not self.can_edit:
			ctk.CTkLabel(
				form_panel,
				text="Solo el rol admin puede capturar o actualizar calificaciones trimestrales.",
				font=self.fonts["small"],
				text_color="#6D7480",
				wraplength=280,
				justify="left",
			).grid(row=1, column=0, padx=18, pady=(0, 10), sticky="w")
			return

		self.inspector_selector = ctk.CTkComboBox(
			form_panel,
			variable=self.inspector_var,
			values=self.controller.get_assignable_inspectors() or ["Sin ejecutivos"],
			height=38,
			fg_color="#FFFFFF",
			border_color="#D5D8DC",
			button_color=self.style["primario"],
			dropdown_hover_color=self.style["primario"],
		)
		self._field(form_panel, 1, "Ejecutivo Tecnico", self.inspector_selector)
		self._field(form_panel, 2, "Anio", ctk.CTkEntry(form_panel, textvariable=self.year_var, height=38, border_color="#D5D8DC"))
		self._field(
			form_panel,
			3,
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
		self._field(form_panel, 4, "Calificacion", ctk.CTkEntry(form_panel, textvariable=self.score_var, height=38, border_color="#D5D8DC"))

		ctk.CTkLabel(
			form_panel,
			text="Notas",
			font=self.fonts["label"],
			text_color=self.style["texto_oscuro"],
		).grid(row=10, column=0, padx=18, pady=(10, 6), sticky="w")
		self.notes_box = ctk.CTkTextbox(form_panel, height=120, corner_radius=16)
		self.notes_box.grid(row=11, column=0, padx=18, sticky="ew")

		actions = ctk.CTkFrame(form_panel, fg_color="transparent")
		actions.grid(row=12, column=0, padx=18, pady=(14, 16), sticky="ew")
		actions.grid_columnconfigure(0, weight=1)
		actions.grid_columnconfigure(1, weight=1)
		actions.grid_columnconfigure(2, weight=1)
		ctk.CTkButton(
			actions,
			text="Limpiar",
			fg_color=self.style["fondo"],
			text_color=self.style["texto_oscuro"],
			hover_color="#E9ECEF",
			command=self.clear_form,
		).grid(row=0, column=0, padx=(0, 6), sticky="ew")
		ctk.CTkButton(
			actions,
			text="Guardar",
			fg_color=self.style["secundario"],
			hover_color="#1D1D1D",
			command=self.save_score,
		).grid(row=0, column=1, padx=3, sticky="ew")
		ctk.CTkButton(
			actions,
			text="Eliminar",
			fg_color=self.style["peligro"],
			hover_color="#B43C31",
			command=self.delete_score,
		).grid(row=0, column=2, padx=(6, 0), sticky="ew")

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

	def refresh(self) -> None:
		if self.tree is None:
			return

		self.row_cache.clear()
		for item in self.tree.get_children():
			self.tree.delete(item)

		for score in self.controller.list_trimestral_scores():
			score_id = score.get("id")
			if not score_id:
				continue
			self.row_cache[score_id] = score
			self.tree.insert(
				"",
				"end",
				iid=score_id,
				values=(
					score.get("inspector", "--"),
					f"{score.get('quarter', '--')} {score.get('year', '--')}",
					f"{float(score.get('score', 0)):.1f}%",
					score.get("evaluator", "--"),
					score.get("updated_at", "--"),
				),
			)

		if self.can_edit and self.inspector_selector is not None:
			self.inspector_selector.configure(values=self.controller.get_assignable_inspectors() or ["Sin ejecutivos"])

	def clear_form(self) -> None:
		self.selected_score_id = None
		self.inspector_var.set("")
		self.year_var.set(str(datetime.now().year))
		self.quarter_var.set("T1")
		self.score_var.set("")
		if self.notes_box is not None:
			self.notes_box.delete("1.0", "end")

	def save_score(self) -> None:
		if not self.can_edit:
			return

		notes = self.notes_box.get("1.0", "end").strip() if self.notes_box is not None else ""
		try:
			self.controller.save_trimestral_score(
				{
					"inspector": self.inspector_var.get(),
					"year": self.year_var.get(),
					"quarter": self.quarter_var.get(),
					"score": self.score_var.get(),
					"notes": notes,
				},
				self.selected_score_id,
			)
		except ValueError as error:
			messagebox.showerror("Trimestral", str(error), parent=self)
			return

		self.clear_form()
		self.refresh()
		messagebox.showinfo("Trimestral", "Calificacion trimestral guardada correctamente.", parent=self)

	def delete_score(self) -> None:
		if not self.can_edit or not self.selected_score_id:
			return
		if not messagebox.askyesno("Trimestral", "Deseas eliminar la calificacion seleccionada?", parent=self):
			return

		self.controller.delete_trimestral_score(self.selected_score_id)
		self.clear_form()
		self.refresh()

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

		self.inspector_var.set(score.get("inspector", ""))
		self.year_var.set(str(score.get("year", "")))
		self.quarter_var.set(score.get("quarter", "T1"))
		self.score_var.set(str(score.get("score", "")))
		if self.notes_box is not None:
			self.notes_box.delete("1.0", "end")
			self.notes_box.insert("1.0", score.get("notes", ""))

