from __future__ import annotations

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

		self.inspector_var = ctk.StringVar()
		self.year_var = ctk.StringVar(value=str(datetime.now().year))
		self.quarter_var = ctk.StringVar(value="T1")
		self.score_var = ctk.StringVar()

		self.tree: ttk.Treeview | None = None
		self.inspector_selector: ctk.CTkComboBox | None = None
		self.notes_box: ctk.CTkTextbox | None = None
		self.score_entry: ctk.CTkEntry | None = None
		self.cards_frame: ctk.CTkScrollableFrame | None = None
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
			"Asigna calificaciones trimestrales, consulta cards y separa el historial por pestana."
			if self.can_edit
			else "Consulta tus calificaciones trimestrales en cards y en historial de solo lectura."
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
			segmented_button_selected_color=self.style["primario"],
			segmented_button_selected_hover_color="#D8C220",
			segmented_button_unselected_color=self.style["fondo"],
			segmented_button_unselected_hover_color="#E9ECEF",
		)
		self.tabview.grid(row=1, column=0, sticky="nsew")
		self.tabview.add(self.capture_tab_name)
		self.tabview.add(self.history_tab_name)

		self._build_capture_tab(self.tabview.tab(self.capture_tab_name))
		self._build_history_tab(self.tabview.tab(self.history_tab_name))

	def _build_capture_tab(self, tab) -> None:
		if self.can_edit:
			tab.grid_columnconfigure(0, weight=3)
			tab.grid_columnconfigure(1, weight=2)
		else:
			tab.grid_columnconfigure(0, weight=1)
		tab.grid_rowconfigure(0, weight=1)

		cards_wrapper = ctk.CTkFrame(tab, fg_color=self.style["surface"], corner_radius=22)
		cards_wrapper.grid(row=0, column=0, sticky="nsew", padx=(0, 12 if self.can_edit else 0))
		cards_wrapper.grid_columnconfigure(0, weight=1)

		cards_title = "Inspectores y normas acreditadas" if self.can_edit else "Tu card trimestral"
		cards_hint = (
			"Usa Ver calificaciones para abrir blandas/tecnicas y Captura trimestral para llenar el formulario."
			if self.can_edit
			else "Aqui aparece tu nombre y las calificaciones trimestrales que te han asignado."
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

		self.cards_frame = ctk.CTkScrollableFrame(cards_wrapper, fg_color="transparent", height=330)
		self.cards_frame.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="nsew")
		for card_col in range(4):
			self.cards_frame.grid_columnconfigure(card_col, weight=1, uniform="trimestral_cards")

		if not self.can_edit:
			return

		form_panel = ctk.CTkFrame(tab, fg_color=self.style["surface"], corner_radius=22)
		form_panel.grid(row=0, column=1, sticky="nsew")
		form_panel.grid_columnconfigure(0, weight=1)

		ctk.CTkLabel(
			form_panel,
			text="Captura trimestral",
			font=self.fonts["label_bold"],
			text_color=self.style["texto_oscuro"],
		).grid(row=0, column=0, padx=18, pady=(14, 8), sticky="w")

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
		self.score_entry = ctk.CTkEntry(form_panel, textvariable=self.score_var, height=38, border_color="#D5D8DC")
		self._field(form_panel, 4, "Calificacion", self.score_entry)

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

	def _build_history_tab(self, tab) -> None:
		tab.grid_columnconfigure(0, weight=1)
		tab.grid_rowconfigure(0, weight=1)

		table_panel = ctk.CTkFrame(tab, fg_color=self.style["surface"], corner_radius=22)
		table_panel.grid(row=0, column=0, sticky="nsew")
		table_panel.grid_columnconfigure(0, weight=1)
		table_panel.grid_rowconfigure(2, weight=1)

		title_text = "Historial trimestral global" if self.can_edit else "Historial trimestral asignado"
		hint_text = (
			"En esta pestana aparecen todos los ejecutivos con sus calificaciones asignadas."
			if self.can_edit
			else "En esta pestana aparece tu historial trimestral asignado."
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

		columns = ("inspector", "periodo", "calificacion", "evaluator", "updated")
		self.tree = ttk.Treeview(container, columns=columns, show="headings", height=18)
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
		self._render_inspector_cards()
		self._refresh_history_table()

		if self.can_edit and self.inspector_selector is not None:
			self.inspector_selector.configure(values=self.controller.get_assignable_inspectors() or ["Sin ejecutivos"])

	def _refresh_history_table(self) -> None:
		if self.tree is None:
			return

		self.row_cache.clear()
		for item in self.tree.get_children():
			self.tree.delete(item)

		target_inspector = None if self.can_edit else str((self.controller.current_user or {}).get("name", "")).strip()
		for score in self.controller.list_trimestral_scores(inspector_name=target_inspector or None):
			score_id = score.get("id")
			if not score_id:
				continue
			self.row_cache[score_id] = score
			raw_score = self._coerce_score(score.get("score"))
			score_text = f"{raw_score:.1f}%" if raw_score is not None else "--"
			self.tree.insert(
				"",
				"end",
				iid=score_id,
				values=(
					score.get("inspector", "--"),
					f"{score.get('quarter', '--')} {score.get('year', '--')}",
					score_text,
					score.get("evaluator", "--"),
					score.get("updated_at", "--"),
				),
			)

	def _render_inspector_cards(self) -> None:
		if self.cards_frame is None:
			return

		for child in self.cards_frame.winfo_children():
			child.destroy()

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

		for inspector_name, values in scores_by_inspector.items():
			values.sort(
				key=lambda item: (
					int(item.get("year", 0)),
					str(item.get("quarter", "")),
					str(item.get("updated_at", "")),
				),
				reverse=True,
			)

		if not rows:
			ctk.CTkLabel(
				self.cards_frame,
				text="Sin inspectores disponibles para mostrar en Trimestral.",
				font=self.fonts["small"],
				text_color="#6D7480",
			).grid(row=0, column=0, padx=10, pady=8, sticky="w")
			return

		for index, row in enumerate(rows):
			inspector_name = str(row.get("name", "")).strip()
			norms_text = str(row.get("norms_text", "Sin acreditaciones")).strip() or "Sin acreditaciones"
			latest = self.controller.get_latest_evaluation(inspector_name)
			soft_score = self._coerce_score(latest.get("soft_skills_score")) if isinstance(latest, dict) else None
			technical_score = self._coerce_score(latest.get("technical_skills_score")) if isinstance(latest, dict) else None
			assigned_scores = scores_by_inspector.get(inspector_name, [])
			latest_trimestral_text, assigned_preview_text = self._format_assigned_scores(assigned_scores)

			card = ctk.CTkFrame(
				self.cards_frame,
				fg_color="#FFFFFF",
				corner_radius=20,
				border_width=1,
				border_color="#E3E6EA",
				height=260,
			)
			card.grid(row=index // 4, column=index % 4, padx=6, pady=6, sticky="nsew")
			card.grid_propagate(False)
			card.grid_columnconfigure(0, weight=1)

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
				text=norms_text,
				font=self.fonts["small"],
				text_color=self.style["texto_oscuro"],
				justify="left",
				wraplength=240,
			).grid(row=2, column=0, padx=12, pady=(0, 6), sticky="ew")

			soft_text = f"{soft_score:.1f}%" if soft_score is not None else "--"
			technical_text = f"{technical_score:.1f}%" if technical_score is not None else "--"
			ctk.CTkLabel(
				card,
				text=f"Blandas: {soft_text}   |   Tecnicas: {technical_text}",
				font=self.fonts["small_bold"],
				text_color="#6D7480",
			).grid(row=3, column=0, padx=12, pady=(0, 8), sticky="w")

			ctk.CTkLabel(
				card,
				text=latest_trimestral_text,
				font=self.fonts["small_bold"],
				text_color="#6D7480",
			).grid(row=4, column=0, padx=12, pady=(0, 4), sticky="w")

			ctk.CTkLabel(
				card,
				text=assigned_preview_text,
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
					text="Captura trimestral",
					fg_color=self.style["fondo"],
					text_color=self.style["texto_oscuro"],
					hover_color="#E9ECEF",
					command=lambda name=inspector_name: self._open_capture_for_inspector(name),
				).grid(row=0, column=1, padx=(4, 0), sticky="ew")

	def _format_assigned_scores(self, scores: list[dict]) -> tuple[str, str]:
		if not scores:
			return "Trimestral actual: --", "Asignadas: sin calificaciones trimestrales"

		latest = scores[0]
		latest_score = self._coerce_score(latest.get("score"))
		latest_period = f"{latest.get('quarter', '--')} {latest.get('year', '--')}"
		if latest_score is None:
			latest_text = f"Trimestral actual: -- ({latest_period})"
		else:
			latest_text = f"Trimestral actual: {latest_score:.1f}% ({latest_period})"

		preview_values: list[str] = []
		for item in scores[:3]:
			item_score = self._coerce_score(item.get("score"))
			item_period = f"{item.get('quarter', '--')} {item.get('year', '--')}"
			if item_score is None:
				preview_values.append(f"{item_period}: --")
			else:
				preview_values.append(f"{item_period}: {item_score:.1f}%")

		preview_text = "Asignadas: " + ", ".join(preview_values)
		return latest_text, preview_text

	def _open_capture_for_inspector(self, inspector_name: str) -> None:
		if not self.can_edit:
			return

		self.selected_score_id = None
		self.inspector_var.set(inspector_name)
		self.year_var.set(str(datetime.now().year))
		self.quarter_var.set("T1")
		self.score_var.set("")
		if self.notes_box is not None:
			self.notes_box.delete("1.0", "end")

		if self.tabview is not None:
			self.tabview.set(self.capture_tab_name)

		if self.score_entry is not None:
			self.after(60, self.score_entry.focus_set)

	def _open_inspector_detail(self, inspector_name: str) -> None:
		profile = self.controller.get_executive_profile(inspector_name)
		accredited_norms = list(profile.get("accredited_norms", [])) if isinstance(profile, dict) else []
		history_entries = self.controller.get_history(inspector_name)
		skill_points = self._build_skill_points(history_entries)

		dialog = ctk.CTkToplevel(self)
		dialog.title(f"Detalle trimestral - {inspector_name}")
		dialog.geometry("1060x620")
		dialog.minsize(940, 520)
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
		).grid(row=0, column=0, columnspan=2, padx=18, pady=(16, 10), sticky="w")

		curve_canvas = tk.Canvas(
			wrapper,
			bg=self.style["surface"],
			highlightthickness=0,
			height=360,
		)
		curve_canvas.grid(row=1, column=0, padx=(18, 10), pady=(0, 16), sticky="nsew")

		info_box = ctk.CTkTextbox(wrapper, corner_radius=16)
		info_box.grid(row=1, column=1, padx=(10, 18), pady=(0, 16), sticky="nsew")

		soft_scores = [point["soft"] for point in skill_points if point.get("soft") is not None]
		technical_scores = [point["technical"] for point in skill_points if point.get("technical") is not None]
		global_scores = [point["score"] for point in skill_points if point.get("score") is not None]

		latest_soft = soft_scores[-1] if soft_scores else None
		latest_technical = technical_scores[-1] if technical_scores else None
		latest_global = global_scores[-1] if global_scores else None

		avg_soft = round(mean(soft_scores), 1) if soft_scores else None
		avg_technical = round(mean(technical_scores), 1) if technical_scores else None
		avg_global = round(mean(global_scores), 1) if global_scores else None

		norms_line = ", ".join(accredited_norms) if accredited_norms else "Sin acreditaciones"
		lines = [
			"Resumen trimestral",
			"",
			f"Ejecutivo Tecnico: {inspector_name}",
			f"Normas acreditadas: {norms_line}",
			"",
			f"Ultima calif. blandas: {latest_soft:.1f}%" if latest_soft is not None else "Ultima calif. blandas: --",
			f"Ultima calif. tecnicas: {latest_technical:.1f}%" if latest_technical is not None else "Ultima calif. tecnicas: --",
			f"Ultima calif. global: {latest_global:.1f}%" if latest_global is not None else "Ultima calif. global: --",
			"",
			f"Promedio blandas: {avg_soft:.1f}%" if avg_soft is not None else "Promedio blandas: --",
			f"Promedio tecnicas: {avg_technical:.1f}%" if avg_technical is not None else "Promedio tecnicas: --",
			f"Promedio global: {avg_global:.1f}%" if avg_global is not None else "Promedio global: --",
			"",
			"Ultimas evaluaciones:",
		]

		for point in skill_points[-8:]:
			date_label = point.get("label", "--")
			soft_label = f"{point['soft']:.1f}%" if point.get("soft") is not None else "--"
			technical_label = f"{point['technical']:.1f}%" if point.get("technical") is not None else "--"
			global_label = f"{point['score']:.1f}%" if point.get("score") is not None else "--"
			lines.append(f"- {date_label} | Blandas: {soft_label} | Tecnicas: {technical_label} | Global: {global_label}")

		info_box.insert("1.0", "\n".join(lines))
		info_box.configure(state="disabled")

		curve_history = [
			{"label": point.get("label", "--"), "score": point.get("score", 0.0)}
			for point in skill_points
			if point.get("score") is not None
		]

		def _redraw_curve(_event=None) -> None:
			self._draw_curve_on_canvas(curve_canvas, curve_history, "Sin curva disponible para este inspector.")

		curve_canvas.bind("<Configure>", _redraw_curve)
		dialog.after(80, _redraw_curve)

		ctk.CTkButton(
			wrapper,
			text="Cerrar",
			fg_color=self.style["fondo"],
			text_color=self.style["texto_oscuro"],
			hover_color="#E9ECEF",
			command=dialog.destroy,
		).grid(row=2, column=1, padx=(10, 18), pady=(0, 14), sticky="e")

	def _build_skill_points(self, history_entries: list[dict]) -> list[dict]:
		points: list[dict] = []
		for entry in history_entries:
			global_score = self._coerce_score(entry.get("score"))
			soft_score, technical_score = self._extract_skill_scores(entry)
			if global_score is None and soft_score is None and technical_score is None:
				continue

			if global_score is None:
				numeric = [value for value in [soft_score, technical_score] if value is not None]
				global_score = round(mean(numeric), 1) if numeric else None

			points.append(
				{
					"label": self._normalize_label(entry.get("visit_date") or entry.get("saved_at", "")),
					"score": global_score,
					"soft": soft_score,
					"technical": technical_score,
				}
			)
		return points

	def _extract_skill_scores(self, entry: dict) -> tuple[float | None, float | None]:
		soft_score = self._coerce_score(entry.get("soft_skills_score"))
		technical_score = self._coerce_score(entry.get("technical_skills_score"))

		if soft_score is not None and technical_score is not None:
			return soft_score, technical_score

		protocol_answers = entry.get("protocol_answers", [])
		process_answers = entry.get("process_answers", [])

		if soft_score is None:
			soft_score = self._score_from_answers(protocol_answers)
		if technical_score is None:
			technical_score = self._score_from_answers(process_answers)

		return soft_score, technical_score

	def _score_from_answers(self, answers: object) -> float | None:
		if not isinstance(answers, list):
			return None

		conforme = 0
		no_conforme = 0
		for answer in answers:
			if not isinstance(answer, dict):
				continue
			result = str(answer.get("result", "")).strip().lower()
			if result == "conforme":
				conforme += 1
			elif result == "no_conforme":
				no_conforme += 1

		aplicables = conforme + no_conforme
		if aplicables <= 0:
			return None
		return round((conforme / aplicables) * 100.0, 1)

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
		if self.tabview is not None:
			self.tabview.set(self.capture_tab_name)

		self.inspector_var.set(score.get("inspector", ""))
		self.year_var.set(str(score.get("year", "")))
		self.quarter_var.set(score.get("quarter", "T1"))
		self.score_var.set(str(score.get("score", "")))
		if self.notes_box is not None:
			self.notes_box.delete("1.0", "end")
			self.notes_box.insert("1.0", score.get("notes", ""))

