from __future__ import annotations

import tkinter as tk

import customtkinter as ctk


class DashboardView(ctk.CTkFrame):
    def __init__(self, master, controller, style: dict, fonts: dict) -> None:
        super().__init__(master, fg_color="transparent")
        self.controller = controller
        self.style = style
        self.fonts = fonts
        self.all_inspectors_label = "Todos los inspectores"
        self._resize_job: str | None = None
        self._last_norm_metrics: list[dict] = []
        self._current_history: list[dict] = []
        self._last_bar_size: tuple[int, int] = (0, 0)
        self._last_line_size: tuple[int, int] = (0, 0)

        self.executive_var = ctk.StringVar()
        self.latest_var = ctk.StringVar(value="--")
        self.average_var = ctk.StringVar(value="--")
        self.focus_var = ctk.StringVar(value="Sin seguimiento")
        self.norm_context_var = ctk.StringVar(value="Vista general de cobertura por norma.")

        self.cards_frame: ctk.CTkScrollableFrame | None = None
        self.accredited_box: ctk.CTkTextbox | None = None
        self.visits_box: ctk.CTkTextbox | None = None
        self.bar_canvas: tk.Canvas | None = None
        self.line_canvas: tk.Canvas | None = None
        self.executive_selector: ctk.CTkComboBox | None = None

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        self._build_ui()

    def _build_ui(self) -> None:
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header,
            text="Dashboard operativo",
            font=self.fonts["subtitle"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            header,
            text=(
                "Consulta la distribucion de inspectores acreditados por norma y el seguimiento individual "
                "para identificar desempenos por debajo del 90%."
            ),
            font=self.fonts["small"],
            text_color="#6D7480",
        ).grid(row=1, column=0, sticky="w", pady=(6, 0))

        cards_wrapper = ctk.CTkFrame(self, fg_color=self.style["surface"], corner_radius=22)
        cards_wrapper.grid(row=1, column=0, sticky="ew")
        cards_wrapper.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            cards_wrapper,
            text="Normas acreditadas y cobertura",
            font=self.fonts["label_bold"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=0, column=0, padx=18, pady=(14, 4), sticky="w")

        ctk.CTkLabel(
            cards_wrapper,
            textvariable=self.norm_context_var,
            font=self.fonts["small"],
            text_color="#6D7480",
        ).grid(row=1, column=0, padx=18, pady=(0, 4), sticky="w")

        self.cards_frame = ctk.CTkScrollableFrame(
            cards_wrapper,
            fg_color="transparent",
            height=92,
            orientation="horizontal",
        )
        self.cards_frame.grid(row=2, column=0, padx=12, pady=(0, 8), sticky="ew")

        content = ctk.CTkFrame(self, fg_color="transparent")
        content.grid(row=2, column=0, sticky="nsew", pady=(10, 0))
        content.grid_columnconfigure(0, weight=3)
        content.grid_columnconfigure(1, weight=2)
        content.grid_rowconfigure(1, weight=1)

        overview_panel = ctk.CTkFrame(content, fg_color=self.style["surface"], corner_radius=22)
        overview_panel.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(0, 12))
        overview_panel.grid_columnconfigure(0, weight=1)
        overview_panel.grid_rowconfigure(3, weight=1)

        top_filter = ctk.CTkFrame(overview_panel, fg_color="transparent")
        top_filter.grid(row=0, column=0, padx=18, pady=(16, 0), sticky="ew")
        top_filter.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            top_filter,
            text="Inspector a analizar",
            font=self.fonts["label_bold"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=0, column=0, sticky="w")

        self.executive_selector = ctk.CTkComboBox(
            top_filter,
            variable=self.executive_var,
            values=[self.all_inspectors_label],
            height=38,
            button_color=self.style["primario"],
            dropdown_hover_color=self.style["primario"],
            fg_color="#FFFFFF",
            border_color="#D5D8DC",
            command=lambda _value: self._update_profile(),
        )
        self.executive_selector.grid(row=1, column=0, pady=(8, 0), sticky="ew")

        metrics_row = ctk.CTkFrame(overview_panel, fg_color="transparent")
        metrics_row.grid(row=1, column=0, padx=18, pady=(18, 12), sticky="ew")
        metrics_row.grid_columnconfigure(0, weight=1)
        metrics_row.grid_columnconfigure(1, weight=1)
        metrics_row.grid_columnconfigure(2, weight=1)

        for index, item in enumerate(
            [
                ("Ultimo puntaje", self.latest_var),
                ("Promedio", self.average_var),
                ("Estado", self.focus_var),
            ]
        ):
            card = ctk.CTkFrame(metrics_row, fg_color=self.style["fondo"], corner_radius=18)
            card.grid(row=0, column=index, padx=(0 if index == 0 else 8, 0), sticky="ew")
            ctk.CTkLabel(
                card,
                text=item[0],
                font=self.fonts["small_bold"],
                text_color="#6D7480",
            ).pack(anchor="w", padx=14, pady=(14, 2))
            ctk.CTkLabel(
                card,
                textvariable=item[1],
                font=self.fonts["subtitle"],
                text_color=self.style["texto_oscuro"],
            ).pack(anchor="w", padx=14, pady=(0, 14))

        ctk.CTkLabel(
            overview_panel,
            text="Cobertura por norma",
            font=self.fonts["label_bold"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=2, column=0, padx=18, sticky="w")

        self.bar_canvas = tk.Canvas(
            overview_panel,
            height=185,
            bg=self.style["surface"],
            highlightthickness=0,
        )
        self.bar_canvas.grid(row=3, column=0, padx=18, pady=(10, 18), sticky="nsew")

        detail_panel = ctk.CTkFrame(content, fg_color=self.style["surface"], corner_radius=22)
        detail_panel.grid(row=0, column=1, sticky="nsew")
        detail_panel.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            detail_panel,
            text="Normas del inspector",
            font=self.fonts["label_bold"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=0, column=0, padx=18, pady=(16, 8), sticky="w")

        self.accredited_box = ctk.CTkTextbox(detail_panel, height=155, corner_radius=18)
        self.accredited_box.grid(row=1, column=0, padx=18, pady=(0, 16), sticky="ew")
        self.accredited_box.configure(state="disabled")

        ctk.CTkLabel(
            detail_panel,
            text="Visitas recientes",
            font=self.fonts["label_bold"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=2, column=0, padx=18, pady=(0, 8), sticky="w")

        self.visits_box = ctk.CTkTextbox(detail_panel, height=130, corner_radius=18)
        self.visits_box.grid(row=3, column=0, padx=18, pady=(0, 18), sticky="ew")
        self.visits_box.configure(state="disabled")

        trend_panel = ctk.CTkFrame(content, fg_color=self.style["surface"], corner_radius=22)
        trend_panel.grid(row=1, column=1, sticky="nsew", pady=(12, 0))
        trend_panel.grid_columnconfigure(0, weight=1)
        trend_panel.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            trend_panel,
            text="Curva de desempeno",
            font=self.fonts["label_bold"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=0, column=0, padx=18, pady=(16, 8), sticky="w")

        self.line_canvas = tk.Canvas(
            trend_panel,
            height=180,
            bg=self.style["surface"],
            highlightthickness=0,
        )
        self.line_canvas.grid(row=1, column=0, padx=18, pady=(0, 18), sticky="nsew")

        if self.bar_canvas is not None:
            self.bar_canvas.bind("<Configure>", self._on_chart_resize)
        if self.line_canvas is not None:
            self.line_canvas.bind("<Configure>", self._on_chart_resize)

    def refresh(self) -> None:
        people = [self.all_inspectors_label] + self.controller.get_assignable_inspectors()
        current = self.executive_var.get()
        if current not in people:
            self.executive_var.set(self.all_inspectors_label)
        if self.executive_selector is not None:
            self.executive_selector.configure(values=people)
        self._update_profile()

    def _render_cards(self, metrics: list[dict]) -> None:
        if self.cards_frame is None:
            return

        for child in self.cards_frame.winfo_children():
            child.destroy()

        if not metrics:
            ctk.CTkLabel(
                self.cards_frame,
                text="Sin normas acreditadas para mostrar.",
                font=self.fonts["small"],
                text_color="#6D7480",
            ).grid(row=0, column=0, padx=10, pady=8, sticky="w")
            return

        for index, item in enumerate(metrics):
            card = ctk.CTkFrame(
                self.cards_frame,
                fg_color=self.style["fondo"],
                corner_radius=18,
                width=210,
                height=74,
            )
            card.grid(row=0, column=index, padx=(0 if index == 0 else 8, 0), pady=4, sticky="n")
            card.grid_propagate(False)
            card.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(
                card,
                text=item["token"],
                font=self.fonts["label_bold"],
                text_color=self.style["texto_oscuro"],
            ).grid(row=0, column=0, padx=12, pady=(8, 0), sticky="w")
            ctk.CTkLabel(
                card,
                text=f"{item['count']} inspectores acreditados",
                font=self.fonts["small_bold"],
                text_color=self.style["exito"],
            ).grid(row=1, column=0, padx=12, pady=(1, 0), sticky="w")
            ctk.CTkLabel(
                card,
                text=self._compact_description(item["description"]),
                font=self.fonts["small"],
                text_color="#6D7480",
            ).grid(row=2, column=0, padx=12, pady=(2, 8), sticky="w")

    @staticmethod
    def _compact_description(text: str, limit: int = 42) -> str:
        compact_text = " ".join(str(text).split())
        if len(compact_text) <= limit:
            return compact_text
        return f"{compact_text[: limit - 3].rstrip()}..."

    def _update_profile(self) -> None:
        name = self.executive_var.get().strip()
        if not name or name == self.all_inspectors_label:
            self._last_norm_metrics = self.controller.get_norm_card_metrics()
            self.norm_context_var.set("Vista general: todas las normas acreditadas registradas.")
            self._render_cards(self._last_norm_metrics)
            self._draw_bar_chart(self._last_norm_metrics)
            self.latest_var.set("--")
            self.average_var.set("--")
            self.focus_var.set("Vista global")
            self._set_textbox(self.accredited_box, "Selecciona un inspector para ver sus normas acreditadas.")
            self._set_textbox(self.visits_box, "Selecciona un inspector para revisar sus visitas recientes.")
            self._current_history = []
            self._draw_line_chart([])
            return

        profile = self.controller.get_executive_profile(name)

        if not profile:
            self._last_norm_metrics = self.controller.get_norm_card_metrics()
            self.norm_context_var.set("Vista general: todas las normas acreditadas registradas.")
            self._render_cards(self._last_norm_metrics)
            self._draw_bar_chart(self._last_norm_metrics)
            self.latest_var.set("--")
            self.average_var.set("--")
            self.focus_var.set("Sin seguimiento")
            self._set_textbox(self.accredited_box, "No hay informacion del inspector seleccionado.")
            self._set_textbox(self.visits_box, "No hay visitas registradas.")
            self._current_history = []
            self._draw_line_chart([])
            return

        self._last_norm_metrics = self._metrics_for_inspector(profile.get("accredited_norms", []))
        self.norm_context_var.set(f"Mostrando cobertura para las normas de {name}.")
        self._render_cards(self._last_norm_metrics)
        self._draw_bar_chart(self._last_norm_metrics)

        latest_score = profile.get("latest_score")
        average_score = profile.get("average_score")
        focus_required = profile.get("focus_required")
        self.latest_var.set(f"{latest_score:.1f}%" if latest_score is not None else "--")
        self.average_var.set(f"{average_score:.1f}%" if average_score is not None else "--")
        self.focus_var.set("Mayor enfoque" if focus_required else "Operacion estable")

        accredited_text = "\n".join(
            f"- {token}" for token in profile.get("accredited_norms", [])
        ) or "Sin normas acreditadas en el registro actual."
        self._set_textbox(self.accredited_box, accredited_text)

        recent_visits = profile.get("recent_visits", [])
        visits_text = "\n".join(
            f"- {visit.get('visit_date', '--')} | {visit.get('client', 'Sin cliente')} | {visit.get('status', 'Sin estado')}"
            for visit in recent_visits
        ) or "Sin visitas registradas para este inspector."
        self._set_textbox(self.visits_box, visits_text)
        self._current_history = profile.get("history", [])
        self._draw_line_chart(self._current_history)

    def _metrics_for_inspector(self, accredited_norms: list[str]) -> list[dict]:
        if not accredited_norms:
            return []
        allowed = set(accredited_norms)
        return [item for item in self.controller.get_norm_card_metrics() if item["token"] in allowed]

    def _on_chart_resize(self, _event=None) -> None:
        if not self.winfo_ismapped():
            return

        current_bar_size = (self.bar_canvas.winfo_width(), self.bar_canvas.winfo_height()) if self.bar_canvas is not None else (0, 0)
        current_line_size = (self.line_canvas.winfo_width(), self.line_canvas.winfo_height()) if self.line_canvas is not None else (0, 0)
        if current_bar_size == self._last_bar_size and current_line_size == self._last_line_size:
            return

        self._last_bar_size = current_bar_size
        self._last_line_size = current_line_size
        if self._resize_job is not None:
            self.after_cancel(self._resize_job)
        self._resize_job = self.after(180, self._redraw_charts)

    def _redraw_charts(self) -> None:
        self._resize_job = None
        self._draw_bar_chart(self._last_norm_metrics)
        self._draw_line_chart(self._current_history)

    def _set_textbox(self, textbox: ctk.CTkTextbox | None, value: str) -> None:
        if textbox is None:
            return
        textbox.configure(state="normal")
        textbox.delete("1.0", "end")
        textbox.insert("1.0", value)
        textbox.configure(state="disabled")

    def _draw_bar_chart(self, metrics: list[dict]) -> None:
        if self.bar_canvas is None:
            return

        canvas = self.bar_canvas
        canvas.delete("all")
        width = max(canvas.winfo_width(), 720)
        height = max(canvas.winfo_height(), 185)
        canvas.configure(width=width, height=height)

        if not metrics:
            canvas.create_text(width / 2, height / 2, text="Sin datos de normas", fill="#6D7480", font=("Arial", 12))
            return

        max_count = max(item["count"] for item in metrics) or 1
        left_margin = 34
        bottom = height - 30
        usable_width = width - 54
        spacing = 12
        bar_width = max(24, usable_width / max(len(metrics), 1) - spacing)

        for index, item in enumerate(metrics):
            x0 = left_margin + index * (bar_width + spacing)
            x1 = x0 + bar_width
            bar_height = (item["count"] / max_count) * (height - 64)
            y0 = bottom - bar_height
            canvas.create_rectangle(x0, y0, x1, bottom, fill=self.style["primario"], outline="")
            canvas.create_text((x0 + x1) / 2, y0 - 10, text=str(item["count"]), fill=self.style["texto_oscuro"], font=("Arial", 10, "bold"))
            canvas.create_text((x0 + x1) / 2, bottom + 12, text=item["token"], fill="#6D7480", font=("Arial", 9))

    def _draw_line_chart(self, history: list[dict]) -> None:
        if self.line_canvas is None:
            return

        canvas = self.line_canvas
        canvas.delete("all")
        width = max(canvas.winfo_width(), 420)
        height = max(canvas.winfo_height(), 180)
        canvas.configure(width=width, height=height)

        left = 40
        right = width - 20
        top = 18
        bottom = height - 28

        for axis_score in [0, 50, 90, 100]:
            y = bottom - ((axis_score / 100) * (bottom - top))
            color = self.style["advertencia"] if axis_score == 90 else "#D5D8DC"
            canvas.create_line(left, y, right, y, fill=color, dash=(4, 4) if axis_score != 100 else ())
            canvas.create_text(left - 18, y, text=str(axis_score), fill="#6D7480", font=("Arial", 9))

        if not history:
            canvas.create_text(width / 2, height / 2, text="Sin historial capturado", fill="#6D7480", font=("Arial", 12))
            return

        usable_width = right - left
        step = usable_width / max(len(history) - 1, 1)
        points = []
        for index, item in enumerate(history):
            x = left + index * step
            score = max(0, min(100, float(item.get("score", 0))))
            y = bottom - ((score / 100) * (bottom - top))
            points.append((x, y, item.get("label", ""), score))

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

        for x, y, label, score in points:
            fill = self.style["exito"] if score >= 90 else self.style["advertencia"]
            canvas.create_oval(x - 5, y - 5, x + 5, y + 5, fill=fill, outline="")
            canvas.create_text(x, y - 12, text=f"{score:.0f}", fill=self.style["texto_oscuro"], font=("Arial", 9, "bold"))
            canvas.create_text(x, bottom + 14, text=label[:10], fill="#6D7480", font=("Arial", 8))

