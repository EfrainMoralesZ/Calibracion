from __future__ import annotations

import re
from statistics import mean
import tkinter as tk

import customtkinter as ctk
from PIL import Image

from runtime_paths import resource_path, writable_path


class DashboardView(ctk.CTkFrame):
    def __init__(self, master, controller, style: dict, fonts: dict) -> None:
        super().__init__(master, fg_color=style["fondo"])
        self.controller = controller
        self.style = style
        self.fonts = fonts
        self.all_inspectors_label = "Todos los ejecutivos tecnicos"

        self._resize_job: str | None = None
        self._last_canvas_size: tuple[int, int] = (0, 0)

        self._last_norm_metrics: list[dict] = []
        self._norm_snapshots: list[dict] = []
        self._learning_history: list[dict] = []
        self._recent_visits: list[dict] = []
        self._active_inspector_name: str | None = None
        self.selected_norm_token: str | None = None
        self._cards_signature: tuple | None = None
        self._visits_text_cache: str | None = None
        self._learning_curve_signature: tuple | None = None

        self.executive_var = ctk.StringVar()
        self.focus_var = ctk.StringVar(value="Sin seguimiento")
        self.medals_var = ctk.StringVar(value="O:0 P:0 B:0")
        self.norm_context_var = ctk.StringVar(value="Vista general de cobertura por norma.")
        self.chart_title_var = ctk.StringVar(value="Curva de aprendizaje general")
        self.visits_mode_var = ctk.StringVar(value="Visitas asignadas")

        self.cards_frame: ctk.CTkScrollableFrame | None = None
        self.visits_box: ctk.CTkTextbox | None = None
        self.learning_canvas: tk.Canvas | None = None
        self.executive_selector: ctk.CTkComboBox | None = None
        self.visits_mode_selector: ctk.CTkOptionMenu | None = None
        self.norm_medal_images = self._load_norm_medal_images()

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self._build_ui()

    def _load_norm_medal_images(self) -> dict[str, ctk.CTkImage | None]:
        files = {
            "ORO": "medalla_oro.png",
            "PLATA": "medalla_plata.png",
            "BRONCE": "medalla_bronce.png",
        }
        images: dict[str, ctk.CTkImage | None] = {}
        for key, file_name in files.items():
            image_path = None
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
            images[key] = ctk.CTkImage(light_image=source, dark_image=source, size=(34, 34))
        return images

    def _build_ui(self) -> None:
        cards_wrapper = ctk.CTkFrame(self, fg_color=self.style["surface"], corner_radius=22)
        cards_wrapper.grid(row=0, column=0, sticky="ew")
        cards_wrapper.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            cards_wrapper,
            text="Normas acreditadas",
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
            height=148,
            orientation="horizontal",
        )
        self.cards_frame.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="ew")

        content = ctk.CTkFrame(self, fg_color="transparent")
        content.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        content.grid_columnconfigure(0, weight=3)
        content.grid_columnconfigure(1, weight=2)
        content.grid_rowconfigure(0, weight=1)

        overview_panel = ctk.CTkFrame(content, fg_color=self.style["surface"], corner_radius=22)
        overview_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        overview_panel.grid_columnconfigure(0, weight=1)
        overview_panel.grid_rowconfigure(3, weight=1)

        top_filter = ctk.CTkFrame(overview_panel, fg_color="transparent")
        top_filter.grid(row=0, column=0, padx=18, pady=(16, 0), sticky="ew")
        top_filter.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            top_filter,
            text="Ejecutivo Tecnico a analizar",
            font=self.fonts["label_bold"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=0, column=0, sticky="w")

        filter_controls = ctk.CTkFrame(top_filter, fg_color="transparent")
        filter_controls.grid(row=1, column=0, pady=(8, 0), sticky="w")

        self.executive_selector = ctk.CTkComboBox(
            filter_controls,
            variable=self.executive_var,
            values=[self.all_inspectors_label],
            height=38,
            width=300,
            button_color=self.style["primario"],
            dropdown_hover_color=self.style["primario"],
            fg_color="#FFFFFF",
            border_color="#D5D8DC",
            command=lambda _value: self._update_profile(),
        )
        self.executive_selector.grid(row=0, column=0, sticky="w")

        ctk.CTkButton(
            filter_controls,
            text="Limpiar",
            width=92,
            fg_color=self.style["fondo"],
            text_color=self.style["texto_oscuro"],
            hover_color="#E9ECEF",
            command=self._clear_inspector_filter,
        ).grid(row=0, column=1, padx=(8, 0), sticky="w")

        state_card = ctk.CTkFrame(filter_controls, fg_color=self.style["fondo"], corner_radius=14)
        state_card.grid(row=0, column=2, padx=(12, 0), sticky="w")
        ctk.CTkLabel(
            state_card,
            text="Estado",
            font=self.fonts["small_bold"],
            text_color="#6D7480",
        ).grid(row=0, column=0, padx=(10, 6), pady=8, sticky="w")
        ctk.CTkLabel(
            state_card,
            textvariable=self.focus_var,
            font=self.fonts["small_bold"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=0, column=1, padx=(0, 10), pady=8, sticky="w")
        ctk.CTkLabel(
            state_card,
            text="Medallas",
            font=self.fonts["small_bold"],
            text_color="#6D7480",
        ).grid(row=0, column=2, padx=(8, 6), pady=8, sticky="w")
        ctk.CTkLabel(
            state_card,
            textvariable=self.medals_var,
            font=self.fonts["small_bold"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=0, column=3, padx=(0, 10), pady=8, sticky="w")

        ctk.CTkLabel(
            overview_panel,
            textvariable=self.chart_title_var,
            font=self.fonts["label_bold"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=2, column=0, padx=18, sticky="w")

        self.learning_canvas = tk.Canvas(
            overview_panel,
            height=300,
            bg=self.style["surface"],
            highlightthickness=0,
        )
        self.learning_canvas.grid(row=3, column=0, padx=18, pady=(8, 18), sticky="nsew")

        detail_panel = ctk.CTkFrame(content, fg_color=self.style["surface"], corner_radius=22)
        detail_panel.grid(row=0, column=1, sticky="nsew")
        detail_panel.grid_columnconfigure(0, weight=1)
        detail_panel.grid_rowconfigure(3, weight=1)

        ctk.CTkLabel(
            detail_panel,
            text="Visitas del ejecutivo tecnico",
            font=self.fonts["label_bold"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=0, column=0, padx=18, pady=(16, 8), sticky="w")

        ctk.CTkLabel(
            detail_panel,
            text="Selecciona el tipo de consulta para las visitas del ejecutivo tecnico elegido.",
            font=self.fonts["small"],
            text_color="#6D7480",
            wraplength=320,
            justify="left",
        ).grid(row=1, column=0, padx=18, sticky="w")

        self.visits_mode_selector = ctk.CTkOptionMenu(
            detail_panel,
            values=["Visitas asignadas", "Recientes"],
            variable=self.visits_mode_var,
            fg_color=self.style["fondo"],
            text_color=self.style["texto_oscuro"],
            button_color=self.style["primario"],
            button_hover_color="#D8C220",
            command=lambda _value: self._refresh_visits_panel(),
        )
        self.visits_mode_selector.grid(row=2, column=0, padx=18, pady=(10, 10), sticky="w")

        self.visits_box = ctk.CTkTextbox(detail_panel, corner_radius=18)
        self.visits_box.grid(row=3, column=0, padx=18, pady=(0, 18), sticky="nsew")
        self.visits_box.configure(state="disabled")

        if self.learning_canvas is not None:
            self.learning_canvas.bind("<Configure>", self._on_chart_resize)

    def refresh(self) -> None:
        people = [self.all_inspectors_label] + self.controller.get_assignable_inspectors()
        current = self.executive_var.get()
        if current not in people:
            self.executive_var.set(self.all_inspectors_label)
        if self.executive_selector is not None:
            self.executive_selector.configure(values=people)
        self._update_profile()

    @staticmethod
    def _format_medals(counts: dict[str, int]) -> str:
        return f"O:{counts.get('ORO', 0)} P:{counts.get('PLATA', 0)} B:{counts.get('BRONCE', 0)}"

    def _render_cards(self, metrics: list[dict], inspector_mode: bool) -> None:
        if self.cards_frame is None:
            return

        signature = (
            inspector_mode,
            self.selected_norm_token if inspector_mode else "",
            tuple(
                (
                    str(item.get("token", "Sin norma")).strip() or "Sin norma",
                    str(item.get("full_nom") or item.get("label") or item.get("token", "Sin norma")),
                    self._compact_description(item.get("description", "Catalogo no definido"), limit=102),
                    self._coerce_score(item.get("average_score")),
                    int(item.get("evaluations", 0)),
                    int(item.get("count", 0)),
                )
                for item in metrics
            ),
        )
        if signature == self._cards_signature and self.cards_frame.winfo_children():
            return
        self._cards_signature = signature

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
            token = str(item.get("token", "Sin norma")).strip() or "Sin norma"
            full_nom = str(item.get("full_nom") or item.get("label") or token)
            description = self._compact_description(item.get("description", "Catalogo no definido"), limit=102)
            is_selected = inspector_mode and token == self.selected_norm_token

            if inspector_mode:
                average = item.get("average_score")
                evaluations = int(item.get("evaluations", 0))
                if average is None:
                    card_color = "#F4F6F8"
                    accent_color = "#6D7480"
                    summary_text = f"Sin evaluaciones | {evaluations} captura(s)"
                elif average < 90:
                    card_color = "#FFF2F0"
                    accent_color = self.style["peligro"]
                    summary_text = f"Promedio {average:.1f}% | {evaluations} captura(s)"
                else:
                    card_color = "#EAF7EF"
                    accent_color = self.style["exito"]
                    summary_text = f"Promedio {average:.1f}% | {evaluations} captura(s)"
            else:
                card_color = self.style["fondo"]
                accent_color = self.style["texto_oscuro"]
                summary_text = f"{item.get('count', 0)} ejecutivos tecnicos acreditados"

            card = ctk.CTkFrame(
                self.cards_frame,
                fg_color=card_color,
                corner_radius=18,
                width=290,
                height=124,
                border_width=2 if is_selected else 1,
                border_color=self.style["secundario"] if is_selected else "#DCE0E5",
            )
            card.grid(row=0, column=index, padx=(0 if index == 0 else 8, 0), pady=6, sticky="n")
            card.grid_propagate(False)
            card.grid_columnconfigure(0, weight=1)

            nom_label = ctk.CTkLabel(
                card,
                text=full_nom,
                font=self.fonts["small_bold"],
                text_color=self.style["texto_oscuro"],
                justify="left",
                wraplength=266,
            )
            nom_label.grid(row=0, column=0, padx=12, pady=(8, 0), sticky="w")

            desc_label = ctk.CTkLabel(
                card,
                text=description,
                font=self.fonts["small"],
                text_color="#545B64",
                wraplength=266,
                justify="left",
            )
            desc_label.grid(row=1, column=0, padx=12, pady=(1, 0), sticky="w")

            summary_label = ctk.CTkLabel(
                card,
                text=summary_text,
                font=self.fonts["small_bold"],
                text_color=accent_color,
            )
            summary_label.grid(row=2, column=0, padx=12, pady=(2, 0), sticky="w")

            if inspector_mode:
                hint_label = ctk.CTkLabel(
                    card,
                    text="Doble clic para vista por norma",
                    font=self.fonts["small"],
                    text_color="#6D7480",
                )
                hint_label.grid(row=3, column=0, padx=12, pady=(0, 6), sticky="w")

                for widget in (card, nom_label, desc_label, summary_label, hint_label):
                    widget.bind(
                        "<Double-Button-1>",
                        lambda _event, norm_token=token: self._open_norm_detail(norm_token),
                    )

    @staticmethod
    def _compact_description(text: str, limit: int = 42) -> str:
        compact_text = " ".join(str(text).split())
        if len(compact_text) <= limit:
            return compact_text
        return f"{compact_text[: limit - 3].rstrip()}..."

    def _update_profile(self) -> None:
        selected_name = self.executive_var.get().strip()
        if not selected_name or selected_name == self.all_inspectors_label:
            self._active_inspector_name = None
            self.selected_norm_token = None
            self._norm_snapshots = []
            self._recent_visits = []

            self._last_norm_metrics = self.controller.get_norm_card_metrics()
            self.norm_context_var.set("Vista general: todas las normas acreditadas registradas.")
            self._render_cards(self._last_norm_metrics, inspector_mode=False)

            overview = self.controller.get_overview_metrics()
            self.focus_var.set(self._state_label(overview.get("average_score")))
            global_medals = self.controller.get_trimestral_medals_summary(include_unsent=True)
            self.medals_var.set(self._format_medals(global_medals.get("counts", {})))
            self.chart_title_var.set("Curva de aprendizaje general (todos los ejecutivos tecnicos)")
            self._learning_history = self._build_global_history()
            self._draw_learning_curve()
            self._refresh_visits_panel()
            return

        profile = self.controller.get_executive_profile(selected_name)
        if not profile:
            self._active_inspector_name = selected_name
            self._recent_visits = []
            self._norm_snapshots = []
            self.selected_norm_token = None

            self._last_norm_metrics = self.controller.get_norm_card_metrics()
            self.norm_context_var.set("No hay informacion disponible para el ejecutivo tecnico seleccionado.")
            self._render_cards(self._last_norm_metrics, inspector_mode=False)

            self.focus_var.set("Sin capturas")
            self.medals_var.set("O:0 P:0 B:0")
            self.chart_title_var.set(f"Curva de aprendizaje general - {selected_name}")
            self._learning_history = []
            self._draw_learning_curve()
            self._refresh_visits_panel()
            return

        self._active_inspector_name = selected_name
        self._recent_visits = list(profile.get("recent_visits", []))
        self._norm_snapshots = self._build_norm_snapshots(selected_name, profile.get("accredited_norms", []))

        self.norm_context_var.set(
            f"Doble clic en una card para abrir la vista por norma de {selected_name}."
        )
        self._render_cards(self._norm_snapshots, inspector_mode=True)

        self.focus_var.set(self._state_label(profile.get("average_score")))
        own_medals = self.controller.get_trimestral_medals_summary(inspector_name=selected_name, include_unsent=True)
        self.medals_var.set(self._format_medals(own_medals.get("counts", {})))
        self.chart_title_var.set(f"Curva de aprendizaje general - {selected_name}")
        self._learning_history = self._normalize_history(profile.get("history", []))
        self._draw_learning_curve()
        self._refresh_visits_panel()

    @staticmethod
    def _state_label(average_score) -> str:
        if average_score is None:
            return "Sin capturas"
        score = float(average_score)
        if score >= 90:
            return "Por arriba del 90%"
        if score >= 80:
            return "Estable"
        return "Bajo del promedio"

    def _clear_inspector_filter(self) -> None:
        self.selected_norm_token = None
        self.executive_var.set(self.all_inspectors_label)
        self._update_profile()

    def _refresh_visits_panel(self) -> None:
        mode = self.visits_mode_var.get().strip() or "Visitas asignadas"
        if not self._active_inspector_name:
            if mode == "Visitas asignadas":
                content = "Selecciona un ejecutivo tecnico para ver sus visitas asignadas."
            else:
                content = "Selecciona un ejecutivo tecnico para ver sus visitas recientes."
        else:
            if mode == "Recientes":
                visits = list(self._recent_visits)
                title = f"Visitas recientes de {self._active_inspector_name}"
            else:
                visits = self.controller.list_visits(name=self._active_inspector_name)
                title = f"Visitas asignadas de {self._active_inspector_name}"

            if not visits:
                content = f"{title}\n\nSin visitas registradas."
            else:
                lines = [title, ""]
                for visit in visits:
                    lines.append(
                        "- "
                        f"{visit.get('visit_date', '--')} | "
                        f"{visit.get('client', 'Sin cliente')} | "
                        f"{visit.get('service', 'Sin servicio')} | "
                        f"{visit.get('status', 'Sin estado')}"
                    )
                content = "\n".join(lines)

        if content == self._visits_text_cache:
            return
        self._visits_text_cache = content
        self._set_textbox(self.visits_box, content)

    def _open_norm_detail(self, token: str) -> None:
        if not self._active_inspector_name:
            return

        snapshot = next((item for item in self._norm_snapshots if item.get("token") == token), None)
        if snapshot is None:
            return

        self.selected_norm_token = token
        self._render_cards(self._norm_snapshots, inspector_mode=True)

        dialog = ctk.CTkToplevel(self)
        dialog.title(f"Vista por norma - {snapshot.get('full_nom', token)}")
        dialog.geometry("980x620")
        dialog.minsize(860, 540)
        dialog.configure(fg_color=self.style["fondo"])
        dialog.transient(self.winfo_toplevel())
        dialog.grab_set()

        wrapper = ctk.CTkFrame(dialog, fg_color=self.style["surface"], corner_radius=20)
        wrapper.pack(fill="both", expand=True, padx=18, pady=18)
        wrapper.grid_columnconfigure(0, weight=3)
        wrapper.grid_columnconfigure(1, weight=2)
        wrapper.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(
            wrapper,
            text=f"Vista por norma: {snapshot.get('full_nom', token)}",
            font=self.fonts["label_bold"],
            text_color=self.style["texto_oscuro"],
            justify="left",
            wraplength=640,
        ).grid(row=0, column=0, columnspan=2, padx=18, pady=(16, 10), sticky="w")

        # --- Date filter row ---
        filter_row = ctk.CTkFrame(wrapper, fg_color="transparent")
        filter_row.grid(row=1, column=0, columnspan=2, padx=18, pady=(0, 6), sticky="ew")

        ctk.CTkLabel(filter_row, text="Desde:", font=self.fonts["small"],
                      text_color=self.style["texto_oscuro"]).pack(side="left", padx=(0, 4))
        date_from_var = ctk.CTkEntry(filter_row, width=110, placeholder_text="AAAA-MM-DD",
                                      font=self.fonts["small"])
        date_from_var.pack(side="left", padx=(0, 10))

        ctk.CTkLabel(filter_row, text="Hasta:", font=self.fonts["small"],
                      text_color=self.style["texto_oscuro"]).pack(side="left", padx=(0, 4))
        date_to_var = ctk.CTkEntry(filter_row, width=110, placeholder_text="AAAA-MM-DD",
                                    font=self.fonts["small"])
        date_to_var.pack(side="left", padx=(0, 10))

        full_history = list(snapshot.get("history", []))

        def _apply_date_filter() -> None:
            d_from = date_from_var.get().strip()
            d_to = date_to_var.get().strip()
            filtered = full_history
            if d_from:
                filtered = [p for p in filtered if str(p.get("label", "")) >= d_from]
            if d_to:
                filtered = [p for p in filtered if str(p.get("label", "")) <= d_to]
            self._draw_curve_on_canvas(curve_canvas, filtered, "Sin historial para el rango seleccionado.")

        def _clear_date_filter() -> None:
            date_from_var.delete(0, "end")
            date_to_var.delete(0, "end")
            self._draw_curve_on_canvas(curve_canvas, full_history, "Sin historial para esta norma.")

        ctk.CTkButton(filter_row, text="Filtrar", width=80, font=self.fonts["small"],
                       fg_color=self.style["primario"], command=_apply_date_filter).pack(side="left", padx=(0, 6))
        ctk.CTkButton(filter_row, text="Limpiar", width=80, font=self.fonts["small"],
                       fg_color="#6D7480", command=_clear_date_filter).pack(side="left")

        curve_canvas = tk.Canvas(
            wrapper,
            bg=self.style["surface"],
            highlightthickness=0,
            height=380,
        )
        curve_canvas.grid(row=2, column=0, padx=(18, 10), pady=(0, 16), sticky="nsew")

        info_panel = ctk.CTkScrollableFrame(wrapper, fg_color="#FFFFFF", corner_radius=16)
        info_panel.grid(row=2, column=1, padx=(10, 18), pady=(0, 16), sticky="nsew")
        info_panel.grid_columnconfigure(0, weight=1)

        average_score = snapshot.get("average_score")
        latest_score = snapshot.get("latest_score")
        latest_medal = self.controller.get_trimestral_medal(latest_score)
        norm_medals = {"ORO": 0, "PLATA": 0, "BRONCE": 0}
        for point in list(snapshot.get("history", [])):
            medal_key = str(self.controller.get_trimestral_medal(point.get("score")).get("key", "")).strip().upper()
            if medal_key in norm_medals:
                norm_medals[medal_key] += 1
        medal_emoji = {"ORO": "🥇", "PLATA": "🥈", "BRONCE": "🥉"}
        medal_titles = {"ORO": "Oro", "PLATA": "Plata", "BRONCE": "Bronce"}
        latest_medal_key = str(latest_medal.get("key", "")).strip().upper()
        latest_medal_title = latest_medal.get("title", "Sin medalla")
        latest_medal_display = f"{medal_emoji.get(latest_medal_key, '🏅')} {latest_medal_title}" if latest_medal_key else latest_medal_title

        ctk.CTkLabel(
            info_panel,
            text="RESUMEN DE NORMA",
            font=self.fonts["label_bold"],
            text_color=self.style["texto_oscuro"],
        ).pack(anchor="w", padx=12, pady=(10, 4))

        ctk.CTkLabel(
            info_panel,
            text=f"Ejecutivo tecnico: {self._active_inspector_name}\n"
            f"Norma: {snapshot.get('full_nom', token)}\n"
            f"Descripcion: {snapshot.get('description', 'Catalogo no definido')}",
            font=self.fonts["small"],
            text_color=self.style["texto_oscuro"],
            justify="left",
            anchor="w",
            wraplength=320,
        ).pack(anchor="w", padx=12)

        ctk.CTkLabel(
            info_panel,
            text="DESEMPEÑO",
            font=self.fonts["label_bold"],
            text_color=self.style["texto_oscuro"],
        ).pack(anchor="w", padx=12, pady=(10, 4))

        ctk.CTkLabel(
            info_panel,
            text=(
                (f"Promedio general: {average_score:.1f}%\n" if average_score is not None else "Promedio general: Sin evaluaciones\n")
                + (f"Ultima captura: {latest_score:.1f}%\n" if latest_score is not None else "Ultima captura: --\n")
                + f"Rango actual (medalla): {latest_medal_display}\n"
                + f"Estado: {self._state_label(average_score)}\n"
                + f"Evaluaciones registradas: {snapshot.get('evaluations', 0)}"
            ),
            font=self.fonts["small"],
            text_color=self.style["texto_oscuro"],
            justify="left",
            anchor="w",
            wraplength=320,
        ).pack(anchor="w", padx=12)

        ctk.CTkLabel(
            info_panel,
            text="MEDALLAS OBTENIDAS EN ESTA NORMA",
            font=self.fonts["label_bold"],
            text_color=self.style["texto_oscuro"],
        ).pack(anchor="w", padx=12, pady=(10, 4))

        earned_medals = [(key, norm_medals.get(key, 0)) for key in ["ORO", "PLATA", "BRONCE"] if norm_medals.get(key, 0) > 0]
        if not earned_medals:
            ctk.CTkLabel(
                info_panel,
                text="Sin medallas registradas en esta norma.",
                font=self.fonts["small"],
                text_color="#6D7480",
                justify="left",
                anchor="w",
            ).pack(anchor="w", padx=12, pady=(0, 8))
        else:
            for key, count in earned_medals:
                row = ctk.CTkFrame(info_panel, fg_color="transparent")
                row.pack(anchor="w", padx=12, pady=(2, 8))
                medal_img = self.norm_medal_images.get(key)
                if medal_img is not None:
                    ctk.CTkLabel(row, text="", image=medal_img).pack(side="left", padx=(0, 10))
                else:
                    ctk.CTkLabel(row, text=medal_emoji.get(key, "🏅"), font=("Inter", 20, "bold")).pack(side="left", padx=(0, 10))
                ctk.CTkLabel(
                    row,
                    text=f"{medal_titles.get(key, key)} x{count}",
                    font=("Inter", 16, "bold"),
                    text_color=self.style["texto_oscuro"],
                ).pack(side="left")

        def _redraw_norm_curve(_event=None) -> None:
            self._draw_curve_on_canvas(
                curve_canvas,
                list(snapshot.get("history", [])),
                "Sin historial para esta norma.",
            )

        curve_canvas.bind("<Configure>", _redraw_norm_curve)
        dialog.after(80, _redraw_norm_curve)

    def _build_norm_snapshots(self, inspector_name: str, accredited_norms: list[str]) -> list[dict]:
        catalog = {item["token"]: item for item in self.controller.get_norm_card_metrics()}
        snapshots: dict[str, dict] = {}

        for token in accredited_norms:
            catalog_item = catalog.get(token, {})
            snapshots[token] = {
                "token": token,
                "full_nom": catalog_item.get("label", token),
                "description": catalog_item.get("description", "Catalogo no definido"),
                "count": catalog_item.get("count", 0),
                "history": [],
                "average_score": None,
                "latest_score": None,
                "evaluations": 0,
            }

        for entry in self.controller.get_history(inspector_name):
            norm_token = self._extract_norm_token(entry.get("selected_norm", ""))
            if not norm_token:
                continue

            if norm_token not in snapshots:
                catalog_item = catalog.get(norm_token, {})
                snapshots[norm_token] = {
                    "token": norm_token,
                    "full_nom": catalog_item.get("label", norm_token),
                    "description": catalog_item.get("description", "Catalogo no definido"),
                    "count": catalog_item.get("count", 0),
                    "history": [],
                    "average_score": None,
                    "latest_score": None,
                    "evaluations": 0,
                }

            score = self._coerce_score(entry.get("score"))
            if score is None:
                continue

            label = self._normalize_label(entry.get("visit_date") or entry.get("saved_at", ""))
            snapshots[norm_token]["history"].append({"label": label, "score": score})

        results: list[dict] = []
        for token, item in snapshots.items():
            history = item.get("history", [])
            scores = [point["score"] for point in history]
            item["evaluations"] = len(scores)
            item["latest_score"] = scores[-1] if scores else None
            item["average_score"] = round(mean(scores), 1) if scores else None
            item["token"] = token
            results.append(item)

        return sorted(results, key=lambda value: self._norm_sort_key(value.get("token", "")))

    def _build_global_history(self) -> list[dict]:
        grouped: dict[str, list[float]] = {}
        for inspector_name in self.controller.get_assignable_inspectors():
            for entry in self.controller.get_history(inspector_name):
                score = self._coerce_score(entry.get("score"))
                if score is None:
                    continue

                label = self._normalize_label(entry.get("visit_date") or entry.get("saved_at", ""))
                if not label:
                    continue
                grouped.setdefault(label, []).append(score)

        history: list[dict] = []
        for label in sorted(grouped.keys()):
            points = grouped[label]
            history.append({"label": label, "score": round(mean(points), 1)})
        return history

    @staticmethod
    def _normalize_label(value) -> str:
        text = str(value or "").strip()
        if not text:
            return "--"
        if len(text) >= 10:
            return text[:10]
        return text

    def _normalize_history(self, history: list[dict]) -> list[dict]:
        normalized: list[dict] = []
        for item in history:
            score = self._coerce_score(item.get("score"))
            if score is None:
                continue
            normalized.append(
                {
                    "label": self._normalize_label(item.get("label", "")),
                    "score": score,
                }
            )
        return normalized

    @staticmethod
    def _extract_norm_token(value: str) -> str | None:
        match = re.search(r"NOM-\d{3}", str(value).upper())
        return match.group(0) if match else None

    @staticmethod
    def _coerce_score(value) -> float | None:
        if value in (None, ""):
            return None
        try:
            return round(float(value), 2)
        except (TypeError, ValueError):
            return None

    @staticmethod
    def _norm_sort_key(token: str) -> tuple[int, str]:
        match = re.search(r"(\d{3})", str(token))
        if match:
            return int(match.group(1)), str(token)
        return 999, str(token)

    def _on_chart_resize(self, _event=None) -> None:
        if not self.winfo_ismapped() or self.learning_canvas is None:
            return

        current_size = (self.learning_canvas.winfo_width(), self.learning_canvas.winfo_height())
        if current_size == self._last_canvas_size:
            return

        self._last_canvas_size = current_size
        if self._resize_job is not None:
            self.after_cancel(self._resize_job)
        self._resize_job = self.after(180, self._redraw_learning_curve)

    def _redraw_learning_curve(self) -> None:
        self._resize_job = None
        self._draw_learning_curve()

    def _draw_learning_curve(self) -> None:
        canvas_size = (0, 0)
        if self.learning_canvas is not None:
            canvas_size = (self.learning_canvas.winfo_width(), self.learning_canvas.winfo_height())

        signature = (
            canvas_size,
            tuple(
                (
                    self._normalize_label(item.get("label", "")),
                    self._coerce_score(item.get("score")),
                )
                for item in self._learning_history
            ),
        )
        if signature == self._learning_curve_signature:
            return
        self._learning_curve_signature = signature
        self._draw_curve_on_canvas(
            self.learning_canvas,
            self._learning_history,
            "Sin historial general disponible.",
        )

    def _set_textbox(self, textbox: ctk.CTkTextbox | None, value: str) -> None:
        if textbox is None:
            return
        textbox.configure(state="normal")
        textbox.delete("1.0", "end")
        textbox.insert("1.0", value)
        textbox.configure(state="disabled")

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
