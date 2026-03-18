from __future__ import annotations

import calendar as pycalendar
from datetime import date, datetime
from tkinter import TclError, messagebox, ttk

import customtkinter as ctk


class CalendarView(ctk.CTkFrame):
    _NO_EXEC2 = "-- Sin segundo ejecutivo --"
    _ALL_EXEC_FILTER = "Todos los ejecutivos"
    _ALL_CLIENT_FILTER = "Todos los clientes"

    def __init__(self, master, controller, style: dict, fonts: dict, can_edit: bool) -> None:
        super().__init__(master, fg_color=style["fondo"])
        self.controller = controller
        self.style = style
        self.fonts = fonts
        self.can_edit = can_edit

        self.selected_visit_id: str | None = None
        self.address_options: list[dict[str, str]] = []
        self.current_month = date.today().replace(day=1)
        self._initial_render_done: bool = False
        self.visible_date_counts: dict[str, int] = {}

        self.inspector_var = ctk.StringVar()
        self.exec2_var = ctk.StringVar()
        self.client_var = ctk.StringVar()
        self.address_var = ctk.StringVar()
        self.service_var = ctk.StringVar(value="Sin servicio")
        self.date_var = ctk.StringVar(value=date.today().strftime("%Y-%m-%d"))
        self.assignment_time_var = ctk.StringVar()
        self.departure_time_var = ctk.StringVar()
        self.status_var = ctk.StringVar(value="Programada")
        self.month_title_var = ctk.StringVar()
        self.search_var = ctk.StringVar()
        self.executive_filter_var = ctk.StringVar(value=self._ALL_EXEC_FILTER)
        self.client_filter_var = ctk.StringVar(value=self._ALL_CLIENT_FILTER)
        self.status_filter_var = ctk.StringVar(value="Todos")

        self.tree: ttk.Treeview | None = None
        self.client_selector: ctk.CTkComboBox | None = None
        self.address_selector: ctk.CTkComboBox | None = None
        self.date_entry: ctk.CTkEntry | None = None
        self.assignment_time_entry: ctk.CTkEntry | None = None
        self.departure_time_entry: ctk.CTkEntry | None = None
        self.status_selector: ctk.CTkComboBox | None = None
        self.notes_box: ctk.CTkTextbox | None = None
        self.calendar_grid_frame: ctk.CTkFrame | None = None
        self.side_panel: ctk.CTkScrollableFrame | None = None
        self.exec1_selector: ctk.CTkComboBox | None = None
        self.exec2_selector: ctk.CTkComboBox | None = None
        self.exec_filter_selector: ctk.CTkComboBox | None = None
        self.client_filter_selector: ctk.CTkComboBox | None = None
        self.status_filter_selector: ctk.CTkComboBox | None = None
        self.norms_display_label: ctk.CTkLabel | None = None
        self.accept_visit_button: ctk.CTkButton | None = None
        self._form_buttons: list[ctk.CTkButton] = []
        self._action_buttons: list[ctk.CTkButton] = []
        self._readonly_notice: ctk.CTkLabel | None = None
        self._viewing_past: bool = False

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self._build_ui()

    def _build_ui(self) -> None:
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        header.grid_columnconfigure(0, weight=1)

        title = "Calendario y visitas" if self.can_edit else "Visitas asignadas"
        subtitle = (
            "Asigna visitas a ejecutivos tecnicos y visualiza el calendario mensual completo con carga por dia."
            if self.can_edit
            else "Consulta tus visitas asignadas en calendario mensual y revisa el detalle de cada sede."
        )
        ctk.CTkLabel(
            header,
            text=title,
            font=self.fonts["subtitle"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(
            header,
            text=subtitle,
            font=self.fonts["small"],
            text_color="#6D7480",
        ).grid(row=1, column=0, sticky="w", pady=(6, 0))

        tabs = ctk.CTkTabview(self, fg_color=self.style["surface"], segmented_button_selected_color=self.style["secundario"])
        tabs.grid(row=1, column=0, sticky="nsew")
        tabs.add("Calendario")
        self._build_calendar_tab(tabs.tab("Calendario"))

        if self.can_edit:
            tabs.add("Visitas asignadas")
            self._build_visits_tab(tabs.tab("Visitas asignadas"))

    def _build_calendar_tab(self, tab) -> None:
        tab.grid_columnconfigure(0, weight=3, minsize=360)
        tab.grid_columnconfigure(1, weight=5)
        tab.grid_rowconfigure(0, weight=1)

        calendar_shell = ctk.CTkFrame(tab, fg_color=self.style["surface"], corner_radius=22)
        calendar_shell.grid(row=0, column=0, sticky="nsew", padx=(0, 12), pady=12)
        calendar_shell.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            calendar_shell,
            text="Agenda operativa",
            font=self.fonts["label_bold"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=0, column=0, padx=18, pady=(16, 8), sticky="w")

        calendar_panel = ctk.CTkFrame(calendar_shell, fg_color=self.style["fondo"], corner_radius=18)
        calendar_panel.grid(row=1, column=0, padx=18, pady=(0, 18), sticky="ew")
        calendar_panel.grid_columnconfigure(0, weight=1)

        month_nav = ctk.CTkFrame(calendar_panel, fg_color="transparent")
        month_nav.grid(row=0, column=0, padx=10, pady=(6, 4), sticky="ew")
        month_nav.grid_columnconfigure(1, weight=1)

        ctk.CTkButton(
            month_nav,
            text="<",
            width=36,
            fg_color=self.style["fondo"],
            text_color=self.style["texto_oscuro"],
            hover_color="#E9ECEF",
            command=self._previous_month,
        ).grid(row=0, column=0, padx=(0, 8))
        ctk.CTkLabel(
            month_nav,
            textvariable=self.month_title_var,
            font=self.fonts["label_bold"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=0, column=1, sticky="w")
        ctk.CTkButton(
            month_nav,
            text="Hoy",
            width=72,
            fg_color=self.style["primario"],
            text_color=self.style["texto_oscuro"],
            hover_color="#D8C220",
            command=self._go_today,
        ).grid(row=0, column=2, padx=8)
        ctk.CTkButton(
            month_nav,
            text=">",
            width=36,
            fg_color=self.style["fondo"],
            text_color=self.style["texto_oscuro"],
            hover_color="#E9ECEF",
            command=self._next_month,
        ).grid(row=0, column=3)

        week_header = ctk.CTkFrame(calendar_panel, fg_color="transparent")
        week_header.grid(row=1, column=0, padx=10, sticky="ew")
        for idx, day_name in enumerate(["Lun", "Mar", "Mie", "Jue", "Vie", "Sab", "Dom"]):
            week_header.grid_columnconfigure(idx, weight=1)
            ctk.CTkLabel(
                week_header,
                text=day_name,
                font=self.fonts["small_bold"],
                text_color="#6D7480",
            ).grid(row=0, column=idx, padx=2, pady=(0, 2), sticky="ew")

        self.calendar_grid_frame = ctk.CTkFrame(calendar_panel, fg_color="transparent")
        self.calendar_grid_frame.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="w")

        # Leyenda de estados
        legend = ctk.CTkFrame(calendar_panel, fg_color="transparent")
        legend.grid(row=3, column=0, padx=10, pady=(6, 10), sticky="w")
        legend.grid_columnconfigure(1, weight=1)

        for col, (label, color) in enumerate([
            ("Asignada", "#FFFFFF"),
            ("Aceptada", "#ECD925"),
            ("Finalizada", "#D1F7D1"),
            ("Cancelada", "#F7D1D1"),
        ]):
            legend_item = ctk.CTkFrame(legend, fg_color=color, corner_radius=6, border_width=1, border_color="#D5D8DC")
            legend_item.grid(row=0, column=col, padx=(0 if col == 0 else 8, 0), sticky="w")
            legend_item.grid_propagate(False)
            legend_item.configure(width=105, height=28)
            text_color = "#282828" if col in [0, 1] else self.style["texto_oscuro"]
            ctk.CTkLabel(
                legend_item,
                text=label,
                font=self.fonts["small"],
                text_color=text_color,
            ).pack(pady=6)

        self.side_panel = ctk.CTkScrollableFrame(tab, fg_color=self.style["surface"], corner_radius=22)
        self.side_panel.grid(row=0, column=1, sticky="nsew", pady=12)
        self.side_panel.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            self.side_panel,
            text="Detalle de visita" if self.can_edit else "Detalle seleccionado",
            font=self.fonts["label_bold"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=0, column=0, padx=18, pady=(16, 4), sticky="w")

        self._readonly_notice = ctk.CTkLabel(
            self.side_panel,
            text="Solo lectura — visita del pasado",
            font=self.fonts["small"],
            text_color="#9AA0A8",
        )
        self._readonly_notice.grid(row=0, column=0, padx=18, pady=(44, 0), sticky="w")
        self._readonly_notice.grid_remove()

        if self.can_edit:
            self._build_form(self.side_panel)
        else:
            message = ctk.CTkLabel(
                self.side_panel,
                text=(
                    "Selecciona una visita para revisar direccion, fecha, servicio y observaciones registradas."
                ),
                font=self.fonts["small"],
                text_color="#6D7480",
                wraplength=280,
                justify="left",
            )
            message.grid(row=1, column=0, padx=18, pady=(0, 12), sticky="w")
            self.notes_box = ctk.CTkTextbox(self.side_panel, height=260, corner_radius=18)
            self.notes_box.grid(row=2, column=0, padx=18, pady=(0, 18), sticky="nsew")
            self.notes_box.configure(state="disabled")

            self.accept_visit_button = ctk.CTkButton(
                self.side_panel,
                text="Aceptar visita",
                command=self._accept_visit_action,
                fg_color="#ECD925",
                text_color="#282828",
                hover_color="#D8C220",
                state="disabled",
            )
            self.accept_visit_button.grid(row=3, column=0, padx=18, pady=(0, 18), sticky="ew")

    def _build_visits_tab(self, tab) -> None:
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)

        filter_row = ctk.CTkFrame(tab, fg_color="transparent")
        filter_row.grid(row=0, column=0, padx=18, pady=(14, 8), sticky="ew")
        filter_row.grid_columnconfigure(0, weight=1)

        search_entry = ctk.CTkEntry(
            filter_row,
            textvariable=self.search_var,
            placeholder_text="Buscar por direccion, servicio, fecha u horario",
            height=38,
            border_color="#D5D8DC",
        )
        search_entry.grid(row=0, column=0, sticky="ew")
        search_entry.bind("<Return>", lambda _event: self.refresh())

        self.exec_filter_selector = ctk.CTkComboBox(
            filter_row,
            variable=self.executive_filter_var,
            values=[self._ALL_EXEC_FILTER],
            width=220,
            height=38,
            fg_color="#FFFFFF",
            border_color="#D5D8DC",
            button_color=self.style["primario"],
            dropdown_hover_color=self.style["primario"],
            command=lambda _value: self.refresh(),
        )
        self.exec_filter_selector.grid(row=0, column=1, padx=(10, 0))

        self.client_filter_selector = ctk.CTkComboBox(
            filter_row,
            variable=self.client_filter_var,
            values=[self._ALL_CLIENT_FILTER],
            width=220,
            height=38,
            fg_color="#FFFFFF",
            border_color="#D5D8DC",
            button_color=self.style["primario"],
            dropdown_hover_color=self.style["primario"],
            command=lambda _value: self.refresh(),
        )
        self.client_filter_selector.grid(row=0, column=2, padx=(10, 0))

        self.status_filter_selector = ctk.CTkComboBox(
            filter_row,
            variable=self.status_filter_var,
            values=["Todos", "Programada", "Realizada", "Reprogramada"],
            width=160,
            height=38,
            fg_color="#FFFFFF",
            border_color="#D5D8DC",
            button_color=self.style["primario"],
            dropdown_hover_color=self.style["primario"],
            command=lambda _value: self.refresh(),
        )
        self.status_filter_selector.grid(row=0, column=3, padx=(10, 0))

        ctk.CTkButton(
            filter_row,
            text="Buscar",
            fg_color=self.style["primario"],
            text_color=self.style["texto_oscuro"],
            hover_color="#D8C220",
            command=self.refresh,
        ).grid(row=0, column=4, padx=(10, 0))

        ctk.CTkButton(
            filter_row,
            text="Limpiar",
            fg_color=self.style["fondo"],
            text_color=self.style["texto_oscuro"],
            hover_color="#E9ECEF",
            command=self._clear_filters,
        ).grid(row=0, column=5, padx=(8, 0))

        if self.can_edit:
            ctk.CTkButton(
                filter_row,
                text="Eliminar asignada",
                fg_color=self.style["peligro"],
                hover_color="#B43C31",
                command=self.delete_selected_visit,
            ).grid(row=0, column=6, padx=(8, 0))

        tree_container = ctk.CTkFrame(tab, fg_color="transparent")
        tree_container.grid(row=1, column=0, padx=18, pady=(0, 18), sticky="nsew")
        tree_container.grid_columnconfigure(0, weight=1)
        tree_container.grid_rowconfigure(0, weight=1)

        columns = ("fecha", "horario", "inspector", "cliente", "servicio", "estado", "direccion")
        self.tree = ttk.Treeview(tree_container, columns=columns, show="headings", height=16)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(tree_container, orient="horizontal", command=self.tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=scrollbar.set, xscrollcommand=scroll_x.set)

        headings = {
            "fecha": "Fecha",
            "horario": "Horario",
            "inspector": "Ejecutivos Tecnicos",
            "cliente": "Cliente",
            "servicio": "Servicio",
            "estado": "Estado",
            "direccion": "Direccion",
        }
        widths = {
            "fecha": 105,
            "horario": 130,
            "inspector": 260,
            "cliente": 190,
            "servicio": 120,
            "estado": 120,
            "direccion": 300,
        }
        for column in columns:
            self.tree.heading(column, text=headings[column])
            self.tree.column(column, width=widths[column], anchor="w")
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

    def _build_form(self, parent) -> None:
        fields = ctk.CTkFrame(parent, fg_color="transparent")
        fields.grid(row=1, column=0, padx=18, pady=(0, 18), sticky="nsew")
        fields.grid_columnconfigure(0, weight=1)

        # ── Ejecutivo Tecnico 1 ─────────────────────────────────────────────
        ctk.CTkLabel(fields, text="Ejecutivo Tecnico 1", font=self.fonts["label"],
                     text_color=self.style["texto_oscuro"]).grid(row=0, column=0, sticky="w", pady=(0, 6))
        all_execs = self.controller.get_assignable_inspectors()
        self.exec1_selector = self._combo(fields, self.inspector_var, all_execs, self._on_exec_change)
        self.exec1_selector.grid(row=1, column=0, sticky="ew")

        # ── Ejecutivo Tecnico 2 (opcional) ───────────────────────────────────
        ctk.CTkLabel(fields, text="Ejecutivo Tecnico 2 (opcional)", font=self.fonts["label"],
                     text_color=self.style["texto_oscuro"]).grid(row=2, column=0, sticky="w", pady=(8, 6))
        exec2_options = [self._NO_EXEC2] + all_execs
        self.exec2_selector = self._combo(fields, self.exec2_var, exec2_options, self._on_exec_change)
        self.exec2_var.set(self._NO_EXEC2)
        self.exec2_selector.grid(row=3, column=0, sticky="ew")

        # ── Normas en comun ─────────────────────────────────────────────────
        ctk.CTkLabel(fields, text="Normas disponibles para la visita", font=self.fonts["label"],
                     text_color=self.style["texto_oscuro"]).grid(row=4, column=0, sticky="w", pady=(8, 4))
        self.norms_display_label = ctk.CTkLabel(
            fields,
            text="Selecciona ejecutivos para ver normas",
            font=self.fonts["small"],
            text_color="#9AA0A8",
            fg_color=self.style["fondo"],
            corner_radius=10,
            wraplength=300,
            justify="left",
            padx=10,
            pady=8,
            anchor="w",
        )
        self.norms_display_label.grid(row=5, column=0, sticky="ew")

        # ── Cliente / Dirección ────────────────────────────────────────────
        self.client_selector = self._combo(fields, self.client_var, self.controller.get_client_names(), self._on_client_change)
        self._label_and_widget_rows(fields, 6, "Cliente", self.client_selector)
        self.address_selector = self._combo(fields, self.address_var, ["Selecciona un cliente"], self._on_address_change)
        self._label_and_widget_rows(fields, 8, "Direccion", self.address_selector)
        self.date_entry = self._entry(fields, self.date_var)
        self._label_and_widget_rows(fields, 10, "Fecha de visita", self.date_entry)
        self.assignment_time_entry = ctk.CTkEntry(fields, textvariable=self.assignment_time_var, height=38, border_color="#D5D8DC", placeholder_text="HH:MM")
        self._label_and_widget_rows(fields, 12, "Hora de asignacion al almacen", self.assignment_time_entry)
        self.departure_time_entry = ctk.CTkEntry(fields, textvariable=self.departure_time_var, height=38, border_color="#D5D8DC", placeholder_text="HH:MM")
        self._label_and_widget_rows(fields, 14, "Hora de salida", self.departure_time_entry)
        self.status_selector = self._combo(fields, self.status_var, ["Programada", "Realizada", "Reprogramada"])
        self._label_and_widget_rows(fields, 16, "Estado", self.status_selector)

        ctk.CTkLabel(fields, text="Servicio", font=self.fonts["label"],
                     text_color=self.style["texto_oscuro"]).grid(row=18, column=0, sticky="w", pady=(8, 6))
        service_label = ctk.CTkLabel(
            fields,
            textvariable=self.service_var,
            font=self.fonts["small_bold"],
            text_color=self.style["exito"],
            fg_color=self.style["fondo"],
            corner_radius=12,
            padx=12,
            pady=8,
        )
        service_label.grid(row=19, column=0, sticky="w")

        ctk.CTkLabel(fields, text="Notas", font=self.fonts["label"],
                     text_color=self.style["texto_oscuro"]).grid(row=20, column=0, sticky="w", pady=(12, 6))
        self.notes_box = ctk.CTkTextbox(fields, height=120, corner_radius=16)
        self.notes_box.grid(row=21, column=0, sticky="ew")

        button_row = ctk.CTkFrame(fields, fg_color="transparent")
        button_row.grid(row=22, column=0, sticky="ew", pady=(16, 0))
        for col in range(5):
            button_row.grid_columnconfigure(col, weight=1)

        self._form_buttons.clear()
        self._action_buttons: list[ctk.CTkButton] = []
        
        if self.can_edit:
            btn_clear = ctk.CTkButton(
                button_row, text="Limpiar", command=self.clear_form,
                fg_color=self.style["fondo"], text_color=self.style["texto_oscuro"], hover_color="#E9ECEF",
            )
            btn_clear.grid(row=0, column=0, padx=(0, 3), sticky="ew")
            self._form_buttons.append(btn_clear)

            btn_save = ctk.CTkButton(
                button_row, text="Guardar", command=self.save_visit,
                fg_color=self.style["secundario"], hover_color="#1D1D1D",
            )
            btn_save.grid(row=0, column=1, padx=3, sticky="ew")
            self._form_buttons.append(btn_save)
            
            btn_cancel = ctk.CTkButton(
                button_row, text="Cancelar visita", command=self._cancel_visit_action,
                fg_color="#F7D1D1", text_color="#D1534E", hover_color="#F0B8B4",
            )
            btn_cancel.grid(row=0, column=2, padx=3, sticky="ew")
            self._form_buttons.append(btn_cancel)
            self._action_buttons.append(btn_cancel)
            
            btn_reassign = ctk.CTkButton(
                button_row, text="Reasignar ejecutivo", command=self._reassign_visit_action,
                fg_color="#FFFFFF", text_color="#282828", hover_color="#F5F5F5",
            )
            btn_reassign.grid(row=0, column=3, padx=3, sticky="ew")
            self._form_buttons.append(btn_reassign)
            self._action_buttons.append(btn_reassign)

            btn_delete = ctk.CTkButton(
                button_row, text="Eliminar", command=self.delete_selected_visit,
                fg_color=self.style["peligro"], hover_color="#B43C31",
            )
            btn_delete.grid(row=0, column=4, padx=(3, 0), sticky="ew")
            self._form_buttons.append(btn_delete)
        else:
            btn_accept = ctk.CTkButton(
                button_row, text="Aceptar visita", command=self._accept_visit_action,
                fg_color="#ECD925", text_color="#282828", hover_color="#D8C220",
            )
            btn_accept.grid(row=0, column=0, padx=(0, 3), sticky="ew")
            self._form_buttons.append(btn_accept)
            self._action_buttons.append(btn_accept)

    def _label_and_widget_rows(self, parent, label_row: int, label_text: str, widget) -> None:
        """Place a label at label_row and its widget at label_row+1."""
        ctk.CTkLabel(parent, text=label_text, font=self.fonts["label"],
                     text_color=self.style["texto_oscuro"]).grid(
            row=label_row, column=0, sticky="w", pady=(8, 6))
        widget.grid(row=label_row + 1, column=0, sticky="ew")

    def _update_norms_display(self) -> None:
        if self.norms_display_label is None:
            return
        exec1 = self.inspector_var.get().strip()
        exec2_raw = self.exec2_var.get().strip()
        exec2 = exec2_raw if exec2_raw and exec2_raw != self._NO_EXEC2 else ""

        if not exec1:
            self.norms_display_label.configure(
                text="Selecciona ejecutivos para ver normas",
                text_color="#9AA0A8",
            )
            return

        norms1 = self.controller.get_accredited_norms(exec1)
        if not exec2:
            text = ("Normas: " + ", ".join(norms1)) if norms1 else "Sin normas acreditadas"
            color = self.style["exito"] if norms1 else "#9AA0A8"
        else:
            set1 = set(norms1)
            set2 = set(self.controller.get_accredited_norms(exec2))
            common = sorted(set1 & set2, key=self.controller._norm_sort_key)
            if common:
                text = "Normas en comun: " + ", ".join(common)
                color = self.style["exito"]
            else:
                text = "Sin normas en comun entre ambos ejecutivos"
                color = self.style["peligro"]

        try:
            self.norms_display_label.configure(text=text, text_color=color)
        except TclError:
            pass

    def _refresh_exec_options(self) -> None:
        """Rebuild exec1 and exec2 dropdown options based on date and current selections."""
        date_iso = self._normalize_date(self.date_var.get())
        busy: set[str] = set()
        if date_iso:
            busy = self.controller.get_busy_executives(date_iso, self.selected_visit_id)

        all_execs = self.controller.get_assignable_inspectors()
        exec1 = self.inspector_var.get().strip()
        exec2_raw = self.exec2_var.get().strip()
        exec2 = exec2_raw if exec2_raw and exec2_raw != self._NO_EXEC2 else ""

        # exec1 options: available (not busy, unless it's their own assigned visit) or the current exec1
        exec1_opts = [e for e in all_execs if e not in busy or e == exec1]
        if exec2:
            exec1_opts = [e for e in exec1_opts if e != exec2]

        # exec2 options: same logic, exclude exec1; always prepend "none" sentinel
        exec2_opts_raw = [e for e in all_execs if e not in busy or e == exec2]
        if exec1:
            exec2_opts_raw = [e for e in exec2_opts_raw if e != exec1]
        exec2_opts = [self._NO_EXEC2] + exec2_opts_raw

        if self.exec1_selector is not None:
            try:
                self.exec1_selector.configure(values=exec1_opts or ["Sin ejecutivos"])
                if exec1 and exec1 not in exec1_opts:
                    self.inspector_var.set(exec1_opts[0] if exec1_opts else "")
            except TclError:
                pass

        if self.exec2_selector is not None:
            try:
                self.exec2_selector.configure(values=exec2_opts)
                if exec2 and exec2 not in exec2_opts:
                    self.exec2_var.set(self._NO_EXEC2)
            except TclError:
                pass

        self._update_norms_display()

    def _on_exec_change(self, _value: str = "") -> None:
        self._refresh_exec_options()

    def _label_and_widget(self, parent, row_index: int, label_text: str, widget) -> None:
        base_row = row_index * 2
        ctk.CTkLabel(parent, text=label_text, font=self.fonts["label"], text_color=self.style["texto_oscuro"]).grid(
            row=base_row,
            column=0,
            sticky="w",
            pady=(0 if row_index == 0 else 8, 6),
        )
        widget.grid(row=base_row + 1, column=0, sticky="ew")

    def _combo(self, parent, variable, values: list[str], command=None):
        return ctk.CTkComboBox(
            parent,
            variable=variable,
            values=values or ["Sin datos"],
            height=38,
            fg_color="#FFFFFF",
            border_color="#D5D8DC",
            button_color=self.style["primario"],
            dropdown_hover_color=self.style["primario"],
            command=command,
        )

    def _entry(self, parent, variable):
        return ctk.CTkEntry(parent, textvariable=variable, height=38, border_color="#D5D8DC")

    def _refresh_visit_filter_options(self, all_visits: list[dict]) -> None:
        executives = sorted(
            {
                inspector_name
                for visit in all_visits
                for inspector_name in visit.get("inspectors", [])
                if str(inspector_name).strip()
            }
        )
        clients = sorted(
            {
                str(visit.get("client", "")).strip()
                for visit in all_visits
                if str(visit.get("client", "")).strip()
            }
        )

        executive_values = [self._ALL_EXEC_FILTER, *executives]
        client_values = [self._ALL_CLIENT_FILTER, *clients]

        if self.exec_filter_selector is not None:
            try:
                self.exec_filter_selector.configure(values=executive_values)
                if self.executive_filter_var.get().strip() not in executive_values:
                    self.executive_filter_var.set(self._ALL_EXEC_FILTER)
            except TclError:
                pass

        if self.client_filter_selector is not None:
            try:
                self.client_filter_selector.configure(values=client_values)
                if self.client_filter_var.get().strip() not in client_values:
                    self.client_filter_var.set(self._ALL_CLIENT_FILTER)
            except TclError:
                pass

    def refresh(self) -> None:
        if self.tree is not None:
            for item in self.tree.get_children():
                self.tree.delete(item)

        all_visits = self.controller.list_visits()
        if self.tree is not None:
            self._refresh_visit_filter_options(all_visits)
        query = self.search_var.get().strip().lower()
        executive_filter = self.executive_filter_var.get().strip() or self._ALL_EXEC_FILTER
        client_filter = self.client_filter_var.get().strip() or self._ALL_CLIENT_FILTER
        status_filter = self.status_filter_var.get().strip() or "Todos"

        filtered_visits: list[dict] = []
        for visit in all_visits:
            if status_filter != "Todos" and str(visit.get("status", "")).strip() != status_filter:
                continue

            if executive_filter != self._ALL_EXEC_FILTER and executive_filter not in visit.get("inspectors", []):
                continue

            if client_filter != self._ALL_CLIENT_FILTER and str(visit.get("client", "")).strip() != client_filter:
                continue

            if query:
                inspectors_text = str(visit.get("inspectors_text", visit.get("inspector", "")))
                searchable = " ".join(
                    [
                        inspectors_text,
                        str(visit.get("client", "")),
                        str(visit.get("service", "")),
                        str(visit.get("address", "")),
                        str(visit.get("visit_date", "")),
                        str(visit.get("assignment_time", "")),
                        str(visit.get("departure_time", "")),
                        str(visit.get("status", "")),
                    ]
                ).lower()
                if query not in searchable:
                    continue

            filtered_visits.append(visit)

        self.visible_date_counts = {}
        for visit in all_visits:
            normalized = self._normalize_date(visit.get("visit_date", ""))
            if normalized:
                self.visible_date_counts[normalized] = self.visible_date_counts.get(normalized, 0) + 1

        if self.tree is not None:
            for visit in filtered_visits:
                self.tree.insert(
                    "",
                    "end",
                    iid=visit["id"],
                    values=(
                        visit.get("visit_date", "--"),
                        visit.get("schedule_text", "--"),
                        visit.get("inspectors_text", visit.get("inspector", "--")),
                        visit.get("client", "--"),
                        visit.get("service", "--"),
                        visit.get("status", "--"),
                        visit.get("address", "--"),
                    ),
                )

        if self.can_edit and self.client_selector is not None and self.client_selector.winfo_exists():
            try:
                self.client_selector.configure(values=self.controller.get_client_names() or ["Sin clientes"])
            except TclError:
                pass
        if self.can_edit and self.address_selector is not None and self.address_selector.winfo_exists():
            try:
                if not self.address_selector.cget("values"):
                    self.address_selector.configure(values=["Selecciona un cliente"])
            except TclError:
                pass

        if self._initial_render_done:
            self._render_month_grid()
        else:
            self.after(150, self._first_render)

    def _first_render(self) -> None:
        self._initial_render_done = True
        self._render_month_grid()

    def _render_month_grid(self) -> None:
        if self.calendar_grid_frame is None:
            return

        for child in self.calendar_grid_frame.winfo_children():
            child.destroy()

        self.month_title_var.set(self.current_month.strftime("%B %Y").title())

        month_matrix = pycalendar.monthcalendar(self.current_month.year, self.current_month.month)
        while len(month_matrix) < 6:
            month_matrix.append([0, 0, 0, 0, 0, 0, 0])

        selected_normalized = self._normalize_date(self.date_var.get())
        today_iso = date.today().strftime("%Y-%m-%d")
        
        # Build map of visit statuses by date for coloring
        visits_by_date = {}
        for visit in self.controller.list_visits():
            v_date = self._normalize_date(visit.get("visit_date", ""))
            if v_date:
                if v_date not in visits_by_date:
                    visits_by_date[v_date] = []
                visits_by_date[v_date].append(visit)

        for row in range(6):
            self.calendar_grid_frame.grid_rowconfigure(row, weight=1)
            for col in range(7):
                self.calendar_grid_frame.grid_columnconfigure(col, weight=1)
                day = month_matrix[row][col]
                if day == 0:
                    cell = ctk.CTkButton(
                        self.calendar_grid_frame,
                        text="",
                        height=28,
                        fg_color="#EEF0F2",
                        hover=False,
                        state="disabled",
                    )
                    cell.grid(row=row, column=col, padx=2, pady=1, sticky="ew")
                    continue

                day_iso = date(self.current_month.year, self.current_month.month, day).strftime("%Y-%m-%d")
                count = self.visible_date_counts.get(day_iso, 0)
                label = f"{day}\n{count} visita(s)" if count else str(day)
                is_past_day = day_iso < today_iso
                has_past_visits = is_past_day and count > 0
                is_sunday = (col == 6)
                
                # Determine color based on acceptance status of visits on this date
                day_visits = visits_by_date.get(day_iso, [])
                has_cancelled = any(v.get("acceptance_status") == "cancelada" for v in day_visits)
                has_finalizada = any(v.get("acceptance_status") == "finalizada" for v in day_visits)
                has_aceptada = any(v.get("acceptance_status") == "aceptada" for v in day_visits)
                has_assigned = any(v.get("acceptance_status") == "asignada" for v in day_visits)

                if (is_past_day and not has_past_visits) or (is_sunday and not count):
                    fg_color = "#EEF0F2"
                    text_color = "#9AA0A8"
                    state = "disabled"
                    hover_color = "#EEF0F2"
                elif has_cancelled:
                    fg_color = "#F7D1D1"
                    text_color = "#D1534E"
                    state = "normal"
                    hover_color = "#F0B8B4"
                elif has_finalizada:
                    fg_color = "#D1F7D1"
                    text_color = "#00A84F"
                    state = "normal"
                    hover_color = "#A8E0A8"
                elif has_aceptada:
                    fg_color = "#ECD925"
                    text_color = "#282828"
                    state = "normal"
                    hover_color = "#D8C220"
                elif has_assigned:
                    fg_color = "#FFFFFF"
                    text_color = "#282828"
                    state = "normal"
                    hover_color = "#F5F5F5"
                elif is_past_day and has_past_visits:
                    fg_color = "#E8EEF8"
                    text_color = "#4A6FA5"
                    state = "normal"
                    hover_color = "#D6E4F0"
                elif selected_normalized == day_iso:
                    fg_color = self.style["primario"]
                    text_color = self.style["texto_oscuro"]
                    state = "normal"
                    hover_color = "#E9ECEF"
                elif count:
                    fg_color = "#EAF7EF"
                    text_color = self.style["texto_oscuro"]
                    state = "normal"
                    hover_color = "#E9ECEF"
                else:
                    fg_color = "#FFFFFF"
                    text_color = self.style["texto_oscuro"]
                    state = "normal"
                    hover_color = "#E9ECEF"

                cell = ctk.CTkButton(
                    self.calendar_grid_frame,
                    text=label,
                    height=28,
                    fg_color=fg_color,
                    text_color=text_color,
                    hover_color=hover_color,
                    state=state,
                    font=self.fonts["small"],
                    command=(lambda iso=day_iso: self._select_calendar_date(iso)) if state == "normal" else None,
                )
                if state == "disabled":
                    cell.configure(hover=False)
                cell.grid(row=row, column=col, padx=2, pady=1, sticky="ew")

    def _select_calendar_date(self, iso_date: str) -> None:
        today_iso = date.today().strftime("%Y-%m-%d")
        is_past = iso_date < today_iso
        self._viewing_past = is_past
        self.date_var.set(iso_date)
        self._render_month_grid()
        self._refresh_exec_options()

        if self.can_edit:
            visits_on_date = [
                v for v in self.controller.list_visits()
                if self._normalize_date(v.get("visit_date", "")) == iso_date
            ]
            if visits_on_date:
                visit = visits_on_date[0]
                self.selected_visit_id = visit["id"]
                inspectors = list(visit.get("inspectors", []))
                # Write to notes_box before locking it
                if self.notes_box is not None:
                    self.notes_box.configure(state="normal")
                    self.notes_box.delete("1.0", "end")
                    self.notes_box.insert("1.0", visit.get("notes", ""))
                self.inspector_var.set(inspectors[0] if inspectors else "")
                self.exec2_var.set(inspectors[1] if len(inspectors) > 1 else self._NO_EXEC2)
                self.client_var.set(visit.get("client", ""))
                self._on_client_change(visit.get("client", ""))
                self.address_var.set(visit.get("address", ""))
                self.service_var.set(visit.get("service", "Sin servicio"))
                self.assignment_time_var.set(visit.get("assignment_time", ""))
                self.departure_time_var.set(visit.get("departure_time", ""))
                self.status_var.set(visit.get("status", "Programada"))
            else:
                self.selected_visit_id = None
        elif not self.can_edit:
            visits_on_date = [
                v for v in self.controller.list_visits()
                if self._normalize_date(v.get("visit_date", "")) == iso_date
            ]
            if visits_on_date:
                self.selected_visit_id = visits_on_date[0].get("id")
                self._show_readonly_visit_details(visits_on_date[0])
            else:
                self.selected_visit_id = None
                self._show_readonly_visit_details(None)

        self._set_form_editable(not is_past)
        # Scroll form panel to top so fields are immediately visible
        if self.side_panel is not None:
            try:
                self.side_panel._parent_canvas.yview_moveto(0)
            except TclError:
                pass

    def _set_form_editable(self, editable: bool) -> None:
        if not self.can_edit:
            if self.notes_box is not None:
                try:
                    self.notes_box.configure(state="disabled")
                except TclError:
                    pass
            return

        state = "normal" if editable else "disabled"
        for widget in [
            self.exec1_selector, self.exec2_selector,
            self.client_selector, self.address_selector,
            self.date_entry, self.assignment_time_entry,
            self.departure_time_entry, self.status_selector,
            self.notes_box,
        ]:
            if widget is not None:
                try:
                    widget.configure(state=state)
                except TclError:
                    pass
        for btn in self._form_buttons:
            try:
                btn.configure(state=state)
            except TclError:
                pass
        # Set action buttons state (only enabled when visit selected)
        action_state = "normal" if self.selected_visit_id else "disabled"
        for btn in getattr(self, "_action_buttons", []):
            try:
                btn.configure(state=action_state)
            except TclError:
                pass
        if self._readonly_notice is not None:
            try:
                if editable:
                    self._readonly_notice.grid_remove()
                else:
                    self._readonly_notice.grid()
            except TclError:
                pass

    def _previous_month(self) -> None:
        month = self.current_month.month - 1
        year = self.current_month.year
        if month == 0:
            month = 12
            year -= 1
        self.current_month = date(year, month, 1)
        self._render_month_grid()

    def _next_month(self) -> None:
        month = self.current_month.month + 1
        year = self.current_month.year
        if month == 13:
            month = 1
            year += 1
        self.current_month = date(year, month, 1)
        self._render_month_grid()

    def _go_today(self) -> None:
        today = date.today()
        self.current_month = today.replace(day=1)
        self._viewing_past = False
        self.date_var.set(today.strftime("%Y-%m-%d"))
        if self.can_edit:
            self._set_form_editable(True)
        self._render_month_grid()

    def clear_form(self) -> None:
        self.selected_visit_id = None
        self._viewing_past = False
        self._set_form_editable(True)
        self.inspector_var.set("")
        self.exec2_var.set(self._NO_EXEC2)
        self.client_var.set("")
        self.address_var.set("")
        self.service_var.set("Sin servicio")
        self.date_var.set(date.today().strftime("%Y-%m-%d"))
        self.assignment_time_var.set("")
        self.departure_time_var.set("")
        self.status_var.set("Programada")
        self.address_options = []
        if self.address_selector is not None:
            self.address_selector.configure(values=["Selecciona un cliente"])
        if self.notes_box is not None:
            self.notes_box.delete("1.0", "end")
        self._go_today()
        self._refresh_exec_options()

    def _collect_executives(self) -> list[str]:
        exec1 = self.inspector_var.get().strip()
        exec2_raw = self.exec2_var.get().strip()
        exec2 = exec2_raw if exec2_raw and exec2_raw != self._NO_EXEC2 else ""

        names: list[str] = []
        if exec1:
            names.append(exec1)
        if exec2 and exec2 != exec1:
            names.append(exec2)

        if not names:
            raise ValueError("Debes seleccionar al menos un ejecutivo tecnico.")

        valid = {value.lower(): value for value in self.controller.get_assignable_inspectors()}
        invalid = [name for name in names if name.lower() not in valid]
        if invalid:
            invalid_text = ", ".join(invalid)
            raise ValueError(f"Los siguientes ejecutivos tecnicos no existen: {invalid_text}")

        return [valid[name.lower()] for name in names]

    def save_visit(self) -> None:
        if not self.can_edit:
            return

        notes = self.notes_box.get("1.0", "end").strip() if self.notes_box is not None else ""
        base_payload = {
            "client": self.client_var.get(),
            "address": self.address_var.get(),
            "service": self.service_var.get(),
            "visit_date": self.date_var.get(),
            "assignment_time": self.assignment_time_var.get(),
            "departure_time": self.departure_time_var.get(),
            "status": self.status_var.get(),
            "notes": notes,
        }

        normalized_date = self._normalize_date(base_payload["visit_date"])
        if not normalized_date:
            messagebox.showerror("Visitas", "La fecha de visita es obligatoria y debe tener formato YYYY-MM-DD.")
            return

        if not self.selected_visit_id and normalized_date < date.today().strftime("%Y-%m-%d"):
            messagebox.showerror("Visitas", "No puedes agendar visitas en fechas anteriores al dia de hoy.")
            return

        base_payload["visit_date"] = normalized_date

        try:
            executives = self._collect_executives()
            if self.selected_visit_id:
                self.controller.save_visit(
                    {
                        **base_payload,
                        "inspectors": executives,
                    },
                    self.selected_visit_id,
                )
            else:
                self.controller.save_visit(
                    {
                        **base_payload,
                        "inspectors": executives,
                    },
                    None,
                )
        except ValueError as error:
            messagebox.showerror("Visitas", str(error))
            return

        self.clear_form()
        self.refresh()
        messagebox.showinfo("Visitas", "La visita fue guardada correctamente.")

    def delete_selected_visit(self) -> None:
        if not self.can_edit:
            return

        if not self.selected_visit_id:
            messagebox.showinfo("Visitas", "Selecciona una visita asignada para eliminarla.")
            return

        if not messagebox.askyesno("Visitas", "Deseas eliminar la visita seleccionada?"):
            return

        self.controller.delete_visit(self.selected_visit_id)
        self.clear_form()
        self.refresh()

    def _cancel_visit_action(self) -> None:
        if not self.can_edit or not self.selected_visit_id:
            messagebox.showwarning("Visitas", "Selecciona una visita para cancelar.")
            return
        
        if not messagebox.askyesno("Visitas", "Deseas cancelar la visita seleccionada?"):
            return
        
        self.controller.cancel_visit(self.selected_visit_id)
        messagebox.showinfo("Visitas", "La visita fue cancelada correctamente.")
        self.clear_form()
        self.refresh()

    def _reassign_visit_action(self) -> None:
        if not self.can_edit or not self.selected_visit_id:
            messagebox.showwarning("Visitas", "Selecciona una visita para reasignar.")
            return
        
        # Show reassignment dialog
        new_inspectors = self.inspector_var.get().strip()
        exec2_raw = self.exec2_var.get().strip()
        if exec2_raw and exec2_raw != self._NO_EXEC2:
            new_inspectors = [new_inspectors, exec2_raw]
        else:
            new_inspectors = [new_inspectors] if new_inspectors else []
        
        if not new_inspectors:
            messagebox.showwarning("Visitas", "Selecciona al menos un ejecutivo para reasignar.")
            return
        
        if not messagebox.askyesno("Visitas", "Deseas reasignar la visita a los ejecutivos seleccionados?"):
            return
        
        try:
            self.controller.reassign_visit(self.selected_visit_id, new_inspectors)
            messagebox.showinfo("Visitas", "La visita fue reasignada correctamente.")
            self.clear_form()
            self.refresh()
        except ValueError as error:
            messagebox.showerror("Visitas", str(error))

    def _accept_visit_action(self) -> None:
        if self.can_edit or not self.selected_visit_id:
            return
        
        if not messagebox.askyesno("Visitas", "Deseas aceptar la visita?"):
            return
        
        self.controller.accept_visit(self.selected_visit_id)
        visit = next((item for item in self.controller.list_visits() if item.get("id") == self.selected_visit_id), None)
        self._show_readonly_visit_details(visit)
        messagebox.showinfo("Visitas", "Visita aceptada correctamente.")
        self._render_month_grid()
        self.refresh()

    def _set_accept_button_state(self, visit: dict | None) -> None:
        if self.accept_visit_button is None:
            return

        if visit is None:
            self.accept_visit_button.configure(state="disabled", text="Aceptar visita")
            return

        status = str(visit.get("acceptance_status", "asignada")).strip().lower() or "asignada"
        if status == "asignada":
            self.accept_visit_button.configure(state="normal", text="Aceptar visita")
        elif status == "aceptada":
            self.accept_visit_button.configure(state="disabled", text="Visita confirmada")
        elif status == "finalizada":
            self.accept_visit_button.configure(state="disabled", text="Visita finalizada")
        elif status == "cancelada":
            self.accept_visit_button.configure(state="disabled", text="Visita cancelada")
        else:
            self.accept_visit_button.configure(state="disabled", text="Aceptar visita")

    def _show_readonly_visit_details(self, visit: dict | None) -> None:
        if self.notes_box is None:
            return

        if visit is None:
            self.notes_box.configure(state="normal")
            self.notes_box.delete("1.0", "end")
            self.notes_box.insert(
                "1.0",
                "Selecciona una visita para revisar direccion, fecha, servicio y observaciones registradas.",
            )
            self.notes_box.configure(state="disabled")
            self._set_accept_button_state(None)
            return

        confirmation_map = {
            "asignada": "Asignada",
            "aceptada": "Confirmada",
            "finalizada": "Finalizada",
            "cancelada": "Cancelada",
        }
        confirmation_status = confirmation_map.get(
            str(visit.get("acceptance_status", "asignada")).strip().lower(),
            "Asignada",
        )
        text = (
            f"Ejecutivos Tecnicos: {visit.get('inspectors_text', visit.get('inspector', '--'))}\n"
            f"Cliente: {visit.get('client', '--')}\n"
            f"Fecha: {visit.get('visit_date', '--')}\n"
            f"Hora de asignacion: {visit.get('assignment_time', '--') or '--'}\n"
            f"Hora de salida: {visit.get('departure_time', '--') or '--'}\n"
            f"Servicio: {visit.get('service', '--')}\n"
            f"Estado operativo: {visit.get('status', '--')}\n"
            f"Confirmacion: {confirmation_status}\n"
            f"Direccion: {visit.get('address', '--')}\n\n"
            f"Notas:\n{visit.get('notes', 'Sin observaciones')}"
        )
        self.notes_box.configure(state="normal")
        self.notes_box.delete("1.0", "end")
        self.notes_box.insert("1.0", text)
        self.notes_box.configure(state="disabled")
        self._set_accept_button_state(visit)

    def _on_client_change(self, _value: str) -> None:
        self.address_options = self.controller.get_client_addresses(self.client_var.get())
        labels = [item["address"] for item in self.address_options] or ["Sin direcciones"]
        if self.address_selector is not None and self.address_selector.winfo_exists():
            try:
                self.address_selector.configure(values=labels)
            except TclError:
                pass
        if self.address_options:
            self.address_var.set(self.address_options[0]["address"])
            self.service_var.set(self.address_options[0]["service"])
        else:
            self.address_var.set("")
            self.service_var.set("Sin servicio")

    def _clear_filters(self) -> None:
        self.search_var.set("")
        self.executive_filter_var.set(self._ALL_EXEC_FILTER)
        self.client_filter_var.set(self._ALL_CLIENT_FILTER)
        self.status_filter_var.set("Todos")
        self.refresh()

    def _on_address_change(self, _value: str) -> None:
        current_address = self.address_var.get()
        match = next((item for item in self.address_options if item["address"] == current_address), None)
        if match:
            self.service_var.set(match["service"])

    def _on_tree_select(self, _event=None) -> None:
        if self.tree is None:
            return

        selected = self.tree.selection()
        if not selected:
            return

        visit_id = selected[0]
        visit = next((item for item in self.controller.list_visits() if item.get("id") == visit_id), None)
        if visit is None:
            return

        normalized_date = self._normalize_date(visit.get("visit_date", ""))
        if normalized_date:
            parsed = datetime.strptime(normalized_date, "%Y-%m-%d").date()
            self.current_month = parsed.replace(day=1)
            self.date_var.set(normalized_date)
            self._render_month_grid()

        is_past = bool(normalized_date and normalized_date < date.today().strftime("%Y-%m-%d"))
        self._viewing_past = is_past

        self.selected_visit_id = visit_id
        if self.can_edit:
            inspectors = list(visit.get("inspectors", []))
            self.inspector_var.set(inspectors[0] if inspectors else "")
            self.exec2_var.set(inspectors[1] if len(inspectors) > 1 else self._NO_EXEC2)
            self.client_var.set(visit.get("client", ""))
            self._on_client_change(visit.get("client", ""))
            self.address_var.set(visit.get("address", ""))
            self.service_var.set(visit.get("service", "Sin servicio"))
            self.date_var.set(normalized_date or visit.get("visit_date", ""))
            self.assignment_time_var.set(visit.get("assignment_time", ""))
            self.departure_time_var.set(visit.get("departure_time", ""))
            self.status_var.set(visit.get("status", "Programada"))
            if self.notes_box is not None:
                self.notes_box.configure(state="normal")
                self.notes_box.delete("1.0", "end")
                self.notes_box.insert("1.0", visit.get("notes", ""))
            self._refresh_exec_options()
            self._set_form_editable(not is_past)
        elif self.notes_box is not None:
            self._show_readonly_visit_details(visit)

    @staticmethod
    def _normalize_date(raw_value: str) -> str:
        value = str(raw_value or "").strip()
        if not value:
            return ""
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
            try:
                return datetime.strptime(value, fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue
        return value if len(value) == 10 else ""
