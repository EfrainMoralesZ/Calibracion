from __future__ import annotations

import calendar as pycalendar
import csv
import io
from datetime import date, datetime
from tkinter import TclError, filedialog, messagebox, ttk
import tkinter as tk

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
        self.warehouse_var = ctk.StringVar(value="Sin almacen")
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
        self.acceptance_details_label: ctk.CTkLabel | None = None
        self.accept_visit_button: ctk.CTkButton | None = None
        self.visit_norm_check_frame: ctk.CTkScrollableFrame | None = None
        self.visit_norm_vars: dict[str, ctk.BooleanVar] = {}
        self.save_norm_report_button: ctk.CTkButton | None = None
        self.norm_report_status_var = ctk.StringVar(value="")
        self.norm_demand_month_var = ctk.StringVar(value=datetime.now().strftime("%Y-%m"))
        self.norm_demand_canvas: tk.Canvas | None = None
        self.norm_demand_summary_label: ctk.CTkLabel | None = None
        self.norm_demand_month_selector: ctk.CTkComboBox | None = None
        self._form_buttons: list[ctk.CTkButton] = []
        self._action_buttons: list[ctk.CTkButton] = []
        self._readonly_notice: ctk.CTkLabel | None = None
        self._viewing_past: bool = False
        self._agreement_banner: ctk.CTkFrame | None = None
        self.saturday_tree: ttk.Treeview | None = None
        self.saturday_year_var = ctk.StringVar(value=str(date.today().year))
        self.saturday_status_var = ctk.StringVar(value="")

        self.vac_exec_var = ctk.StringVar()
        self.vac_start_var = ctk.StringVar()
        self.vac_end_var = ctk.StringVar()
        self.ws_title_var = ctk.StringVar()
        self.ws_date_var = ctk.StringVar()
        self.ws_desc_var = ctk.StringVar()
        self.vacation_tree: ttk.Treeview | None = None
        self.workshop_tree: ttk.Treeview | None = None

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self._build_ui()

    def _build_ui(self) -> None:
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        header.grid_columnconfigure(0, weight=1)

        title = "Agenda operativa" if self.can_edit else "Visitas asignadas"
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

        if not self.controller.is_executive_role(self.controller.current_user or {}):
            tabs.add("Reporte Sábado")
            self._build_saturday_report_tab(tabs.tab("Reporte Sábado"))

        if self.can_edit:
            tabs.add("Vacaciones y Talleres")
            self._build_vacations_workshops_tab(tabs.tab("Vacaciones y Talleres"))

    def _build_calendar_tab(self, tab) -> None:
        # Para ejecutivos: 3 columnas (calendario grande, detalle, normas compactas)
        # Para admin: 2 columnas (calendario, detalle con formulario)
        if not self.can_edit:
            tab.grid_columnconfigure(0, weight=5, minsize=600)  # Calendario mucho más grande
            tab.grid_columnconfigure(1, weight=2, minsize=330)  # Detalle de visita
            tab.grid_columnconfigure(2, weight=1, minsize=90)  # Normas aplicadas más pequeña
        else:
            tab.grid_columnconfigure(0, weight=5, minsize=600)
            tab.grid_columnconfigure(1, weight=2, minsize=320)
        tab.grid_rowconfigure(0, weight=1)

        calendar_shell = ctk.CTkFrame(tab, fg_color=self.style["surface"], corner_radius=22)
        calendar_shell.grid(row=0, column=0, sticky="nsew", padx=(0, 0), pady=12)
        calendar_shell.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            calendar_shell,
            text="Calendario Operativo",
            font=self.fonts["label_bold"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=0, column=0, padx=18, pady=(16, 8), sticky="w")

        calendar_panel = ctk.CTkFrame(calendar_shell, fg_color=self.style["fondo"], corner_radius=18)
        calendar_panel.grid(row=1, column=0, padx=0, pady=(0, 0), sticky="nsew")
        calendar_panel.grid_columnconfigure(0, weight=1)
        calendar_shell.grid_rowconfigure(1, weight=1)

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
            week_header.grid_columnconfigure(idx, weight=1, minsize=96)
            ctk.CTkLabel(
                week_header,
                text=day_name,
                font=self.fonts["small_bold"],
                text_color="#6D7480",
            ).grid(row=0, column=idx, padx=2, pady=(0, 2), sticky="nsew")

        self.calendar_grid_frame = ctk.CTkFrame(calendar_panel, fg_color="transparent")
        self.calendar_grid_frame.grid(row=2, column=0, padx=0, pady=(0, 0), sticky="nsew")
        calendar_panel.grid_rowconfigure(2, weight=1)

        # Leyenda de estados
        legend = ctk.CTkFrame(calendar_panel, fg_color="transparent")
        legend.grid(row=3, column=0, padx=10, pady=(6, 10), sticky="w")

        for col, (label, color) in enumerate([
            ("Asignada", "#FFFFFF"),
            ("Aceptada", "#ECD925"),
            ("Reasignada", "#E4E7EB"),
            ("Finalizada", "#D1F7D1"),
            ("Cancelada", "#F7D1D1"),
            ("Vacaciones", "#FFE0B2"),
            ("Taller", "#D4EDFC"),
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


        # Contenedor para detalle y botón fijo
        self.side_panel_container = ctk.CTkFrame(tab, fg_color=self.style["surface"], corner_radius=22)
        self.side_panel_container.grid(row=0, column=1, sticky="nsew", pady=12)
        self.side_panel_container.grid_rowconfigure(0, weight=1)
        self.side_panel_container.grid_rowconfigure(1, weight=0)
        self.side_panel_container.grid_columnconfigure(0, weight=1)

        # Frame scrollable para el contenido de detalle
        self.side_panel = ctk.CTkScrollableFrame(self.side_panel_container, fg_color="transparent", corner_radius=0, scrollbar_button_color=None, scrollbar_button_hover_color=None)
        self.side_panel.grid(row=0, column=0, sticky="nsew")
        self.side_panel.grid_columnconfigure(0, weight=1)

        # Frame fijo inferior para el botón
        self.bottom_button_frame = ctk.CTkFrame(self.side_panel_container, fg_color="transparent")
        self.bottom_button_frame.grid(row=1, column=0, sticky="ew", padx=0, pady=(0, 10))
        self.bottom_button_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            self.side_panel,
            text="Detalle de visita",
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
            # Agreement banner for admin — shown below the form when client has agreements
            self._agreement_banner = ctk.CTkFrame(
                self.side_panel,
                fg_color="#FFF3CD",
                corner_radius=10,
                border_width=1,
                border_color="#F0C040",
            )
            self._agreement_banner.grid(row=2, column=0, padx=18, pady=(0, 10), sticky="ew")
            self._agreement_banner.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(
                self._agreement_banner,
                text="⚠️  Este cliente tiene acuerdos registrados. Revísalos antes de realizar la visita.",
                font=self.fonts["small"],
                text_color="#7A5A00",
                wraplength=280,
                justify="left",
            ).grid(row=0, column=0, padx=12, pady=10, sticky="w")
            self._agreement_banner.grid_remove()
        else:
            # Vista para ejecutivos: Columna 2 = Detalle de visita
            message = ctk.CTkLabel(
                self.side_panel,
                text=(
                    "Selecciona una visita para consultar la informacion principal y observaciones registradas."
                ),
                font=self.fonts["small"],
                text_color="#6D7480",
                wraplength=330,
                justify="left",
            )
            message.grid(row=1, column=0, padx=18, pady=(0, 12), sticky="w")

            # Agreement banner — shown above notes when client has agreements
            self._agreement_banner = ctk.CTkFrame(
                self.side_panel,
                fg_color="#FFF3CD",
                corner_radius=10,
                border_width=1,
                border_color="#F0C040",
            )
            self._agreement_banner.grid(row=2, column=0, padx=18, pady=(0, 8), sticky="ew")
            self._agreement_banner.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(
                self._agreement_banner,
                text="⚠️  Este cliente tiene acuerdos registrados. Revísalos antes de realizar la visita.",
                font=self.fonts["small"],
                text_color="#7A5A00",
                wraplength=330,
                justify="left",
            ).grid(row=0, column=0, padx=12, pady=10, sticky="w")
            self._agreement_banner.grid_remove()

            self.notes_box = ctk.CTkTextbox(self.side_panel, height=420, corner_radius=18)
            self.notes_box.grid(row=3, column=0, padx=18, pady=(0, 12), sticky="nsew")
            self.notes_box.configure(state="disabled")

            # Columna 3: Panel de normas (columna separada para ejecutivos)
            self.normas_panel = ctk.CTkFrame(tab, fg_color=self.style["surface"], corner_radius=22)
            self.normas_panel.grid(row=0, column=2, sticky="nsew", padx=(0, 0), pady=12)
            self.normas_panel.grid_columnconfigure(0, weight=1)
            self.normas_panel.grid_rowconfigure(1, weight=1)

            self.accept_visit_button = ctk.CTkButton(
                self.bottom_button_frame,
                text="Aceptar visita",
                command=self._accept_visit_action,
                fg_color="#ECD925",
                text_color="#282828",
                hover_color="#D8C220",
                state="disabled",
            )
            self.accept_visit_button.grid(row=0, column=0, padx=18, pady=(0, 0), sticky="ew")

            ctk.CTkLabel(
                self.normas_panel,
                text="Normas aplicadas",
                font=self.fonts["label_bold"],
                text_color=self.style["texto_oscuro"],
            ).grid(row=0, column=0, padx=12, pady=(16, 8), sticky="w")

            self.visit_norm_check_frame = ctk.CTkScrollableFrame(
                self.normas_panel,
                fg_color=self.style["fondo"],
                corner_radius=12,
                scrollbar_button_color=None,
                scrollbar_button_hover_color=None,
            )
            self.visit_norm_check_frame.grid(row=1, column=0, padx=12, pady=(0, 8), sticky="nsew")
            self.visit_norm_check_frame.grid_columnconfigure(0, weight=1)

            ctk.CTkLabel(
                self.normas_panel,
                textvariable=self.norm_report_status_var,
                font=self.fonts["small"],
                text_color="#6D7480",
                justify="left",
                wraplength=160,
            ).grid(row=2, column=0, padx=12, pady=(0, 8), sticky="w")

            self.save_norm_report_button = ctk.CTkButton(
                self.normas_panel,
                text="Finalizar visita",
                command=self._save_visit_norm_report,
                fg_color=self.style["secundario"],
                hover_color="#1D1D1D",
                state="disabled",
            )
            self.save_norm_report_button.grid(row=3, column=0, padx=12, pady=(0, 12), sticky="ew")

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

    def _build_dashboard_tab(self, tab) -> None:
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)

        toolbar = ctk.CTkFrame(tab, fg_color="transparent")
        toolbar.grid(row=0, column=0, padx=18, pady=(14, 10), sticky="ew")
        toolbar.grid_columnconfigure(3, weight=1)

        ctk.CTkLabel(
            toolbar,
            text="Mes",
            font=self.fonts["label"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=0, column=0, padx=(0, 8), sticky="w")

        month_values = self.controller.get_norm_report_months() or [datetime.now().strftime("%Y-%m")]
        if self.norm_demand_month_var.get() not in month_values:
            self.norm_demand_month_var.set(month_values[-1])
        self.norm_demand_month_selector = ctk.CTkComboBox(
            toolbar,
            variable=self.norm_demand_month_var,
            values=month_values,
            width=150,
            height=36,
            fg_color="#FFFFFF",
            border_color="#D5D8DC",
            button_color=self.style["primario"],
            dropdown_hover_color=self.style["primario"],
            command=lambda _value: self._refresh_norm_dashboard(),
        )
        self.norm_demand_month_selector.grid(row=0, column=1, sticky="w")

        ctk.CTkButton(
            toolbar,
            text="Actualizar",
            fg_color=self.style["primario"],
            text_color=self.style["texto_oscuro"],
            hover_color="#D8C220",
            command=self._refresh_norm_dashboard,
        ).grid(row=0, column=2, padx=(10, 0), sticky="w")

        panel = ctk.CTkFrame(tab, fg_color="#FFFFFF", corner_radius=18, border_width=1, border_color="#E3E6EA")
        panel.grid(row=1, column=0, padx=18, pady=(0, 18), sticky="nsew")
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(
            panel,
            text="Demanda de normas por mes",
            font=self.fonts["label_bold"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=0, column=0, padx=14, pady=(12, 4), sticky="w")

        self.norm_demand_summary_label = ctk.CTkLabel(
            panel,
            text="Sin reportes de normas para el mes seleccionado.",
            font=self.fonts["small"],
            text_color="#6D7480",
            justify="left",
        )
        self.norm_demand_summary_label.grid(row=1, column=0, padx=14, pady=(0, 6), sticky="w")

        self.norm_demand_canvas = tk.Canvas(panel, bg="#FFFFFF", highlightthickness=0)
        self.norm_demand_canvas.grid(row=2, column=0, padx=14, pady=(0, 14), sticky="nsew")
        self.norm_demand_canvas.bind("<Configure>", lambda _event: self._refresh_norm_dashboard())
        self._refresh_norm_dashboard()

    def _refresh_norm_dashboard(self) -> None:
        if not self.can_edit or self.norm_demand_canvas is None:
            return

        month_values = self.controller.get_norm_report_months() or [datetime.now().strftime("%Y-%m")]
        if self.norm_demand_month_selector is not None:
            try:
                self.norm_demand_month_selector.configure(values=month_values)
            except TclError:
                pass
        if self.norm_demand_month_var.get() not in month_values:
            self.norm_demand_month_var.set(month_values[-1])

        selected_month = self.norm_demand_month_var.get().strip()
        demand_rows = self.controller.get_monthly_norm_demand(selected_month)
        self._draw_norm_demand_chart(demand_rows)

        if self.norm_demand_summary_label is not None:
            if not demand_rows:
                self.norm_demand_summary_label.configure(text="Sin reportes de normas para el mes seleccionado.")
            else:
                top = demand_rows[0]
                bottom = demand_rows[-1]
                self.norm_demand_summary_label.configure(
                    text=(
                        f"Mes: {selected_month} | Norma mas demandada: {top['norm']} ({top['count']}) | "
                        f"Menor demanda: {bottom['norm']} ({bottom['count']})."
                    )
                )

    def _draw_norm_demand_chart(self, demand_rows: list[dict]) -> None:
        if self.norm_demand_canvas is None:
            return

        canvas = self.norm_demand_canvas
        canvas.delete("all")
        width = max(canvas.winfo_width(), 620)
        height = max(canvas.winfo_height(), 260)

        if not demand_rows:
            canvas.create_text(
                width / 2,
                height / 2,
                text="No hay datos para graficar en este mes.",
                fill="#6D7480",
                font=("Arial", 12),
            )
            return

        visible = demand_rows[:12]
        max_count = max(int(item.get("count", 0)) for item in visible) or 1
        left_margin = 170
        top_margin = 16
        row_height = max(20, (height - 24) // len(visible))
        bar_space = max(120, width - left_margin - 70)

        for index, item in enumerate(visible):
            y = top_margin + index * row_height
            norm = str(item.get("norm", "")).strip() or "SIN_NORMA"
            count = int(item.get("count", 0))
            bar_width = int((count / max_count) * bar_space)
            if bar_width <= 0 and count > 0:
                bar_width = 2

            canvas.create_text(
                left_margin - 8,
                y + row_height // 2,
                text=norm,
                anchor="e",
                fill=self.style["texto_oscuro"],
                font=("Arial", 9),
            )
            fill_color = self.style["primario"] if index == 0 else "#BFC7D1"
            canvas.create_rectangle(
                left_margin,
                y + 4,
                left_margin + bar_width,
                y + row_height - 4,
                fill=fill_color,
                outline="",
            )
            canvas.create_text(
                left_margin + bar_width + 6,
                y + row_height // 2,
                text=str(count),
                anchor="w",
                fill=self.style["texto_oscuro"],
                font=("Arial", 9, "bold"),
            )

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

        # ── Normas disponibles para ambos ejecutivos ───────────────────────
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

        ctk.CTkLabel(fields, text="Confirmaciones de ejecutivos", font=self.fonts["label"],
                     text_color=self.style["texto_oscuro"]).grid(row=6, column=0, sticky="w", pady=(8, 4))
        self.acceptance_details_label = ctk.CTkLabel(
            fields,
            text="Sin confirmaciones registradas.",
            font=self.fonts["small"],
            text_color="#6D7480",
            fg_color=self.style["fondo"],
            corner_radius=10,
            wraplength=420,
            justify="left",
            padx=10,
            pady=8,
            anchor="w",
        )
        self.acceptance_details_label.grid(row=7, column=0, sticky="ew")

        # ── Cliente / Dirección ────────────────────────────────────────────
        self.client_selector = self._combo(fields, self.client_var, self.controller.get_client_names(), self._on_client_change)
        self._label_and_widget_rows(fields, 8, "Cliente", self.client_selector)
        self.address_selector = self._combo(fields, self.address_var, ["Selecciona un cliente"], self._on_address_change)
        self._label_and_widget_rows(fields, 10, "Direccion", self.address_selector)
        warehouse_label = ctk.CTkLabel(
            fields,
            textvariable=self.warehouse_var,
            font=self.fonts["small_bold"],
            text_color=self.style["texto_oscuro"],
            fg_color=self.style["fondo"],
            corner_radius=12,
            padx=12,
            pady=8,
            anchor="w",
            justify="left",
        )
        self._label_and_widget_rows(fields, 12, "Almacen", warehouse_label)
        self.date_entry = self._entry(fields, self.date_var)
        self._label_and_widget_rows(fields, 14, "Fecha de visita", self.date_entry)
        self.assignment_time_entry = ctk.CTkEntry(fields, textvariable=self.assignment_time_var, height=38, border_color="#D5D8DC", placeholder_text="08:00")
        self.assignment_time_entry.bind("<FocusOut>", lambda _event: self._normalize_time_var(self.assignment_time_var))
        self.assignment_time_entry.bind("<Return>", lambda _event: self._normalize_time_var(self.assignment_time_var))
        self._label_and_widget_rows(fields, 16, "Hora de asignacion al almacen (24 hrs)", self.assignment_time_entry)
        self.departure_time_entry = ctk.CTkEntry(fields, textvariable=self.departure_time_var, height=38, border_color="#D5D8DC", placeholder_text="18:00")
        self.departure_time_entry.bind("<FocusOut>", lambda _event: self._normalize_time_var(self.departure_time_var))
        self.departure_time_entry.bind("<Return>", lambda _event: self._normalize_time_var(self.departure_time_var))
        self._label_and_widget_rows(fields, 18, "Hora de salida (24 hrs)", self.departure_time_entry)
        self.status_selector = self._combo(fields, self.status_var, ["Programada", "Realizada", "Reprogramada"])
        self._label_and_widget_rows(fields, 20, "Estado", self.status_selector)

        ctk.CTkLabel(fields, text="Usa formato 24 horas. Ejemplos: 08:00, 14:30, 18:45", font=self.fonts["small"],
                     text_color="#6D7480").grid(row=22, column=0, sticky="w", pady=(4, 6))

        ctk.CTkLabel(fields, text="Notas", font=self.fonts["label"],
                     text_color=self.style["texto_oscuro"]).grid(row=23, column=0, sticky="w", pady=(12, 6))
        self.notes_box = ctk.CTkTextbox(fields, height=120, corner_radius=16)
        self.notes_box.grid(row=24, column=0, sticky="ew")

        button_row = ctk.CTkFrame(fields, fg_color="transparent")
        button_row.grid(row=25, column=0, sticky="ew", pady=(16, 0))

        self._form_buttons.clear()
        self._action_buttons: list[ctk.CTkButton] = []

        if self.can_edit:
            primary_row = ctk.CTkFrame(button_row, fg_color="transparent")
            primary_row.grid(row=0, column=0, sticky="ew")
            primary_row.grid_columnconfigure(0, weight=1)
            primary_row.grid_columnconfigure(1, weight=1)

            danger_row = ctk.CTkFrame(button_row, fg_color="transparent")
            danger_row.grid(row=1, column=0, sticky="ew", pady=(8, 0))
            for col in range(3):
                danger_row.grid_columnconfigure(col, weight=1)

            btn_clear = ctk.CTkButton(
                primary_row, text="Limpiar", command=self.clear_form,
                fg_color=self.style["fondo"], text_color=self.style["texto_oscuro"], hover_color="#E9ECEF",
            )
            btn_clear.grid(row=0, column=0, padx=(0, 6), sticky="ew")
            self._form_buttons.append(btn_clear)

            btn_save = ctk.CTkButton(
                primary_row, text="Guardar", command=self.save_visit,
                fg_color=self.style["secundario"], hover_color="#1D1D1D",
            )
            btn_save.grid(row=0, column=1, padx=(6, 0), sticky="ew")
            self._form_buttons.append(btn_save)

            btn_cancel = ctk.CTkButton(
                danger_row, text="Cancelar visita", command=self._cancel_visit_action,
                fg_color="#F7D1D1", text_color="#D1534E", hover_color="#F0B8B4",
            )
            btn_cancel.grid(row=0, column=0, padx=(0, 4), sticky="ew")
            self._form_buttons.append(btn_cancel)
            self._action_buttons.append(btn_cancel)

            btn_reassign = ctk.CTkButton(
                danger_row, text="Reasignar ejecutivo", command=self._reassign_visit_action,
                fg_color="#FFFFFF", text_color="#282828", hover_color="#F5F5F5",
            )
            btn_reassign.grid(row=0, column=1, padx=4, sticky="ew")
            self._form_buttons.append(btn_reassign)
            self._action_buttons.append(btn_reassign)

            btn_delete = ctk.CTkButton(
                danger_row, text="Eliminar", command=self.delete_selected_visit,
                fg_color=self.style["peligro"], hover_color="#B43C31",
            )
            btn_delete.grid(row=0, column=2, padx=(4, 0), sticky="ew")
            self._form_buttons.append(btn_delete)
            self._action_buttons.append(btn_delete)
        # El botón 'Aceptar visita' fijo en la parte inferior del panel de detalle
        if hasattr(self, 'side_panel') and self.side_panel is not None:
            self.side_panel.grid_rowconfigure(98, weight=1)
            bottom_frame = ctk.CTkFrame(self.side_panel, fg_color="transparent")
            bottom_frame.grid(row=99, column=0, sticky="sew", padx=0, pady=(0, 12))
            self.accept_visit_button = ctk.CTkButton(
                bottom_frame,
                text="Aceptar visita",
                command=self._accept_visit_action,
                fg_color="#ECD925",
                text_color="#282828",
                hover_color="#D8C220",
                state="disabled",
            )
            self.accept_visit_button.pack(fill="x", padx=18)

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
            norms2 = self.controller.get_accredited_norms(exec2)
            joined_1 = ", ".join(norms1) if norms1 else "Sin normas"
            joined_2 = ", ".join(norms2) if norms2 else "Sin normas"
            text = f"{exec1}: {joined_1}\n{exec2}: {joined_2}"
            color = self.style["exito"] if (norms1 or norms2) else "#9AA0A8"

        try:
            self.norms_display_label.configure(text=text, text_color=color)
        except TclError:
            pass

    @staticmethod
    def _identity_key(raw_value: object) -> str:
        return str(raw_value or "").strip().casefold()

    def _get_visit_acceptance_map(self, visit: dict | None) -> dict[str, str]:
        if not isinstance(visit, dict):
            return {}

        raw_responses = visit.get("acceptance_responses", {})
        if not isinstance(raw_responses, dict):
            return {}

        normalized: dict[str, str] = {}
        for inspector_name, payload in raw_responses.items():
            clean_name = str(inspector_name or "").strip()
            if not clean_name:
                continue

            if isinstance(payload, dict):
                confirmed_at = str(payload.get("confirmed_at", "")).strip()
            else:
                confirmed_at = str(payload or "").strip()

            if confirmed_at:
                normalized[clean_name] = confirmed_at
        return normalized

    def _build_acceptance_details_text(self, visit: dict | None) -> str:
        if not isinstance(visit, dict):
            return "Sin confirmaciones registradas."

        # Mejor formato para el detalle de visita
        info = []
        info.append("[b]INFORMACIÓN PRINCIPAL[/b]")
        info.append(f"• Ejecutivos técnicos: {', '.join(visit.get('inspectors', [])) or 'Sin asignar'}")
        info.append(f"• Cliente: {visit.get('client', 'Sin cliente')}")
        info.append(f"• Fecha: {visit.get('date', 'Sin fecha')}")
        info.append(f"• Estado operativo: {visit.get('status', 'Sin estado')}")

        info.append("\n[b]HORARIOS[/b]")
        info.append(f"• Hora de asignación: {visit.get('assignment_time', 'Sin hora')}")
        info.append(f"• Hora de salida: {visit.get('departure_time', 'Sin hora')}")

        info.append("\n[b]UBICACIÓN[/b]")
        info.append(f"• Dirección: {visit.get('address', 'Sin dirección')}")

        obs = visit.get('notes', '').strip()
        if obs:
            info.append("\n[b]OBSERVACIONES[/b]")
            info.append(f"• {obs}")

        # Confirmaciones
        inspectors = list(visit.get("inspectors", []))
        confirmations = self._get_visit_acceptance_map(visit)
        confirmations_by_identity = {
            self._identity_key(name): confirmed_at
            for name, confirmed_at in confirmations.items()
            if self._identity_key(name)
        }
        if inspectors:
            lines = []
            confirmed_count = 0
            for inspector_name in inspectors:
                identity = self._identity_key(inspector_name)
                confirmed_at = confirmations_by_identity.get(identity, "")
                if confirmed_at:
                    confirmed_count += 1
                    lines.append(f"- {inspector_name}: Confirmó el {confirmed_at}")
                else:
                    lines.append(f"- {inspector_name}: Pendiente de confirmar")
            info.append(f"\n[b]Confirmaciones:[/b] {confirmed_count}/{len(inspectors)}")
            info.extend(lines)

        finalized_at = str(visit.get("finalized_at", "")).strip() if isinstance(visit, dict) else ""
        if finalized_at:
            info.append(f"\n[b]Finalizada a las:[/b] {finalized_at}")

        # Unir todo con saltos de línea
        return "\n".join(info)

    def _update_acceptance_details(self, visit: dict | None) -> None:
        if self.acceptance_details_label is None:
            return

        details_text = self._build_acceptance_details_text(visit)
        try:
            self.acceptance_details_label.configure(text=details_text)
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

    @staticmethod
    def _normalize_time_text(raw_value: str) -> str:
        value = str(raw_value or "").strip().upper().replace(".", "")
        if not value:
            return ""

        for fmt in ("%H:%M", "%H%M", "%I:%M %p", "%I:%M%p", "%I %p"):
            try:
                return datetime.strptime(value, fmt).strftime("%H:%M")
            except ValueError:
                continue
        return value

    def _normalize_time_var(self, variable: ctk.StringVar) -> None:
        variable.set(self._normalize_time_text(variable.get()))

    def _sync_address_metadata(self) -> None:
        current_address = self.address_var.get().strip()
        current_client = self.client_var.get().strip()
        if not current_address or not current_client:
            self.warehouse_var.set("Sin almacen")
            self.service_var.set("Sin servicio")
            return

        match = next((item for item in self.address_options if item.get("address", "") == current_address), None)
        if match is None:
            match = next((item for item in self.address_options if item.get("location", "") == current_address), None)
            if match is not None:
                self.address_var.set(match.get("address", current_address))

        warehouse = ""
        service = "Sin servicio"
        if match is not None:
            warehouse = str(match.get("warehouse", "")).strip()
            service = str(match.get("service", "Sin servicio")).strip() or "Sin servicio"

        if not warehouse:
            warehouse = self.controller.get_client_warehouse_for_address(current_client, self.address_var.get())

        self.warehouse_var.set(warehouse or "Sin almacen")
        self.service_var.set(service)

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
            self._first_render()

        if self.can_edit:
            self._refresh_norm_dashboard()
        else:
            selected_visit = None
            if self.selected_visit_id:
                selected_visit = next(
                    (item for item in self.controller.list_visits() if item.get("id") == self.selected_visit_id),
                    None,
                )
            self._render_visit_norm_checklist(selected_visit)

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
            self.calendar_grid_frame.grid_rowconfigure(row, weight=1, minsize=44)
            for col in range(7):
                self.calendar_grid_frame.grid_columnconfigure(col, weight=1, minsize=96)
                day = month_matrix[row][col]
                if day == 0:
                    cell = ctk.CTkButton(
                        self.calendar_grid_frame,
                        text="",
                        width=96,
                        height=42,
                        fg_color="#EEF0F2",
                        hover=False,
                        state="disabled",
                    )
                    cell.grid(row=row, column=col, padx=2, pady=1, sticky="ew")
                    continue

                day_iso = date(self.current_month.year, self.current_month.month, day).strftime("%Y-%m-%d")
                count = self.visible_date_counts.get(day_iso, 0)
                is_past_day = day_iso < today_iso
                has_past_visits = is_past_day and count > 0
                is_sunday = (col == 6)

                # Check vacation / workshop events for this date
                day_vacations = self.controller.get_vacations_for_date(day_iso)
                day_workshops = self.controller.get_workshops_for_date(day_iso)

                # Build label with event info
                label_parts = [str(day)]
                if count:
                    label_parts.append(f"{count} visita(s)")
                if day_vacations:
                    vac_names = ", ".join(sorted({v.get("executive", "") for v in day_vacations}))
                    label_parts.append(f"🏖 {vac_names}" if len(vac_names) <= 12 else "🏖 Vacaciones")
                if day_workshops:
                    label_parts.append(f"📋 Taller")
                label = "\n".join(label_parts)
                
                # Determine color based on acceptance status of visits on this date
                day_visits = visits_by_date.get(day_iso, [])
                has_cancelled = any(v.get("acceptance_status") == "cancelada" for v in day_visits)
                has_finalizada = any(v.get("acceptance_status") == "finalizada" for v in day_visits)
                has_aceptada = any(v.get("acceptance_status") == "aceptada" for v in day_visits)
                has_reasignada = any(
                    v.get("acceptance_status") == "reasignada"
                    or (v.get("acceptance_status") == "asignada" and v.get("reassigned_from"))
                    for v in day_visits
                )
                has_assigned = any(v.get("acceptance_status") == "asignada" for v in day_visits)

                if (is_past_day and not has_past_visits and not day_vacations and not day_workshops) or (is_sunday and not count and not day_vacations and not day_workshops):
                    fg_color = "#EEF0F2"
                    text_color = "#9AA0A8"
                    state = "disabled"
                    hover_color = "#EEF0F2"
                elif day_workshops and day_vacations:
                    fg_color = "#E0D4F5"
                    text_color = "#5B3A8C"
                    state = "normal"
                    hover_color = "#D0C0EB"
                elif day_workshops:
                    fg_color = "#D4EDFC"
                    text_color = "#1A6FA0"
                    state = "normal"
                    hover_color = "#B8DFFA"
                elif day_vacations:
                    fg_color = "#FFE0B2"
                    text_color = "#BF6C00"
                    state = "normal"
                    hover_color = "#FFD18C"
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
                elif has_reasignada:
                    fg_color = "#E4E7EB"
                    text_color = "#3C4657"
                    state = "normal"
                    hover_color = "#D6DBE2"
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
                    width=96,
                    height=42,
                    fg_color=fg_color,
                    text_color=text_color,
                    hover_color=hover_color,
                    state=state,
                    font=self.fonts["small"],
                    command=(lambda iso=day_iso: self._select_calendar_date(iso)) if state == "normal" else None,
                )
                if state == "disabled":
                    cell.configure(hover=False)
                cell.grid(row=row, column=col, padx=2, pady=1, sticky="nsew")

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
                visit = self._pick_visit_for_date(visits_on_date, is_past)
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
                self._sync_address_metadata()
                self.assignment_time_var.set(visit.get("assignment_time", ""))
                self.departure_time_var.set(visit.get("departure_time", ""))
                self.status_var.set(visit.get("status", "Programada"))
                self._update_acceptance_details(visit)
                # Show agreement banner for admin if this client has recorded agreements
                if self._agreement_banner is not None:
                    try:
                        client_name = visit.get("client", "").strip()
                        has_agreements = bool(
                            client_name and self.controller.get_client_agreements(client_name)
                        )
                        if has_agreements:
                            self._agreement_banner.grid()
                        else:
                            self._agreement_banner.grid_remove()
                    except Exception:
                        self._agreement_banner.grid_remove()
            else:
                self.selected_visit_id = None
                self._update_acceptance_details(None)
                if self._agreement_banner is not None:
                    self._agreement_banner.grid_remove()
        elif not self.can_edit:
            visits_on_date = [
                v for v in self.controller.list_visits()
                if self._normalize_date(v.get("visit_date", "")) == iso_date
            ]
            if visits_on_date:
                visit = self._pick_visit_for_date(visits_on_date, is_past)
                self.selected_visit_id = visit.get("id")
                self._show_readonly_visit_details(visit)
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

    @staticmethod
    def _pick_visit_for_date(visits_on_date: list[dict], is_past: bool) -> dict:
        """Choose the most relevant visit for a date.

        For past dates, prioritize finalized visits so they can be reviewed quickly.
        """

        def _rank(visit: dict) -> tuple[int, str, str]:
            acceptance_status = str(visit.get("acceptance_status", "")).strip().lower()
            operational_status = str(visit.get("status", "")).strip().lower()
            is_finalized = acceptance_status == "finalizada" or operational_status == "finalizada"
            priority = 1 if (is_past and is_finalized) else 0
            return (
                priority,
                str(visit.get("updated_at", "")),
                str(visit.get("assignment_time", "")),
            )

        return sorted(visits_on_date, key=_rank, reverse=True)[0]

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
        self.norm_report_status_var.set("")
        self._update_acceptance_details(None)
        self._set_form_editable(True)
        self.inspector_var.set("")
        self.exec2_var.set(self._NO_EXEC2)
        self.client_var.set("")
        self.address_var.set("")
        self.warehouse_var.set("Sin almacen")
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
        if not self.can_edit:
            self._render_visit_norm_checklist(None)

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

        self._normalize_time_var(self.assignment_time_var)
        self._normalize_time_var(self.departure_time_var)
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

        try:
            self.controller.accept_visit(self.selected_visit_id)
        except ValueError as error:
            messagebox.showwarning("Visitas", str(error))
            return
        visit = next((item for item in self.controller.list_visits() if item.get("id") == self.selected_visit_id), None)
        self._show_readonly_visit_details(visit)
        messagebox.showinfo("Visitas", "Confirmacion registrada correctamente.")
        self._render_month_grid()
        self.refresh()

    def _set_accept_button_state(self, visit: dict | None) -> None:
        if self.accept_visit_button is None:
            return

        if visit is None:
            self.accept_visit_button.configure(
                state="disabled",
                text="Aceptar visita",
                fg_color="#E9ECEF",
                hover=False,
                border_width=0,
                border_color="#E9ECEF",
                text_color="#7A828C",
                text_color_disabled="#7A828C",
            )
            return

        status = str(visit.get("acceptance_status", "asignada")).strip().lower() or "asignada"
        if status == "finalizada":
            self.accept_visit_button.configure(
                state="disabled",
                text="Visita finalizada",
                fg_color="#FFF0A8",
                hover=False,
                border_width=1,
                border_color="#D7BE20",
                text_color="#5F4B00",
                text_color_disabled="#5F4B00",
            )
        elif status == "cancelada":
            self.accept_visit_button.configure(
                state="disabled",
                text="Visita cancelada",
                fg_color="#F7D1D1",
                hover=False,
                border_width=1,
                border_color="#E1A4A0",
                text_color="#9C2F20",
                text_color_disabled="#9C2F20",
            )
        else:
            self.accept_visit_button.configure(
                fg_color="#ECD925",
                hover=True,
                hover_color="#D8C220",
                border_width=0,
                border_color="#ECD925",
                text_color="#282828",
                text_color_disabled="#282828",
            )
            viewer_name = str((self.controller.current_user or {}).get("name", "")).strip()
            viewer_identity = self._identity_key(viewer_name)
            inspectors = list(visit.get("inspectors", []))
            inspector_identities = {self._identity_key(name) for name in inspectors}
            if viewer_identity not in inspector_identities:
                self.accept_visit_button.configure(
                    state="disabled",
                    text="Sin permiso",
                    fg_color="#E9ECEF",
                    hover=False,
                    text_color="#6D7480",
                    text_color_disabled="#6D7480",
                )
                return

            confirmations = self._get_visit_acceptance_map(visit)
            confirmed_identities = {self._identity_key(name) for name in confirmations}
            if viewer_identity in confirmed_identities:
                self.accept_visit_button.configure(
                    state="disabled",
                    text="Ya confirmaste",
                    fg_color="#D1F7D1",
                    hover=False,
                    border_width=1,
                    border_color="#A8E0A8",
                    text_color="#0B7A4D",
                    text_color_disabled="#0B7A4D",
                )
            else:
                self.accept_visit_button.configure(state="normal", text="Aceptar visita")

    def _render_visit_norm_checklist(self, visit: dict | None) -> None:
        if self.visit_norm_check_frame is None:
            return

        for child in self.visit_norm_check_frame.winfo_children():
            child.destroy()
        self.visit_norm_vars.clear()

        if visit is None:
            self.norm_report_status_var.set("Selecciona una visita para reportar normas aplicadas.")
            if self.save_norm_report_button is not None:
                self.save_norm_report_button.configure(state="disabled")
            return

        visit_id = str(visit.get("id", "")).strip()
        viewer_name = str((self.controller.current_user or {}).get("name", "")).strip()
        catalog_tokens = sorted(
            {
                str(item.get("token", "")).strip()
                for item in self.controller.get_catalog_norms()
                if str(item.get("token", "")).strip()
            },
            key=self.controller._norm_sort_key,
        )
        available_norms = catalog_tokens or self.controller.get_visit_available_norms(visit_id, viewer_name)
        selected_norms = set(self.controller.get_visit_reported_norms(visit_id, viewer_name))
        norm_display_map = {
            str(item.get("token", "")).strip(): str(item.get("token", "")).strip()
            for item in self.controller.get_catalog_norms()
            if str(item.get("token", "")).strip()
        }

        if not available_norms:
            ctk.CTkLabel(
                self.visit_norm_check_frame,
                text="No hay normas disponibles para esta visita.",
                font=self.fonts["small"],
                text_color="#6D7480",
                justify="left",
            ).grid(row=0, column=0, padx=8, pady=8, sticky="w")
        else:
            for index, token in enumerate(available_norms):
                variable = ctk.BooleanVar(value=token in selected_norms)
                self.visit_norm_vars[token] = variable
                ctk.CTkCheckBox(
                    self.visit_norm_check_frame,
                    text=token,
                    variable=variable,
                    text_color=self.style["texto_oscuro"],
                    fg_color=self.style["primario"],
                    hover_color="#D8C220",
                    checkmark_color=self.style["secundario"],
                ).grid(row=index, column=0, padx=8, pady=(8 if index == 0 else 4, 4), sticky="w")

        confirmation_status = str(visit.get("acceptance_status", "asignada")).strip().lower()
        if selected_norms:
            selected_labels = [
                norm_display_map.get(token, token)
                for token in sorted(selected_norms, key=self.controller._norm_sort_key)
            ]
            self.norm_report_status_var.set(
                f"Normas reportadas: {'; '.join(selected_labels)}."
            )
        else:
            self.norm_report_status_var.set("Marca las normas aplicadas y guarda el reporte de la visita.")

        if self.save_norm_report_button is not None:
            finalized_at = str(visit.get("finalized_at", "")).strip()
            enabled = confirmation_status not in {"cancelada"} and not finalized_at
            self.save_norm_report_button.configure(state="normal" if enabled else "disabled")

    def _save_visit_norm_report(self) -> None:
        if self.can_edit:
            return
        if not self.selected_visit_id:
            messagebox.showinfo("Visitas", "Selecciona una visita para reportar normas.")
            return

        selected_norms = [token for token, variable in self.visit_norm_vars.items() if variable.get()]
        if not selected_norms:
            messagebox.showwarning("Visitas", "Selecciona al menos una norma aplicada.")
            return

        viewer_name = str((self.controller.current_user or {}).get("name", "")).strip()
        try:
            self.controller.save_visit_norm_report(self.selected_visit_id, selected_norms, viewer_name)
        except ValueError as error:
            messagebox.showerror("Visitas", str(error))
            return

        now_str = datetime.now().strftime("%H:%M")
        self.controller.mark_visit_finalized(self.selected_visit_id, now_str)

        self.refresh()
        visit = next((item for item in self.controller.list_visits() if item.get("id") == self.selected_visit_id), None)
        self._show_readonly_visit_details(visit)
        messagebox.showinfo("Visitas", f"Visita finalizada y reporte de normas guardado.\nHora de finalización: {now_str}")

    def _show_readonly_visit_details(self, visit: dict | None) -> None:
        if self.notes_box is None:
            return

        if visit is None:
            self.notes_box.configure(state="normal")
            self.notes_box.delete("1.0", "end")
            self.notes_box.insert(
                "1.0",
                "Selecciona una visita para consultar la informacion principal y observaciones registradas.",
            )
            self.notes_box.configure(state="disabled")
            self._set_accept_button_state(None)
            self._render_visit_norm_checklist(None)
            if self._agreement_banner is not None:
                self._agreement_banner.grid_remove()
            return

        assignment_time = str(visit.get("assignment_time", "")).strip() or "--"
        departure_time = str(visit.get("departure_time", "")).strip() or "--"
        notes_text = str(visit.get("notes", "")).strip() or "Sin observaciones"
        inspectors_text = str(visit.get("inspectors_text", visit.get("inspector", "--"))).strip() or "--"
        client_text = str(visit.get("client", "--")).strip() or "--"
        date_text = str(visit.get("visit_date", "--")).strip() or "--"
        status_text = str(visit.get("status", "--")).strip() or "--"
        address_text = str(visit.get("address", "--")).strip() or "--"
        warehouse_text = self.controller.get_client_warehouse_for_address(client_text, address_text)

        current_user = self.controller.current_user or {}
        is_exec = self.controller.is_executive_role(current_user)
        finalized_at = str(visit.get("finalized_at", "")).strip()

        finalization_section = ""
        if not is_exec and finalized_at:
            finalization_section = f"\n\n5) FINALIZACIÓN\n• Visita finalizada a las: {finalized_at}"

        warehouse_line = ""
        if warehouse_text and warehouse_text.casefold() not in address_text.casefold():
            warehouse_line = f"• Almacen: {warehouse_text}\n"

        text = (
            "1) INFORMACION PRINCIPAL\n"
            f"• Ejecutivos tecnicos: {inspectors_text}\n"
            f"• Cliente: {client_text}\n"
            f"• Fecha: {date_text}\n"
            f"• Estado operativo: {status_text}\n\n"
            "2) HORARIOS\n"
            f"• Hora de asignacion: {assignment_time}\n"
            f"• Hora de salida: {departure_time}\n\n"
            "3) UBICACION\n"
              f"{warehouse_line}"
            f"• Direccion: {address_text}\n\n"
            "4) OBSERVACIONES\n"
            f"{notes_text}"
            f"{finalization_section}"
        )
        self.notes_box.configure(state="normal")
        self.notes_box.delete("1.0", "end")
        self.notes_box.insert("1.0", text)
        self.notes_box.configure(state="disabled")
        self._set_accept_button_state(visit)
        self._render_visit_norm_checklist(visit)
        # Show agreement banner if this client has registered agreements
        if self._agreement_banner is not None:
            try:
                client_name = visit.get("client", "").strip()
                has_agreements = bool(
                    client_name and self.controller.get_client_agreements(client_name)
                )
                if has_agreements:
                    self._agreement_banner.grid()
                else:
                    self._agreement_banner.grid_remove()
            except Exception:
                self._agreement_banner.grid_remove()

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
            self._sync_address_metadata()
        else:
            self.address_var.set("")
            self.warehouse_var.set("Sin almacen")
            self.service_var.set("Sin servicio")

    def _clear_filters(self) -> None:
        self.search_var.set("")
        self.executive_filter_var.set(self._ALL_EXEC_FILTER)
        self.client_filter_var.set(self._ALL_CLIENT_FILTER)
        self.status_filter_var.set("Todos")
        self.refresh()

    def _on_address_change(self, _value: str) -> None:
        self._sync_address_metadata()

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
            self._sync_address_metadata()
            self.date_var.set(normalized_date or visit.get("visit_date", ""))
            self.assignment_time_var.set(visit.get("assignment_time", ""))
            self.departure_time_var.set(visit.get("departure_time", ""))
            self.status_var.set(visit.get("status", "Programada"))
            if self.notes_box is not None:
                self.notes_box.configure(state="normal")
                self.notes_box.delete("1.0", "end")
                self.notes_box.insert("1.0", visit.get("notes", ""))
            self._update_acceptance_details(visit)
            self._refresh_exec_options()
            self._set_form_editable(not is_past)
            # Show agreement banner for admin if this client has recorded agreements
            if self._agreement_banner is not None:
                try:
                    client_name = visit.get("client", "").strip()
                    has_agreements = bool(
                        client_name and self.controller.get_client_agreements(client_name)
                    )
                    if has_agreements:
                        self._agreement_banner.grid()
                    else:
                        self._agreement_banner.grid_remove()
                except Exception:
                    self._agreement_banner.grid_remove()
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

    # ─── Reporte Sábado ────────────────────────────────────────────────────────

    def _build_saturday_report_tab(self, tab) -> None:
        """Tab exclusivo para roles no ejecutivos: reporte de ejecutivos que
        laboraron en sábado (para cálculo de bono de almacén)."""
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)

        # ── Barra de filtros ──────────────────────────────────────────────────
        filter_row = ctk.CTkFrame(tab, fg_color="transparent")
        filter_row.grid(row=0, column=0, padx=18, pady=(14, 8), sticky="ew")

        ctk.CTkLabel(
            filter_row,
            text="Año:",
            font=self.fonts["small_bold"],
            text_color=self.style["texto_oscuro"],
        ).pack(side="left", padx=(0, 6))

        years = [str(y) for y in range(date.today().year - 2, date.today().year + 2)]
        ctk.CTkComboBox(
            filter_row,
            variable=self.saturday_year_var,
            values=years,
            width=110,
            height=36,
            fg_color="#FFFFFF",
            border_color="#D5D8DC",
            button_color=self.style["primario"],
            dropdown_hover_color=self.style["primario"],
        ).pack(side="left", padx=(0, 12))

        ctk.CTkButton(
            filter_row,
            text="Actualizar",
            fg_color=self.style["primario"],
            text_color=self.style["texto_oscuro"],
            hover_color="#D8C220",
            height=36,
            width=110,
            command=self._refresh_saturday_report,
        ).pack(side="left", padx=(0, 12))

        ctk.CTkButton(
            filter_row,
            text="⬇  Descargar CSV",
            fg_color=self.style["secundario"],
            text_color=self.style["texto_claro"],
            hover_color="#1D1D1D",
            height=36,
            width=160,
            command=self._export_saturday_report_csv,
        ).pack(side="left", padx=(0, 12))

        self.saturday_status_label = ctk.CTkLabel(
            filter_row,
            textvariable=self.saturday_status_var,
            font=self.fonts["small"],
            text_color="#6D7480",
        )
        self.saturday_status_label.pack(side="left", padx=(8, 0))

        # ── Tabla ─────────────────────────────────────────────────────────────
        tree_container = ctk.CTkFrame(tab, fg_color="transparent")
        tree_container.grid(row=1, column=0, padx=18, pady=(0, 18), sticky="nsew")
        tree_container.grid_columnconfigure(0, weight=1)
        tree_container.grid_rowconfigure(0, weight=1)

        columns = ("fecha", "dia", "ejecutivo", "cliente", "servicio", "estado", "direccion")
        self.saturday_tree = ttk.Treeview(tree_container, columns=columns, show="headings", height=18)
        self.saturday_tree.grid(row=0, column=0, sticky="nsew")

        sb_v = ttk.Scrollbar(tree_container, orient="vertical", command=self.saturday_tree.yview)
        sb_v.grid(row=0, column=1, sticky="ns")
        sb_h = ttk.Scrollbar(tree_container, orient="horizontal", command=self.saturday_tree.xview)
        sb_h.grid(row=1, column=0, sticky="ew")
        self.saturday_tree.configure(yscrollcommand=sb_v.set, xscrollcommand=sb_h.set)

        headings = {
            "fecha": "Fecha",
            "dia": "Día",
            "ejecutivo": "Ejecutivo Técnico",
            "cliente": "Cliente",
            "servicio": "Servicio",
            "estado": "Estado",
            "direccion": "Dirección",
        }
        widths = {
            "fecha": 105,
            "dia": 80,
            "ejecutivo": 230,
            "cliente": 190,
            "servicio": 110,
            "estado": 110,
            "direccion": 320,
        }
        for col in columns:
            self.saturday_tree.heading(col, text=headings[col])
            self.saturday_tree.column(col, width=widths[col], anchor="w")

        self._refresh_saturday_report()

    # ─── Vacaciones y Talleres ──────────────────────────────────────────────

    def _build_vacations_workshops_tab(self, tab) -> None:
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_columnconfigure(1, weight=1)
        tab.grid_rowconfigure(1, weight=1)

        # ── Vacaciones (izquierda) ────────────────────────────────────────
        vac_frame = ctk.CTkFrame(tab, fg_color=self.style["surface"], corner_radius=18)
        vac_frame.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(18, 8), pady=12)
        vac_frame.grid_columnconfigure(0, weight=1)
        vac_frame.grid_rowconfigure(3, weight=1)

        ctk.CTkLabel(vac_frame, text="Vacaciones de ejecutivos", font=self.fonts["label_bold"],
                     text_color=self.style["texto_oscuro"]).grid(row=0, column=0, padx=14, pady=(14, 8), sticky="w")

        vac_form = ctk.CTkFrame(vac_frame, fg_color="transparent")
        vac_form.grid(row=1, column=0, padx=14, pady=(0, 8), sticky="ew")
        vac_form.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(vac_form, text="Ejecutivo", font=self.fonts["small"], text_color=self.style["texto_oscuro"]).grid(row=0, column=0, sticky="w", padx=(0, 8), pady=2)
        vac_exec_combo = ctk.CTkComboBox(vac_form, variable=self.vac_exec_var,
                                          values=self.controller.get_assignable_inspectors() or ["Sin ejecutivos"],
                                          height=34, fg_color="#FFFFFF", border_color="#D5D8DC",
                                          button_color=self.style["primario"], dropdown_hover_color=self.style["primario"])
        vac_exec_combo.grid(row=0, column=1, sticky="ew", pady=2)

        ctk.CTkLabel(vac_form, text="Inicio", font=self.fonts["small"], text_color=self.style["texto_oscuro"]).grid(row=1, column=0, sticky="w", padx=(0, 8), pady=2)
        ctk.CTkEntry(vac_form, textvariable=self.vac_start_var, height=34, border_color="#D5D8DC",
                     placeholder_text="YYYY-MM-DD").grid(row=1, column=1, sticky="ew", pady=2)

        ctk.CTkLabel(vac_form, text="Fin", font=self.fonts["small"], text_color=self.style["texto_oscuro"]).grid(row=2, column=0, sticky="w", padx=(0, 8), pady=2)
        ctk.CTkEntry(vac_form, textvariable=self.vac_end_var, height=34, border_color="#D5D8DC",
                     placeholder_text="YYYY-MM-DD").grid(row=2, column=1, sticky="ew", pady=2)

        vac_btn_row = ctk.CTkFrame(vac_frame, fg_color="transparent")
        vac_btn_row.grid(row=2, column=0, padx=14, pady=(0, 8), sticky="ew")
        vac_btn_row.grid_columnconfigure(0, weight=1)
        vac_btn_row.grid_columnconfigure(1, weight=1)

        ctk.CTkButton(vac_btn_row, text="Agregar vacacion", fg_color=self.style["primario"],
                       text_color=self.style["texto_oscuro"], hover_color="#D8C220",
                       command=self._add_vacation).grid(row=0, column=0, padx=(0, 4), sticky="ew")
        ctk.CTkButton(vac_btn_row, text="Eliminar seleccion", fg_color=self.style["peligro"],
                       hover_color="#B43C31", command=self._delete_vacation).grid(row=0, column=1, padx=(4, 0), sticky="ew")

        vac_tree_container = ctk.CTkFrame(vac_frame, fg_color="transparent")
        vac_tree_container.grid(row=3, column=0, padx=14, pady=(0, 14), sticky="nsew")
        vac_tree_container.grid_columnconfigure(0, weight=1)
        vac_tree_container.grid_rowconfigure(0, weight=1)

        vac_cols = ("ejecutivo", "inicio", "fin")
        self.vacation_tree = ttk.Treeview(vac_tree_container, columns=vac_cols, show="headings", height=10)
        self.vacation_tree.grid(row=0, column=0, sticky="nsew")
        vac_sb = ttk.Scrollbar(vac_tree_container, orient="vertical", command=self.vacation_tree.yview)
        vac_sb.grid(row=0, column=1, sticky="ns")
        self.vacation_tree.configure(yscrollcommand=vac_sb.set)
        for col, heading, w in [("ejecutivo", "Ejecutivo", 200), ("inicio", "Inicio", 110), ("fin", "Fin", 110)]:
            self.vacation_tree.heading(col, text=heading)
            self.vacation_tree.column(col, width=w, anchor="w")

        # ── Talleres (derecha) ────────────────────────────────────────────
        ws_frame = ctk.CTkFrame(tab, fg_color=self.style["surface"], corner_radius=18)
        ws_frame.grid(row=0, column=1, rowspan=2, sticky="nsew", padx=(8, 18), pady=12)
        ws_frame.grid_columnconfigure(0, weight=1)
        ws_frame.grid_rowconfigure(3, weight=1)

        ctk.CTkLabel(ws_frame, text="Talleres para ejecutivos", font=self.fonts["label_bold"],
                     text_color=self.style["texto_oscuro"]).grid(row=0, column=0, padx=14, pady=(14, 8), sticky="w")

        ws_form = ctk.CTkFrame(ws_frame, fg_color="transparent")
        ws_form.grid(row=1, column=0, padx=14, pady=(0, 8), sticky="ew")
        ws_form.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(ws_form, text="Titulo", font=self.fonts["small"], text_color=self.style["texto_oscuro"]).grid(row=0, column=0, sticky="w", padx=(0, 8), pady=2)
        ctk.CTkEntry(ws_form, textvariable=self.ws_title_var, height=34, border_color="#D5D8DC",
                     placeholder_text="Nombre del taller").grid(row=0, column=1, sticky="ew", pady=2)

        ctk.CTkLabel(ws_form, text="Fecha", font=self.fonts["small"], text_color=self.style["texto_oscuro"]).grid(row=1, column=0, sticky="w", padx=(0, 8), pady=2)
        ctk.CTkEntry(ws_form, textvariable=self.ws_date_var, height=34, border_color="#D5D8DC",
                     placeholder_text="YYYY-MM-DD").grid(row=1, column=1, sticky="ew", pady=2)

        ctk.CTkLabel(ws_form, text="Descripcion", font=self.fonts["small"], text_color=self.style["texto_oscuro"]).grid(row=2, column=0, sticky="w", padx=(0, 8), pady=2)
        ctk.CTkEntry(ws_form, textvariable=self.ws_desc_var, height=34, border_color="#D5D8DC",
                     placeholder_text="Descripcion breve (opcional)").grid(row=2, column=1, sticky="ew", pady=2)

        ws_btn_row = ctk.CTkFrame(ws_frame, fg_color="transparent")
        ws_btn_row.grid(row=2, column=0, padx=14, pady=(0, 8), sticky="ew")
        ws_btn_row.grid_columnconfigure(0, weight=1)
        ws_btn_row.grid_columnconfigure(1, weight=1)

        ctk.CTkButton(ws_btn_row, text="Agregar taller", fg_color=self.style["primario"],
                       text_color=self.style["texto_oscuro"], hover_color="#D8C220",
                       command=self._add_workshop).grid(row=0, column=0, padx=(0, 4), sticky="ew")
        ctk.CTkButton(ws_btn_row, text="Eliminar seleccion", fg_color=self.style["peligro"],
                       hover_color="#B43C31", command=self._delete_workshop).grid(row=0, column=1, padx=(4, 0), sticky="ew")

        ws_tree_container = ctk.CTkFrame(ws_frame, fg_color="transparent")
        ws_tree_container.grid(row=3, column=0, padx=14, pady=(0, 14), sticky="nsew")
        ws_tree_container.grid_columnconfigure(0, weight=1)
        ws_tree_container.grid_rowconfigure(0, weight=1)

        ws_cols = ("titulo", "fecha", "descripcion")
        self.workshop_tree = ttk.Treeview(ws_tree_container, columns=ws_cols, show="headings", height=10)
        self.workshop_tree.grid(row=0, column=0, sticky="nsew")
        ws_sb = ttk.Scrollbar(ws_tree_container, orient="vertical", command=self.workshop_tree.yview)
        ws_sb.grid(row=0, column=1, sticky="ns")
        self.workshop_tree.configure(yscrollcommand=ws_sb.set)
        for col, heading, w in [("titulo", "Titulo", 200), ("fecha", "Fecha", 110), ("descripcion", "Descripcion", 260)]:
            self.workshop_tree.heading(col, text=heading)
            self.workshop_tree.column(col, width=w, anchor="w")

        self._refresh_vacations_table()
        self._refresh_workshops_table()

    def _refresh_vacations_table(self) -> None:
        if self.vacation_tree is None:
            return
        for row in self.vacation_tree.get_children():
            self.vacation_tree.delete(row)
        for v in self.controller.list_vacations():
            self.vacation_tree.insert("", "end", iid=v["id"], values=(
                v.get("executive", ""), v.get("start_date", ""), v.get("end_date", "")))

    def _refresh_workshops_table(self) -> None:
        if self.workshop_tree is None:
            return
        for row in self.workshop_tree.get_children():
            self.workshop_tree.delete(row)
        for w in self.controller.list_workshops():
            self.workshop_tree.insert("", "end", iid=w["id"], values=(
                w.get("title", ""), w.get("date", ""), w.get("description", "")))

    def _add_vacation(self) -> None:
        executive = self.vac_exec_var.get().strip()
        start = self._normalize_date(self.vac_start_var.get())
        end = self._normalize_date(self.vac_end_var.get())
        if not executive or not start or not end:
            messagebox.showerror("Vacaciones", "Completa ejecutivo, fecha inicio y fecha fin (YYYY-MM-DD).")
            return
        try:
            self.controller.save_vacation(executive, start, end)
        except ValueError as e:
            messagebox.showerror("Vacaciones", str(e))
            return
        self.vac_start_var.set("")
        self.vac_end_var.set("")
        self._refresh_vacations_table()
        self._render_month_grid()
        messagebox.showinfo("Vacaciones", f"Vacaciones registradas para {executive}.")

    def _delete_vacation(self) -> None:
        if self.vacation_tree is None:
            return
        selected = self.vacation_tree.selection()
        if not selected:
            messagebox.showinfo("Vacaciones", "Selecciona una vacacion para eliminar.")
            return
        if not messagebox.askyesno("Vacaciones", "Deseas eliminar la vacacion seleccionada?"):
            return
        self.controller.delete_vacation(selected[0])
        self._refresh_vacations_table()
        self._render_month_grid()

    def _add_workshop(self) -> None:
        title = self.ws_title_var.get().strip()
        ws_date = self._normalize_date(self.ws_date_var.get())
        desc = self.ws_desc_var.get().strip()
        if not title or not ws_date:
            messagebox.showerror("Talleres", "Completa titulo y fecha del taller (YYYY-MM-DD).")
            return
        try:
            self.controller.save_workshop(title, ws_date, desc)
        except ValueError as e:
            messagebox.showerror("Talleres", str(e))
            return
        self.ws_title_var.set("")
        self.ws_date_var.set("")
        self.ws_desc_var.set("")
        self._refresh_workshops_table()
        self._render_month_grid()
        messagebox.showinfo("Talleres", f"Taller '{title}' registrado.")

    def _delete_workshop(self) -> None:
        if self.workshop_tree is None:
            return
        selected = self.workshop_tree.selection()
        if not selected:
            messagebox.showinfo("Talleres", "Selecciona un taller para eliminar.")
            return
        if not messagebox.askyesno("Talleres", "Deseas eliminar el taller seleccionado?"):
            return
        self.controller.delete_workshop(selected[0])
        self._refresh_workshops_table()
        self._render_month_grid()

    def _get_saturday_visits(self) -> list[dict]:
        """Devuelve todas las visitas cuya fecha cae en sábado para el año seleccionado."""
        try:
            target_year = int(self.saturday_year_var.get())
        except ValueError:
            target_year = date.today().year

        result: list[dict] = []
        for visit in self.controller.list_visits():
            raw_date = str(visit.get("visit_date", "")).strip()
            if not raw_date:
                continue
            try:
                d = date.fromisoformat(raw_date)
            except ValueError:
                continue
            if d.year != target_year:
                continue
            if d.weekday() != 5:  # 5 = sábado
                continue
            result.append(visit)
        return result

    def _refresh_saturday_report(self) -> None:
        if self.saturday_tree is None:
            return
        for row in self.saturday_tree.get_children():
            self.saturday_tree.delete(row)

        visits = self._get_saturday_visits()

        # Expandir por ejecutivo (un registro por ejecutivo en la visita)
        rows_added = 0
        for visit in sorted(visits, key=lambda v: v.get("visit_date", "")):
            raw_date = str(visit.get("visit_date", ""))
            try:
                d = date.fromisoformat(raw_date)
                dia_str = d.strftime("%d/%m/%Y")
            except ValueError:
                dia_str = raw_date

            inspectors = visit.get("inspectors") or [visit.get("inspector", "--")]
            for inspector in inspectors:
                self.saturday_tree.insert(
                    "",
                    "end",
                    values=(
                        raw_date,
                        dia_str,
                        inspector,
                        visit.get("client", "--"),
                        visit.get("service", "--"),
                        visit.get("status", "--"),
                        visit.get("address", "--"),
                    ),
                )
                rows_added += 1

        if rows_added:
            self.saturday_status_var.set(f"{rows_added} registro(s) encontrado(s).")
        else:
            self.saturday_status_var.set("Sin visitas en sábado para el año seleccionado.")

    def _export_saturday_report_csv(self) -> None:
        visits = self._get_saturday_visits()
        if not visits:
            messagebox.showinfo(
                "Reporte Sábado",
                "No hay visitas en sábado para el año seleccionado.",
                parent=self,
            )
            return

        save_path = filedialog.asksaveasfilename(
            parent=self,
            title="Guardar reporte de sábado",
            defaultextension=".csv",
            filetypes=[("Archivo CSV", "*.csv"), ("Todos los archivos", "*.*")],
            initialfile=f"reporte_sabado_{self.saturday_year_var.get()}.csv",
        )
        if not save_path:
            return

        # Expandir por ejecutivo
        rows: list[dict] = []
        for visit in sorted(visits, key=lambda v: v.get("visit_date", "")):
            inspectors = visit.get("inspectors") or [visit.get("inspector", "")]
            for inspector in inspectors:
                rows.append({
                    "Fecha": visit.get("visit_date", ""),
                    "Ejecutivo Técnico": inspector,
                    "Cliente": visit.get("client", ""),
                    "Servicio": visit.get("service", ""),
                    "Estado": visit.get("status", ""),
                    "Dirección": visit.get("address", ""),
                    "Notas": visit.get("notes", ""),
                })

        fieldnames = ["Fecha", "Ejecutivo Técnico", "Cliente", "Servicio", "Estado", "Dirección", "Notas"]
        try:
            with open(save_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(rows)
            messagebox.showinfo(
                "Reporte Sábado",
                f"Reporte guardado en:\n{save_path}",
                parent=self,
            )
        except OSError as exc:
            messagebox.showerror("Reporte Sábado", f"Error al guardar el archivo:\n{exc}", parent=self)
