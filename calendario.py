from __future__ import annotations

import calendar as pycalendar
from datetime import date, datetime
from tkinter import TclError, messagebox, ttk

import customtkinter as ctk


class CalendarView(ctk.CTkFrame):
    def __init__(self, master, controller, style: dict, fonts: dict, can_edit: bool) -> None:
        super().__init__(master, fg_color="transparent")
        self.controller = controller
        self.style = style
        self.fonts = fonts
        self.can_edit = can_edit

        self.selected_visit_id: str | None = None
        self.address_options: list[dict[str, str]] = []
        self.current_month = date.today().replace(day=1)
        self.visible_date_counts: dict[str, int] = {}

        self.inspector_var = ctk.StringVar()
        self.client_var = ctk.StringVar()
        self.address_var = ctk.StringVar()
        self.service_var = ctk.StringVar(value="Sin servicio")
        self.date_var = ctk.StringVar(value=date.today().strftime("%Y-%m-%d"))
        self.status_var = ctk.StringVar(value="Programada")
        self.month_title_var = ctk.StringVar()
        self.search_var = ctk.StringVar()
        self.status_filter_var = ctk.StringVar(value="Todos")

        self.tree: ttk.Treeview | None = None
        self.client_selector: ctk.CTkComboBox | None = None
        self.address_selector: ctk.CTkComboBox | None = None
        self.notes_box: ctk.CTkTextbox | None = None
        self.calendar_grid_frame: ctk.CTkFrame | None = None

        self.grid_columnconfigure(0, weight=3)
        self.grid_columnconfigure(1, weight=2)
        self.grid_rowconfigure(1, weight=1)
        self._build_ui()

    def _build_ui(self) -> None:
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 18))
        header.grid_columnconfigure(0, weight=1)

        title = "Calendario y visitas" if self.can_edit else "Visitas asignadas"
        subtitle = (
            "Asigna visitas a inspectores y visualiza el calendario mensual completo con carga por dia."
            if self.can_edit
            else "Consulta tus visitas asignadas en calendario mensual y revisa el detalle de cada sede."
        )
        ctk.CTkLabel(header, text=title, font=self.fonts["subtitle"], text_color=self.style["texto_oscuro"]).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(header, text=subtitle, font=self.fonts["small"], text_color="#6D7480").grid(row=1, column=0, sticky="w", pady=(6, 0))

        table_panel = ctk.CTkFrame(self, fg_color=self.style["surface"], corner_radius=22)
        table_panel.grid(row=1, column=0, sticky="nsew", padx=(0, 12))
        table_panel.grid_columnconfigure(0, weight=1)
        table_panel.grid_rowconfigure(4, weight=1)

        ctk.CTkLabel(
            table_panel,
            text="Agenda operativa",
            font=self.fonts["label_bold"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=0, column=0, padx=18, pady=(16, 8), sticky="w")

        calendar_panel = ctk.CTkFrame(table_panel, fg_color=self.style["fondo"], corner_radius=18)
        calendar_panel.grid(row=1, column=0, padx=18, pady=(0, 12), sticky="ew")
        calendar_panel.grid_columnconfigure(0, weight=1)

        month_nav = ctk.CTkFrame(calendar_panel, fg_color="transparent")
        month_nav.grid(row=0, column=0, padx=10, pady=(10, 6), sticky="ew")
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
            ).grid(row=0, column=idx, padx=2, pady=(0, 4), sticky="ew")

        self.calendar_grid_frame = ctk.CTkFrame(calendar_panel, fg_color="transparent")
        self.calendar_grid_frame.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="ew")

        filter_row = ctk.CTkFrame(table_panel, fg_color="transparent")
        filter_row.grid(row=2, column=0, padx=18, pady=(0, 10), sticky="ew")
        filter_row.grid_columnconfigure(0, weight=1)

        search_entry = ctk.CTkEntry(
            filter_row,
            textvariable=self.search_var,
            placeholder_text="Buscar por inspector, cliente, direccion o servicio",
            height=38,
            border_color="#D5D8DC",
        )
        search_entry.grid(row=0, column=0, sticky="ew")
        search_entry.bind("<Return>", lambda _event: self.refresh())

        ctk.CTkComboBox(
            filter_row,
            variable=self.status_filter_var,
            values=["Todos", "Programada", "En ruta", "Realizada", "Reprogramada"],
            width=150,
            height=38,
            fg_color="#FFFFFF",
            border_color="#D5D8DC",
            button_color=self.style["primario"],
            dropdown_hover_color=self.style["primario"],
            command=lambda _value: self.refresh(),
        ).grid(row=0, column=1, padx=(10, 0))

        ctk.CTkButton(
            filter_row,
            text="Buscar",
            fg_color=self.style["primario"],
            text_color=self.style["texto_oscuro"],
            hover_color="#D8C220",
            command=self.refresh,
        ).grid(row=0, column=2, padx=(10, 0))

        ctk.CTkButton(
            filter_row,
            text="Limpiar",
            fg_color=self.style["fondo"],
            text_color=self.style["texto_oscuro"],
            hover_color="#E9ECEF",
            command=self._clear_filters,
        ).grid(row=0, column=3, padx=(8, 0))

        tree_container = ctk.CTkFrame(table_panel, fg_color="transparent")
        tree_container.grid(row=4, column=0, padx=18, pady=(0, 18), sticky="nsew")
        tree_container.grid_columnconfigure(0, weight=1)
        tree_container.grid_rowconfigure(0, weight=1)

        columns = ("fecha", "inspector", "cliente", "servicio", "estado", "direccion")
        self.tree = ttk.Treeview(tree_container, columns=columns, show="headings", height=14)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(tree_container, orient="horizontal", command=self.tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=scrollbar.set, xscrollcommand=scroll_x.set)

        headings = {
            "fecha": "Fecha",
            "inspector": "Inspector",
            "cliente": "Cliente",
            "servicio": "Servicio",
            "estado": "Estado",
            "direccion": "Direccion",
        }
        widths = {
            "fecha": 95,
            "inspector": 160,
            "cliente": 190,
            "servicio": 90,
            "estado": 90,
            "direccion": 280,
        }
        for column in columns:
            self.tree.heading(column, text=headings[column])
            self.tree.column(column, width=widths[column], anchor="w")
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        side_panel = ctk.CTkScrollableFrame(self, fg_color=self.style["surface"], corner_radius=22)
        side_panel.grid(row=1, column=1, sticky="nsew")
        side_panel.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            side_panel,
            text="Detalle de visita" if self.can_edit else "Detalle seleccionado",
            font=self.fonts["label_bold"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=0, column=0, padx=18, pady=(16, 14), sticky="w")

        if self.can_edit:
            self._build_form(side_panel)
        else:
            message = ctk.CTkLabel(
                side_panel,
                text=(
                    "Selecciona una visita para revisar direccion, fecha, servicio y observaciones registradas por el admin."
                ),
                font=self.fonts["small"],
                text_color="#6D7480",
                wraplength=280,
                justify="left",
            )
            message.grid(row=1, column=0, padx=18, pady=(0, 12), sticky="w")
            self.notes_box = ctk.CTkTextbox(side_panel, height=260, corner_radius=18)
            self.notes_box.grid(row=2, column=0, padx=18, pady=(0, 18), sticky="nsew")
            self.notes_box.configure(state="disabled")

    def _build_form(self, parent) -> None:
        fields = ctk.CTkFrame(parent, fg_color="transparent")
        fields.grid(row=1, column=0, padx=18, pady=(0, 18), sticky="nsew")
        fields.grid_columnconfigure(0, weight=1)

        self._label_and_widget(fields, 0, "Inspector", self._combo(fields, self.inspector_var, self.controller.get_assignable_inspectors()))
        self.client_selector = self._combo(fields, self.client_var, self.controller.get_client_names(), self._on_client_change)
        self._label_and_widget(fields, 1, "Cliente", self.client_selector)
        self.address_selector = self._combo(fields, self.address_var, ["Selecciona un cliente"], self._on_address_change)
        self._label_and_widget(fields, 2, "Direccion", self.address_selector)
        self._label_and_widget(fields, 3, "Fecha de visita", self._entry(fields, self.date_var))
        self._label_and_widget(fields, 4, "Estado", self._combo(fields, self.status_var, ["Programada", "En ruta", "Realizada", "Reprogramada"]))

        ctk.CTkLabel(fields, text="Servicio", font=self.fonts["label"], text_color=self.style["texto_oscuro"]).grid(row=10, column=0, sticky="w", pady=(8, 6))
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
        service_label.grid(row=11, column=0, sticky="w")

        ctk.CTkLabel(fields, text="Notas", font=self.fonts["label"], text_color=self.style["texto_oscuro"]).grid(row=12, column=0, sticky="w", pady=(12, 6))
        self.notes_box = ctk.CTkTextbox(fields, height=110, corner_radius=16)
        self.notes_box.grid(row=13, column=0, sticky="ew")

        button_row = ctk.CTkFrame(fields, fg_color="transparent")
        button_row.grid(row=14, column=0, sticky="ew", pady=(16, 0))
        button_row.grid_columnconfigure(0, weight=1)
        button_row.grid_columnconfigure(1, weight=1)
        button_row.grid_columnconfigure(2, weight=1)

        ctk.CTkButton(
            button_row,
            text="Nuevo",
            command=self.clear_form,
            fg_color=self.style["fondo"],
            text_color=self.style["texto_oscuro"],
            hover_color="#E9ECEF",
        ).grid(row=0, column=0, padx=(0, 8), sticky="ew")
        ctk.CTkButton(
            button_row,
            text="Guardar",
            command=self.save_visit,
            fg_color=self.style["secundario"],
            hover_color="#1D1D1D",
        ).grid(row=0, column=1, padx=4, sticky="ew")
        ctk.CTkButton(
            button_row,
            text="Eliminar",
            command=self.delete_selected_visit,
            fg_color=self.style["peligro"],
            hover_color="#B43C31",
        ).grid(row=0, column=2, padx=(8, 0), sticky="ew")

    def _label_and_widget(self, parent, row_index: int, label_text: str, widget) -> None:
        base_row = row_index * 2
        ctk.CTkLabel(parent, text=label_text, font=self.fonts["label"], text_color=self.style["texto_oscuro"]).grid(row=base_row, column=0, sticky="w", pady=(0 if row_index == 0 else 8, 6))
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

    def refresh(self) -> None:
        if self.tree is None:
            return

        for item in self.tree.get_children():
            self.tree.delete(item)

        visits = self.controller.list_visits()
        query = self.search_var.get().strip().lower()
        status_filter = self.status_filter_var.get().strip() or "Todos"

        filtered_visits: list[dict] = []
        for visit in visits:
            if status_filter != "Todos" and str(visit.get("status", "")).strip() != status_filter:
                continue

            if query:
                searchable = " ".join(
                    [
                        str(visit.get("inspector", "")),
                        str(visit.get("client", "")),
                        str(visit.get("service", "")),
                        str(visit.get("address", "")),
                        str(visit.get("visit_date", "")),
                        str(visit.get("status", "")),
                    ]
                ).lower()
                if query not in searchable:
                    continue

            filtered_visits.append(visit)

        self.visible_date_counts = {}
        for visit in filtered_visits:
            normalized = self._normalize_date(visit.get("visit_date", ""))
            if normalized:
                self.visible_date_counts[normalized] = self.visible_date_counts.get(normalized, 0) + 1

            self.tree.insert(
                "",
                "end",
                iid=visit["id"],
                values=(
                    visit.get("visit_date", "--"),
                    visit.get("inspector", "--"),
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

        for row in range(6):
            self.calendar_grid_frame.grid_rowconfigure(row, weight=1)
            for col in range(7):
                self.calendar_grid_frame.grid_columnconfigure(col, weight=1)
                day = month_matrix[row][col]
                if day == 0:
                    cell = ctk.CTkButton(
                        self.calendar_grid_frame,
                        text="",
                        height=44,
                        fg_color="#EEF0F2",
                        hover=False,
                        state="disabled",
                    )
                    cell.grid(row=row, column=col, padx=2, pady=2, sticky="ew")
                    continue

                day_iso = date(self.current_month.year, self.current_month.month, day).strftime("%Y-%m-%d")
                count = self.visible_date_counts.get(day_iso, 0)
                label = f"{day}\n{count} visita(s)" if count else str(day)

                if selected_normalized == day_iso:
                    fg_color = self.style["primario"]
                    text_color = self.style["texto_oscuro"]
                elif count:
                    fg_color = "#EAF7EF"
                    text_color = self.style["texto_oscuro"]
                else:
                    fg_color = "#FFFFFF"
                    text_color = self.style["texto_oscuro"]

                cell = ctk.CTkButton(
                    self.calendar_grid_frame,
                    text=label,
                    height=44,
                    fg_color=fg_color,
                    text_color=text_color,
                    hover_color="#E9ECEF",
                    font=self.fonts["small"],
                    command=lambda iso=day_iso: self._select_calendar_date(iso),
                )
                cell.grid(row=row, column=col, padx=2, pady=2, sticky="ew")

    def _select_calendar_date(self, iso_date: str) -> None:
        self.date_var.set(iso_date)
        self._render_month_grid()

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
        self.date_var.set(today.strftime("%Y-%m-%d"))
        self._render_month_grid()

    def clear_form(self) -> None:
        self.selected_visit_id = None
        self.inspector_var.set("")
        self.client_var.set("")
        self.address_var.set("")
        self.service_var.set("Sin servicio")
        self.date_var.set(date.today().strftime("%Y-%m-%d"))
        self.status_var.set("Programada")
        self.address_options = []
        if self.address_selector is not None:
            self.address_selector.configure(values=["Selecciona un cliente"])
        if self.notes_box is not None:
            self.notes_box.delete("1.0", "end")
        self._go_today()

    def save_visit(self) -> None:
        if not self.can_edit:
            return

        notes = self.notes_box.get("1.0", "end").strip() if self.notes_box is not None else ""
        try:
            self.controller.save_visit(
                {
                    "inspector": self.inspector_var.get(),
                    "client": self.client_var.get(),
                    "address": self.address_var.get(),
                    "service": self.service_var.get(),
                    "visit_date": self.date_var.get(),
                    "status": self.status_var.get(),
                    "notes": notes,
                },
                self.selected_visit_id,
            )
        except ValueError as error:
            messagebox.showerror("Visitas", str(error))
            return

        self.clear_form()
        self.refresh()
        messagebox.showinfo("Visitas", "La visita fue guardada correctamente.")

    def delete_selected_visit(self) -> None:
        if not self.can_edit or not self.selected_visit_id:
            return

        if not messagebox.askyesno("Visitas", "Deseas eliminar la visita seleccionada?"):
            return

        self.controller.delete_visit(self.selected_visit_id)
        self.clear_form()
        self.refresh()

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

        self.selected_visit_id = visit_id
        if self.can_edit:
            self.inspector_var.set(visit.get("inspector", ""))
            self.client_var.set(visit.get("client", ""))
            self._on_client_change(visit.get("client", ""))
            self.address_var.set(visit.get("address", ""))
            self.service_var.set(visit.get("service", "Sin servicio"))
            self.date_var.set(normalized_date or visit.get("visit_date", ""))
            self.status_var.set(visit.get("status", "Programada"))
            if self.notes_box is not None:
                self.notes_box.delete("1.0", "end")
                self.notes_box.insert("1.0", visit.get("notes", ""))
        elif self.notes_box is not None:
            text = (
                f"Inspector: {visit.get('inspector', '--')}\n"
                f"Cliente: {visit.get('client', '--')}\n"
                f"Fecha: {visit.get('visit_date', '--')}\n"
                f"Servicio: {visit.get('service', '--')}\n"
                f"Estado: {visit.get('status', '--')}\n"
                f"Direccion: {visit.get('address', '--')}\n\n"
                f"Notas:\n{visit.get('notes', 'Sin observaciones')}"
            )
            self.notes_box.configure(state="normal")
            self.notes_box.delete("1.0", "end")
            self.notes_box.insert("1.0", text)
            self.notes_box.configure(state="disabled")

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

