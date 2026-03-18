from __future__ import annotations

from tkinter import messagebox, ttk

import customtkinter as ctk


class ConfigurationView(ctk.CTkFrame):
    def __init__(self, master, controller, style: dict, fonts: dict, can_edit: bool) -> None:
        super().__init__(master, fg_color=style["fondo"])
        self.controller = controller
        self.style = style
        self.fonts = fonts
        self.can_edit = can_edit

        self.selected_norm: str | None = None
        self.selected_user: str | None = None
        self.selected_exec_name: str | None = None

        self.norm_nom_var = ctk.StringVar()
        self.norm_name_var = ctk.StringVar()
        self.norm_section_var = ctk.StringVar()
        self.exec_name_var = ctk.StringVar(value="")
        self.exec_status_var = ctk.StringVar(value="Selecciona un ejecutivo tecnico para actualizar sus normas acreditadas.")

        self.user_name_var = ctk.StringVar()
        self.user_username_var = ctk.StringVar()
        self.user_password_var = ctk.StringVar()
        self.user_role_var = ctk.StringVar(value="ejecutivo")
        self.norm_search_var = ctk.StringVar()
        self.user_search_var = ctk.StringVar()

        self.norm_tree: ttk.Treeview | None = None
        self.user_tree: ttk.Treeview | None = None
        self.exec_tree: ttk.Treeview | None = None
        self.exec_norm_frame: ctk.CTkScrollableFrame | None = None
        self.exec_check_vars: dict[str, ctk.BooleanVar] = {}

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self._build_ui()

    def _build_ui(self) -> None:
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 18))
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(header, text="Configuraciones", font=self.fonts["subtitle"], text_color=self.style["texto_oscuro"]).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(
            header,
            text="Administra los catalogos de normas y usuarios que alimentan el sistema.",
            font=self.fonts["small"],
            text_color="#6D7480",
        ).grid(row=1, column=0, sticky="w", pady=(6, 0))

        if not self.can_edit:
            ctk.CTkLabel(
                self,
                text="Este perfil no tiene permisos para modificar configuraciones.",
                font=self.fonts["label_bold"],
                text_color=self.style["advertencia"],
            ).grid(row=1, column=0, sticky="w")
            return

        tabs = ctk.CTkTabview(self, fg_color=self.style["surface"], segmented_button_selected_color=self.style["secundario"])
        tabs.grid(row=1, column=0, sticky="nsew")
        tabs.add("Normas")
        tabs.add("Usuarios")
        tabs.add("Ejecutivos Tecnicos")

        self._build_norm_tab(tabs.tab("Normas"))
        self._build_user_tab(tabs.tab("Usuarios"))
        self._build_executive_tab(tabs.tab("Ejecutivos Tecnicos"))

    def _build_norm_tab(self, tab) -> None:
        tab.grid_columnconfigure(0, weight=3)
        tab.grid_columnconfigure(1, weight=2)
        tab.grid_rowconfigure(0, weight=1)

        table_panel = ctk.CTkFrame(tab, fg_color=self.style["surface"], corner_radius=20)
        table_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 12), pady=12)
        table_panel.grid_columnconfigure(0, weight=1)
        table_panel.grid_rowconfigure(1, weight=1)

        filter_row = ctk.CTkFrame(table_panel, fg_color="transparent")
        filter_row.grid(row=0, column=0, padx=16, pady=(16, 0), sticky="ew")
        filter_row.grid_columnconfigure(0, weight=1)

        norm_search = ctk.CTkEntry(
            filter_row,
            textvariable=self.norm_search_var,
            placeholder_text="Buscar norma, nombre o capitulo",
            height=36,
            border_color="#D5D8DC",
        )
        norm_search.grid(row=0, column=0, sticky="ew")
        norm_search.bind("<Return>", lambda _event: self._refresh_norms())

        ctk.CTkButton(
            filter_row,
            text="Buscar",
            fg_color=self.style["primario"],
            text_color=self.style["texto_oscuro"],
            hover_color="#D8C220",
            command=self._refresh_norms,
        ).grid(row=0, column=1, padx=(10, 0))

        ctk.CTkButton(
            filter_row,
            text="Limpiar",
            fg_color=self.style["fondo"],
            text_color=self.style["texto_oscuro"],
            hover_color="#E9ECEF",
            command=self._clear_norm_search,
        ).grid(row=0, column=2, padx=(8, 0))

        self.norm_tree = ttk.Treeview(table_panel, columns=("nom", "nombre", "capitulo"), show="headings", height=16)
        self.norm_tree.grid(row=1, column=0, sticky="nsew", padx=16, pady=16)
        self.norm_tree.heading("nom", text="Norma")
        self.norm_tree.heading("nombre", text="Nombre")
        self.norm_tree.heading("capitulo", text="Capitulo")
        self.norm_tree.column("nom", width=150, anchor="w")
        self.norm_tree.column("nombre", width=320, anchor="w")
        self.norm_tree.column("capitulo", width=260, anchor="w")
        self.norm_tree.bind("<<TreeviewSelect>>", self._on_norm_select)

        form_panel = ctk.CTkFrame(tab, fg_color=self.style["fondo"], corner_radius=20)
        form_panel.grid(row=0, column=1, sticky="nsew", pady=12)
        form_panel.grid_columnconfigure(0, weight=1)

        self._form_field(form_panel, 0, "Clave NOM", ctk.CTkEntry(form_panel, textvariable=self.norm_nom_var, height=38))
        self._form_field(form_panel, 1, "Nombre", ctk.CTkEntry(form_panel, textvariable=self.norm_name_var, height=38))
        self._form_field(form_panel, 2, "Capitulo", ctk.CTkEntry(form_panel, textvariable=self.norm_section_var, height=38))
        self._action_buttons(form_panel, 6, self.save_norm, self.clear_norm_form, self.delete_norm)

    def _build_user_tab(self, tab) -> None:
        tab.grid_columnconfigure(0, weight=3)
        tab.grid_columnconfigure(1, weight=2)
        tab.grid_rowconfigure(0, weight=1)

        table_panel = ctk.CTkFrame(tab, fg_color=self.style["surface"], corner_radius=20)
        table_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 12), pady=12)
        table_panel.grid_columnconfigure(0, weight=1)
        table_panel.grid_rowconfigure(1, weight=1)

        filter_row = ctk.CTkFrame(table_panel, fg_color="transparent")
        filter_row.grid(row=0, column=0, padx=16, pady=(16, 0), sticky="ew")
        filter_row.grid_columnconfigure(0, weight=1)

        user_search = ctk.CTkEntry(
            filter_row,
            textvariable=self.user_search_var,
            placeholder_text="Buscar nombre, usuario o rol",
            height=36,
            border_color="#D5D8DC",
        )
        user_search.grid(row=0, column=0, sticky="ew")
        user_search.bind("<Return>", lambda _event: self._refresh_users())

        ctk.CTkButton(
            filter_row,
            text="Buscar",
            fg_color=self.style["primario"],
            text_color=self.style["texto_oscuro"],
            hover_color="#D8C220",
            command=self._refresh_users,
        ).grid(row=0, column=1, padx=(10, 0))

        ctk.CTkButton(
            filter_row,
            text="Limpiar",
            fg_color=self.style["fondo"],
            text_color=self.style["texto_oscuro"],
            hover_color="#E9ECEF",
            command=self._clear_user_search,
        ).grid(row=0, column=2, padx=(8, 0))

        self.user_tree = ttk.Treeview(table_panel, columns=("name", "username", "role"), show="headings", height=16)
        self.user_tree.grid(row=1, column=0, sticky="nsew", padx=16, pady=16)
        self.user_tree.heading("name", text="Nombre")
        self.user_tree.heading("username", text="Usuario")
        self.user_tree.heading("role", text="Rol")
        self.user_tree.column("name", width=260, anchor="w")
        self.user_tree.column("username", width=140, anchor="w")
        self.user_tree.column("role", width=100, anchor="w")
        self.user_tree.bind("<<TreeviewSelect>>", self._on_user_select)

        form_panel = ctk.CTkFrame(tab, fg_color=self.style["fondo"], corner_radius=20)
        form_panel.grid(row=0, column=1, sticky="nsew", pady=12)
        form_panel.grid_columnconfigure(0, weight=1)

        self._form_field(form_panel, 0, "Nombre", ctk.CTkEntry(form_panel, textvariable=self.user_name_var, height=38))
        self._form_field(form_panel, 1, "Usuario", ctk.CTkEntry(form_panel, textvariable=self.user_username_var, height=38))
        self._form_field(form_panel, 2, "Contrasena", ctk.CTkEntry(form_panel, textvariable=self.user_password_var, height=38))
        role_combo = ctk.CTkComboBox(
            form_panel,
            variable=self.user_role_var,
            values=["admin", "ejecutivo"],
            height=38,
            fg_color="#FFFFFF",
            border_color="#D5D8DC",
            button_color=self.style["primario"],
            dropdown_hover_color=self.style["primario"],
        )
        self._form_field(form_panel, 3, "Rol", role_combo)
        self._action_buttons(form_panel, 8, self.save_user, self.clear_user_form, self.delete_user)

    def _build_executive_tab(self, tab) -> None:
        tab.grid_columnconfigure(0, weight=3)
        tab.grid_columnconfigure(1, weight=2)
        tab.grid_rowconfigure(0, weight=1)

        table_panel = ctk.CTkFrame(tab, fg_color=self.style["surface"], corner_radius=20)
        table_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 12), pady=12)
        table_panel.grid_columnconfigure(0, weight=1)
        table_panel.grid_rowconfigure(0, weight=1)

        self.exec_tree = ttk.Treeview(table_panel, columns=("name", "norm_count", "norms"), show="headings", height=16)
        self.exec_tree.grid(row=0, column=0, sticky="nsew", padx=16, pady=16)
        self.exec_tree.heading("name", text="Ejecutivo Tecnico")
        self.exec_tree.heading("norm_count", text="Total normas")
        self.exec_tree.heading("norms", text="Normas acreditadas")
        self.exec_tree.column("name", width=240, anchor="w")
        self.exec_tree.column("norm_count", width=110, anchor="center")
        self.exec_tree.column("norms", width=360, anchor="w")
        self.exec_tree.bind("<<TreeviewSelect>>", self._on_exec_select)

        form_panel = ctk.CTkFrame(tab, fg_color=self.style["fondo"], corner_radius=20)
        form_panel.grid(row=0, column=1, sticky="nsew", pady=12)
        form_panel.grid_columnconfigure(0, weight=1)
        form_panel.grid_rowconfigure(5, weight=1)

        ctk.CTkLabel(
            form_panel,
            text="Ejecutivo tecnico seleccionado",
            font=self.fonts["label"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=0, column=0, padx=16, pady=(16, 6), sticky="w")
        ctk.CTkLabel(
            form_panel,
            textvariable=self.exec_name_var,
            font=self.fonts["label_bold"],
            text_color=self.style["secundario"],
        ).grid(row=1, column=0, padx=16, sticky="w")

        ctk.CTkLabel(
            form_panel,
            textvariable=self.exec_status_var,
            font=self.fonts["small"],
            text_color="#6D7480",
            wraplength=320,
            justify="left",
        ).grid(row=2, column=0, padx=16, pady=(8, 10), sticky="w")

        actions = ctk.CTkFrame(form_panel, fg_color="transparent")
        actions.grid(row=3, column=0, padx=16, pady=(10, 12), sticky="ew")
        actions.grid_columnconfigure(0, weight=1)
        actions.grid_columnconfigure(1, weight=1)
        actions.grid_columnconfigure(2, weight=1)

        ctk.CTkButton(
            actions,
            text="Guardar",
            fg_color=self.style["secundario"],
            hover_color="#1D1D1D",
            command=self.save_executive_norms,
        ).grid(row=0, column=0, padx=(0, 6), sticky="ew")
        ctk.CTkButton(
            actions,
            text="Limpiar",
            fg_color=self.style["fondo"],
            text_color=self.style["texto_oscuro"],
            hover_color="#E9ECEF",
            command=self.clear_executive_selection,
        ).grid(row=0, column=1, padx=3, sticky="ew")
        ctk.CTkButton(
            actions,
            text="Eliminar",
            fg_color=self.style["peligro"],
            hover_color="#B43C31",
            command=self.delete_executive_norms,
        ).grid(row=0, column=2, padx=(6, 0), sticky="ew")

        ctk.CTkLabel(
            form_panel,
            text="Normas acreditadas",
            font=self.fonts["label"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=4, column=0, padx=16, pady=(0, 6), sticky="w")

        self.exec_norm_frame = ctk.CTkScrollableFrame(form_panel, fg_color=self.style["surface"], height=300)
        self.exec_norm_frame.grid(row=5, column=0, padx=16, pady=(0, 16), sticky="nsew")
        self.exec_norm_frame.grid_columnconfigure(0, weight=1)

        self._render_executive_norms()

    def _form_field(self, parent, row_index: int, label_text: str, widget) -> None:
        base = row_index * 2
        ctk.CTkLabel(parent, text=label_text, font=self.fonts["label"], text_color=self.style["texto_oscuro"]).grid(row=base, column=0, padx=16, pady=(16 if row_index == 0 else 10, 6), sticky="w")
        widget.grid(row=base + 1, column=0, padx=16, sticky="ew")

    def _action_buttons(self, parent, row_index: int, save_command, clear_command, delete_command) -> None:
        actions = ctk.CTkFrame(parent, fg_color="transparent")
        actions.grid(row=row_index, column=0, padx=16, pady=(18, 16), sticky="ew")
        actions.grid_columnconfigure(0, weight=1)
        actions.grid_columnconfigure(1, weight=1)
        actions.grid_columnconfigure(2, weight=1)

        ctk.CTkButton(actions, text="Guardar", command=save_command, fg_color=self.style["secundario"], hover_color="#1D1D1D").grid(row=0, column=0, padx=(0, 6), sticky="ew")
        ctk.CTkButton(actions, text="Limpiar", command=clear_command, fg_color=self.style["fondo"], text_color=self.style["texto_oscuro"], hover_color="#E9ECEF").grid(row=0, column=1, padx=3, sticky="ew")
        ctk.CTkButton(actions, text="Eliminar", command=delete_command, fg_color=self.style["peligro"], hover_color="#B43C31").grid(row=0, column=2, padx=(6, 0), sticky="ew")

    def refresh(self) -> None:
        self._refresh_norms()
        self._refresh_users()
        self._render_executive_norms()
        self._refresh_executives()

        if self.selected_exec_name:
            current_norms = set(self.controller.get_accredited_norms(self.selected_exec_name))
            for token, variable in self.exec_check_vars.items():
                variable.set(token in current_norms)

    def _refresh_norms(self) -> None:
        if self.norm_tree is None:
            return
        for item in self.norm_tree.get_children():
            self.norm_tree.delete(item)
        query = self.norm_search_var.get().strip().lower()
        for norm in self.controller.norms_catalog:
            if query:
                searchable = " ".join(
                    [
                        str(norm.get("NOM", "")),
                        str(norm.get("NOMBRE", "")),
                        str(norm.get("CAPITULO", "")),
                    ]
                ).lower()
                if query not in searchable:
                    continue
            self.norm_tree.insert("", "end", iid=norm["NOM"], values=(norm.get("NOM", ""), norm.get("NOMBRE", ""), norm.get("CAPITULO", "")))

    def _refresh_users(self) -> None:
        if self.user_tree is None:
            return
        for item in self.user_tree.get_children():
            self.user_tree.delete(item)
        query = self.user_search_var.get().strip().lower()
        for user in self.controller.users_catalog:
            if query:
                searchable = " ".join(
                    [
                        str(user.get("name", "")),
                        str(user.get("username", "")),
                        str(user.get("role", "")),
                    ]
                ).lower()
                if query not in searchable:
                    continue
            self.user_tree.insert("", "end", iid=user["username"], values=(user.get("name", ""), user.get("username", ""), user.get("role", "")))

    def _clear_norm_search(self) -> None:
        self.norm_search_var.set("")
        self._refresh_norms()

    def _clear_user_search(self) -> None:
        self.user_search_var.set("")
        self._refresh_users()

    def _refresh_executives(self) -> None:
        if self.exec_tree is None:
            return

        for item in self.exec_tree.get_children():
            self.exec_tree.delete(item)

        executives = self.controller.get_assignable_inspectors()
        for name in executives:
            norms = self.controller.get_accredited_norms(name)
            norms_text = ", ".join(norms) if norms else "Sin acreditaciones"
            self.exec_tree.insert("", "end", iid=name, values=(name, len(norms), norms_text))

    def _render_executive_norms(self) -> None:
        if self.exec_norm_frame is None:
            return

        for child in self.exec_norm_frame.winfo_children():
            child.destroy()

        self.exec_check_vars.clear()
        for index, item in enumerate(self.controller.get_catalog_norms()):
            token = item["token"]
            variable = ctk.BooleanVar(value=False)
            self.exec_check_vars[token] = variable

            row = ctk.CTkFrame(self.exec_norm_frame, fg_color="transparent")
            row.grid(row=index, column=0, padx=6, pady=4, sticky="ew")
            row.grid_columnconfigure(1, weight=1)

            checkbox = ctk.CTkCheckBox(
                row,
                text="",
                variable=variable,
                font=self.fonts["small"],
                text_color=self.style["texto_oscuro"],
                checkmark_color=self.style["secundario"],
                fg_color=self.style["primario"],
                hover_color="#D8C220",
            )
            checkbox.grid(row=0, column=0, padx=(0, 8), sticky="nw")

            label = ctk.CTkLabel(
                row,
                text=f"{token} | {item['nombre']}",
                font=self.fonts["small"],
                text_color=self.style["texto_oscuro"],
                justify="left",
                anchor="w",
                wraplength=420,
            )
            label.grid(row=0, column=1, sticky="ew")
            label.bind("<Button-1>", lambda _event, var=variable: var.set(not var.get()))

    def clear_executive_selection(self) -> None:
        self.selected_exec_name = None
        self.exec_name_var.set("")
        self.exec_status_var.set("Selecciona un ejecutivo tecnico para actualizar sus normas acreditadas.")
        for variable in self.exec_check_vars.values():
            variable.set(False)

    def save_executive_norms(self) -> None:
        if not self.selected_exec_name:
            messagebox.showinfo("Ejecutivos", "Selecciona un ejecutivo tecnico para guardar sus normas.")
            return

        selected_norms = [token for token, variable in self.exec_check_vars.items() if variable.get()]
        try:
            existing = self.controller.get_record(self.selected_exec_name)
            self.controller.save_principal_record(
                self.selected_exec_name,
                selected_norms,
                self.selected_exec_name if existing else None,
            )
        except ValueError as error:
            messagebox.showerror("Ejecutivos", str(error))
            return

        self.exec_status_var.set("Normas acreditadas actualizadas correctamente.")
        self.refresh()

    def delete_executive_norms(self) -> None:
        if not self.selected_exec_name:
            messagebox.showinfo("Ejecutivos", "Selecciona un ejecutivo tecnico para eliminar sus normas.")
            return

        if not messagebox.askyesno(
            "Ejecutivos",
            "Deseas eliminar todas las normas acreditadas del ejecutivo tecnico seleccionado?",
        ):
            return

        try:
            existing = self.controller.get_record(self.selected_exec_name)
            if existing is None:
                self.exec_status_var.set("No hay normas acreditadas registradas para este ejecutivo tecnico.")
                return

            self.controller.save_principal_record(
                self.selected_exec_name,
                [],
                self.selected_exec_name,
            )
        except ValueError as error:
            messagebox.showerror("Ejecutivos", str(error))
            return

        self.exec_status_var.set("Se eliminaron todas las normas acreditadas del ejecutivo tecnico.")
        self.refresh()

    def clear_norm_form(self) -> None:
        self.selected_norm = None
        self.norm_nom_var.set("")
        self.norm_name_var.set("")
        self.norm_section_var.set("")

    def clear_user_form(self) -> None:
        self.selected_user = None
        self.user_name_var.set("")
        self.user_username_var.set("")
        self.user_password_var.set("")
        self.user_role_var.set("ejecutivo")

    def save_norm(self) -> None:
        try:
            self.controller.save_norm(
                {
                    "NOM": self.norm_nom_var.get(),
                    "NOMBRE": self.norm_name_var.get(),
                    "CAPITULO": self.norm_section_var.get(),
                },
                self.selected_norm,
            )
        except ValueError as error:
            messagebox.showerror("Normas", str(error))
            return
        self.clear_norm_form()
        self.refresh()

    def delete_norm(self) -> None:
        if not self.selected_norm:
            return
        if not messagebox.askyesno("Normas", "Deseas eliminar la norma seleccionada?"):
            return
        self.controller.delete_norm(self.selected_norm)
        self.clear_norm_form()
        self.refresh()

    def save_user(self) -> None:
        try:
            self.controller.save_user(
                {
                    "name": self.user_name_var.get(),
                    "username": self.user_username_var.get(),
                    "password": self.user_password_var.get(),
                    "role": self.user_role_var.get(),
                },
                self.selected_user,
            )
        except ValueError as error:
            messagebox.showerror("Usuarios", str(error))
            return
        self.clear_user_form()
        self.refresh()

    def delete_user(self) -> None:
        if not self.selected_user:
            return
        if not messagebox.askyesno("Usuarios", "Deseas eliminar el usuario seleccionado?"):
            return
        self.controller.delete_user(self.selected_user)
        self.clear_user_form()
        self.refresh()

    def _on_norm_select(self, _event=None) -> None:
        if self.norm_tree is None:
            return
        selected = self.norm_tree.selection()
        if not selected:
            return
        self.selected_norm = selected[0]
        norm = next((item for item in self.controller.norms_catalog if item.get("NOM") == self.selected_norm), None)
        if norm is None:
            return
        self.norm_nom_var.set(norm.get("NOM", ""))
        self.norm_name_var.set(norm.get("NOMBRE", ""))
        self.norm_section_var.set(norm.get("CAPITULO", ""))

    def _on_user_select(self, _event=None) -> None:
        if self.user_tree is None:
            return
        selected = self.user_tree.selection()
        if not selected:
            return
        self.selected_user = selected[0]
        user = next((item for item in self.controller.users_catalog if item.get("username") == self.selected_user), None)
        if user is None:
            return
        self.user_name_var.set(user.get("name", ""))
        self.user_username_var.set(user.get("username", ""))
        self.user_password_var.set(user.get("password", ""))
        self.user_role_var.set(user.get("role", "ejecutivo"))

    def _on_exec_select(self, _event=None) -> None:
        if self.exec_tree is None:
            return
        selected = self.exec_tree.selection()
        if not selected:
            return

        self.selected_exec_name = selected[0]
        self.exec_name_var.set(self.selected_exec_name)
        current_norms = set(self.controller.get_accredited_norms(self.selected_exec_name))
        for token, variable in self.exec_check_vars.items():
            variable.set(token in current_norms)

        self.exec_status_var.set(
            "Actualiza las normas del ejecutivo tecnico y guarda para reflejar cambios en Principal y Dashboard."
        )

