from __future__ import annotations

import customtkinter as ctk


class LoginView(ctk.CTkFrame):
    def __init__(self, master, style: dict, fonts: dict, on_login) -> None:
        super().__init__(master, fg_color=style["surface"], corner_radius=28)
        self.style = style
        self.fonts = fonts
        self.on_login = on_login

        self.username_var = ctk.StringVar()
        self.password_var = ctk.StringVar()
        self.status_var = ctk.StringVar(value="Ingresa con tu usuario y contrasena asignados.")

        self.grid_columnconfigure(0, weight=1)
        self._build_ui()

    def _build_ui(self) -> None:
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.grid(row=0, column=0, padx=42, pady=42, sticky="nsew")
        content.grid_columnconfigure(0, weight=1)

        badge = ctk.CTkLabel(
            content,
            text="V&C CALIBRACION",
            font=self.fonts["small_bold"],
            text_color=self.style["secundario"],
            fg_color=self.style["primario"],
            corner_radius=18,
            padx=16,
            pady=8,
        )
        badge.grid(row=0, column=0, sticky="w")

        title = ctk.CTkLabel(
            content,
            text="Control central de inspecciones y seguimiento",
            font=self.fonts["title"],
            text_color=self.style["texto_oscuro"],
            justify="left",
        )
        title.grid(row=1, column=0, pady=(18, 10), sticky="w")

        subtitle = ctk.CTkLabel(
            content,
            text=(
                "Accede para consultar la base principal, revisar visitas, mantener catalogos y generar "
                "documentos despues de completar el formulario de supervision."
            ),
            font=self.fonts["small"],
            text_color="#5B616A",
            justify="left",
            wraplength=420,
        )
        subtitle.grid(row=2, column=0, sticky="w")

        form = ctk.CTkFrame(content, fg_color=self.style["fondo"], corner_radius=20)
        form.grid(row=3, column=0, pady=(28, 0), sticky="ew")
        form.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            form,
            text="Usuario",
            font=self.fonts["label"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=0, column=0, padx=20, pady=(20, 8), sticky="w")
        username_entry = ctk.CTkEntry(
            form,
            textvariable=self.username_var,
            height=42,
            corner_radius=14,
            border_color="#D5D8DC",
        )
        username_entry.grid(row=1, column=0, padx=20, sticky="ew")

        ctk.CTkLabel(
            form,
            text="Contrasena",
            font=self.fonts["label"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=2, column=0, padx=20, pady=(16, 8), sticky="w")
        password_entry = ctk.CTkEntry(
            form,
            textvariable=self.password_var,
            show="*",
            height=42,
            corner_radius=14,
            border_color="#D5D8DC",
        )
        password_entry.grid(row=3, column=0, padx=20, sticky="ew")

        ctk.CTkLabel(
            form,
            textvariable=self.status_var,
            font=self.fonts["small"],
            text_color="#6D7480",
            justify="left",
            wraplength=360,
        ).grid(row=4, column=0, padx=20, pady=(16, 12), sticky="w")

        submit_button = ctk.CTkButton(
            form,
            text="Entrar al sistema",
            height=44,
            corner_radius=14,
            fg_color=self.style["secundario"],
            hover_color="#1D1D1D",
            text_color=self.style["texto_claro"],
            font=self.fonts["label_bold"],
            command=self._submit,
        )
        submit_button.grid(row=5, column=0, padx=20, pady=(0, 20), sticky="ew")

        username_entry.bind("<Return>", lambda _event: self._submit())
        password_entry.bind("<Return>", lambda _event: self._submit())
        username_entry.focus_set()

    def _submit(self) -> None:
        username = self.username_var.get().strip()
        password = self.password_var.get().strip()
        if not username or not password:
            self.status_var.set("Captura usuario y contrasena para continuar.")
            return

        error_message = self.on_login(username, password)
        if error_message:
            self.status_var.set(error_message)
