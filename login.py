from __future__ import annotations

from pathlib import Path
from tkinter import TclError

import customtkinter as ctk
from PIL import Image


class LoginView(ctk.CTkFrame):
    def __init__(self, master, style: dict, fonts: dict, on_login) -> None:
        super().__init__(master, fg_color=style["surface"], corner_radius=28)
        self.style = style
        self.fonts = fonts
        self.on_login = on_login

        self.username_var = ctk.StringVar()
        self.password_var = ctk.StringVar()
        self.status_var = ctk.StringVar(value="Ingresa con tu usuario y contrasena asignados.")
        self.logo_image: ctk.CTkImage | None = self._load_logo_image()

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._build_ui()

    def _load_logo_image(self) -> ctk.CTkImage | None:
        image_dir = Path(__file__).resolve().parent / "img"
        candidates = [image_dir / "logo.png", image_dir / "Logo.png"]

        for image_path in candidates:
            if not image_path.exists():
                continue
            try:
                source = Image.open(image_path)
            except OSError:
                continue

            max_width = 360
            max_height = 320
            width, height = source.size
            ratio = min(max_width / width, max_height / height, 1)
            target_size = (max(1, int(width * ratio)), max(1, int(height * ratio)))
            prepared = source.copy()
            source.close()
            return ctk.CTkImage(light_image=prepared, dark_image=prepared, size=target_size)

        return None

    def _focus_widget(self, widget) -> None:
        def _apply_focus() -> None:
            try:
                if widget is not None and widget.winfo_exists() and widget.winfo_toplevel().winfo_exists():
                    widget.focus_set()
            except TclError:
                return

        self.after(30, _apply_focus)

    def _build_ui(self) -> None:
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.grid(row=0, column=0, padx=42, pady=42, sticky="nsew")
        content.grid_columnconfigure(0, weight=5)
        content.grid_columnconfigure(1, weight=4)
        content.grid_rowconfigure(0, weight=1)

        visual_panel = ctk.CTkFrame(content, fg_color="#FFFFFF", corner_radius=24, border_width=1, border_color="#E3E6EA")
        visual_panel.grid(row=0, column=0, padx=(0, 18), sticky="nsew")
        visual_panel.grid_columnconfigure(0, weight=1)
        visual_panel.grid_rowconfigure(3, weight=1)

        ctk.CTkLabel(
            visual_panel,
            text="V&C CALIBRACION",
            font=self.fonts["small_bold"],
            text_color=self.style["secundario"],
            fg_color=self.style["primario"],
            corner_radius=18,
            padx=16,
            pady=8,
        ).grid(row=0, column=0, padx=26, pady=(26, 18), sticky="w")

        ctk.CTkLabel(
            visual_panel,
            text="Control central de inspecciones y seguimiento",
            font=self.fonts["title"],
            text_color=self.style["texto_oscuro"],
            justify="left",
            wraplength=430,
        ).grid(row=1, column=0, padx=26, pady=(0, 10), sticky="w")

        ctk.CTkLabel(
            visual_panel,
            text=(
                "Accede para consultar la base principal, revisar visitas, mantener catalogos y generar "
                "documentos despues de completar el formulario de supervision."
            ),
            font=self.fonts["small"],
            text_color="#5B616A",
            justify="left",
            wraplength=430,
        ).grid(row=2, column=0, padx=26, sticky="w")

        if self.logo_image is not None:
            ctk.CTkLabel(
                visual_panel,
                text="",
                image=self.logo_image,
                fg_color="transparent",
            ).grid(row=3, column=0, padx=26, pady=(24, 26), sticky="n")
        else:
            ctk.CTkLabel(
                visual_panel,
                text="V&C",
                font=(self.fonts["title"][0], 48, "bold"),
                text_color=self.style["secundario"],
                fg_color=self.style["primario"],
                corner_radius=32,
                width=240,
                height=180,
            ).grid(row=3, column=0, padx=26, pady=(24, 26), sticky="n")

        form_panel = ctk.CTkFrame(content, fg_color=self.style["fondo"], corner_radius=24)
        form_panel.grid(row=0, column=1, sticky="nsew")
        form_panel.grid_columnconfigure(0, weight=1)
        form_panel.grid_rowconfigure(0, weight=1)
        form_panel.grid_rowconfigure(2, weight=1)

        form_content = ctk.CTkFrame(form_panel, fg_color="transparent")
        form_content.grid(row=1, column=0, sticky="ew")
        form_content.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            form_content,
            text="Acceso al sistema",
            font=self.fonts["subtitle"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=0, column=0, padx=24, pady=(24, 6), sticky="w")

        ctk.CTkLabel(
            form_content,
            text="Ingresa tu usuario y contrasena para continuar.",
            font=self.fonts["small"],
            text_color="#6D7480",
            justify="left",
            wraplength=340,
        ).grid(row=1, column=0, padx=24, pady=(0, 18), sticky="w")

        ctk.CTkLabel(
            form_content,
            text="Usuario",
            font=self.fonts["label"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=2, column=0, padx=24, pady=(0, 8), sticky="w")
        username_entry = ctk.CTkEntry(
            form_content,
            textvariable=self.username_var,
            height=44,
            corner_radius=14,
            border_color="#D5D8DC",
        )
        username_entry.grid(row=3, column=0, padx=24, sticky="ew")

        ctk.CTkLabel(
            form_content,
            text="Contrasena",
            font=self.fonts["label"],
            text_color=self.style["texto_oscuro"],
        ).grid(row=4, column=0, padx=24, pady=(16, 8), sticky="w")
        password_entry = ctk.CTkEntry(
            form_content,
            textvariable=self.password_var,
            show="*",
            height=44,
            corner_radius=14,
            border_color="#D5D8DC",
        )
        password_entry.grid(row=5, column=0, padx=24, sticky="ew")

        ctk.CTkLabel(
            form_content,
            textvariable=self.status_var,
            font=self.fonts["small"],
            text_color="#6D7480",
            justify="left",
            wraplength=340,
        ).grid(row=6, column=0, padx=24, pady=(16, 12), sticky="w")

        ctk.CTkButton(
            form_content,
            text="Ingresar",
            height=46,
            corner_radius=14,
            fg_color=self.style["secundario"],
            hover_color="#1D1D1D",
            text_color=self.style["texto_claro"],
            font=self.fonts["label_bold"],
            command=self._submit,
        ).grid(row=7, column=0, padx=24, pady=(0, 24), sticky="ew")

        username_entry.bind("<Return>", lambda _event: self._submit())
        password_entry.bind("<Return>", lambda _event: self._submit())
        self._focus_widget(username_entry)

    def _submit(self) -> None:
        username = self.username_var.get().strip()
        password = self.password_var.get().strip()
        if not username or not password:
            self.status_var.set("Captura usuario y contrasena para continuar.")
            return

        error_message = self.on_login(username, password)
        if error_message:
            self.status_var.set(error_message)
