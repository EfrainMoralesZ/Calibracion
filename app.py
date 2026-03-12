from __future__ import annotations

import os
from datetime import datetime
from tkinter import TclError, filedialog, messagebox, ttk

import customtkinter as ctk

from calendario import CalendarView
from configuraciones import ConfigurationView
from dashboard import DashboardView
from index import CalibrationController
from login import LoginView
from trimestral import TrimestralView


STYLE = {
    "primario": "#ECD925",
    "secundario": "#282828",
    "exito": "#008D53",
    "advertencia": "#ff1500",
    "peligro": "#d74a3d",
    "fondo": "#F8F9FA",
    "surface": "#F8F9FA",
    "texto_oscuro": "#282828",
    "texto_claro": "#ffffff",
    "borde": "#F8F9FA",
}

FONT_TITLE = ("Inter", 22, "bold")
FONT_SUBTITLE = ("Inter", 17, "bold")
FONT_LABEL = ("Inter", 13)
FONT_SMALL = ("Inter", 12)

FONTS = {
    "title": FONT_TITLE,
    "subtitle": FONT_SUBTITLE,
    "label": FONT_LABEL,
    "label_bold": ("Inter", 13, "bold"),
    "small": FONT_SMALL,
    "small_bold": ("Inter", 12, "bold"),
}


def _safe_focus(widget) -> None:
    def _apply_focus() -> None:
        try:
            if widget is not None and widget.winfo_exists():
                widget.focus_set()
        except TclError:
            return

    try:
        if widget is not None and widget.winfo_exists():
            widget.after(20, _apply_focus)
    except TclError:
        return


class InspectorEditDialog(ctk.CTkToplevel):
    def __init__(self, master, controller, inspector_name: str | None, on_saved) -> None:
        super().__init__(master)
        self.controller = controller
        self.inspector_name = inspector_name
        self.on_saved = on_saved
        self.name_var = ctk.StringVar(value=inspector_name or "")
        self.check_vars: dict[str, ctk.BooleanVar] = {}

        self.title("Inspector")
        self.geometry("720x620")
        self.resizable(False, False)
        self.configure(fg_color=STYLE["fondo"])
        self.transient(master)
        self.grab_set()

        record = controller.get_record(inspector_name) if inspector_name else None
        selected_norms = set(controller.get_accredited_norms(record)) if record else set()
        catalog = controller.get_catalog_norms()

        wrapper = ctk.CTkFrame(self, fg_color=STYLE["surface"], corner_radius=24)
        wrapper.pack(fill="both", expand=True, padx=20, pady=20)
        wrapper.grid_columnconfigure(0, weight=1)

        title = "Editar inspector" if inspector_name else "Nuevo inspector"
        ctk.CTkLabel(wrapper, text=title, font=FONTS["subtitle"], text_color=STYLE["texto_oscuro"]).grid(row=0, column=0, padx=20, pady=(20, 6), sticky="w")
        ctk.CTkLabel(
            wrapper,
            text="Define el nombre y las normas acreditadas que deben mostrarse en la base principal.",
            font=FONTS["small"],
            text_color="#6D7480",
        ).grid(row=1, column=0, padx=20, sticky="w")

        ctk.CTkLabel(wrapper, text="Nombre del inspector", font=FONTS["label"], text_color=STYLE["texto_oscuro"]).grid(row=2, column=0, padx=20, pady=(20, 8), sticky="w")
        name_entry = ctk.CTkEntry(wrapper, textvariable=self.name_var, height=40, border_color="#D5D8DC")
        name_entry.grid(row=3, column=0, padx=20, sticky="ew")

        ctk.CTkLabel(wrapper, text="Normas acreditadas", font=FONTS["label"], text_color=STYLE["texto_oscuro"]).grid(row=4, column=0, padx=20, pady=(18, 8), sticky="w")
        norms_frame = ctk.CTkScrollableFrame(wrapper, fg_color=STYLE["fondo"], height=360)
        norms_frame.grid(row=5, column=0, padx=20, sticky="ew")
        norms_frame.grid_columnconfigure(0, weight=1)
        norms_frame.grid_columnconfigure(1, weight=1)

        for index, norm in enumerate(catalog):
            variable = ctk.BooleanVar(value=norm["token"] in selected_norms)
            self.check_vars[norm["token"]] = variable
            checkbox = ctk.CTkCheckBox(
                norms_frame,
                text=f"{norm['token']}  |  {norm['nombre']}",
                variable=variable,
                font=FONTS["small"],
                text_color=STYLE["texto_oscuro"],
                checkmark_color=STYLE["secundario"],
                fg_color=STYLE["primario"],
                hover_color="#D8C220",
            )
            checkbox.grid(row=index, column=0, sticky="w", pady=6, padx=8)

        actions = ctk.CTkFrame(wrapper, fg_color="transparent")
        actions.grid(row=6, column=0, padx=20, pady=20, sticky="ew")
        actions.grid_columnconfigure(0, weight=1)
        actions.grid_columnconfigure(1, weight=1)

        ctk.CTkButton(
            actions,
            text="Cancelar",
            fg_color=STYLE["fondo"],
            text_color=STYLE["texto_oscuro"],
            hover_color="#E9ECEF",
            command=self.destroy,
        ).grid(row=0, column=0, padx=(0, 8), sticky="ew")
        ctk.CTkButton(
            actions,
            text="Guardar inspector",
            fg_color=STYLE["secundario"],
            hover_color="#1D1D1D",
            command=self._save,
        ).grid(row=0, column=1, padx=(8, 0), sticky="ew")

        _safe_focus(name_entry)

    def _save(self) -> None:
        selected_norms = [token for token, variable in self.check_vars.items() if variable.get()]
        try:
            self.controller.save_principal_record(
                self.name_var.get(),
                selected_norms,
                self.inspector_name,
            )
        except ValueError as error:
            messagebox.showerror("Principal", str(error), parent=self)
            return

        self.on_saved()
        self.destroy()


class NormSelectionDialog(ctk.CTkToplevel):
    def __init__(
        self,
        master,
        inspector_name: str,
        norms: list[str],
        selected_norm: str | None,
        on_select,
    ) -> None:
        super().__init__(master)
        self.on_select = on_select
        self.norms = norms or ["Sin norma"]
        default_norm = selected_norm if selected_norm in self.norms else self.norms[0]
        self.norm_var = ctk.StringVar(value=default_norm)

        self.title(f"Seleccion de formulario - {inspector_name}")
        self.geometry("540x260")
        self.resizable(False, False)
        self.configure(fg_color=STYLE["fondo"])
        self.transient(master)
        self.grab_set()

        wrapper = ctk.CTkFrame(self, fg_color=STYLE["surface"], corner_radius=20)
        wrapper.pack(fill="both", expand=True, padx=18, pady=18)
        wrapper.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            wrapper,
            text="Selecciona la norma para elaborar el formulario",
            font=FONTS["label_bold"],
            text_color=STYLE["texto_oscuro"],
        ).grid(row=0, column=0, padx=18, pady=(18, 8), sticky="w")

        norm_combo = ctk.CTkComboBox(
            wrapper,
            variable=self.norm_var,
            values=self.norms,
            height=40,
            fg_color="#FFFFFF",
            border_color="#D5D8DC",
            button_color=STYLE["primario"],
            dropdown_hover_color=STYLE["primario"],
        )
        norm_combo.grid(row=1, column=0, padx=18, sticky="ew")

        actions = ctk.CTkFrame(wrapper, fg_color="transparent")
        actions.grid(row=2, column=0, padx=18, pady=(18, 18), sticky="ew")
        actions.grid_columnconfigure(0, weight=1)
        actions.grid_columnconfigure(1, weight=1)

        ctk.CTkButton(
            actions,
            text="Cancelar",
            fg_color=STYLE["fondo"],
            text_color=STYLE["texto_oscuro"],
            hover_color="#E9ECEF",
            command=self.destroy,
        ).grid(row=0, column=0, padx=(0, 8), sticky="ew")
        ctk.CTkButton(
            actions,
            text="Continuar",
            fg_color=STYLE["secundario"],
            hover_color="#1D1D1D",
            command=self._confirm,
        ).grid(row=0, column=1, padx=(8, 0), sticky="ew")

        _safe_focus(norm_combo)

    def _confirm(self) -> None:
        self.on_select(self.norm_var.get().strip() or "Sin norma")
        self.destroy()


class EvaluationDialog(ctk.CTkToplevel):
    def __init__(
        self,
        master,
        controller,
        inspector_name: str,
        can_edit: bool,
        on_saved,
        initial_norm: str | None = None,
    ) -> None:
        super().__init__(master)
        self.controller = controller
        self.inspector_name = inspector_name
        self.can_edit = can_edit
        self.on_saved = on_saved
        self.initial_norm = (initial_norm or "").strip()

        self.norm_var = ctk.StringVar()
        self.client_var = ctk.StringVar()
        self.date_var = ctk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.score_var = ctk.StringVar()
        self.status_var = ctk.StringVar(value="Estable")
        self.form_status_var = ctk.StringVar()

        self.notes_box: ctk.CTkTextbox | None = None
        self.actions_box: ctk.CTkTextbox | None = None
        self.norm_selector: ctk.CTkComboBox | None = None
        self.form_widgets: list = []
        self.download_buttons: list[ctk.CTkButton] = []

        self.title(f"Acciones - {inspector_name}")
        self.geometry("860x680")
        self.configure(fg_color=STYLE["fondo"])
        self.transient(master)
        self.grab_set()

        self._build_ui()
        self._load_latest()

    def _build_ui(self) -> None:
        wrapper = ctk.CTkFrame(self, fg_color=STYLE["surface"], corner_radius=24)
        wrapper.pack(fill="both", expand=True, padx=20, pady=20)
        wrapper.grid_columnconfigure(0, weight=3)
        wrapper.grid_columnconfigure(1, weight=2)
        wrapper.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(wrapper, text=self.inspector_name, font=FONTS["subtitle"], text_color=STYLE["texto_oscuro"]).grid(row=0, column=0, columnspan=2, padx=20, pady=(20, 6), sticky="w")
        ctk.CTkLabel(
            wrapper,
            text="Completa el formulario de supervision para liberar la descarga de los PDF operativos.",
            font=FONTS["small"],
            text_color="#6D7480",
        ).grid(row=1, column=0, columnspan=2, padx=20, sticky="w")

        form_panel = ctk.CTkFrame(wrapper, fg_color=STYLE["fondo"], corner_radius=22)
        form_panel.grid(row=2, column=0, padx=(20, 10), pady=20, sticky="nsew")
        form_panel.grid_columnconfigure(0, weight=1)

        norms = self.controller.get_accredited_norms(self.inspector_name) or self.controller.get_norm_tokens() or ["Sin norma"]
        clients = self.controller.get_client_names() or ["Sin cliente"]
        if self.initial_norm and self.initial_norm not in norms:
            norms.insert(0, self.initial_norm)
        self.norm_var.set(self.initial_norm if self.initial_norm else norms[0])

        self.norm_selector = ctk.CTkComboBox(
            form_panel,
            variable=self.norm_var,
            values=norms,
            height=38,
            fg_color="#FFFFFF",
            border_color="#D5D8DC",
            button_color=STYLE["primario"],
            dropdown_hover_color=STYLE["primario"],
            command=lambda _value: self._on_norm_change(),
        )
        self._field(form_panel, 0, "Norma evaluada", self.norm_selector)
        self._field(form_panel, 1, "Cliente", ctk.CTkComboBox(
            form_panel,
            variable=self.client_var,
            values=clients,
            height=38,
            fg_color="#FFFFFF",
            border_color="#D5D8DC",
            button_color=STYLE["primario"],
            dropdown_hover_color=STYLE["primario"],
        ))
        self._field(form_panel, 2, "Fecha", ctk.CTkEntry(form_panel, textvariable=self.date_var, height=38, border_color="#D5D8DC"))
        self._field(form_panel, 3, "Puntaje (%)", ctk.CTkEntry(form_panel, textvariable=self.score_var, height=38, border_color="#D5D8DC"))
        self._field(form_panel, 4, "Estado", ctk.CTkComboBox(
            form_panel,
            variable=self.status_var,
            values=["Estable", "En seguimiento", "Critico"],
            height=38,
            fg_color="#FFFFFF",
            border_color="#D5D8DC",
            button_color=STYLE["primario"],
            dropdown_hover_color=STYLE["primario"],
        ))

        ctk.CTkLabel(form_panel, text="Observaciones", font=FONTS["label"], text_color=STYLE["texto_oscuro"]).grid(row=10, column=0, padx=18, pady=(10, 6), sticky="w")
        self.notes_box = ctk.CTkTextbox(form_panel, height=120, corner_radius=16)
        self.notes_box.grid(row=11, column=0, padx=18, sticky="ew")
        ctk.CTkLabel(form_panel, text="Acciones correctivas", font=FONTS["label"], text_color=STYLE["texto_oscuro"]).grid(row=12, column=0, padx=18, pady=(10, 6), sticky="w")
        self.actions_box = ctk.CTkTextbox(form_panel, height=120, corner_radius=16)
        self.actions_box.grid(row=13, column=0, padx=18, sticky="ew")

        self.form_widgets.extend(
            [
                child
                for child in form_panel.winfo_children()
                if isinstance(child, (ctk.CTkEntry, ctk.CTkComboBox, ctk.CTkTextbox))
            ]
        )

        button_row = ctk.CTkFrame(form_panel, fg_color="transparent")
        button_row.grid(row=14, column=0, padx=18, pady=18, sticky="ew")
        button_row.grid_columnconfigure(0, weight=1)
        button_row.grid_columnconfigure(1, weight=1)

        save_button = ctk.CTkButton(
            button_row,
            text="Guardar formulario",
            fg_color=STYLE["secundario"],
            hover_color="#1D1D1D",
            command=self._save_form,
        )
        save_button.grid(row=0, column=0, padx=(0, 8), sticky="ew")
        close_button = ctk.CTkButton(
            button_row,
            text="Cerrar",
            fg_color=STYLE["fondo"],
            text_color=STYLE["texto_oscuro"],
            hover_color="#E9ECEF",
            command=self.destroy,
        )
        close_button.grid(row=0, column=1, padx=(8, 0), sticky="ew")

        side_panel = ctk.CTkFrame(wrapper, fg_color=STYLE["fondo"], corner_radius=22)
        side_panel.grid(row=2, column=1, padx=(10, 20), pady=20, sticky="nsew")
        side_panel.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(side_panel, text="Descarga de documentos", font=FONTS["label_bold"], text_color=STYLE["texto_oscuro"]).grid(row=0, column=0, padx=18, pady=(18, 8), sticky="w")
        ctk.CTkLabel(
            side_panel,
            textvariable=self.form_status_var,
            font=FONTS["small"],
            text_color="#6D7480",
            wraplength=260,
            justify="left",
        ).grid(row=1, column=0, padx=18, sticky="w")

        formato_button = ctk.CTkButton(
            side_panel,
            text="Descargar FormatoSupervicion",
            command=lambda: self._download_document("formato"),
            fg_color=STYLE["primario"],
            text_color=STYLE["texto_oscuro"],
            hover_color="#D8C220",
        )
        formato_button.grid(row=2, column=0, padx=18, pady=(18, 10), sticky="ew")

        criterio_button = ctk.CTkButton(
            side_panel,
            text="Descargar CriterioEvaluacionTecnica",
            command=lambda: self._download_document("criterio"),
            fg_color=STYLE["secundario"],
            hover_color="#1D1D1D",
        )
        criterio_button.grid(row=3, column=0, padx=18, sticky="ew")
        self.download_buttons = [formato_button, criterio_button]

        ctk.CTkLabel(side_panel, text="Regla", font=FONTS["label_bold"], text_color=STYLE["texto_oscuro"]).grid(row=4, column=0, padx=18, pady=(22, 8), sticky="w")
        rule_text = (
            "Los documentos solo se habilitan cuando el formulario queda guardado. "
            "El archivo se almacena en el historico del ejecutivo y puedes descargarlo a otra ruta si lo requieres."
        )
        rule_box = ctk.CTkTextbox(side_panel, height=180, corner_radius=18)
        rule_box.grid(row=5, column=0, padx=18, pady=(0, 18), sticky="ew")
        rule_box.insert("1.0", rule_text)
        rule_box.configure(state="disabled")

        if not self.can_edit:
            save_button.configure(state="disabled")
            for widget in self.form_widgets:
                try:
                    widget.configure(state="disabled")
                except Exception:
                    pass

    def _field(self, parent, index: int, label: str, widget) -> None:
        base_row = index * 2
        ctk.CTkLabel(parent, text=label, font=FONTS["label"], text_color=STYLE["texto_oscuro"]).grid(row=base_row, column=0, padx=18, pady=(18 if index == 0 else 10, 6), sticky="w")
        widget.grid(row=base_row + 1, column=0, padx=18, sticky="ew")

    def _load_latest(self) -> None:
        selected_norm = self.norm_var.get().strip() or self.initial_norm
        latest = self.controller.get_latest_evaluation(self.inspector_name, selected_norm)
        if latest:
            self.norm_var.set(latest.get("selected_norm", self.norm_var.get()))
            self.client_var.set(latest.get("client", self.client_var.get()))
            self.date_var.set(latest.get("visit_date", self.date_var.get()))
            self.score_var.set(str(latest.get("score", "")))
            self.status_var.set(latest.get("status", self.status_var.get()))
            if self.notes_box is not None:
                self.notes_box.delete("1.0", "end")
                self.notes_box.insert("1.0", latest.get("observations", ""))
            if self.actions_box is not None:
                self.actions_box.delete("1.0", "end")
                self.actions_box.insert("1.0", latest.get("corrective_actions", ""))

        self._sync_download_state()

    def _save_form(self) -> None:
        observations = self.notes_box.get("1.0", "end").strip() if self.notes_box is not None else ""
        corrective_actions = self.actions_box.get("1.0", "end").strip() if self.actions_box is not None else ""
        try:
            self.controller.save_evaluation(
                self.inspector_name,
                {
                    "selected_norm": self.norm_var.get(),
                    "client": self.client_var.get(),
                    "visit_date": self.date_var.get(),
                    "score": self.score_var.get(),
                    "status": self.status_var.get(),
                    "observations": observations,
                    "corrective_actions": corrective_actions,
                    "evaluator": (self.controller.current_user or {}).get("name", "Sistema"),
                },
            )
        except ValueError as error:
            messagebox.showerror("Formulario", str(error), parent=self)
            return

        self._sync_download_state()
        self.on_saved()
        messagebox.showinfo("Formulario", "El formulario fue guardado y los documentos quedaron habilitados.", parent=self)

    def _sync_download_state(self) -> None:
        enabled = self.controller.has_completed_form(self.inspector_name, self.norm_var.get())
        message = (
            "Formulario completo. Ya puedes descargar ambos documentos."
            if enabled
            else "Formulario pendiente. Los PDF permanecen bloqueados hasta completar y guardar la captura."
        )
        self.form_status_var.set(message)
        for button in self.download_buttons:
            button.configure(state="normal" if enabled else "disabled")

    def _download_document(self, kind: str) -> None:
        selected_norm = self.norm_var.get().strip() or self.initial_norm
        default_path = self.controller.get_default_document_path(self.inspector_name, kind, selected_norm)
        destination = filedialog.asksaveasfilename(
            parent=self,
            title="Guardar documento",
            defaultextension=".pdf",
            initialdir=str(default_path.parent),
            initialfile=default_path.name,
            filetypes=[("PDF", "*.pdf")],
        )
        if not destination:
            return

        try:
            output = self.controller.generate_document(self.inspector_name, kind, destination, selected_norm)
        except ValueError as error:
            messagebox.showerror("Documentos", str(error), parent=self)
            return

        if hasattr(os, "startfile"):
            try:
                os.startfile(output)
            except OSError:
                pass
        messagebox.showinfo("Documentos", f"Archivo generado en:\n{output}", parent=self)

    def _on_norm_change(self) -> None:
        self._load_latest()


class PrincipalRowActionsDialog(ctk.CTkToplevel):
    def __init__(
        self,
        master,
        inspector_name: str,
        on_form,
        on_edit,
        on_delete,
        can_manage: bool,
    ) -> None:
        super().__init__(master)
        self.on_form = on_form
        self.on_edit = on_edit
        self.on_delete = on_delete
        self.can_manage = can_manage

        self.title("Acciones del inspector")
        self.geometry("380x380")
        self.resizable(False, False)
        self.configure(fg_color=STYLE["fondo"])
        self.transient(master)
        self.grab_set()

        # ── Avatar strip ─────────────────────────────────────────────────
        top_bar = ctk.CTkFrame(self, fg_color=STYLE["secundario"], corner_radius=0, height=72)
        top_bar.pack(fill="x")
        top_bar.pack_propagate(False)
        top_bar.grid_columnconfigure(1, weight=1)

        avatar = ctk.CTkLabel(
            top_bar,
            text="👤",
            font=("Segoe UI Emoji", 28),
            text_color=STYLE["primario"],
            fg_color="#3A3A3A",
            corner_radius=30,
            width=52,
            height=52,
        )
        avatar.grid(row=0, column=0, padx=(18, 12), pady=10)

        name_block = ctk.CTkFrame(top_bar, fg_color="transparent")
        name_block.grid(row=0, column=1, sticky="w")

        ctk.CTkLabel(
            name_block,
            text=inspector_name,
            font=FONTS["label_bold"],
            text_color=STYLE["texto_claro"],
        ).pack(anchor="w")
        ctk.CTkLabel(
            name_block,
            text="Selecciona una accion",
            font=FONTS["small"],
            text_color="#AAAAAA",
        ).pack(anchor="w")

        # ── Action buttons ────────────────────────────────────────────────
        panel = ctk.CTkFrame(self, fg_color="transparent")
        panel.pack(fill="both", expand=True, padx=20, pady=16)
        panel.grid_columnconfigure(0, weight=1)

        def _action_row(row, icon, label, desc, fg, hover, text_color, cmd, enabled=True):
            btn_frame = ctk.CTkFrame(
                panel,
                fg_color=fg,
                corner_radius=14,
                cursor="hand2" if enabled else "arrow",
            )
            btn_frame.grid(row=row, column=0, sticky="ew", pady=5)
            btn_frame.grid_columnconfigure(1, weight=1)

            ctk.CTkLabel(
                btn_frame,
                text=icon,
                font=("Segoe UI Emoji", 20),
                text_color=text_color if enabled else "#AAAAAA",
                width=40,
            ).grid(row=0, column=0, rowspan=2, padx=(12, 8), pady=10)

            ctk.CTkLabel(
                btn_frame,
                text=label,
                font=FONTS["label_bold"],
                text_color=text_color if enabled else "#AAAAAA",
            ).grid(row=0, column=1, sticky="w", pady=(10, 1))

            ctk.CTkLabel(
                btn_frame,
                text=desc,
                font=FONTS["small"],
                text_color=("#CCCCCC" if fg == STYLE["secundario"] else "#888888") if enabled else "#BBBBBB",
            ).grid(row=1, column=1, sticky="w", pady=(0, 10))

            if enabled:
                for child in (btn_frame, *btn_frame.winfo_children()):
                    child.bind("<Enter>", lambda _e, f=btn_frame: f.configure(fg_color=hover))
                    child.bind("<Leave>", lambda _e, f=btn_frame, c=fg: f.configure(fg_color=c))
                    child.bind("<Button-1>", lambda _e, c=cmd: c())

        _action_row(
            0, "📋", "Formulario / PDFs",
            "Captura supervision y descarga documentos",
            STYLE["secundario"], "#3A3A3A", STYLE["texto_claro"],
            self._run_form,
        )
        _action_row(
            1, "✏️", "Editar inspector",
            "Modifica nombre y normas acreditadas",
            "#FFFFFF", "#F0F0F0", STYLE["texto_oscuro"],
            self._run_edit, enabled=self.can_manage,
        )
        _action_row(
            2, "🗑️", "Borrar inspector",
            "Elimina el registro permanentemente",
            "#FFF0EE", "#FADADD", STYLE["peligro"],
            self._run_delete, enabled=self.can_manage,
        )

        # ── Close link ────────────────────────────────────────────────────
        ctk.CTkButton(
            panel,
            text="✕  Cancelar",
            fg_color="transparent",
            text_color="#888888",
            hover_color="#EEEEEE",
            height=32,
            command=self.destroy,
        ).grid(row=3, column=0, pady=(6, 0), sticky="ew")

    def _run_form(self) -> None:
        self.on_form()
        self.destroy()

    def _run_edit(self) -> None:
        if not self.can_manage:
            return
        self.on_edit()
        self.destroy()

    def _run_delete(self) -> None:
        if not self.can_manage:
            return
        self.on_delete()
        self.destroy()


class PrincipalView(ctk.CTkFrame):
    def __init__(self, master, controller, can_edit: bool, on_change) -> None:
        super().__init__(master, fg_color="transparent")
        self.controller = controller
        self.can_edit = can_edit
        self.on_change = on_change

        self.selection_var = ctk.StringVar(value="Selecciona una fila y usa la columna Acciones para operar el inspector.")
        self.tree: ttk.Treeview | None = None
        self.row_cache: dict[str, dict] = {}

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self._build_ui()

    def _build_ui(self) -> None:
        header = ctk.CTkFrame(self, fg_color=STYLE["surface"], corner_radius=24)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(header, text="Principal", font=FONTS["subtitle"], text_color=STYLE["texto_oscuro"]).grid(row=0, column=0, padx=20, pady=(18, 6), sticky="w")
        ctk.CTkLabel(
            header,
            text=(
                "La pantalla principal muestra la tabla de BD-Calibracion.json, el estado del formulario y la columna de acciones para gestionar supervision y documentos."
            ),
            font=FONTS["small"],
            text_color="#6D7480",
            justify="left",
            wraplength=880,
        ).grid(row=1, column=0, padx=20, pady=(0, 14), sticky="w")

        table_panel = ctk.CTkFrame(self, fg_color=STYLE["surface"], corner_radius=24)
        table_panel.grid(row=1, column=0, sticky="nsew")
        table_panel.grid_columnconfigure(0, weight=1)
        table_panel.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(table_panel, textvariable=self.selection_var, font=FONTS["small"], text_color="#6D7480").grid(row=0, column=0, padx=20, pady=(16, 10), sticky="w")

        tree_container = ctk.CTkFrame(table_panel, fg_color="transparent")
        tree_container.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="nsew")
        tree_container.grid_columnconfigure(0, weight=1)
        tree_container.grid_rowconfigure(0, weight=1)

        columns = ("nombre", "norma", "total", "puntaje", "estado", "fecha", "acciones")
        self.tree = ttk.Treeview(tree_container, columns=columns, show="headings", height=18)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)

        headings = {
            "nombre": "Inspector",
            "norma": "Norma",
            "total": "Normas totales",
            "puntaje": "Ultimo puntaje",
            "estado": "Estado",
            "fecha": "Ultima captura",
            "acciones": "Acciones",
        }
        widths = {
            "nombre": 230,
            "norma": 360,
            "total": 110,
            "puntaje": 110,
            "estado": 110,
            "fecha": 120,
            "acciones": 200,
        }
        for column in columns:
            self.tree.heading(column, text=headings[column])
            self.tree.column(column, width=widths[column], anchor="w")

        self.tree.bind("<<TreeviewSelect>>", self._on_selection)
        self.tree.bind("<ButtonRelease-1>", self._on_tree_click)
        self.tree.tag_configure("stable", background="#F0FBF6")
        self.tree.tag_configure("focus", background="#FFF2F0")
        self.tree.tag_configure("pending", background="#FFFBEA")

    def refresh(self) -> None:
        if self.tree is None:
            return

        self.row_cache.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)

        rows = self.controller.get_principal_rows()
        for row in rows:
            if row["status"] == "Estable":
                tag = "stable"
            elif row["status"] == "En enfoque":
                tag = "focus"
            else:
                tag = "pending"
            row_id = row.get("row_id", row["name"])
            self.row_cache[row_id] = row
            self.tree.insert(
                "",
                "end",
                iid=row_id,
                values=(
                    row["name"],
                    row["norms_text"],
                    row["norm_count"],
                    row["latest_score_text"],
                    row["status"],
                    row["latest_date"],
                    row["actions_text"],
                ),
                tags=(tag,),
            )

        if rows:
            first = rows[0].get("row_id", rows[0]["name"])
            self.tree.selection_set(first)
            first_row = self.row_cache.get(first, rows[0])
            self.selection_var.set(f"Registro activo: {first_row['name']} | {first_row.get('norms_text', 'Sin norma')}")
        else:
            self.selection_var.set("No hay registros disponibles en BD-Calibracion.")

    def get_selected_row(self) -> dict | None:
        if self.tree is None:
            return None
        selected = self.tree.selection()
        if not selected:
            return None
        return self.row_cache.get(selected[0])

    def get_selected_name(self) -> str | None:
        row = self.get_selected_row()
        return row["name"] if row else None

    def _on_selection(self, _event=None) -> None:
        selected_row = self.get_selected_row()
        if selected_row:
            self.selection_var.set(f"Registro activo: {selected_row['name']} | {selected_row.get('norms_text', 'Sin norma')}")

    def _on_tree_click(self, event) -> None:
        if self.tree is None:
            return

        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        row_id = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        if not row_id:
            return

        self.tree.selection_set(row_id)
        if column_id != "#7":
            return

        row = self.row_cache.get(row_id)
        if row is None:
            return
        self._open_actions_dialog(row)

    def open_editor(self, inspector_name: str | None) -> None:
        if not self.can_edit:
            return
        InspectorEditDialog(self, self.controller, inspector_name, self._handle_change)

    def open_actions(self, selected_row: dict | None = None) -> None:
        selected_row = selected_row or self.get_selected_row()
        if not selected_row:
            messagebox.showinfo("Principal", "Selecciona un inspector de la tabla.")
            return

        inspector_name = selected_row["name"]
        available_norms = self.controller.get_accredited_norms(inspector_name) or ["Sin norma"]

        def _open_form(norm_token: str) -> None:
            EvaluationDialog(
                self,
                self.controller,
                inspector_name,
                self.can_edit,
                self._handle_change,
                initial_norm=norm_token,
            )

        if len(available_norms) == 1:
            _open_form(available_norms[0])
            return

        NormSelectionDialog(
            self,
            inspector_name,
            available_norms,
            None,
            _open_form,
        )

    def _open_actions_dialog(self, selected_row: dict) -> None:
        inspector_name = selected_row["name"]
        PrincipalRowActionsDialog(
            self,
            inspector_name,
            on_form=lambda: self.open_actions(selected_row),
            on_edit=lambda: self.open_editor(inspector_name),
            on_delete=lambda: self.delete_selected(inspector_name),
            can_manage=self.can_edit,
        )

    def delete_selected(self, inspector_name: str | None = None) -> None:
        selected_name = inspector_name or self.get_selected_name()
        if not self.can_edit or not selected_name:
            return
        if not messagebox.askyesno("Principal", f"Deseas eliminar a {selected_name}?"):
            return
        self.controller.delete_principal_record(selected_name)
        self._handle_change()

    def _handle_change(self) -> None:
        self.refresh()
        self.on_change()


class CalibrationApp(ctk.CTk):
    def __init__(self) -> None:
        ctk.set_appearance_mode("light")
        super().__init__(fg_color=STYLE["fondo"])

        self.controller = CalibrationController()
        self.pages: dict[str, ctk.CTkFrame] = {}
        self.nav_buttons: dict[str, ctk.CTkButton] = {}
        self.summary_labels: dict[str, ctk.CTkLabel] = {}
        self.main_frame: ctk.CTkFrame | None = None
        self.content_frame: ctk.CTkFrame | None = None
        self.current_section = ""

        self.title("Sistema de Calibracion V&C")
        self.geometry("1480x920")
        self.minsize(1200, 760)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._configure_treeview_style()
        self._show_login()

    def _configure_treeview_style(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Treeview",
            background="#FFFFFF",
            fieldbackground="#FFFFFF",
            foreground=STYLE["texto_oscuro"],
            borderwidth=0,
            rowheight=34,
        )
        style.configure(
            "Treeview.Heading",
            background=STYLE["secundario"],
            foreground=STYLE["texto_claro"],
            relief="flat",
            padding=(8, 8),
            font=("Inter", 11, "bold"),
        )
        style.map("Treeview", background=[("selected", "#F2F2F2")], foreground=[("selected", STYLE["texto_oscuro"])])

    def _clear_root(self) -> None:
        for child in self.winfo_children():
            child.destroy()

    def _show_login(self) -> None:
        self._clear_root()
        shell = ctk.CTkFrame(self, fg_color="transparent")
        shell.grid(row=0, column=0, sticky="nsew", padx=28, pady=28)
        shell.grid_columnconfigure(0, weight=1)
        shell.grid_columnconfigure(1, weight=1)
        shell.grid_rowconfigure(0, weight=1)

        hero = ctk.CTkFrame(shell, fg_color=STYLE["secundario"], corner_radius=30)
        hero.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        hero.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(hero, text="V&C", font=("Inter", 38, "bold"), text_color=STYLE["primario"]).grid(row=0, column=0, padx=32, pady=(42, 8), sticky="w")
        ctk.CTkLabel(hero, text="Operacion de calibracion con seguimiento por inspector", font=("Inter", 24, "bold"), text_color=STYLE["texto_claro"], wraplength=440, justify="left").grid(row=1, column=0, padx=32, sticky="w")
        ctk.CTkLabel(
            hero,
            text=(
                "La interfaz concentra la base principal, dashboard por desempeno, asignacion de visitas y catalogos de normas y usuarios. "
                "Cada ejecutivo conserva historial propio y genera documentos solo despues de completar su formulario."
            ),
            font=FONTS["small"],
            text_color="#DADADA",
            wraplength=480,
            justify="left",
        ).grid(row=2, column=0, padx=32, pady=(16, 0), sticky="w")

        stripe = ctk.CTkFrame(hero, fg_color=STYLE["primario"], corner_radius=24, height=180)
        stripe.grid(row=3, column=0, padx=32, pady=(32, 42), sticky="ew")
        stripe.grid_columnconfigure(0, weight=1)
        stripe.grid_columnconfigure(1, weight=1)
        for index, text in enumerate([
            "Principal con tabla de BD-Calibracion",
            "Dashboard con foco en desempeno < 90%",
            "Calendario para visitas con Clientes.json",
            "Configuraciones sobre Normas.json y Usuarios.json",
        ]):
            card = ctk.CTkFrame(stripe, fg_color=STYLE["fondo"], corner_radius=20)
            card.grid(row=index // 2, column=index % 2, padx=12, pady=12, sticky="nsew")
            ctk.CTkLabel(card, text=text, font=FONTS["small_bold"], text_color=STYLE["texto_oscuro"], wraplength=180, justify="left").pack(anchor="w", padx=14, pady=14)

        login_card = LoginView(shell, STYLE, FONTS, self._handle_login)
        login_card.grid(row=0, column=1, sticky="nsew", padx=(14, 0), pady=80)

    def _handle_login(self, username: str, password: str) -> str | None:
        user = self.controller.authenticate(username, password)
        if not user:
            return "Usuario o contrasena incorrectos. Revisa Usuarios.json para validar el acceso."

        self._show_main_shell()
        return None

    def _show_main_shell(self) -> None:
        self._clear_root()
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.grid(row=0, column=0, sticky="nsew", padx=24, pady=24)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(3, weight=1)

        self._build_header(self.main_frame)
        self._build_navigation(self.main_frame)
        self._build_content(self.main_frame)
        self.refresh_all_views()

        available = self.controller.available_sections()
        self.show_section(available[0])

    def _build_header(self, parent) -> None:
        header = ctk.CTkFrame(parent, fg_color=STYLE["surface"], corner_radius=24)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)
        self.summary_labels.clear()

        current_user = self.controller.current_user or {}
        role_text = current_user.get("role", "usuario").upper()

        top_row = ctk.CTkFrame(header, fg_color="transparent")
        top_row.grid(row=0, column=0, padx=18, pady=(14, 8), sticky="ew")
        top_row.grid_columnconfigure(0, weight=1)

        title_group = ctk.CTkFrame(top_row, fg_color="transparent")
        title_group.grid(row=0, column=0, sticky="w")
        title_group.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            title_group,
            text="Sistema de Calibracion V&C",
            font=("Inter", 22, "bold"),
            text_color=STYLE["texto_oscuro"],
        ).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(
            title_group,
            text="Resumen operativo y acceso rapido a las secciones principales.",
            font=FONTS["small"],
            text_color="#7A808B",
        ).grid(row=1, column=0, pady=(2, 0), sticky="w")

        controls_row = ctk.CTkFrame(top_row, fg_color="transparent")
        controls_row.grid(row=0, column=1, sticky="e")

        user_badge = ctk.CTkLabel(
            controls_row,
            text=f"{current_user.get('name', 'Sin sesion')}  |  {role_text}",
            font=FONTS["small_bold"],
            text_color=STYLE["texto_oscuro"],
            fg_color=STYLE["primario"],
            corner_radius=18,
            padx=14,
            pady=6,
        )
        user_badge.grid(row=0, column=0, padx=(0, 10))

        ctk.CTkButton(
            controls_row,
            text="Cerrar sesion",
            fg_color=STYLE["fondo"],
            text_color=STYLE["texto_oscuro"],
            hover_color="#E9ECEF",
            width=112,
            height=34,
            command=self._logout,
        ).grid(row=0, column=1)

        if not self.controller.is_admin():
            return

        metrics_row = ctk.CTkFrame(header, fg_color="transparent")
        metrics_row.grid(row=2, column=0, padx=18, pady=(8, 14), sticky="ew")
        metrics_row.grid_columnconfigure(0, weight=1)
        metrics_row.grid_columnconfigure(1, weight=1)
        metrics_row.grid_columnconfigure(2, weight=1)

        for index, key in enumerate(["inspectors", "average_score", "alerts"]):
            card = ctk.CTkFrame(
                metrics_row,
                fg_color="#FFFFFF",
                corner_radius=16,
                border_width=1,
                border_color="#E3E6EA",
            )
            card.grid(row=0, column=index, padx=(0 if index == 0 else 10, 0), sticky="ew")
            card.grid_columnconfigure(0, weight=1)
            title_map = {
                "inspectors": "Inspectores",
                "average_score": "Promedio global",
                "alerts": "Alertas < 90%",
            }
            ctk.CTkLabel(
                card,
                text=title_map[key],
                font=FONTS["small"],
                text_color="#6D7480",
            ).grid(row=0, column=0, padx=(12, 6), pady=10, sticky="w")
            value_label = ctk.CTkLabel(card, text="--", font=("Inter", 16, "bold"), text_color=STYLE["texto_oscuro"])
            value_label.grid(row=0, column=1, padx=(6, 12), sticky="e")
            self.summary_labels[key] = value_label

    def _build_navigation(self, parent) -> None:
        nav = ctk.CTkFrame(parent, fg_color="transparent")
        nav.grid(row=1, column=0, sticky="ew", pady=(10, 12))
        nav.grid_columnconfigure(0, weight=1)

        button_row = ctk.CTkFrame(nav, fg_color="transparent")
        button_row.grid(row=0, column=0, sticky="w")

        self.nav_buttons.clear()
        for index, section in enumerate(self.controller.available_sections()):
            button = ctk.CTkButton(
                button_row,
                text=section,
                fg_color=STYLE["fondo"],
                text_color=STYLE["texto_oscuro"],
                hover_color="#E9ECEF",
                width=132,
                height=38,
                command=lambda section_name=section: self.show_section(section_name),
            )
            button.grid(row=0, column=index, padx=(0 if index == 0 else 8, 0), pady=0)
            self.nav_buttons[section] = button

    def _build_content(self, parent) -> None:
        self.content_frame = ctk.CTkFrame(parent, fg_color="transparent")
        self.content_frame.grid(row=3, column=0, sticky="nsew")
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(0, weight=1)

        can_edit = self.controller.is_admin()
        self.pages = {
            "Principal": PrincipalView(self.content_frame, self.controller, can_edit, self.refresh_all_views),
            "Calendario": CalendarView(self.content_frame, self.controller, STYLE, FONTS, can_edit),
        }
        if can_edit:
            self.pages["Dashboard"] = DashboardView(self.content_frame, self.controller, STYLE, FONTS)
            self.pages["Trimestral"] = TrimestralView(self.content_frame, self.controller, STYLE, FONTS, True)
            self.pages["Configuraciones"] = ConfigurationView(self.content_frame, self.controller, STYLE, FONTS, True)

        for page in self.pages.values():
            page.grid(row=0, column=0, sticky="nsew")

    def show_section(self, section: str) -> None:
        if section not in self.pages:
            return
        self.current_section = section
        self.pages[section].tkraise()
        for name, button in self.nav_buttons.items():
            if name == section:
                button.configure(fg_color=STYLE["primario"], text_color=STYLE["texto_oscuro"], hover_color="#D8C220")
            else:
                button.configure(fg_color=STYLE["fondo"], text_color=STYLE["texto_oscuro"], hover_color="#E9ECEF")
        self.after_idle(self._refresh_current_section)

    def refresh_all_views(self) -> None:
        self.controller.reload()
        metrics = self.controller.get_overview_metrics()
        self._set_summary_value("inspectors", str(metrics.get("inspectors", 0)))
        average = metrics.get("average_score")
        self._set_summary_value("average_score", f"{average:.1f}%" if average is not None else "--")
        self._set_summary_value("alerts", str(metrics.get("alerts", 0)))

        for page in self.pages.values():
            if not hasattr(page, "refresh"):
                continue
            try:
                if page.winfo_exists():
                    page.refresh()
            except TclError:
                continue

    def _set_summary_value(self, key: str, value: str) -> None:
        label = self.summary_labels.get(key)
        if label is not None:
            label.configure(text=value)

    def _refresh_current_section(self) -> None:
        if not self.current_section:
            return
        page = self.pages.get(self.current_section)
        if page is None or not hasattr(page, "refresh"):
            return
        try:
            if page.winfo_exists():
                page.refresh()
        except TclError:
            return

    def _logout(self) -> None:
        if not messagebox.askyesno("Sesion", "Deseas cerrar la sesion actual?"):
            return
        self.controller.logout()
        self._show_login()


def main() -> None:
    app = CalibrationApp()
    app.mainloop()


if __name__ == "__main__":
    main()


