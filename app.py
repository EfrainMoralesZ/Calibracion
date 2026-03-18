from __future__ import annotations

import os
import re
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

BASE_FONTS = {
    "title": FONT_TITLE,
    "subtitle": FONT_SUBTITLE,
    "label": FONT_LABEL,
    "label_bold": ("Inter", 13, "bold"),
    "small": FONT_SMALL,
    "small_bold": ("Inter", 12, "bold"),
}

FONTS = dict(BASE_FONTS)


# ╔═══════════════════════════════════════════════════════════════════════════════╗
# ║ Preguntas del Formulario de Supervisión                                      ║
# ╚═══════════════════════════════════════════════════════════════════════════════╝

PROTOCOL_QUESTIONS = [
    "¿El ejecutivo cumple con su horario establecido?",
    "¿Se cuenta con cobertura completa durante toda la jornada?",
    "¿Porta chaleco limpio y planchado?",
    "¿Se encuentra rasurado y con buena presentación personal?",
    "¿Cumple con el protocolo de vestimenta ejecutiva de V&C?",
    "¿Lleva las botas limpias y en buen estado?",
    "¿Viste la camisa asignada por V&C?",
    "¿Saluda al cliente principal en campo de manera cordial?",
    "¿Se despide del cliente principal en campo de manera cordial? 'Me retiro, algo más que te pueda apoyar'",
    "¿El colaborador escucha las necesidades del cliente?",
    "¿Usa un lenguaje claro sin tecnicismos evitando ambigüedades?",
    "¿Personaliza la comunicación de acuerdo con el perfil?",
    "¿Responde con rapidez, mantiene informado al cliente y responde de manera concisa?",
    "¿Muestra comprensión hacia la situación del cliente?",
    "¿No espera a que el cliente pregunte, se anticipa?",
    "¿Es cordial formal, respetuoso, evita hacer comentarios de los proveedores y del cliente?",
    "¿Cierra confirmando acuerdos?",
    "¿El ejecutivo sabe clasificar entre los diferentes campos de aplicación de las NOMs?",
    "¿Tiene el suficiente conocimiento técnico para evaluar de manera correcta los criterios normativos?",
]

PROCESS_QUESTIONS = [
    "¿El ejecutivo realiza muestreo y revisión de mercancías? ¿Si, cuantas veces?",
    "¿Se está realizando el escaneo de todas las cajas para la correcta identificación de artículos?",
    "¿El ejecutivo valida la correcta colocación de etiquetas en muestreo?",
    "¿Se genera de manera completa y precisa la base de etiquetado NOM para el maquilador?",
    "¿Se valida que la información entregada esté actualizada y libre de errores?",
    "¿Se solicita y archiva la documentación requerida (pedimento y facturas) para cada embarque?",
    "¿Se están tomando fotografías claras y suficientes como evidencia del proceso?",
    "¿La documentación se cargan a la nube conforme a las reglas establecidas (nomenclatura, carpeta, fecha, evidencia)?",
    "Colocar el nombre del técnico que está liberando, así como colocar en el cuerpo del correo el tipo de liberación (preliberación o liberación)",
]

TECHNICAL_RESULT_OPTIONS = ["C", "NC"]


def _scaled_font(font_value: tuple, factor: float) -> tuple:
    family = font_value[0]
    size = int(font_value[1])
    traits = tuple(font_value[2:])
    scaled_size = max(9, int(round(size * factor)))
    return (family, scaled_size, *traits)


def _safe_focus(widget) -> None:
    def _apply_focus() -> None:
        try:
            if widget is not None and widget.winfo_exists() and widget.winfo_toplevel().winfo_exists():
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

        self.title("Ejecutivo Tecnico")
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

        title = "Editar ejecutivo tecnico" if inspector_name else "Nuevo ejecutivo tecnico"
        ctk.CTkLabel(wrapper, text=title, font=FONTS["subtitle"], text_color=STYLE["texto_oscuro"]).grid(row=0, column=0, padx=20, pady=(20, 6), sticky="w")
        ctk.CTkLabel(
            wrapper,
            text="Define el nombre y las normas acreditadas que deben mostrarse en el sistema.",
            font=FONTS["small"],
            text_color="#6D7480",
        ).grid(row=1, column=0, padx=20, sticky="w")

        ctk.CTkLabel(wrapper, text="Nombre del ejecutivo tecnico", font=FONTS["label"], text_color=STYLE["texto_oscuro"]).grid(row=2, column=0, padx=20, pady=(20, 8), sticky="w")
        name_entry = ctk.CTkEntry(wrapper, textvariable=self.name_var, height=40, border_color="#94A3B8")
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
                text=f"{' '.join(str(norm.get('nom', norm['token'])).split())}",
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
            text="Guardar ejecutivo tecnico",
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
        self.controller = getattr(master, "controller", None)
        self.inspector_name = inspector_name
        self.on_select = on_select
        self.norms = norms or ["Sin norma"]
        self.selected_norm = selected_norm if selected_norm in self.norms else ""
        catalog_norms = self.controller.get_catalog_norms() if self.controller is not None else []
        self.norm_labels = {
            item.get("token", ""): item.get("nombre", "")
            for item in catalog_norms
        }
        self.norm_full_codes = {
            item.get("token", ""): str(item.get("nom", item.get("token", ""))).strip()
            for item in catalog_norms
        }

        self.title(f"Seleccion de formulario - {inspector_name}")
        self.geometry("940x640")
        self.minsize(820, 520)
        self.configure(fg_color=STYLE["fondo"])
        self.transient(master)
        self.grab_set()

        wrapper = ctk.CTkFrame(self, fg_color=STYLE["surface"], corner_radius=24)
        wrapper.pack(fill="both", expand=True, padx=18, pady=18)
        wrapper.grid_columnconfigure(0, weight=1)
        wrapper.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(
            wrapper,
            text=f"Formularios disponibles para {inspector_name}",
            font=FONTS["subtitle"],
            text_color=STYLE["texto_oscuro"],
        ).grid(row=0, column=0, padx=20, pady=(20, 6), sticky="w")
        ctk.CTkLabel(
            wrapper,
            text="Cada norma acreditada tiene su propia tarjeta. Usa el boton Formulario para abrir la captura correspondiente.",
            font=FONTS["small"],
            text_color="#6D7480",
            wraplength=780,
            justify="left",
        ).grid(row=1, column=0, padx=20, pady=(0, 12), sticky="w")

        cards_frame = ctk.CTkScrollableFrame(wrapper, fg_color=STYLE["fondo"], corner_radius=18)
        cards_frame.grid(row=2, column=0, padx=20, pady=(0, 16), sticky="nsew")
        for col_index in range(4):
            cards_frame.grid_columnconfigure(col_index, weight=1, uniform="norm_cards")

        full_titles = {token: self._norm_nom(token) for token in self.norms}

        for index, norm_token in enumerate(self.norms):
            row = index // 4
            col = index % 4
            cards_frame.grid_rowconfigure(row, weight=1)
            self._build_selection_card(
                parent=cards_frame,
                row=row,
                col=col,
                icon=self._norm_icon(norm_token),
                badge_text="Formulario",
                title=full_titles[norm_token],
                description="Abre la captura operativa de esta norma acreditada.",
                button_text="Formulario",
                command=lambda token=norm_token: self._open_form(token),
                accent=False,
            )

        history_index = len(self.norms)
        history_row = history_index // 4
        history_col = history_index % 4
        cards_frame.grid_rowconfigure(history_row, weight=1)
        self._build_selection_card(
            parent=cards_frame,
            row=history_row,
            col=history_col,
            icon="📊",
            badge_text="Calificaciones",
            title="Historial de calificaciones",
            description="Consulta registros por norma, fecha, supervisor y estatus.",
            button_text="Ver historial",
            command=self._open_score_history,
            accent=True,
        )

        actions = ctk.CTkFrame(wrapper, fg_color="transparent")
        actions.grid(row=3, column=0, padx=20, pady=(0, 18), sticky="ew")
        actions.grid_columnconfigure(0, weight=1)

        ctk.CTkButton(
            actions,
            text="Cancelar",
            fg_color=STYLE["fondo"],
            text_color=STYLE["texto_oscuro"],
            hover_color="#E9ECEF",
            command=self.destroy,
        ).grid(row=0, column=0, sticky="e")

    def _norm_name(self, norm_token: str) -> str:
        return self.norm_labels.get(norm_token, "Formulario operativo disponible para esta norma.")

    def _norm_nom(self, norm_token: str) -> str:
        return " ".join(str(self.norm_full_codes.get(norm_token, norm_token)).split())

    def _build_selection_card(
        self,
        parent,
        row: int,
        col: int,
        icon: str,
        badge_text: str,
        title: str,
        description: str,
        button_text: str,
        command,
        accent: bool = False,
    ) -> None:
        card_fg = "#FFFBEA" if accent else "#FFFFFF"
        border_color = STYLE["primario"] if accent else "#E3E6EA"
        icon_fg = STYLE["secundario"] if accent else STYLE["primario"]
        icon_text = STYLE["texto_claro"] if accent else STYLE["texto_oscuro"]
        badge_fg = STYLE["primario"] if accent else "#F1F3F5"
        badge_text_color = STYLE["texto_oscuro"] if accent else "#6D7480"
        button_fg = STYLE["primario"] if accent else STYLE["secundario"]
        button_text_color = STYLE["texto_oscuro"] if accent else STYLE["texto_claro"]
        button_hover = "#D8C220" if accent else "#1D1D1D"

        card = ctk.CTkFrame(
            parent,
            fg_color=card_fg,
            corner_radius=18,
            border_width=1,
            border_color=border_color,
        )
        card.grid(row=row, column=col, padx=8, pady=8, sticky="nsew")
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(2, weight=1)

        top_row = ctk.CTkFrame(card, fg_color="transparent")
        top_row.grid(row=0, column=0, padx=14, pady=(14, 10), sticky="ew")
        top_row.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            top_row,
            text=icon,
            font=("Segoe UI Emoji", 22),
            text_color=icon_text,
            fg_color=icon_fg,
            corner_radius=16,
            width=44,
            height=44,
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            top_row,
            text=badge_text,
            font=FONTS["small_bold"],
            text_color=badge_text_color,
            fg_color=badge_fg,
            corner_radius=10,
            padx=8,
            pady=4,
        ).grid(row=0, column=1, sticky="e")

        ctk.CTkLabel(
            card,
            text=title,
            font=FONTS["small_bold"],
            text_color=STYLE["texto_oscuro"],
            justify="left",
            wraplength=170,
            anchor="w",
        ).grid(row=1, column=0, padx=14, pady=(0, 8), sticky="ew")

        ctk.CTkLabel(
            card,
            text=description,
            font=FONTS["small"],
            text_color="#6D7480",
            justify="left",
            wraplength=170,
            anchor="nw",
        ).grid(row=2, column=0, padx=14, pady=(0, 12), sticky="new")

        ctk.CTkButton(
            card,
            text=button_text,
            fg_color=button_fg,
            text_color=button_text_color,
            hover_color=button_hover,
            height=34,
            command=command,
        ).grid(row=3, column=0, padx=14, pady=(0, 14), sticky="ew")

    def _render_score_history(self, parent: ctk.CTkScrollableFrame) -> None:
        headers = ["Norma", "Calificación", "Fecha", "Supervisor", "Estatus"]
        for col, title in enumerate(headers):
            ctk.CTkLabel(
                parent,
                text=title,
                font=FONTS["small_bold"],
                text_color=STYLE["texto_oscuro"],
                fg_color=STYLE["primario"],
                corner_radius=8,
                padx=8,
                pady=4,
            ).grid(row=0, column=col, padx=(0 if col == 0 else 6, 0), pady=(0, 8), sticky="ew")

        if self.controller is None:
            ctk.CTkLabel(
                parent,
                text="No se pudo cargar el historial de calificaciones.",
                font=FONTS["small"],
                text_color="#6D7480",
            ).grid(row=1, column=0, columnspan=5, padx=8, pady=8, sticky="w")
            return

        history_rows = self.controller.get_norm_score_history(self.inspector_name)
        if not history_rows:
            ctk.CTkLabel(
                parent,
                text="Sin calificaciones guardadas para este ejecutivo técnico.",
                font=FONTS["small"],
                text_color="#6D7480",
            ).grid(row=1, column=0, columnspan=5, padx=8, pady=8, sticky="w")
            return

        for idx, row in enumerate(history_rows, start=1):
            visit_date = str(row.get("visit_date", "")).strip() or str(row.get("saved_at", "")).strip() or "--"
            values = [
                str(row.get("norm", "Sin norma")).strip() or "Sin norma",
                f"{float(row.get('score', 0.0)):.1f}%",
                visit_date,
                str(row.get("evaluator", "Sin supervisor")).strip() or "Sin supervisor",
                str(row.get("status", "Sin estatus")).strip() or "Sin estatus",
            ]

            for col, value in enumerate(values):
                anchor = "w" if col in {0, 3, 4} else "center"
                padx = (4, 0) if col == 0 else 6
                ctk.CTkLabel(
                    parent,
                    text=value,
                    font=FONTS["small"],
                    text_color=STYLE["texto_oscuro"],
                    anchor=anchor,
                    justify="left",
                ).grid(row=idx, column=col, padx=padx, pady=(0, 8), sticky="ew")

    def _open_score_history(self) -> None:
        dialog = ctk.CTkToplevel(self)
        dialog.title(f"Historial de calificaciones - {self.inspector_name}")
        dialog.geometry("980x540")
        dialog.minsize(860, 420)
        dialog.configure(fg_color=STYLE["fondo"])
        dialog.transient(self)
        dialog.grab_set()

        wrapper = ctk.CTkFrame(dialog, fg_color=STYLE["surface"], corner_radius=24)
        wrapper.pack(fill="both", expand=True, padx=18, pady=18)
        wrapper.grid_columnconfigure(0, weight=1)
        wrapper.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(
            wrapper,
            text=f"Historial de calificaciones de {self.inspector_name}",
            font=FONTS["subtitle"],
            text_color=STYLE["texto_oscuro"],
        ).grid(row=0, column=0, padx=20, pady=(20, 6), sticky="w")
        ctk.CTkLabel(
            wrapper,
            text="Consulta el detalle guardado por norma, fecha, supervisor y estatus.",
            font=FONTS["small"],
            text_color="#6D7480",
            wraplength=860,
            justify="left",
        ).grid(row=1, column=0, padx=20, pady=(0, 12), sticky="w")

        history_scroll = ctk.CTkScrollableFrame(wrapper, fg_color=STYLE["fondo"], corner_radius=18)
        history_scroll.grid(row=2, column=0, padx=20, pady=(0, 16), sticky="nsew")
        history_scroll.grid_columnconfigure(0, weight=3)
        history_scroll.grid_columnconfigure(1, weight=1)
        history_scroll.grid_columnconfigure(2, weight=1)
        history_scroll.grid_columnconfigure(3, weight=2)
        history_scroll.grid_columnconfigure(4, weight=2)
        self._render_score_history(history_scroll)

        actions = ctk.CTkFrame(wrapper, fg_color="transparent")
        actions.grid(row=3, column=0, padx=20, pady=(0, 18), sticky="ew")
        actions.grid_columnconfigure(0, weight=1)

        ctk.CTkButton(
            actions,
            text="Cerrar",
            fg_color=STYLE["fondo"],
            text_color=STYLE["texto_oscuro"],
            hover_color="#E9ECEF",
            command=dialog.destroy,
        ).grid(row=0, column=0, sticky="e")

    def _norm_icon(self, norm_token: str) -> str:
        norm_name = self._norm_name(norm_token).lower()
        if "textil" in norm_name or "vestir" in norm_name:
            return "👕"
        if "juguete" in norm_name:
            return "🧸"
        if "alimento" in norm_name or "bebida" in norm_name or "atun" in norm_name:
            return "🥫"
        if "cosm" in norm_name:
            return "💄"
        if "aseo" in norm_name or "sanitario" in norm_name:
            return "🧴"
        if "electron" in norm_name or "electrodom" in norm_name:
            return "📦"
        return "📋"

    def _open_form(self, norm_token: str) -> None:
        self.on_select(norm_token.strip() or "Sin norma")
        self.withdraw()
        self.after(120, self.destroy)


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
        self.inspector_var = ctk.StringVar(value=inspector_name)
        self.supervisor_var = ctk.StringVar(value=(controller.current_user or {}).get("name", ""))
        self.score_var = ctk.StringVar(value="")
        self.status_var = ctk.StringVar(value="Estable")
        self.form_status_var = ctk.StringVar(value="")

        self.norm_selector: ctk.CTkComboBox | None = None
        self.client_selector: ctk.CTkComboBox | None = None
        self.supervisor_entry: ctk.CTkEntry | None = None
        self.download_button: ctk.CTkButton | None = None
        self.close_button: ctk.CTkButton | None = None
        self.norm_token_to_display: dict[str, str] = {}
        self.norm_display_to_token: dict[str, str] = {}
        self.applicable_norm_values: list[str] = []
        self._is_loading_latest = False
        self._observed_vars: list[ctk.StringVar] = []

        self.protocol_result_vars: list[ctk.StringVar] = []
        self.protocol_obs_vars: list[ctk.StringVar] = []
        self.process_result_vars: list[ctk.StringVar] = []
        self.process_obs_vars: list[ctk.StringVar] = []

        self.technical_rows_container: ctk.CTkScrollableFrame | None = None
        self.technical_rows: list[dict[str, object]] = []

        self.title(f"Acciones - {inspector_name}")
        self.geometry("960x660")
        self.configure(fg_color=STYLE["fondo"])
        self.transient(master)
        self.grab_set()

        self._build_ui()
        self._configure_form_state_tracking()
        self._load_latest(preload_saved=False)

        # Centre on screen so nothing is clipped on smaller displays
        self.update_idletasks()
        _w, _h = 960, 660
        _x = max(0, (self.winfo_screenwidth() - _w) // 2)
        _y = max(30, (self.winfo_screenheight() - _h) // 2)
        self.geometry(f"{_w}x{_h}+{_x}+{_y}")

    def _build_ui(self) -> None:
        wrapper = ctk.CTkFrame(self, fg_color=STYLE["surface"], corner_radius=24)
        wrapper.pack(fill="both", expand=True, padx=20, pady=20)
        wrapper.grid_columnconfigure(0, weight=1)
        wrapper.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(wrapper, text=f"Supervisión - {self.inspector_name}", font=FONTS["subtitle"], text_color=STYLE["texto_oscuro"]).grid(row=0, column=0, padx=20, pady=(20, 6), sticky="w")
        ctk.CTkLabel(
            wrapper,
            text="Completa los 4 apartados del formulario para generar los documentos operativos.",
            font=FONTS["small"],
            text_color="#6D7480",
        ).grid(row=1, column=0, padx=20, sticky="w")

        tabview = ctk.CTkTabview(wrapper, fg_color=STYLE["fondo"], corner_radius=22)
        tabview.grid(row=2, column=0, padx=20, pady=20, sticky="nsew")
        tabview.add("Información")
        tabview.add("Protocolo y Habilidades")
        tabview.add("Procesos")
        tabview.add("Técnico Normativa")

        self._build_info_tab(tabview.tab("Información"))
        self._build_answers_tab(
            tabview.tab("Protocolo y Habilidades"),
            "SUPERVISIÓN DE PROTOCOLO Y HABILIDADES BLANDAS",
            PROTOCOL_QUESTIONS,
            self.protocol_result_vars,
            self.protocol_obs_vars,
        )
        self._build_answers_tab(
            tabview.tab("Procesos"),
            "SUPERVISIÓN DE PROCESOS",
            PROCESS_QUESTIONS,
            self.process_result_vars,
            self.process_obs_vars,
        )
        self._build_technical_tab(tabview.tab("Técnico Normativa"))

        button_panel = ctk.CTkFrame(wrapper, fg_color="transparent")
        button_panel.grid(row=4, column=0, padx=20, pady=(4, 10), sticky="ew")
        button_panel.grid_columnconfigure(0, weight=1)

        self.close_button = ctk.CTkButton(
            button_panel,
            text="Cerrar",
            fg_color=STYLE["fondo"],
            text_color=STYLE["texto_oscuro"],
            hover_color="#E9ECEF",
            command=self.destroy,
        )
        self.close_button.grid(row=0, column=0, sticky="e")

    def _build_info_tab(self, parent: ctk.CTkFrame) -> None:
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        form_scroll = ctk.CTkScrollableFrame(parent, fg_color=STYLE["fondo"], corner_radius=0)
        form_scroll.grid(row=0, column=0, sticky="nsew")
        form_scroll.grid_columnconfigure(0, weight=1)

        norms = self.controller.get_accredited_norms(self.inspector_name) or self.controller.get_norm_tokens() or ["Sin norma"]
        clients = self.controller.get_client_names() or ["Sin cliente"]
        if self.initial_norm and self.initial_norm not in norms:
            norms.insert(0, self.initial_norm)

        self.norm_token_to_display.clear()
        self.norm_display_to_token.clear()
        catalog = self.controller.get_catalog_norms()
        for item in catalog:
            token = str(item.get("token", "")).strip()
            if not token:
                continue
            nom_value = " ".join(str(item.get("nom", token)).split())
            full_label = nom_value or token
            self.norm_token_to_display[token] = full_label
            self.norm_display_to_token.setdefault(full_label, token)

        options: list[str] = []
        seen_options: set[str] = set()
        for token in norms:
            full_label = self._norm_label_from_value(token)
            if full_label in seen_options:
                continue
            options.append(full_label)
            seen_options.add(full_label)

        if not options:
            options = ["Sin norma"]

        initial_norm_label = self._norm_label_from_value(self.initial_norm or options[0])
        if initial_norm_label not in seen_options:
            options.insert(0, initial_norm_label)
            seen_options.add(initial_norm_label)
        self.norm_var.set(initial_norm_label)
        self._refresh_applicable_norm_values()

        row = 0
        self.norm_selector = ctk.CTkComboBox(
            form_scroll,
            variable=self.norm_var,
            values=options,
            height=38,
            fg_color="#FFFFFF",
            border_color="#94A3B8",
            button_color=STYLE["primario"],
            dropdown_hover_color=STYLE["primario"],
            command=lambda _value: self._on_norm_change(),
        )
        self._add_field(form_scroll, row, "Norma evaluada", self.norm_selector)

        row += 1
        self._add_field(
            form_scroll,
            row,
            "Fecha de supervisión",
            ctk.CTkEntry(form_scroll, textvariable=self.date_var, height=38, border_color="#94A3B8", state="readonly"),
        )

        row += 1
        self.client_selector = ctk.CTkComboBox(
            form_scroll,
            variable=self.client_var,
            values=clients,
            height=38,
            fg_color="#FFFFFF",
            border_color="#94A3B8",
            button_color=STYLE["primario"],
            dropdown_hover_color=STYLE["primario"],
        )
        self._add_field(form_scroll, row, "Cliente / Almacén", self.client_selector)

        row += 1
        self._add_field(
            form_scroll,
            row,
            "Ejecutivo supervisado",
            ctk.CTkEntry(form_scroll, height=38, border_color="#94A3B8", state="readonly", textvariable=self.inspector_var),
        )

        row += 1
        self.supervisor_entry = ctk.CTkEntry(form_scroll, textvariable=self.supervisor_var, height=38, border_color="#94A3B8")
        self._add_field(form_scroll, row, "Nombre del supervisor", self.supervisor_entry)

        ctk.CTkLabel(
            form_scroll,
            text="La calificación final ya no se captura aquí. Revisa las calificaciones por norma en la pantalla de Formularios disponibles.",
            font=FONTS["small"],
            text_color="#6D7480",
            justify="left",
            wraplength=980,
        ).grid(row=(row + 1) * 2, column=0, padx=18, pady=(14, 8), sticky="w")

        if not self.can_edit:
            self.norm_selector.configure(state="disabled")
            self.client_selector.configure(state="disabled")
            self.supervisor_entry.configure(state="disabled")

    def _build_answers_tab(
        self,
        parent: ctk.CTkFrame,
        title: str,
        questions: list[str],
        result_vars: list[ctk.StringVar],
        obs_vars: list[ctk.StringVar],
    ) -> None:
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(parent, text=title, font=FONTS["label_bold"], text_color=STYLE["texto_oscuro"]).grid(row=0, column=0, padx=18, pady=(16, 8), sticky="w")

        scroll_frame = ctk.CTkScrollableFrame(parent, fg_color=STYLE["fondo"], corner_radius=0)
        scroll_frame.grid(row=1, column=0, padx=18, pady=(0, 18), sticky="nsew")
        scroll_frame.grid_columnconfigure(0, weight=1)

        for idx, question in enumerate(questions):
            result_var = ctk.StringVar(value="")
            obs_var = ctk.StringVar(value="")
            result_vars.append(result_var)
            obs_vars.append(obs_var)

            row = idx
            question_card = ctk.CTkFrame(scroll_frame, fg_color="#FFFFFF", border_width=1, border_color="#E3E6EA", corner_radius=10)
            question_card.grid(row=row, column=0, pady=(0, 10), sticky="ew")
            question_card.grid_columnconfigure(0, weight=1)

            ctk.CTkLabel(
                question_card,
                text=f"{idx + 1}. {question}",
                font=FONTS["small_bold"],
                text_color=STYLE["texto_oscuro"],
                wraplength=980,
                justify="left",
                anchor="w",
            ).grid(row=0, column=0, padx=12, pady=(10, 6), sticky="ew")

            control_grid = ctk.CTkFrame(question_card, fg_color="transparent")
            control_grid.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="ew")
            control_grid.grid_columnconfigure(0, weight=1)
            control_grid.grid_columnconfigure(1, weight=1)
            control_grid.grid_columnconfigure(2, weight=1)
            control_grid.grid_columnconfigure(3, weight=3)

            option_specs = [
                ("Conforme", "conforme"),
                ("No conforme", "no_conforme"),
                ("No aplica", "no_aplica"),
            ]
            radio_buttons: list[ctk.CTkRadioButton] = []
            for col, (label_text, value) in enumerate(option_specs):
                option_cell = ctk.CTkFrame(control_grid, fg_color="#F5F7FA", corner_radius=8)
                option_cell.grid(row=0, column=col, padx=(0 if col == 0 else 6, 6), sticky="nsew")

                ctk.CTkLabel(
                    option_cell,
                    text=label_text,
                    font=FONTS["small_bold"],
                    text_color=STYLE["texto_oscuro"],
                ).pack(pady=(6, 2))

                radio = ctk.CTkRadioButton(option_cell, text="", variable=result_var, value=value, width=20)
                radio.pack(pady=(0, 6))
                radio_buttons.append(radio)

            observations_cell = ctk.CTkFrame(control_grid, fg_color="#F5F7FA", corner_radius=8)
            observations_cell.grid(row=0, column=3, padx=(6, 0), sticky="nsew")

            ctk.CTkLabel(
                observations_cell,
                text="Observaciones",
                font=FONTS["small_bold"],
                text_color=STYLE["texto_oscuro"],
            ).pack(anchor="w", padx=8, pady=(6, 2))

            obs_entry = ctk.CTkEntry(
                observations_cell,
                textvariable=obs_var,
                height=32,
                border_color="#94A3B8",
                placeholder_text="Observaciones",
            )
            obs_entry.pack(fill="x", padx=8, pady=(0, 6))

            if not self.can_edit:
                for radio in radio_buttons:
                    radio.configure(state="disabled")
                obs_entry.configure(state="disabled")

    def _build_technical_tab(self, parent: ctk.CTkFrame) -> None:
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(3, weight=1)

        ctk.CTkLabel(parent, text="SUPERVISIÓN TÉCNICO NORMATIVA", font=FONTS["label_bold"], text_color=STYLE["texto_oscuro"]).grid(row=0, column=0, padx=18, pady=(16, 6), sticky="w")
        ctk.CTkLabel(
            parent,
            text="Completa este apartado para habilitar la descarga del PDF. La calificación se calcula por norma y se guarda en el historial del inspector.",
            font=FONTS["small"],
            text_color="#6D7480",
            wraplength=980,
            justify="left",
        ).grid(row=1, column=0, padx=18, sticky="w")

        header = ctk.CTkFrame(parent, fg_color=STYLE["primario"], corner_radius=10)
        header.grid(row=2, column=0, padx=18, pady=(10, 8), sticky="ew")
        header.grid_columnconfigure(0, weight=4)
        header.grid_columnconfigure(1, weight=3)
        header.grid_columnconfigure(2, weight=1)
        header.grid_columnconfigure(3, weight=3)
        header.grid_columnconfigure(4, weight=1)

        ctk.CTkLabel(header, text="SKU/ITEM/CÓDIGO/UPC INSPECCIONADO", font=FONTS["small_bold"], text_color=STYLE["texto_oscuro"]).grid(row=0, column=0, padx=8, pady=8, sticky="w")
        ctk.CTkLabel(header, text="NOM APLICABLE", font=FONTS["small_bold"], text_color=STYLE["texto_oscuro"]).grid(row=0, column=1, padx=8, pady=8, sticky="w")
        ctk.CTkLabel(header, text="C/NC", font=FONTS["small_bold"], text_color=STYLE["texto_oscuro"]).grid(row=0, column=2, padx=8, pady=8)
        ctk.CTkLabel(header, text="OBSERVACIONES", font=FONTS["small_bold"], text_color=STYLE["texto_oscuro"]).grid(row=0, column=3, padx=8, pady=8, sticky="w")

        add_btn = ctk.CTkButton(
            header,
            text="+ Fila",
            width=78,
            fg_color=STYLE["secundario"],
            hover_color="#1D1D1D",
            command=self._add_technical_row,
        )
        add_btn.grid(row=0, column=4, padx=(6, 8), pady=6)

        if not self.can_edit:
            add_btn.configure(state="disabled")

        self.technical_rows_container = ctk.CTkScrollableFrame(parent, fg_color=STYLE["fondo"], corner_radius=0)
        self.technical_rows_container.grid(row=3, column=0, padx=18, pady=(0, 12), sticky="nsew")
        self.technical_rows_container.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            parent,
            textvariable=self.form_status_var,
            font=FONTS["small"],
            text_color="#6D7480",
            justify="left",
        ).grid(row=4, column=0, padx=18, pady=(0, 8), sticky="w")

        actions = ctk.CTkFrame(parent, fg_color="transparent")
        actions.grid(row=5, column=0, padx=18, pady=(0, 12), sticky="ew")
        actions.grid_columnconfigure(0, weight=1)

        self.download_button = ctk.CTkButton(
            actions,
            text="Formato de Supervisión",
            fg_color=STYLE["primario"],
            text_color=STYLE["texto_oscuro"],
            hover_color="#D8C220",
            command=lambda: self._download_document("formato"),
        )
        self.download_button.grid(row=0, column=0, sticky="ew")

        if not self.can_edit:
            self.download_button.configure(state="disabled")

        self._add_technical_row()

    def _add_technical_row(self, initial_values: dict[str, str] | None = None) -> None:
        if self.technical_rows_container is None:
            return

        row_frame = ctk.CTkFrame(self.technical_rows_container, fg_color="#FFFFFF", border_width=1, border_color="#E3E6EA", corner_radius=10)
        row_frame.grid_columnconfigure(0, weight=4)
        row_frame.grid_columnconfigure(1, weight=3)
        row_frame.grid_columnconfigure(2, weight=1)
        row_frame.grid_columnconfigure(3, weight=3)
        row_frame.grid_columnconfigure(4, weight=1)

        sku_var = ctk.StringVar(value=(initial_values or {}).get("sku", ""))
        default_norm_label = self.norm_var.get().strip() or self._norm_label_from_value(self.initial_norm)
        initial_norm_value = self._norm_label_from_value((initial_values or {}).get("applicable_norm", ""))
        applicable_norm_var = ctk.StringVar(value=initial_norm_value or default_norm_label)

        raw_result = str((initial_values or {}).get("result", "")).strip().lower()
        c_nc_value = str((initial_values or {}).get("c_nc", "")).strip().upper()
        if not c_nc_value and raw_result == "conforme":
            c_nc_value = "C"
        elif not c_nc_value and raw_result == "no_conforme":
            c_nc_value = "NC"

        result_var = ctk.StringVar(value=c_nc_value)
        observations_var = ctk.StringVar(value=(initial_values or {}).get("observations", ""))

        sku_entry = ctk.CTkEntry(row_frame, textvariable=sku_var, height=34, border_color="#94A3B8", placeholder_text="SKU / Código / UPC")
        sku_entry.grid(row=0, column=0, padx=8, pady=8, sticky="ew")

        norm_values = self.applicable_norm_values or [default_norm_label]
        norm_combo = ctk.CTkComboBox(
            row_frame,
            variable=applicable_norm_var,
            values=norm_values,
            height=34,
            fg_color="#FFFFFF",
            border_color="#94A3B8",
            button_color=STYLE["primario"],
            dropdown_hover_color=STYLE["primario"],
            state="readonly",
        )
        norm_combo.grid(row=0, column=1, padx=8, pady=8, sticky="ew")

        result_combo = ctk.CTkComboBox(
            row_frame,
            variable=result_var,
            values=TECHNICAL_RESULT_OPTIONS,
            width=90,
            height=34,
            fg_color="#FFFFFF",
            border_color="#94A3B8",
            button_color=STYLE["primario"],
            dropdown_hover_color=STYLE["primario"],
            state="readonly",
        )
        result_combo.grid(row=0, column=2, padx=8, pady=8)

        obs_entry = ctk.CTkEntry(row_frame, textvariable=observations_var, height=34, border_color="#94A3B8", placeholder_text="Observaciones")
        obs_entry.grid(row=0, column=3, padx=8, pady=8, sticky="ew")

        remove_button = ctk.CTkButton(
            row_frame,
            text="-",
            width=34,
            fg_color="#F7E0DE",
            text_color=STYLE["peligro"],
            hover_color="#F2C9C5",
            command=lambda: self._remove_technical_row(row_frame),
        )
        remove_button.grid(row=0, column=4, padx=(4, 8), pady=8)

        if not self.can_edit:
            sku_entry.configure(state="disabled")
            norm_combo.configure(state="disabled")
            result_combo.configure(state="disabled")
            obs_entry.configure(state="disabled")
            remove_button.configure(state="disabled")

        self.technical_rows.append(
            {
                "frame": row_frame,
                "sku_var": sku_var,
                "applicable_norm_var": applicable_norm_var,
                "norm_combo": norm_combo,
                "result_var": result_var,
                "observations_var": observations_var,
                "remove_button": remove_button,
            }
        )
        self._observe_state_var(sku_var)
        self._observe_state_var(applicable_norm_var)
        self._observe_state_var(result_var)
        self._observe_state_var(observations_var)
        self._refresh_technical_rows_layout()
        self._sync_download_state()

    def _remove_technical_row(self, row_frame) -> None:
        if not self.can_edit:
            return

        if len(self.technical_rows) == 1:
            first = self.technical_rows[0]
            first["sku_var"].set("")
            first["applicable_norm_var"].set(self._norm_label_from_value(self.norm_var.get() or self.initial_norm))
            first["result_var"].set("")
            first["observations_var"].set("")
            self._sync_download_state()
            return

        self.technical_rows = [item for item in self.technical_rows if item.get("frame") != row_frame]
        try:
            row_frame.destroy()
        except Exception:
            pass
        self._refresh_technical_rows_layout()
        self._sync_download_state()

    def _refresh_technical_rows_layout(self) -> None:
        for index, row in enumerate(self.technical_rows):
            frame = row["frame"]
            frame.grid(row=index, column=0, pady=(0, 8), sticky="ew")

        if self.can_edit:
            only_one = len(self.technical_rows) == 1
            for row in self.technical_rows:
                row["remove_button"].configure(state="disabled" if only_one else "normal")

    def _reset_technical_rows(self, rows: list[dict[str, str]] | None = None) -> None:
        for row in self.technical_rows:
            try:
                row["frame"].destroy()
            except Exception:
                pass
        self.technical_rows.clear()

        if rows:
            for item in rows:
                self._add_technical_row(item)
        else:
            self._add_technical_row()
        self._sync_download_state()

    def _add_field(self, parent: ctk.CTkFrame, row: int, label: str, widget) -> None:
        ctk.CTkLabel(parent, text=label, font=FONTS["label"], text_color=STYLE["texto_oscuro"]).grid(row=row * 2, column=0, padx=18, pady=(18 if row == 0 else 10, 6), sticky="w")
        widget.grid(row=row * 2 + 1, column=0, padx=18, sticky="ew")

    def _norm_label_from_value(self, raw_value: str) -> str:
        clean_value = " ".join(str(raw_value or "").split())
        if not clean_value:
            return "Sin norma"

        if clean_value in self.norm_display_to_token:
            return clean_value

        token = self.norm_display_to_token.get(clean_value)
        if token and token in self.norm_token_to_display:
            return self.norm_token_to_display[token]

        extracted = re.search(r"NOM-\d{3}", clean_value.upper())
        if extracted:
            token = extracted.group(0)
            if token in self.norm_token_to_display:
                return self.norm_token_to_display[token]

        return clean_value

    def _get_selected_norm_token(self) -> str:
        selected_display = " ".join(self.norm_var.get().split())
        if not selected_display:
            return ""

        mapped = self.norm_display_to_token.get(selected_display)
        if mapped:
            return mapped

        extracted = re.search(r"NOM-\d{3}", selected_display.upper())
        return extracted.group(0) if extracted else selected_display

    def _refresh_applicable_norm_values(self) -> None:
        values: list[str] = []
        seen: set[str] = set()

        for item in self.controller.get_catalog_norms():
            token = str(item.get("token", "")).strip()
            if not token:
                continue
            label = self._norm_label_from_value(token)
            if label in seen:
                continue
            values.append(label)
            seen.add(label)

        selected_norm = self._norm_label_from_value(self.norm_var.get())
        if selected_norm and selected_norm not in seen:
            values.insert(0, selected_norm)
            seen.add(selected_norm)

        self.applicable_norm_values = values or ["Sin norma"]
        for row in self.technical_rows:
            norm_combo = row.get("norm_combo")
            if norm_combo is None:
                continue
            norm_combo.configure(values=self.applicable_norm_values)

    def _collect_supervision_answers(
        self,
        questions: list[str],
        result_vars: list[ctk.StringVar],
        obs_vars: list[ctk.StringVar],
    ) -> list[dict[str, str]]:
        answers: list[dict[str, str]] = []
        for index, question in enumerate(questions):
            answers.append(
                {
                    "activity": question,
                    "result": result_vars[index].get().strip().lower(),
                    "observations": obs_vars[index].get().strip(),
                }
            )
        return answers

    def _collect_technical_rows(self) -> list[dict[str, str]]:
        rows: list[dict[str, str]] = []
        default_norm = self._norm_label_from_value(self.norm_var.get() or self.initial_norm)
        for index, row in enumerate(self.technical_rows):
            sku = str(row["sku_var"].get()).strip()
            applicable_norm = str(row["applicable_norm_var"].get()).strip()
            c_nc = str(row["result_var"].get()).strip().upper()
            observations = str(row["observations_var"].get()).strip()

            has_any_value = bool(sku or c_nc or observations)
            if not has_any_value and applicable_norm and applicable_norm != default_norm:
                has_any_value = True
            if not has_any_value:
                continue

            if not sku:
                raise ValueError(f"Completa SKU/ITEM/CÓDIGO/UPC en la fila técnica {index + 1}.")
            if not applicable_norm:
                raise ValueError(f"Completa NOM aplicable en la fila técnica {index + 1}.")
            if c_nc not in {"C", "NC"}:
                raise ValueError(f"Selecciona C o NC en la fila técnica {index + 1}.")

            rows.append(
                {
                    "sku": sku,
                    "applicable_norm": applicable_norm,
                    "c_nc": c_nc,
                    "result": "conforme" if c_nc == "C" else "no_conforme",
                    "observations": observations,
                }
            )
        return rows

    def _calculate_scores(
        self,
        selected_norm: str,
        protocol_answers: list[dict[str, str]],
        process_answers: list[dict[str, str]],
        technical_rows: list[dict[str, str]],
    ) -> tuple[float, str, dict[str, int], dict[str, float]]:
        norm_stats: dict[str, dict[str, int]] = {}

        def ensure_norm(norm_name: str) -> dict[str, int]:
            normalized = norm_name.strip() or "Sin norma"
            if normalized not in norm_stats:
                norm_stats[normalized] = {"conforme": 0, "no_conforme": 0, "no_aplica": 0}
            return norm_stats[normalized]

        selected_norm_name = selected_norm.strip() or "Sin norma"
        for answer in [*protocol_answers, *process_answers]:
            result = str(answer.get("result", "")).strip().lower()
            bucket = ensure_norm(selected_norm_name)
            if result == "conforme":
                bucket["conforme"] += 1
            elif result == "no_conforme":
                bucket["no_conforme"] += 1
            elif result == "no_aplica":
                bucket["no_aplica"] += 1

        for row in technical_rows:
            bucket = ensure_norm(str(row.get("applicable_norm", "")))
            result = str(row.get("result", "")).strip().lower()
            if result == "conforme":
                bucket["conforme"] += 1
            elif result == "no_conforme":
                bucket["no_conforme"] += 1

        total_conforme = 0
        total_no_conforme = 0
        total_no_aplica = 0
        score_by_norm: dict[str, float] = {}

        for norm_name, stats in norm_stats.items():
            conforme = int(stats.get("conforme", 0))
            no_conforme = int(stats.get("no_conforme", 0))
            no_aplica = int(stats.get("no_aplica", 0))
            aplicables = conforme + no_conforme
            norm_score = (conforme / aplicables * 100.0) if aplicables > 0 else 0.0
            score_by_norm[norm_name] = round(norm_score, 1)

            total_conforme += conforme
            total_no_conforme += no_conforme
            total_no_aplica += no_aplica

        final_score = round(sum(score_by_norm.values()) / len(score_by_norm), 1) if score_by_norm else 0.0
        status = "Estable" if final_score >= 90 else ("En seguimiento" if final_score >= 70 else "Critico")
        breakdown = {
            "conforme": total_conforme,
            "no_conforme": total_no_conforme,
            "no_aplica": total_no_aplica,
            "aplicables": total_conforme + total_no_conforme,
        }
        return final_score, status, breakdown, score_by_norm

    def _load_latest(self, preload_saved: bool = False) -> None:
        if self._is_loading_latest:
            return

        self._is_loading_latest = True
        try:
            if not preload_saved:
                self._reset_form_fields()
                self._sync_download_state()
                return

            selected_norm_token = self._get_selected_norm_token() or self.initial_norm
            latest = self.controller.get_latest_evaluation(self.inspector_name, selected_norm_token)

            if latest:
                saved_norm = self._norm_label_from_value(latest.get("selected_norm", self.norm_var.get()))
                if saved_norm:
                    self.norm_var.set(saved_norm)

                self.client_var.set(latest.get("client", self.client_var.get()))
                self.date_var.set(latest.get("visit_date", self.date_var.get()))
                self.supervisor_var.set(latest.get("evaluator", self.supervisor_var.get()))
                self.score_var.set(str(latest.get("score", "")))
                self.status_var.set(str(latest.get("status", self.status_var.get())))

                protocol_map = {
                    str(item.get("activity", "")).strip(): item
                    for item in latest.get("protocol_answers", [])
                    if isinstance(item, dict)
                }
                for idx, question in enumerate(PROTOCOL_QUESTIONS):
                    item = protocol_map.get(question, {})
                    self.protocol_result_vars[idx].set(str(item.get("result", "")).strip().lower())
                    self.protocol_obs_vars[idx].set(str(item.get("observations", "")).strip())

                process_map = {
                    str(item.get("activity", "")).strip(): item
                    for item in latest.get("process_answers", [])
                    if isinstance(item, dict)
                }
                for idx, question in enumerate(PROCESS_QUESTIONS):
                    item = process_map.get(question, {})
                    self.process_result_vars[idx].set(str(item.get("result", "")).strip().lower())
                    self.process_obs_vars[idx].set(str(item.get("observations", "")).strip())

                technical_saved = latest.get("technical_normative_rows", [])
                normalized_technical: list[dict[str, str]] = []
                if isinstance(technical_saved, list):
                    for row in technical_saved:
                        if not isinstance(row, dict):
                            continue
                        normalized_technical.append(
                            {
                                "sku": str(row.get("sku", "")).strip(),
                                "applicable_norm": self._norm_label_from_value(str(row.get("applicable_norm", "")).strip()),
                                "result": str(row.get("result", "")).strip().lower(),
                                "c_nc": str(row.get("c_nc", "")).strip().upper(),
                                "observations": str(row.get("observations", "")).strip(),
                            }
                        )
                self._refresh_applicable_norm_values()
                self._reset_technical_rows(normalized_technical)
            else:
                self.date_var.set(datetime.now().strftime("%Y-%m-%d"))
                self.score_var.set("")
                for idx in range(len(PROTOCOL_QUESTIONS)):
                    self.protocol_result_vars[idx].set("")
                    self.protocol_obs_vars[idx].set("")
                for idx in range(len(PROCESS_QUESTIONS)):
                    self.process_result_vars[idx].set("")
                    self.process_obs_vars[idx].set("")
                self._refresh_applicable_norm_values()
                self._reset_technical_rows()

            self._sync_download_state()
        finally:
            self._is_loading_latest = False

    def _configure_form_state_tracking(self) -> None:
        tracked_vars = [
            self.norm_var,
            self.client_var,
            self.date_var,
            self.supervisor_var,
            *self.protocol_result_vars,
            *self.process_result_vars,
        ]
        for variable in tracked_vars:
            self._observe_state_var(variable)

    def _observe_state_var(self, variable: ctk.StringVar) -> None:
        try:
            variable.trace_add("write", self._handle_form_value_change)
            self._observed_vars.append(variable)
        except Exception:
            return

    def _handle_form_value_change(self, *_args) -> None:
        if self._is_loading_latest:
            return
        self._sync_download_state()

    @staticmethod
    def _answer_completed(value: str) -> bool:
        return value.strip().lower() in {"conforme", "no_conforme", "no_aplica"}

    def _get_missing_answers(self, result_vars: list[ctk.StringVar]) -> list[str]:
        return [
            str(index + 1)
            for index, var in enumerate(result_vars)
            if not self._answer_completed(var.get())
        ]

    def _has_complete_technical_rows(self) -> bool:
        default_norm = self._norm_label_from_value(self.norm_var.get() or self.initial_norm)
        completed_rows = 0
        for row in self.technical_rows:
            sku = str(row["sku_var"].get()).strip()
            applicable_norm = str(row["applicable_norm_var"].get()).strip()
            c_nc = str(row["result_var"].get()).strip().upper()
            observations = str(row["observations_var"].get()).strip()

            has_any_value = bool(sku or c_nc or observations)
            if not has_any_value and applicable_norm and applicable_norm != default_norm:
                has_any_value = True
            if not has_any_value:
                continue

            if not sku or not applicable_norm or c_nc not in {"C", "NC"}:
                return False
            completed_rows += 1

        return completed_rows > 0

    def _is_form_complete(self) -> bool:
        if not self.client_var.get().strip() or not self.supervisor_var.get().strip():
            return False
        if self._get_missing_answers(self.protocol_result_vars):
            return False
        if self._get_missing_answers(self.process_result_vars):
            return False
        return self._has_complete_technical_rows()

    def _build_evaluation_payload(self, show_errors: bool) -> dict[str, object] | None:
        def _fail(message: str) -> None:
            if show_errors:
                messagebox.showerror("Formulario", message, parent=self)

        client_name = self.client_var.get().strip()
        if not client_name:
            _fail("Debes seleccionar Cliente / Almacén.")
            return None

        supervisor_name = self.supervisor_var.get().strip()
        if not supervisor_name:
            _fail("Debes capturar el Nombre del supervisor.")
            return None

        missing_protocol = self._get_missing_answers(self.protocol_result_vars)
        if missing_protocol:
            _fail("Completa Protocolo y Habilidades. Faltan preguntas: " + ", ".join(missing_protocol))
            return None

        missing_process = self._get_missing_answers(self.process_result_vars)
        if missing_process:
            _fail("Completa Procesos. Faltan preguntas: " + ", ".join(missing_process))
            return None

        try:
            technical_rows = self._collect_technical_rows()
        except ValueError as error:
            _fail(str(error))
            return None

        if not technical_rows:
            _fail("Debes capturar al menos una fila en Supervisión Técnico Normativa.")
            return None

        protocol_answers = self._collect_supervision_answers(PROTOCOL_QUESTIONS, self.protocol_result_vars, self.protocol_obs_vars)
        process_answers = self._collect_supervision_answers(PROCESS_QUESTIONS, self.process_result_vars, self.process_obs_vars)
        selected_norm_label = self._norm_label_from_value(self.norm_var.get())
        score, status, score_breakdown, score_by_norm = self._calculate_scores(
            selected_norm_label,
            protocol_answers,
            process_answers,
            technical_rows,
        )

        self.score_var.set(f"{score:.1f}")
        self.status_var.set(status)

        return {
            "selected_norm": selected_norm_label,
            "client": client_name,
            "visit_date": self.date_var.get().strip(),
            "score": f"{score:.1f}",
            "status": status,
            "observations": "",
            "corrective_actions": "",
            "evaluator": supervisor_name,
            "inspector_supervised": self.inspector_name,
            "protocol_answers": protocol_answers,
            "process_answers": process_answers,
            "technical_normative_rows": technical_rows,
            "score_breakdown": score_breakdown,
            "score_by_norm": score_by_norm,
        }

    def _persist_evaluation(self, payload: dict[str, object]) -> dict | None:
        try:
            saved = self.controller.save_evaluation(self.inspector_name, payload)
        except ValueError as error:
            messagebox.showerror("Formulario", str(error), parent=self)
            return None

        self.on_saved()
        return saved

    def _sync_download_state(self) -> None:
        enabled = self._is_form_complete()
        self.form_status_var.set(
            "Formulario completo. Ya puedes generar el Formato de Supervisión."
            if enabled
            else "Completa los 4 apartados para habilitar el Formato de Supervisión."
        )
        if self.download_button is None:
            return

        if enabled:
            self.download_button.grid()
            self.download_button.configure(state="normal" if self.can_edit else "disabled")
        else:
            self.download_button.grid_remove()

    def _download_document(self, kind: str) -> None:
        if not self.can_edit:
            return

        payload = self._build_evaluation_payload(show_errors=True)
        if payload is None:
            return

        selected_norm = self._get_selected_norm_token() or self.initial_norm
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

        if self._persist_evaluation(payload) is None:
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
        self._clear_form()

    def _reset_form_fields(self) -> None:
        self.client_var.set("")
        self.date_var.set(datetime.now().strftime("%Y-%m-%d"))
        self.supervisor_var.set((self.controller.current_user or {}).get("name", ""))
        self.score_var.set("")
        self.status_var.set("Estable")

        for variable in self.protocol_result_vars:
            variable.set("")
        for variable in self.protocol_obs_vars:
            variable.set("")
        for variable in self.process_result_vars:
            variable.set("")
        for variable in self.process_obs_vars:
            variable.set("")

        self._refresh_applicable_norm_values()
        self._reset_technical_rows()

    def _clear_form(self) -> None:
        self._reset_form_fields()
        self._sync_download_state()

    def _on_norm_change(self) -> None:
        self._refresh_applicable_norm_values()
        self._load_latest(preload_saved=False)


class PrincipalView(ctk.CTkFrame):
    FILTER_OPTIONS = ["Todos", "Pendientes", "En enfoque", "Completos"]
    PAGE_SIZE = 8

    def __init__(self, master, controller, can_edit: bool, on_change) -> None:
        super().__init__(master, fg_color=STYLE["fondo"])
        self.controller = controller
        self.can_edit = can_edit
        self.on_change = on_change

        self.cards_frame: ctk.CTkScrollableFrame | None = None
        self.search_var = ctk.StringVar(value="")
        self.status_filter_var = ctk.StringVar(value="Todos")
        self.results_var = ctk.StringVar(value="0 ejecutivos visibles")
        self.selected_row_id: str | None = None
        self.row_cache: dict[str, dict] = {}
        self._filters_refresh_job: str | None = None
        self._all_rows: list[dict] = []
        self._current_page = 0
        self._pager_frame: ctk.CTkFrame | None = None

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._build_ui()
        self.search_var.trace_add("write", lambda *_args: self._schedule_refresh())
        self.status_filter_var.trace_add("write", lambda *_args: self._schedule_refresh())

    def _build_ui(self) -> None:
        panel = ctk.CTkFrame(self, fg_color=STYLE["surface"], corner_radius=24)
        panel.grid(row=0, column=0, sticky="nsew")
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(1, weight=1)

        toolbar = ctk.CTkFrame(
            panel,
            fg_color="#FFFFFF",
            corner_radius=20,
            border_width=1,
            border_color="#E3E6EA",
        )
        toolbar.grid(row=0, column=0, padx=20, pady=(20, 12), sticky="ew")
        toolbar.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            toolbar,
            text="Buscar",
            font=FONTS["small_bold"],
            text_color=STYLE["texto_oscuro"],
        ).grid(row=0, column=0, padx=(18, 10), pady=(16, 6), sticky="w")

        search_entry = ctk.CTkEntry(
            toolbar,
            textvariable=self.search_var,
            height=38,
            border_color="#94A3B8",
        )
        search_entry.grid(row=0, column=1, padx=(0, 10), pady=(16, 6), sticky="ew")

        ctk.CTkButton(
            toolbar,
            text="Limpiar",
            width=92,
            height=36,
            fg_color=STYLE["fondo"],
            text_color=STYLE["texto_oscuro"],
            hover_color="#E9ECEF",
            command=self._clear_filters,
        ).grid(row=0, column=2, padx=(0, 14), pady=(16, 6), sticky="ew")

        filter_group = ctk.CTkFrame(toolbar, fg_color="transparent")
        filter_group.grid(row=0, column=3, padx=(0, 18), pady=(16, 6), sticky="e")
        filter_group.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            filter_group,
            text="Estado",
            font=FONTS["small_bold"],
            text_color=STYLE["texto_oscuro"],
        ).grid(row=0, column=0, padx=(0, 10), sticky="w")

        ctk.CTkComboBox(
            filter_group,
            variable=self.status_filter_var,
            values=self.FILTER_OPTIONS,
            width=170,
            height=38,
            fg_color="#FFFFFF",
            border_color="#94A3B8",
            button_color=STYLE["primario"],
            dropdown_hover_color=STYLE["primario"],
        ).grid(row=0, column=1, sticky="ew")

        ctk.CTkLabel(
            toolbar,
            textvariable=self.results_var,
            font=FONTS["small"],
            text_color="#6D7480",
        ).grid(row=1, column=0, columnspan=4, padx=18, pady=(0, 16), sticky="w")

        self.cards_frame = ctk.CTkScrollableFrame(
            panel,
            fg_color=STYLE["fondo"],
            corner_radius=20,
            scrollbar_button_color="#CDD2D9",
            scrollbar_button_hover_color="#B7BEC8",
        )
        self.cards_frame.grid(row=1, column=0, padx=20, pady=(0, 8), sticky="nsew")
        self.cards_frame.grid_columnconfigure(0, weight=1, uniform="principal_cards")
        self.cards_frame.grid_columnconfigure(1, weight=1, uniform="principal_cards")
        self.cards_frame.grid_columnconfigure(2, weight=1, uniform="principal_cards")
        self.cards_frame.grid_columnconfigure(3, weight=1, uniform="principal_cards")

        self._pager_frame = ctk.CTkFrame(panel, fg_color="transparent")
        self._pager_frame.grid(row=2, column=0, padx=20, pady=(0, 16), sticky="ew")

    def refresh(self) -> None:
        if self.cards_frame is None:
            return

        if self._filters_refresh_job is not None:
            self.after_cancel(self._filters_refresh_job)
            self._filters_refresh_job = None

        rows = self.controller.get_principal_rows(
            self.search_var.get(),
            self.status_filter_var.get(),
        )
        self.row_cache = {row.get("row_id", row["name"]): row for row in rows}
        self._all_rows = rows
        self._current_page = 0
        self.results_var.set(
            f"{len(rows)} ejecutivos visibles. Usa el boton Formulario para abrir la captura y sus documentos."
        )

        if not rows:
            self.selected_row_id = None
        elif self.selected_row_id not in self.row_cache:
            self.selected_row_id = rows[0].get("row_id", rows[0]["name"])

        self._render_page()

    def _render_page(self) -> None:
        if self.cards_frame is None:
            return

        # hide the scrollable frame before destroying children to avoid
        # the Windows canvas transparency artifact during widget churn
        self.cards_frame.grid_remove()

        for child in self.cards_frame.winfo_children():
            child.destroy()

        start = self._current_page * self.PAGE_SIZE
        page_rows = self._all_rows[start : start + self.PAGE_SIZE]

        if not page_rows:
            self._render_empty_state()
        else:
            for index, row in enumerate(page_rows):
                self._render_card(index, row)

        self.cards_frame.grid()  # restore
        self._rebuild_pager()

    def _rebuild_pager(self) -> None:
        if self._pager_frame is None:
            return

        for child in self._pager_frame.winfo_children():
            child.destroy()

        total = len(self._all_rows)
        if total == 0:
            return

        total_pages = max(1, (total + self.PAGE_SIZE - 1) // self.PAGE_SIZE)
        if total_pages <= 1:
            return

        inner = ctk.CTkFrame(self._pager_frame, fg_color="transparent")
        inner.pack(anchor="center")

        ctk.CTkButton(
            inner,
            text="←  Anterior",
            width=110,
            height=34,
            fg_color=STYLE["fondo"],
            text_color=STYLE["texto_oscuro"],
            hover_color="#E9ECEF",
            state="normal" if self._current_page > 0 else "disabled",
            command=lambda: self._go_page(-1),
        ).pack(side="left", padx=(0, 8))

        ctk.CTkLabel(
            inner,
            text=f"Pagina {self._current_page + 1} de {total_pages}  —  {total} ejecutivos",
            font=FONTS["small"],
            text_color="#6D7480",
        ).pack(side="left", padx=12)

        ctk.CTkButton(
            inner,
            text="Siguiente  →",
            width=110,
            height=34,
            fg_color=STYLE["fondo"],
            text_color=STYLE["texto_oscuro"],
            hover_color="#E9ECEF",
            state="normal" if self._current_page < total_pages - 1 else "disabled",
            command=lambda: self._go_page(1),
        ).pack(side="left", padx=(8, 0))

    def _go_page(self, delta: int) -> None:
        total = len(self._all_rows)
        total_pages = max(1, (total + self.PAGE_SIZE - 1) // self.PAGE_SIZE)
        new_page = max(0, min(total_pages - 1, self._current_page + delta))
        if new_page == self._current_page:
            return
        self._current_page = new_page
        self._render_page()
        try:
            if self.cards_frame is not None:
                self.cards_frame._parent_canvas.yview_moveto(0)
        except Exception:
            pass

    def get_selected_row(self) -> dict | None:
        if not self.selected_row_id:
            return None
        return self.row_cache.get(self.selected_row_id)

    def get_selected_name(self) -> str | None:
        row = self.get_selected_row()
        return row["name"] if row else None

    def _schedule_refresh(self, delay_ms: int = 120) -> None:
        if self._filters_refresh_job is not None:
            self.after_cancel(self._filters_refresh_job)
        self._filters_refresh_job = self.after(delay_ms, self.refresh)

    def _clear_filters(self) -> None:
        self.search_var.set("")
        self.status_filter_var.set("Todos")
        self._schedule_refresh(10)

    def _render_empty_state(self) -> None:
        if self.cards_frame is None:
            return

        empty_card = ctk.CTkFrame(
            self.cards_frame,
            fg_color="#FFFFFF",
            corner_radius=20,
            border_width=1,
            border_color="#E3E6EA",
        )
        empty_card.grid(row=0, column=0, columnspan=4, padx=8, pady=8, sticky="ew")
        empty_card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            empty_card,
            text="No hay ejecutivos para los filtros seleccionados.",
            font=FONTS["label_bold"],
            text_color=STYLE["texto_oscuro"],
        ).grid(row=0, column=0, padx=20, pady=(22, 8), sticky="w")
        ctk.CTkLabel(
            empty_card,
            text="Ajusta la busqueda o cambia el estado para volver a mostrar tarjetas.",
            font=FONTS["small"],
            text_color="#6D7480",
        ).grid(row=1, column=0, padx=20, pady=(0, 22), sticky="w")

    def _render_card(self, index: int, row: dict) -> None:
        if self.cards_frame is None:
            return

        row_id = row.get("row_id", row["name"])
        palette = self._status_palette(row.get("status", "Pendiente"))
        latest_date = self._compact_date(str(row.get("latest_date", "--")))
        form_label = "Captura guardada" if row.get("form_completed") else "Formulario pendiente"
        norms_preview = self._truncate_text(str(row.get("norms_text", "Sin acreditacion")), 72)
        name_text = self._truncate_text(row["name"], 32)

        card = ctk.CTkFrame(
            self.cards_frame,
            fg_color="#FFFFFF",
            corner_radius=20,
            border_width=1,
            border_color="#E3E6EA",
            height=248,
        )
        card.grid(row=index // 4, column=index % 4, padx=6, pady=6, sticky="nsew")
        card.grid_propagate(False)
        card.grid_columnconfigure(1, weight=1)
        # row 2 (norms box) absorbs extra vertical space so the button stays at bottom
        card.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(
            card,
            text="👤",
            font=("Segoe UI Emoji", 18),
            text_color=STYLE["texto_oscuro"],
            fg_color="#F5F1CD",
            corner_radius=12,
            width=34,
            height=34,
        ).grid(row=0, column=0, padx=(12, 8), pady=(12, 0), sticky="nw")

        ctk.CTkLabel(
            card,
            text=name_text,
            font=FONTS["label_bold"],
            text_color=STYLE["texto_oscuro"],
            anchor="w",
            justify="left",
            wraplength=160,
        ).grid(row=0, column=1, pady=(12, 0), sticky="w")

        ctk.CTkLabel(
            card,
            text=row["status"],
            font=FONTS["small_bold"],
            text_color=palette["text"],
            fg_color=palette["soft"],
            corner_radius=10,
            padx=8,
            pady=3,
        ).grid(row=0, column=2, padx=(4, 12), pady=(12, 0), sticky="ne")

        ctk.CTkLabel(
            card,
            text=f"{row['norm_count']} normas  |  Puntaje: {row.get('latest_score_text', '--')}  |  {latest_date}",
            font=FONTS["small"],
            text_color="#6D7480",
            anchor="w",
        ).grid(row=1, column=0, columnspan=3, padx=12, pady=(4, 6), sticky="w")

        ctk.CTkLabel(
            card,
            text=norms_preview,
            font=FONTS["small"],
            text_color="#5F6671",
            fg_color="#F8F9FA",
            corner_radius=10,
            anchor="nw",
            justify="left",
            wraplength=240,
            padx=10,
            pady=6,
        ).grid(row=2, column=0, columnspan=3, padx=12, pady=(0, 8), sticky="nsew")

        ctk.CTkLabel(
            card,
            text=form_label,
            font=FONTS["small"],
            text_color="#6D7480",
        ).grid(row=3, column=0, columnspan=3, padx=12, pady=(0, 4), sticky="w")

        ctk.CTkButton(
            card,
            text="📋  Formulario",
            height=34,
            font=FONTS["small_bold"],
            fg_color=STYLE["primario"],
            text_color=STYLE["texto_oscuro"],
            hover_color="#D8C220",
            command=lambda current=row_id, current_row=row: self._open_row_actions(current, current_row),
        ).grid(row=4, column=0, columnspan=3, padx=12, pady=(0, 12), sticky="ew")

    def _set_selected_row(self, row_id: str) -> None:
        self.selected_row_id = row_id

    def _open_row_actions(self, row_id: str, row: dict) -> None:
        self._set_selected_row(row_id)
        self.open_actions(row)

    def _status_palette(self, status: str) -> dict[str, str]:
        if status == "Estable":
            return {"soft": "#E5F6ED", "text": "#0D6B42"}
        if status == "En enfoque":
            return {"soft": "#FDE7E2", "text": "#B84A33"}
        return {"soft": "#FFF4D9", "text": "#8A6A17"}

    def _compact_date(self, raw_value: str) -> str:
        value = raw_value.strip() or "--"
        if "T" in value:
            return value.split("T", 1)[0]
        if len(value) >= 10 and value[4:5] == "-" and value[7:8] == "-":
            return value[:10]
        return value

    def _truncate_text(self, text: str, limit: int) -> str:
        clean_text = text.strip()
        if len(clean_text) <= limit:
            return clean_text
        return clean_text[: limit - 3].rstrip(", ") + "..."

    def open_editor(self, inspector_name: str | None) -> None:
        if not self.can_edit:
            return
        InspectorEditDialog(self, self.controller, inspector_name, self._handle_change)

    def open_actions(self, selected_row: dict | None = None) -> None:
        selected_row = selected_row or self.get_selected_row()
        if not selected_row:
            messagebox.showinfo("Principal", "Selecciona un ejecutivo tecnico.")
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

        NormSelectionDialog(
            self,
            inspector_name,
            available_norms,
            None,
            _open_form,
        )

    def _open_actions_dialog(self, selected_row: dict) -> None:
        self.open_actions(selected_row)

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
        self.page_dirty: dict[str, bool] = {}
        self._section_refresh_job: str | None = None
        self.ui_scale = 1.0

        self.title("Sistema de Calibracion V&C")
        self.ui_scale = self._configure_ui_scale()
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._configure_treeview_style()
        self._show_login()

    def _configure_ui_scale(self) -> float:
        screen_width = max(1, int(self.winfo_screenwidth()))
        screen_height = max(1, int(self.winfo_screenheight()))

        if screen_width <= 1366 or screen_height <= 768:
            scale = 0.88
        elif screen_width <= 1536:
            scale = 0.94
        else:
            scale = 1.0

        self._apply_font_scale(scale)
        ctk.set_widget_scaling(scale)
        ctk.set_window_scaling(scale)

        default_width = min(1480, int(screen_width * 0.9))
        default_height = min(920, int(screen_height * 0.9))
        min_width = min(1200, max(960, int(screen_width * 0.75)))
        min_height = min(760, max(640, int(screen_height * 0.72)))

        self.geometry(f"{default_width}x{default_height}")
        self.minsize(min_width, min_height)
        return scale

    def _apply_font_scale(self, scale: float) -> None:
        for key, font_value in BASE_FONTS.items():
            FONTS[key] = _scaled_font(font_value, scale)

    def _configure_treeview_style(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")
        heading_size = int(FONTS["small_bold"][1])
        row_height = max(26, int(round(34 * self.ui_scale)))
        heading_padding = max(6, int(round(8 * self.ui_scale)))
        style.configure(
            "Treeview",
            background="#FFFFFF",
            fieldbackground="#FFFFFF",
            foreground=STYLE["texto_oscuro"],
            borderwidth=0,
            rowheight=row_height,
        )
        style.configure(
            "Treeview.Heading",
            background=STYLE["secundario"],
            foreground=STYLE["texto_claro"],
            relief="flat",
            padding=(heading_padding, heading_padding),
            font=("Inter", heading_size, "bold"),
        )
        style.map(
            "Treeview.Heading",
            background=[("active", STYLE["secundario"]), ("pressed", STYLE["secundario"])],
            foreground=[("active", STYLE["texto_claro"]), ("pressed", STYLE["texto_claro"])],
        )
        style.map("Treeview", background=[("selected", "#F2F2F2")], foreground=[("selected", STYLE["texto_oscuro"])])

    def _clear_root(self) -> None:
        for child in self.winfo_children():
            child.destroy()

    def _show_login(self) -> None:
        self._clear_root()
        shell = ctk.CTkFrame(self, fg_color="transparent")
        shell_pad_x = 16 if self.ui_scale < 0.95 else 24
        shell_pad_y = 14 if self.ui_scale < 0.95 else 22
        shell.grid(row=0, column=0, sticky="nsew", padx=shell_pad_x, pady=shell_pad_y)
        shell.grid_columnconfigure(0, weight=1)
        shell.grid_rowconfigure(0, weight=1)

        login_card = LoginView(shell, STYLE, FONTS, self._handle_login)
        login_card.grid(row=0, column=0, sticky="nsew")

    def _handle_login(self, username: str, password: str) -> str | None:
        user = self.controller.authenticate(username, password)
        if not user:
            return "Usuario o contrasena incorrectos. Revisa Usuarios.json para validar el acceso."

        self._show_main_shell()
        return None

    def _show_main_shell(self) -> None:
        self._clear_root()
        self.main_frame = ctk.CTkFrame(self, fg_color=STYLE["fondo"])
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
            font=FONTS["title"],
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

    def _build_navigation(self, parent) -> None:
        nav = ctk.CTkFrame(parent, fg_color="transparent")
        nav.grid(row=1, column=0, sticky="ew", pady=(10, 12))
        nav.grid_columnconfigure(0, weight=1)
        nav.grid_columnconfigure(1, weight=0)

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
                width=120,
                height=38,
                command=lambda section_name=section: self.show_section(section_name),
            )
            button.grid(row=0, column=index, padx=(0 if index == 0 else 8, 0), pady=0)
            self.nav_buttons[section] = button

        if not self.controller.is_admin():
            return

        summary_row = ctk.CTkFrame(nav, fg_color="transparent")
        summary_row.grid(row=0, column=1, sticky="e")

        title_map = {
            "inspectors": "Ejecutivos tecnicos",
            "average_score": "Promedio",
            "alerts": "Alertas < 90%",
        }
        for index, key in enumerate(["inspectors", "average_score", "alerts"]):
            card = ctk.CTkFrame(
                summary_row,
                fg_color="#FFFFFF",
                corner_radius=12,
                border_width=1,
                border_color="#E3E6EA",
                width=148,
                height=40,
            )
            card.grid(row=0, column=index, padx=(0 if index == 0 else 6, 0), sticky="e")
            card.grid_propagate(False)
            card.grid_columnconfigure(0, weight=1)

            ctk.CTkLabel(
                card,
                text=title_map[key],
                font=FONTS["small"],
                text_color="#6D7480",
            ).grid(row=0, column=0, padx=(8, 4), pady=9, sticky="w")

            value_label = ctk.CTkLabel(
                card,
                text="--",
                font=("Inter", int(FONTS["small_bold"][1]), "bold"),
                text_color=STYLE["texto_oscuro"],
            )
            value_label.grid(row=0, column=1, padx=(4, 8), sticky="e")
            self.summary_labels[key] = value_label

    def _build_content(self, parent) -> None:
        self.content_frame = ctk.CTkFrame(parent, fg_color=STYLE["fondo"], corner_radius=0)
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

        self.page_dirty = {name: True for name in self.pages}

        for page in self.pages.values():
            page.grid(row=0, column=0, sticky="nsew")

    def show_section(self, section: str) -> None:
        if section not in self.pages:
            return
        self.current_section = section
        self.pages[section].tkraise()
        self.update()
        for name, button in self.nav_buttons.items():
            if name == section:
                button.configure(fg_color=STYLE["primario"], text_color=STYLE["texto_oscuro"], hover_color="#D8C220")
            else:
                button.configure(fg_color=STYLE["fondo"], text_color=STYLE["texto_oscuro"], hover_color="#E9ECEF")
        self._schedule_refresh_section(section)

    def _schedule_refresh_section(self, section: str) -> None:
        if self._section_refresh_job is not None:
            self.after_cancel(self._section_refresh_job)
        self._section_refresh_job = self.after(50, lambda section_name=section: self._refresh_section(section_name))

    def refresh_all_views(self) -> None:
        self.controller.reload()
        metrics = self.controller.get_overview_metrics()
        self._set_summary_value("inspectors", str(metrics.get("inspectors", 0)))
        average = metrics.get("average_score")
        self._set_summary_value("average_score", f"{average:.1f}%" if average is not None else "--")
        self._set_summary_value("alerts", str(metrics.get("alerts", 0)))

        for section in self.pages:
            self.page_dirty[section] = True
        if self.current_section:
            self._schedule_refresh_section(self.current_section)

    def _set_summary_value(self, key: str, value: str) -> None:
        label = self.summary_labels.get(key)
        if label is not None:
            label.configure(text=value)

    def _refresh_current_section(self) -> None:
        if not self.current_section:
            return
        self._refresh_section(self.current_section)

    def _refresh_section(self, section: str) -> None:
        self._section_refresh_job = None
        page = self.pages.get(section)
        if page is None or not hasattr(page, "refresh"):
            self.page_dirty[section] = False
            return
        if not self.page_dirty.get(section, True):
            return
        try:
            if page.winfo_exists():
                page.refresh()
                self.page_dirty[section] = False
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



