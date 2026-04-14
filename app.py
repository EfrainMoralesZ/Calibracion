from __future__ import annotations

import ctypes
import os
import re
import threading
from collections.abc import Callable
from datetime import datetime
from tkinter import TclError, filedialog, messagebox, ttk

import customtkinter as ctk
from PIL import Image
import ui_shared

from calendario import CalendarView
from configuraciones import ConfigurationView
from dashboard import DashboardView
from calibration_controller import CalibrationController
from login import LoginView
from runtime_paths import resource_path, writable_path
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


_MONITOR_DEFAULTTONEAREST = 2


class _WinRect(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]


class _MonitorInfo(ctypes.Structure):
    _fields_ = [
        ("cbSize", ctypes.c_ulong),
        ("rcMonitor", _WinRect),
        ("rcWork", _WinRect),
        ("dwFlags", ctypes.c_ulong),
    ]


def _configure_windows_dpi_behavior() -> None:
    if os.name != "nt":
        return

    try:
        ctk.deactivate_automatic_dpi_awareness()
    except AttributeError:
        pass

    try:
        user32 = ctypes.windll.user32
    except AttributeError:
        return

    for context in (ctypes.c_void_p(-2),):
        try:
            if user32.SetProcessDpiAwarenessContext(context):
                return
        except (AttributeError, OSError):
            pass

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        return
    except (AttributeError, OSError):
        pass

    try:
        user32.SetProcessDPIAware()
    except (AttributeError, OSError):
        return


def _get_window_work_area(window) -> tuple[int, int, int, int] | None:
    if os.name != "nt" or window is None:
        return None

    try:
        hwnd = int(window.winfo_id())
    except (AttributeError, TclError, ValueError):
        return None

    try:
        user32 = ctypes.windll.user32
        monitor = user32.MonitorFromWindow(hwnd, _MONITOR_DEFAULTTONEAREST)
        if not monitor:
            return None

        monitor_info = _MonitorInfo()
        monitor_info.cbSize = ctypes.sizeof(_MonitorInfo)
        if not user32.GetMonitorInfoW(monitor, ctypes.byref(monitor_info)):
            return None

        work_area = monitor_info.rcWork
        return work_area.left, work_area.top, work_area.right, work_area.bottom
    except (AttributeError, OSError):
        return None


def _position_toplevel(window, parent, width: int, height: int) -> None:
    try:
        window.update_idletasks()
    except TclError:
        return

    work_area = _get_window_work_area(parent or window)
    if work_area is None:
        left = 0
        top = 0
        right = max(width, int(window.winfo_screenwidth()))
        bottom = max(height, int(window.winfo_screenheight()))
    else:
        left, top, right, bottom = work_area

    x = left + max(0, ((right - left) - width) // 2)
    y = top + max(0, ((bottom - top) - height) // 2)

    if parent is not None:
        try:
            if parent.winfo_exists():
                parent.update_idletasks()
                parent_width = max(parent.winfo_width(), parent.winfo_reqwidth())
                parent_height = max(parent.winfo_height(), parent.winfo_reqheight())
                if parent_width > 1 and parent_height > 1:
                    x = parent.winfo_rootx() + (parent_width - width) // 2
                    y = parent.winfo_rooty() + (parent_height - height) // 2
        except TclError:
            pass

    padding = 16
    max_x = max(left + padding, right - width - padding)
    max_y = max(top + padding, bottom - height - padding)
    x = min(max(left + padding, x), max_x)
    y = min(max(top + padding, y), max_y)
    window.geometry(f"{width}x{height}+{x}+{y}")


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


 # supervision.py contiene toda la lógica de la sección Supervisión.
# Se importa aquí, después de que STYLE, FONTS y utilidades están definidos,
# para que el import circular (supervision → app) funcione correctamente.
from supervision import PrincipalView  # noqa: E402
from criterioEvaluacion import CriteriaEvaluationView  # noqa: E402



class CalibrationApp(ctk.CTk):
    def __init__(self) -> None:
        _configure_windows_dpi_behavior()
        ctk.set_appearance_mode("light")
        super().__init__(fg_color=STYLE["fondo"])

        self.controller = CalibrationController()
        self.pages: dict[str, ctk.CTkFrame] = {}
        self.page_factories: dict[str, Callable[[], ctk.CTkFrame]] = {}
        self.nav_buttons: dict[str, ctk.CTkButton] = {}
        self.summary_labels: dict[str, ctk.CTkLabel] = {}
        self.header_identity_label: ctk.CTkLabel | None = None
        self.header_greeting_label: ctk.CTkLabel | None = None
        self.header_message_label: ctk.CTkLabel | None = None
        self.main_frame: ctk.CTkFrame | None = None
        self.content_frame: ctk.CTkFrame | None = None
        self.current_section = ""
        self.page_dirty: dict[str, bool] = {}
        self._section_refresh_job: str | None = None
        self.ui_scale = 1.0
        self._kpi_medal_images: dict[str, ctk.CTkImage | None] = {}
        self._kpi_medal_images_exec: dict[str, ctk.CTkImage | None] = {}

        self.title("Sistema de fiabilidad técnica y calidad en el servicio")
        self._set_window_icon()
        self.ui_scale = self._configure_ui_scale()
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._configure_treeview_style()
        self._show_login()

    def _load_kpi_medal_images(self, size: tuple[int, int] = (26, 26)) -> dict[str, ctk.CTkImage | None]:
        files = {
            "ORO": "medalla_oro.png",
            "PLATINO": "medalla_plata.png",
            "BRONCE": "medalla_bronce.png",
        }
        images: dict[str, ctk.CTkImage | None] = {}
        for key, fname in files.items():
            img_path = None
            for candidate in [resource_path("img", fname), writable_path("img", fname)]:
                if candidate.exists():
                    img_path = candidate
                    break
            if img_path is None:
                images[key] = None
                continue
            try:
                src = Image.open(img_path)
                images[key] = ctk.CTkImage(light_image=src, dark_image=src, size=size)
            except OSError:
                images[key] = None
        return images

    @staticmethod
    def _part_of_day_greeting() -> tuple[str, str]:
        hour = datetime.now().hour
        if hour < 12:
            return "Buenos dias", "☀️"
        if hour < 19:
            return "Buenas tardes", "🌤️"
        return "Buenas noches", "🌙"

    @staticmethod
    def _format_medals_text(counts: dict[str, int]) -> str:
        return f"O:{counts.get('ORO', 0)} P:{counts.get('PLATINO', 0)} B:{counts.get('BRONCE', 0)}"

    @staticmethod
    def _build_yearly_phrase_catalog(openings: list[str], focuses: list[str]) -> list[str]:
        phrases: list[str] = []
        for opening in openings:
            for focus in focuses:
                phrases.append(f"{opening} {focus}")
                if len(phrases) == 366:
                    return phrases
        return phrases

    @staticmethod
    def _message_for_day(messages: list[str], fallback: str) -> str:
        if not messages:
            return fallback
        day_index = max(0, datetime.now().timetuple().tm_yday - 1)
        return messages[day_index % len(messages)]

    def _build_header_messages(self) -> tuple[str, str]:
        current_user = self.controller.current_user or {}
        greeting, emoji = self._part_of_day_greeting()
        greeting_line = f"{emoji}  {greeting}"

        if not self.controller.is_admin(current_user):
            encouragements = self._build_yearly_phrase_catalog(
                [
                    "Hoy tu trabajo en campo",
                    "Este dia tu disciplina operativa",
                    "Tu constancia en cada visita",
                    "La calidad de tu seguimiento",
                    "Tu atencion en cada cliente",
                    "El enfoque que llevas a cada jornada",
                ],
                [
                    "marca una diferencia real en el servicio.",
                    "refuerza la confianza del cliente.",
                    "sostiene el nivel operativo del equipo.",
                    "eleva la calidad de cada revision.",
                    "fortalece la ejecucion en campo.",
                    "mejora la experiencia de supervision.",
                    "da valor a cada norma revisada.",
                    "mantiene el orden en cada proceso.",
                    "ayuda a prevenir errores antes de que crezcan.",
                    "convierte el detalle en resultados visibles.",
                    "hace mas solida la entrega al cliente.",
                    "impulsa el avance diario del equipo.",
                    "demuestra compromiso con el resultado final.",
                    "ayuda a cerrar cada visita con claridad.",
                    "refleja profesionalismo en cada decision.",
                    "convierte cada inspeccion en una oportunidad de mejora.",
                    "aporta orden y criterio a la operacion.",
                    "mantiene el ritmo correcto de trabajo.",
                    "genera confianza en cada seguimiento.",
                    "refuerza la imagen del equipo frente al cliente.",
                    "hace visible tu nivel de compromiso.",
                    "suma precision a cada evaluacion.",
                    "hace que cada evidencia cuente.",
                    "te acerca a resultados mas consistentes.",
                    "fortalece el cumplimiento diario.",
                    "ayuda a sostener una ejecucion limpia y ordenada.",
                    "mejora la lectura tecnica de cada caso.",
                    "te distingue por tu enfoque profesional.",
                    "abre espacio para mejores resultados.",
                    "convierte la constancia en desempeno.",
                    "mejora la forma en que el cliente percibe el servicio.",
                    "mantiene el control sobre cada detalle.",
                    "te permite construir mejores cierres de visita.",
                    "refuerza el valor de cada accion en campo.",
                    "demuestra solidez en la operacion.",
                    "ayuda a que el trabajo del equipo se note.",
                    "hace mas fuerte cada entrega realizada.",
                    "suma claridad a cada intervencion.",
                    "convierte la preparacion en confianza.",
                    "hace que cada jornada deje aprendizaje.",
                    "mantiene la calidad aun en dias exigentes.",
                    "te ayuda a responder mejor ante cada reto.",
                    "sostiene una ejecucion profesional de principio a fin.",
                    "impulsa mejoras que el cliente si percibe.",
                    "fortalece el criterio con el que operas.",
                    "te permite avanzar con seguridad y orden.",
                    "mantiene enfocado lo importante de cada visita.",
                    "genera resultados que se construyen desde el detalle.",
                    "ayuda a que cada revision tenga mayor valor.",
                    "demuestra que la consistencia si hace diferencia.",
                    "protege la calidad en cada entrega.",
                    "suma confianza a cada paso del proceso.",
                    "hace que tu trabajo hable por si mismo.",
                    "fortalece la forma en que representas al equipo.",
                    "impulsa una operacion mas clara y estable.",
                    "te posiciona como parte clave del resultado.",
                    "refuerza cada avance conseguido durante el dia.",
                    "te permite sostener un buen nivel de ejecucion.",
                    "hace de cada visita una oportunidad para destacar.",
                    "mantiene el servicio alineado con lo esperado.",
                    "hace que el esfuerzo diario se convierta en valor.",
                ],
            )
            message = self._message_for_day(encouragements, "Tu trabajo en campo hace la diferencia.")
            return greeting_line, message

        admin_messages = self._build_yearly_phrase_catalog(
            [
                "Hoy tu liderazgo",
                "La lectura que haces del desempeno",
                "Tu seguimiento del equipo",
                "La forma en que diriges la operacion",
                "Tu criterio para priorizar",
                "El control que mantienes del sistema",
            ],
            [
                "impulsa mejores decisiones para el equipo.",
                "ayuda a traducir datos en accion.",
                "fortalece el rumbo operativo del dia.",
                "convierte el seguimiento en resultados medibles.",
                "mantiene al equipo enfocado en lo importante.",
                "sostiene una operacion mas ordenada y clara.",
                "te da visibilidad para actuar a tiempo.",
                "hace posible corregir con rapidez.",
                "refuerza el desempeno colectivo.",
                "marca el paso para un mejor cierre diario.",
                "te permite liderar con datos y contexto.",
                "genera claridad en cada frente operativo.",
                "ayuda a detectar oportunidades reales de mejora.",
                "hace mas precisa la toma de decisiones.",
                "convierte la supervision en una ventaja del equipo.",
                "refuerza el seguimiento constante del resultado.",
                "da forma a una operacion mas solida.",
                "te permite anticiparte a los riesgos.",
                "ayuda a mantener el nivel del servicio.",
                "hace visible donde intervenir primero.",
                "mantiene alineadas las prioridades del equipo.",
                "suma direccion y enfoque al trabajo diario.",
                "te permite decidir con mayor confianza.",
                "fortalece la consistencia del sistema.",
                "impulsa una ejecucion mejor coordinada.",
                "hace que cada indicador tenga utilidad practica.",
                "te ayuda a liderar con mas claridad.",
                "mantiene viva la mejora continua.",
                "genera una operacion mas estable y predecible.",
                "hace mas visibles los avances del equipo.",
                "refuerza el impacto de tus decisiones.",
                "te ayuda a mover al equipo con precision.",
                "mejora la forma en que respondes a cada reto.",
                "convierte el monitoreo en direccion efectiva.",
                "hace mas clara la ruta de trabajo.",
                "sostiene el ritmo correcto de la operacion.",
                "te permite identificar focos de accion con rapidez.",
                "impulsa el balance entre seguimiento y ejecucion.",
                "fortalece la confianza del equipo en el rumbo.",
                "te da base para liderar con objetividad.",
                "mantiene enfocado el esfuerzo colectivo.",
                "te ayuda a proteger la calidad del servicio.",
                "hace que cada revision tenga proposito.",
                "convierte la informacion en decisiones concretas.",
                "te permite sostener una mejora mas consistente.",
                "ayuda a elevar el nivel de respuesta del equipo.",
                "hace del seguimiento una ventaja operativa.",
                "refuerza una cultura de ejecucion con criterio.",
                "te permite ver tendencias antes de que escalen.",
                "mantiene conectado el dato con la accion correcta.",
                "genera direccion en cada cierre del dia.",
                "hace que el sistema trabaje a tu favor.",
                "fortalece la disciplina operativa del equipo.",
                "te ayuda a mantener el control sin perder agilidad.",
                "impulsa resultados mas claros y sostenibles.",
                "convierte cada lectura del sistema en una oportunidad.",
                "te permite liderar con enfoque y consistencia.",
                "mantiene la operacion orientada al resultado.",
                "hace que la supervision se traduzca en desempeno.",
                "refuerza el valor de cada decision tomada hoy.",
                "da contexto para liderar mejor.",
            ],
        )
        message = self._message_for_day(admin_messages, "Analiza tendencias y lidera con datos.")
        return greeting_line, message

    def _set_window_icon(self) -> None:
        if os.name != "nt":
            return

        candidate_paths = [
            resource_path("img", "icono.ico"),
            writable_path("img", "icono.ico"),
        ]

        for icon_path in candidate_paths:
            if not icon_path.exists():
                continue
            try:
                self.iconbitmap(str(icon_path))
                return
            except (TclError, OSError):
                continue

    def _configure_ui_scale(self) -> float:
        work_area = _get_window_work_area(self)
        if work_area is None:
            left = 0
            top = 0
            screen_width = max(1, int(self.winfo_screenwidth()))
            screen_height = max(1, int(self.winfo_screenheight()))
        else:
            left, top, right, bottom = work_area
            screen_width = max(1, int(right - left))
            screen_height = max(1, int(bottom - top))

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

        origin_x = left + max(0, (screen_width - default_width) // 2)
        origin_y = top + max(0, (screen_height - default_height) // 2)
        self.geometry(f"{default_width}x{default_height}+{origin_x}+{origin_y}")
        self.minsize(min_width, min_height)
        return scale

    def _apply_font_scale(self, scale: float) -> None:
        for key, font_value in BASE_FONTS.items():
            FONTS[key] = _scaled_font(font_value, scale)
        for key, font_value in ui_shared.BASE_FONTS.items():
            ui_shared.FONTS[key] = ui_shared._scaled_font(font_value, scale)

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
        if self._section_refresh_job is not None:
            self.after_cancel(self._section_refresh_job)
            self._section_refresh_job = None
        for child in self.winfo_children():
            child.destroy()

    def _show_login(self) -> None:
        self.current_section = ""
        self.pages = {}
        self.page_factories = {}
        self.page_dirty = {}
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
            return "Usuario o contraseña incorrectos."

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
        is_executive = self.controller.is_executive_role(current_user)

        top_row = ctk.CTkFrame(header, fg_color="transparent")
        top_row.grid(row=0, column=0, padx=18, pady=(14, 8), sticky="ew")
        top_row.grid_columnconfigure(0, weight=0)
        top_row.grid_columnconfigure(1, weight=1)  # Centro para las cards de ejecutivos
        top_row.grid_columnconfigure(2, weight=0)

        title_group = ctk.CTkFrame(top_row, fg_color="transparent")
        title_group.grid(row=0, column=0, sticky="w")
        title_group.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            title_group,
            text="Sistema de fiabilidad técnica y calidad en el servicio",
            font=FONTS["title"],
            text_color=STYLE["texto_oscuro"],
        ).grid(row=0, column=0, sticky="w")

        # Para ejecutivos: Cards de estadísticas centradas a nivel del título
        if is_executive:
            self._build_executive_stats_row(top_row, current_user)

        controls_row = ctk.CTkFrame(top_row, fg_color="transparent")
        controls_row.grid(row=0, column=2, sticky="e")

        greeting_text, message_text = self._build_header_messages()

        greeting_card = ctk.CTkFrame(
            controls_row,
            fg_color="#FFF4B2",
            corner_radius=18,
            border_width=2,
            border_color="#E5C100",
        )
        greeting_card.grid(row=0, column=0, padx=(0, 8), sticky="e")

        self.header_greeting_label = ctk.CTkLabel(
            greeting_card,
            text=greeting_text,
            font=("Inter", int(FONTS["small_bold"][1]) + 2, "bold"),
            text_color="#5A4800",
            anchor="e",
            justify="right",
        )
        self.header_greeting_label.grid(row=0, column=0, padx=(14, 14), pady=(8, 0), sticky="e")

        self.header_identity_label = ctk.CTkLabel(
            greeting_card,
            text=f"{current_user.get('name', 'Sin sesion')}  |  {role_text}",
            font=FONTS["small_bold"],
            text_color="#4A3B00",
            anchor="e",
            justify="right",
        )
        self.header_identity_label.grid(row=1, column=0, padx=(14, 14), pady=(0, 0), sticky="e")

        self.header_message_label = ctk.CTkLabel(
            greeting_card,
            text=message_text,
            font=("Inter", int(FONTS["small_bold"][1]) + 1, "bold"),
            text_color="#7A6000",
            anchor="e",
            justify="right",
            wraplength=320,
        )
        self.header_message_label.grid(row=2, column=0, padx=(14, 14), pady=(0, 8), sticky="e")

        ctk.CTkButton(
            controls_row,
            text="Cerrar sesion",
            fg_color=STYLE["fondo"],
            text_color=STYLE["texto_oscuro"],
            hover_color="#E9ECEF",
            width=112,
            height=34,
            command=self._logout,
        ).grid(row=0, column=1, sticky="e")

    def _build_executive_stats_row(self, parent, current_user: dict) -> None:
        """Construye las cards de estadísticas para ejecutivos al nivel del título."""
        if not self._kpi_medal_images:
            self._kpi_medal_images = self._load_kpi_medal_images(size=(26, 26))
        if not self._kpi_medal_images_exec:
            self._kpi_medal_images_exec = self._load_kpi_medal_images(size=(42, 42))

        stats_container = ctk.CTkFrame(parent, fg_color="transparent")
        stats_container.grid(row=0, column=1, sticky="nsew", padx=20)

        medal_colors = {"ORO": "#B98500", "PLATINO": "#4F5D73", "BRONCE": "#8C4B20"}
        title_map = {
            "average_score": "Mi promedio",
            "alerts": "Mis alertas",
            "medals": "Mis medallas",
        }

        # Card Mi promedio - diseño prominente
        avg_card = ctk.CTkFrame(
            stats_container,
            fg_color="#FFFFFF",
            corner_radius=14,
            border_width=2,
            border_color="#E5C100",
            width=190,
            height=65,
        )
        avg_card.grid(row=0, column=0, padx=(0, 12), sticky="nsew")
        avg_card.grid_propagate(False)
        avg_card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            avg_card,
            text=title_map["average_score"],
            font=("Inter", 13, "bold"),
            text_color="#6D7480",
        ).grid(row=0, column=0, padx=14, pady=(10, 0), sticky="w")

        avg_value = ctk.CTkLabel(
            avg_card,
            text="--",
            font=("Inter", 18, "bold"),
            text_color=STYLE["texto_oscuro"],
        )
        avg_value.grid(row=0, column=1, padx=(0, 14), pady=(10, 0), sticky="e")
        self.summary_labels["average_score"] = avg_value

        # Bind click handler
        for widget in (avg_card, avg_value):
            widget.bind("<Button-1>", lambda _e: self._show_average_detail_popup())
            widget.configure(cursor="hand2")
        for child in avg_card.winfo_children():
            child.bind("<Button-1>", lambda _e: self._show_average_detail_popup())
            child.configure(cursor="hand2")

        # Card Mis alertas - diseño prominente
        alerts_card = ctk.CTkFrame(
            stats_container,
            fg_color="#FFFFFF",
            corner_radius=14,
            border_width=2,
            border_color="#E5C100",
            width=175,
            height=65,
        )
        alerts_card.grid(row=0, column=1, padx=(0, 12), sticky="nsew")
        alerts_card.grid_propagate(False)
        alerts_card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            alerts_card,
            text=title_map["alerts"],
            font=("Inter", 13, "bold"),
            text_color="#6D7480",
        ).grid(row=0, column=0, padx=14, pady=(10, 0), sticky="w")

        alerts_value = ctk.CTkLabel(
            alerts_card,
            text="0",
            font=("Inter", 18, "bold"),
            text_color=STYLE["texto_oscuro"],
        )
        alerts_value.grid(row=0, column=1, padx=(0, 14), pady=(10, 0), sticky="e")
        self.summary_labels["alerts"] = alerts_value

        # Bind click handler
        for widget in (alerts_card, alerts_value):
            widget.bind("<Button-1>", lambda _e: self._show_alerts_detail_popup())
            widget.configure(cursor="hand2")
        for child in alerts_card.winfo_children():
            child.bind("<Button-1>", lambda _e: self._show_alerts_detail_popup())
            child.configure(cursor="hand2")

        # Card Mis medallas - diseño prominente y destacado
        medals_card = ctk.CTkFrame(
            stats_container,
            fg_color="#FFF8CC",
            corner_radius=14,
            border_width=3,
            border_color="#E5C100",
            width=320,
            height=65,
        )
        medals_card.grid(row=0, column=2, sticky="nsew")
        medals_card.grid_propagate(False)

        ctk.CTkLabel(
            medals_card,
            text=title_map["medals"],
            font=("Inter", 14, "bold"),
            text_color="#7A6000",
        ).grid(row=0, column=0, padx=(14, 12), pady=12, sticky="w")

        for m_idx, medal_key in enumerate(["ORO", "PLATINO", "BRONCE"]):
            col_base = 1 + m_idx * 2
            img = self._kpi_medal_images_exec.get(medal_key)
            if img is not None:
                ctk.CTkLabel(medals_card, text="", image=img).grid(
                    row=0,
                    column=col_base,
                    padx=(0, 4),
                    pady=8,
                )
            count_lbl = ctk.CTkLabel(
                medals_card,
                text="0",
                font=("Inter", 20, "bold"),
                text_color=medal_colors[medal_key],
            )
            count_lbl.grid(
                row=0,
                column=col_base + 1,
                padx=(0, 10),
                pady=8,
            )
            self.summary_labels[f"medals_{medal_key}"] = count_lbl

    def _build_navigation(self, parent) -> None:
        current_user = self.controller.current_user or {}
        is_executive = self.controller.is_executive_role(current_user)
        
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

        # Para ejecutivos las cards se muestran en el header, no aqui
        if is_executive:
            return

        summary_row = ctk.CTkFrame(nav, fg_color="transparent")
        summary_row.grid(row=0, column=1, sticky="e")

        if self.controller.is_admin():
            title_map = {
                "inspectors": "Ejecutivos tecnicos",
                "average_score": "Promedio",
                "alerts": "Alertas < 90%",
                "medals": "Medallas",
            }
            summary_keys = ["inspectors", "average_score", "alerts", "medals"]
        else:
            title_map = {
                "average_score": "Mi promedio",
                "alerts": "Mis alertas",
                "medals": "Mis medallas",
            }
            summary_keys = ["average_score", "alerts", "medals"]

        if not self._kpi_medal_images:
            self._kpi_medal_images = self._load_kpi_medal_images(size=(26, 26))
        if not self._kpi_medal_images_exec:
            self._kpi_medal_images_exec = self._load_kpi_medal_images(size=(38, 38))

        medal_colors = {"ORO": "#B98500", "PLATINO": "#4F5D73", "BRONCE": "#8C4B20"}

        for index, key in enumerate(summary_keys):
            if key == "medals":
                card = ctk.CTkFrame(
                    summary_row,
                    fg_color="#FFFFFF",
                    corner_radius=12,
                    border_width=1,
                    border_color="#E3E6EA",
                    width=232,
                    height=56,
                )
                card.grid(row=0, column=index, padx=(0 if index == 0 else 6, 0), sticky="e")
                card.grid_propagate(False)
                ctk.CTkLabel(
                    card,
                    text=title_map[key],
                    font=FONTS["small_bold"],
                    text_color="#6D7480",
                ).grid(row=0, column=0, padx=(12, 8), pady=12, sticky="w")
                for m_idx, medal_key in enumerate(["ORO", "PLATINO", "BRONCE"]):
                    col_base = 1 + m_idx * 2
                    image_set = self._kpi_medal_images
                    img = image_set.get(medal_key)
                    if img is not None:
                        ctk.CTkLabel(card, text="", image=img).grid(
                            row=0,
                            column=col_base,
                            padx=(0, 3),
                            pady=10,
                        )
                    count_lbl = ctk.CTkLabel(
                        card,
                        text="0",
                        font=("Inter", int(FONTS["small_bold"][1]) + 3, "bold"),
                        text_color=medal_colors[medal_key],
                    )
                    count_lbl.grid(
                        row=0,
                        column=col_base + 1,
                        padx=(0, 7),
                        pady=10,
                    )
                    self.summary_labels[f"medals_{medal_key}"] = count_lbl
                continue

            card = ctk.CTkFrame(
                summary_row,
                fg_color="#FFFFFF",
                corner_radius=12,
                border_width=1,
                border_color="#E3E6EA",
                width=170,
                height=52,
            )
            card.grid(row=0, column=index, padx=(0 if index == 0 else 6, 0), sticky="e")
            card.grid_propagate(False)
            card.grid_columnconfigure(0, weight=1)

            ctk.CTkLabel(
                card,
                text=title_map[key],
                font=FONTS["small_bold"],
                text_color="#6D7480",
            ).grid(row=0, column=0, padx=(10, 4), pady=12, sticky="w")

            value_label = ctk.CTkLabel(
                card,
                text="--",
                font=("Inter", int(FONTS["small_bold"][1]) + 2, "bold"),
                text_color=STYLE["texto_oscuro"],
            )
            value_label.grid(row=0, column=1, padx=(4, 10), sticky="e")
            self.summary_labels[key] = value_label

            # Bind click handler for interactive cards
            if key in ("average_score", "alerts"):
                handler = (
                    self._show_average_detail_popup
                    if key == "average_score"
                    else self._show_alerts_detail_popup
                )
                for widget in (card, value_label):
                    widget.bind("<Button-1>", lambda _e, h=handler: h())
                    widget.configure(cursor="hand2")
                # also bind the title label (last created child)
                for child in card.winfo_children():
                    child.bind("<Button-1>", lambda _e, h=handler: h())
                    child.configure(cursor="hand2")

    def _build_content(self, parent) -> None:
        self.content_frame = ctk.CTkFrame(parent, fg_color=STYLE["fondo"], corner_radius=0)
        self.content_frame.grid(row=3, column=0, sticky="nsew")
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(0, weight=1)

        # Permitir edición completa en calendario para talento humano
        can_edit = self.controller.is_admin() or self.controller._role_name() == "talento humano"
        self.pages = {}
        self.page_factories = self._build_page_factories(can_edit)
        self.page_dirty = {name: True for name in self.page_factories}

    def _build_page_factories(self, can_edit: bool) -> dict[str, Callable[[], ctk.CTkFrame]]:
        if can_edit:
            return {
                "Supervisión": lambda: PrincipalView(self.content_frame, self.controller, can_edit, self.refresh_all_views),
                "Criterios": lambda: CriteriaEvaluationView(self.content_frame, self.controller, can_edit, self.refresh_all_views),
                "Dashboard": lambda: DashboardView(self.content_frame, self.controller, STYLE, FONTS),
                "Calendario": lambda: CalendarView(self.content_frame, self.controller, STYLE, FONTS, can_edit),
                "Trimestral": lambda: TrimestralView(self.content_frame, self.controller, STYLE, FONTS, True),
                "Configuraciones": lambda: ConfigurationView(self.content_frame, self.controller, STYLE, FONTS, True, self.refresh_all_views),
            }

        is_supervisor = self.controller._role_name() == "supervisor"

        non_admin_factories = {
            "Supervisión": lambda: PrincipalView(self.content_frame, self.controller, is_supervisor, self.refresh_all_views),
            "Criterios": lambda: CriteriaEvaluationView(self.content_frame, self.controller, can_edit, self.refresh_all_views),
            "Dashboard": lambda: DashboardView(self.content_frame, self.controller, STYLE, FONTS),
            "Calendario": lambda: CalendarView(self.content_frame, self.controller, STYLE, FONTS, is_supervisor),
            "Trimestral": lambda: TrimestralView(self.content_frame, self.controller, STYLE, FONTS, False),
        }
        allowed_sections = set(self.controller.available_sections())
        return {section: factory for section, factory in non_admin_factories.items() if section in allowed_sections}

    def _get_or_create_page(self, section: str) -> ctk.CTkFrame | None:
        page = self.pages.get(section)
        if page is not None and page.winfo_exists():
            return page

        factory = self.page_factories.get(section)
        if factory is None or self.content_frame is None:
            return None

        page = factory()
        page.grid(row=0, column=0, sticky="nsew")
        page.grid_remove()
        self.pages[section] = page
        return page

    def show_section(self, section: str) -> None:
        if section not in self.page_factories:
            return

        previous_section = self.current_section
        page = self._get_or_create_page(section)
        if page is None:
            return

        previous_page = self.pages.get(previous_section)
        if previous_page is not None and previous_page is not page:
            try:
                if previous_page.winfo_exists():
                    previous_page.grid_remove()
            except TclError:
                pass

        self.current_section = section
        page.grid()
        page.tkraise()
        self.update_idletasks()
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
        greeting_text, message_text = self._build_header_messages()
        role_text = str((self.controller.current_user or {}).get("role", "usuario")).upper()
        full_name = str((self.controller.current_user or {}).get("name", "Sin sesion"))
        if self.header_identity_label is not None:
            self.header_identity_label.configure(text=f"{full_name}  |  {role_text}")
        if self.header_greeting_label is not None:
            self.header_greeting_label.configure(text=greeting_text)
        if self.header_message_label is not None:
            self.header_message_label.configure(text=message_text)
        current_user = self.controller.current_user or {}
        if self.controller.is_admin(current_user):
            metrics = self.controller.get_overview_metrics()
            medal_summary = self.controller.get_trimestral_medals_summary(include_unsent=True)
            self._set_summary_value("inspectors", str(metrics.get("inspectors", 0)))
            average = metrics.get("average_score")
            self._set_summary_value("average_score", f"{average:.1f}%" if average is not None else "--")
            self._set_summary_value("alerts", str(metrics.get("alerts", 0)))
            counts = medal_summary.get("counts", {"ORO": 0, "PLATINO": 0, "BRONCE": 0})
            for medal_key in ["ORO", "PLATINO", "BRONCE"]:
                self._set_summary_value(f"medals_{medal_key}", str(counts.get(medal_key, 0)))
        else:
            viewer_name = str(current_user.get("name", "")).strip()
            profile = self.controller.get_executive_profile(viewer_name) if viewer_name else {}
            own_scores = self.controller.list_trimestral_scores(inspector_name=viewer_name) if viewer_name else []
            own_alerts = 0
            for row in own_scores:
                score_value = row.get("score")
                try:
                    if float(score_value) < 90:
                        own_alerts += 1
                except (TypeError, ValueError):
                    continue
            own_medals = self.controller.get_trimestral_medals_summary(inspector_name=viewer_name, include_unsent=True)
            average = profile.get("average_score") if isinstance(profile, dict) else None
            self._set_summary_value("average_score", f"{average:.1f}%" if average is not None else "--")
            self._set_summary_value("alerts", str(own_alerts))
            own_counts = own_medals.get("counts", {"ORO": 0, "PLATINO": 0, "BRONCE": 0})
            for medal_key in ["ORO", "PLATINO", "BRONCE"]:
                self._set_summary_value(f"medals_{medal_key}", str(own_counts.get(medal_key, 0)))

        for section in self.page_factories:
            self.page_dirty[section] = True
        if self.current_section:
            self._schedule_refresh_section(self.current_section)

    def _set_summary_value(self, key: str, value: str) -> None:
        label = self.summary_labels.get(key)
        if label is not None:
            label.configure(text=value)

    def _show_average_detail_popup(self) -> None:
        current_user = self.controller.current_user or {}
        popup = ctk.CTkToplevel(self)
        popup.title("Detalle de promedio")
        popup.geometry("420x420")
        popup.resizable(False, False)
        popup.grab_set()
        popup.lift()

        ctk.CTkLabel(
            popup,
            text="¿De dónde viene mi promedio?",
            font=("Inter", 14, "bold"),
            text_color=STYLE["texto_oscuro"],
        ).pack(padx=20, pady=(18, 4), anchor="w")
        ctk.CTkLabel(
            popup,
            text="El promedio se calcula con base en las supervisiones registradas en tu historial.",
            font=FONTS["small"],
            text_color="#6D7480",
            wraplength=380,
            justify="left",
        ).pack(padx=20, pady=(0, 4), anchor="w")
        ctk.CTkLabel(
            popup,
            text="Cada fila representa una supervisión realizada.",
            font=FONTS["small"],
            text_color="#6D7480",
            wraplength=380,
            justify="left",
        ).pack(padx=20, pady=(0, 12), anchor="w")

        scroll = ctk.CTkFrame(popup, fg_color="#F7F8FA", corner_radius=10)
        scroll.pack(padx=20, pady=(0, 16), fill="both", expand=True)
        scroll.grid_columnconfigure(0, weight=1)
        scroll.grid_columnconfigure(1, weight=0)

        if self.controller.is_admin(current_user):
            rows = self.controller.get_principal_rows()
            scores_with_names = [
                (r["name"], r["latest_score"], r.get("latest_date", "--"))
                for r in rows
                if r.get("latest_score") is not None
            ]
            if not scores_with_names:
                ctk.CTkLabel(scroll, text="Sin evaluaciones registradas.", font=FONTS["small"], text_color="#6D7480").grid(row=0, column=0, padx=12, pady=8, sticky="w")
            else:
                ctk.CTkLabel(scroll, text="Ejecutivo", font=FONTS["small_bold"], text_color="#6D7480").grid(row=0, column=0, padx=12, pady=(8, 4), sticky="w")
                ctk.CTkLabel(scroll, text="Puntaje", font=FONTS["small_bold"], text_color="#6D7480").grid(row=0, column=1, padx=(4, 12), pady=(8, 4), sticky="e")
                for i, (name, score, _date) in enumerate(scores_with_names, start=1):
                    color = "#0D6B42" if score >= 90 else "#B84A33"
                    ctk.CTkLabel(scroll, text=name, font=FONTS["small"], text_color=STYLE["texto_oscuro"]).grid(row=i, column=0, padx=12, pady=2, sticky="w")
                    ctk.CTkLabel(scroll, text=f"{score:.1f}%", font=FONTS["small_bold"], text_color=color).grid(row=i, column=1, padx=(4, 12), pady=2, sticky="e")
        else:
            viewer_name = str(current_user.get("name", "")).strip()
            profile = self.controller.get_executive_profile(viewer_name) if viewer_name else {}
            history = profile.get("history", [])
            if not history:
                ctk.CTkLabel(scroll, text="Sin evaluaciones registradas en tu historial.", font=FONTS["small"], text_color="#6D7480").grid(row=0, column=0, padx=12, pady=8, sticky="w")
            else:
                ctk.CTkLabel(scroll, text="Supervisión", font=FONTS["small_bold"], text_color="#6D7480").grid(row=0, column=0, padx=12, pady=(8, 4), sticky="w")
                ctk.CTkLabel(scroll, text="Puntaje", font=FONTS["small_bold"], text_color="#6D7480").grid(row=0, column=1, padx=(4, 12), pady=(8, 4), sticky="e")
                for i, point in enumerate(history, start=1):
                    score = point.get("score", 0)
                    color = "#0D6B42" if score >= 90 else "#B84A33"
                    ctk.CTkLabel(scroll, text=str(point.get("label", "--")), font=FONTS["small"], text_color=STYLE["texto_oscuro"]).grid(row=i, column=0, padx=12, pady=2, sticky="w")
                    ctk.CTkLabel(scroll, text=f"{score:.1f}%", font=FONTS["small_bold"], text_color=color).grid(row=i, column=1, padx=(4, 12), pady=2, sticky="e")

    def _show_alerts_detail_popup(self) -> None:
        current_user = self.controller.current_user or {}
        popup = ctk.CTkToplevel(self)
        popup.title("Detalle de alertas")
        popup.geometry("460x440")
        popup.resizable(False, False)
        popup.grab_set()
        popup.lift()

        ctk.CTkLabel(
            popup,
            text="¿Por qué tengo alertas?",
            font=("Inter", 14, "bold"),
            text_color=STYLE["texto_oscuro"],
        ).pack(padx=20, pady=(18, 4), anchor="w")
        ctk.CTkLabel(
            popup,
            text="Una alerta se genera cuando una evaluación tiene un puntaje menor a 90%. Cada alerta indica una oportunidad de mejora.",
            font=FONTS["small"],
            text_color="#6D7480",
            wraplength=420,
            justify="left",
        ).pack(padx=20, pady=(0, 12), anchor="w")

        scroll = ctk.CTkFrame(popup, fg_color="#F7F8FA", corner_radius=10)
        scroll.pack(padx=20, pady=(0, 16), fill="both", expand=True)
        scroll.grid_columnconfigure(0, weight=1)
        scroll.grid_columnconfigure(1, weight=0)

        if self.controller.is_admin(current_user):
            alert_rows = [r for r in self.controller.get_principal_rows() if r.get("latest_score") is not None and r["latest_score"] < 90]
            if not alert_rows:
                ctk.CTkLabel(scroll, text="Sin alertas activas. Todos los puntajes están sobre 90%.", font=FONTS["small"], text_color="#0D6B42").grid(row=0, column=0, padx=12, pady=8, sticky="w")
            else:
                ctk.CTkLabel(scroll, text="Ejecutivo", font=FONTS["small_bold"], text_color="#6D7480").grid(row=0, column=0, padx=12, pady=(8, 4), sticky="w")
                ctk.CTkLabel(scroll, text="Puntaje", font=FONTS["small_bold"], text_color="#6D7480").grid(row=0, column=1, padx=(4, 12), pady=(8, 4), sticky="e")
                for i, row in enumerate(alert_rows, start=1):
                    ctk.CTkLabel(scroll, text=row["name"], font=FONTS["small"], text_color=STYLE["texto_oscuro"]).grid(row=i, column=0, padx=12, pady=2, sticky="w")
                    ctk.CTkLabel(scroll, text=f"{row['latest_score']:.1f}%", font=FONTS["small_bold"], text_color="#B84A33").grid(row=i, column=1, padx=(4, 12), pady=2, sticky="e")
        else:
            viewer_name = str(current_user.get("name", "")).strip()
            own_scores = self.controller.list_trimestral_scores(inspector_name=viewer_name, include_unsent=True) if viewer_name else []
            alerts = [s for s in own_scores if s.get("score") is not None and float(s["score"]) < 90]
            if not alerts:
                ctk.CTkLabel(scroll, text="Sin alertas activas. Todos tus puntajes están sobre 90%.", font=FONTS["small"], text_color="#0D6B42").grid(row=0, column=0, padx=12, pady=8, sticky="w")
            else:
                headers = ["Norma", "Trimestre", "Año", "Puntaje"]
                for col, h in enumerate(headers):
                    scroll.grid_columnconfigure(col, weight=1)
                    ctk.CTkLabel(scroll, text=h, font=FONTS["small_bold"], text_color="#6D7480").grid(row=0, column=col, padx=8, pady=(8, 4), sticky="w")
                for i, s in enumerate(alerts, start=1):
                    score_val = float(s["score"])
                    ctk.CTkLabel(scroll, text=str(s.get("norm", "--")), font=FONTS["small"], text_color=STYLE["texto_oscuro"]).grid(row=i, column=0, padx=8, pady=2, sticky="w")
                    ctk.CTkLabel(scroll, text=str(s.get("quarter", "--")), font=FONTS["small"], text_color=STYLE["texto_oscuro"]).grid(row=i, column=1, padx=8, pady=2, sticky="w")
                    ctk.CTkLabel(scroll, text=str(s.get("year", "--")), font=FONTS["small"], text_color=STYLE["texto_oscuro"]).grid(row=i, column=2, padx=8, pady=2, sticky="w")
                    ctk.CTkLabel(scroll, text=f"{score_val:.1f}%", font=FONTS["small_bold"], text_color="#B84A33").grid(row=i, column=3, padx=8, pady=2, sticky="w")

    def _refresh_current_section(self) -> None:
        if not self.current_section:
            return
        self._refresh_section(self.current_section)

    def _refresh_section(self, section: str) -> None:
        self._section_refresh_job = None
        page = self.pages.get(section)
        if page is None:
            if section != self.current_section:
                return
            page = self._get_or_create_page(section)
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
