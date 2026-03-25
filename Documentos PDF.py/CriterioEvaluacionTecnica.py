from __future__ import annotations

from pathlib import Path
from typing import Any

from reportlab.lib import colors
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from runtime_paths import resource_path, writable_path


ACCENT = colors.HexColor("#1B2A4A")
DARK = colors.HexColor("#282828")
GRID = colors.HexColor("#D5D8DC")
ALLOWED_IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".webp"}
HEADER_LINE_1 = "FICHA DE CONSULTAS NORMATIVAS DE"
HEADER_LINE_2 = "VERIFICACIÓN & CONTROL UVA, S.C."


def _text(value: Any, fallback: str = "--") -> str:
    clean_value = str(value or "").strip()
    return clean_value or fallback


def _template_path() -> Path | None:
    for candidate in (resource_path("img", "plantilla.png"), writable_path("img", "plantilla.png")):
        if candidate.exists():
            return candidate
    return None


def _collect_image_paths(folder: str) -> list[Path]:
    base = Path(str(folder or "").strip())
    if not base.exists() or not base.is_dir():
        return []

    files: list[Path] = []
    for item in sorted(base.iterdir()):
        if item.is_file() and item.suffix.lower() in ALLOWED_IMAGE_SUFFIXES:
            files.append(item)
    return files


def _normalize_evidence_files(payload: dict[str, Any]) -> list[Path]:
    raw_files = payload.get("evidence_files", [])
    images: list[Path] = []
    if isinstance(raw_files, list):
        for raw_path in raw_files:
            path = Path(str(raw_path or "").strip())
            if path.exists() and path.is_file() and path.suffix.lower() in ALLOWED_IMAGE_SUFFIXES:
                images.append(path)

    if images:
        return images

    return _collect_image_paths(_text(payload.get("image_folder"), ""))


def _build_styles():
    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="VCBody",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=9,
            leading=12,
            textColor=DARK,
        )
    )
    styles.add(
        ParagraphStyle(
            name="VCBodySmall",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=8,
            leading=11,
            textColor=DARK,
        )
    )
    styles.add(
        ParagraphStyle(
            name="VCSection",
            parent=styles["Heading3"],
            fontName="Helvetica-Bold",
            fontSize=10,
            leading=12,
            textColor=DARK,
            spaceAfter=4,
        )
    )
    styles.add(
        ParagraphStyle(
            name="VCHeaderCell",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=8.5,
            leading=10,
            textColor=DARK,
            alignment=1,
        )
    )
    styles.add(
        ParagraphStyle(
            name="VCSignature",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=9,
            leading=14,
            textColor=DARK,
        )
    )
    return styles


def _header_cell(text: str, styles) -> Paragraph:
    return Paragraph(text, styles["VCHeaderCell"])


def _build_summary_table(payload: dict[str, Any]) -> Table:
    summary_rows = [
        ["Número de resolución", _text(payload.get("resolution_number"))],
        ["Fecha", _text(payload.get("visit_date") or payload.get("generated_at"))],
        ["Cliente", _text(payload.get("client"))],
        ["Nombre del ejecutivo", _text(payload.get("executive_name") or payload.get("inspector_supervised"))],
        ["Norma aplicable", _text(payload.get("selected_norm"))],
        ["Producto evaluado", _text(payload.get("evaluated_product"))],
    ]

    table = Table(summary_rows, colWidths=[5.1 * cm, 11.9 * cm])
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#F0F2F5")),
                ("BACKGROUND", (1, 0), (1, -1), colors.white),
                ("TEXTCOLOR", (0, 0), (-1, -1), DARK),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTNAME", (1, 0), (1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("GRID", (0, 0), (-1, -1), 0.5, GRID),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    return table


def _scaled_image(image_path: Path, width: float, height: float) -> Image | None:
    try:
        flow = Image(str(image_path))
    except OSError:
        return None
    flow._restrictSize(width, height)
    return flow


def _build_evidence_cell(image_paths: list[Path], styles):
    if not image_paths:
        return Paragraph("Sin evidencia fotografica adjunta.", styles["VCBody"])

    rows: list[list[object]] = []

    for image_path in image_paths:
        image_flow = _scaled_image(image_path, 4.5 * cm, 3.0 * cm)
        if image_flow is not None:
            rows.append([image_flow])

    if not rows:
        return Paragraph("Sin imagenes validas cargadas.", styles["VCBody"])

    nested = Table(rows, colWidths=[4.6 * cm])
    nested.setStyle(
        TableStyle(
            [
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 1),
                ("RIGHTPADDING", (0, 0), (-1, -1), 1),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    return nested


def _build_consultation_table(payload: dict[str, Any], styles) -> Table:
    comment_cell = Paragraph(_text(payload.get("comment")), styles["VCBody"])
    resolution_cell = Paragraph(_text(payload.get("resolution_text")), styles["VCBody"])
    evidence_cell = _build_evidence_cell(_normalize_evidence_files(payload), styles)

    table = Table(
        [
            [
                _header_cell("COMENTARIO", styles),
                _header_cell("EVIDENCIA", styles),
                _header_cell("RESOLUCIÓN", styles),
            ],
            [comment_cell, evidence_cell, resolution_cell],
        ],
        colWidths=[5.4 * cm, 5.4 * cm, 5.4 * cm],
    )
    table.setStyle(
        TableStyle(
            [
                ("TEXTCOLOR", (0, 0), (-1, 0), DARK),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("LINEBELOW", (0, 0), (-1, 0), 1.5, colors.HexColor("#282828")),
                ("GRID", (0, 0), (-1, -1), 0.5, GRID),
                ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
                ("VALIGN", (0, 1), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
            ]
        )
    )
    return table


class _NumberedCanvas(pdf_canvas.Canvas):
    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        self._saved_page_states: list[dict[str, object]] = []

    def showPage(self) -> None:
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self) -> None:
        if getattr(self, "_code", None):
            self._saved_page_states.append(dict(self.__dict__))
        total_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            _draw_page_number(self, total_pages)
            super().showPage()
        super().save()


def _draw_page_static_decorations(pdf, _doc=None) -> None:
    page_width, page_height = LETTER
    pdf.saveState()
    template = _template_path()
    if template is not None:
        pdf.drawImage(str(template), 0, 0, width=page_width, height=page_height, mask="auto")
    pdf.setFillColor(DARK)
    pdf.setFont("Helvetica-Bold", 11)
    pdf.drawCentredString(page_width / 2, page_height - 34, HEADER_LINE_1)
    pdf.drawCentredString(page_width / 2, page_height - 48, HEADER_LINE_2)
    pdf.restoreState()


def _draw_page_number(pdf, total_pages: int) -> None:
    page_width, page_height = LETTER
    pdf.saveState()
    pdf.setFillColor(DARK)
    pdf.setFont("Helvetica", 9)
    pdf.drawRightString(page_width - 36, page_height - 28, f"Pagina {pdf._pageNumber} de {total_pages}")
    pdf.restoreState()


def build_criterio_evaluacion_pdf(output_path: str | Path, payload: dict[str, Any]) -> Path:
    destination = Path(output_path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    styles = _build_styles()

    document = SimpleDocTemplate(
        str(destination),
        pagesize=LETTER,
        leftMargin=1.15 * cm,
        rightMargin=1.15 * cm,
        topMargin=3.3 * cm,
        bottomMargin=2.2 * cm,
        title="Ficha de consultas normativas",
        author="Sistema de Calibracion V&C",
    )

    story: list[object] = [Spacer(1, 0.15 * cm)]
    story.append(_build_summary_table(payload))
    story.append(Spacer(1, 0.45 * cm))
    story.append(_build_consultation_table(payload, styles))
    story.append(Spacer(1, 0.6 * cm))

    executive_name = _text(payload.get("executive_name") or payload.get("inspector_supervised"), "Sin ejecutivo")
    signature_lines = [
        f"Elaboro: {executive_name} / Ejecutivo Tecnico / Firma __________________",
        "Reviso: Nombre / Cargo / Firma __________________",
        "Autorizo cliente: Nombre / Cargo / Firma __________________",
    ]
    for line in signature_lines:
        story.append(Paragraph(line, styles["VCSignature"]))
        story.append(Spacer(1, 0.18 * cm))

    document.build(
        story,
        onFirstPage=_draw_page_static_decorations,
        onLaterPages=_draw_page_static_decorations,
        canvasmaker=_NumberedCanvas,
    )
    return destination
