from __future__ import annotations

from pathlib import Path
from statistics import mean

from reportlab.lib import colors
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


ACCENT = colors.HexColor("#ECD925")
DARK = colors.HexColor("#282828")
LIGHT = colors.HexColor("#F8F9FA")
SUCCESS = colors.HexColor("#008D53")
WARNING = colors.HexColor("#ff1500")
ROOT_DIR = Path(__file__).resolve().parents[1]
TEMPLATE_PATH = ROOT_DIR / "img" / "plantilla.png"


def _safe(value: object, fallback: str = "No capturado") -> str:
    text = str(value).strip() if value not in (None, "") else ""
    return text or fallback


def _to_float(value: object, fallback: float = 0.0) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return fallback


def _to_int(value: object, fallback: int = 0) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return fallback


def _normalize_answers(raw_answers: object) -> list[dict[str, str]]:
    normalized: list[dict[str, str]] = []
    if not isinstance(raw_answers, list):
        return normalized

    for item in raw_answers:
        if not isinstance(item, dict):
            continue
        normalized.append(
            {
                "activity": _safe(item.get("activity"), "Actividad sin descripcion"),
                "result": str(item.get("result", "")).strip().lower(),
                "observations": _safe(item.get("observations"), "--"),
            }
        )

    return normalized


def _normalize_technical_rows(raw_rows: object) -> list[dict[str, str]]:
    normalized: list[dict[str, str]] = []
    if not isinstance(raw_rows, list):
        return normalized

    for item in raw_rows:
        if not isinstance(item, dict):
            continue

        sku = _safe(item.get("sku"), "--")
        applicable_norm = _safe(item.get("applicable_norm"), "Sin norma")
        observations = _safe(item.get("observations"), "--")

        c_nc = str(item.get("c_nc", "")).strip().upper()
        result = str(item.get("result", "")).strip().lower()
        if c_nc not in {"C", "NC"}:
            if result == "conforme":
                c_nc = "C"
            elif result == "no_conforme":
                c_nc = "NC"
            else:
                c_nc = "--"

        normalized.append(
            {
                "sku": sku,
                "applicable_norm": applicable_norm,
                "c_nc": c_nc,
                "observations": observations,
            }
        )

    return normalized


def _normalize_score_by_norm(raw_scores: object) -> dict[str, float]:
    normalized: dict[str, float] = {}
    if not isinstance(raw_scores, dict):
        return normalized

    for norm_name, raw_score in raw_scores.items():
        clean_norm = str(norm_name).strip()
        if not clean_norm:
            continue
        score = _to_float(raw_score, -1.0)
        if score < 0:
            continue
        normalized[clean_norm] = round(score, 1)

    return normalized


def _header_cell(text: str, styles) -> Paragraph:
    return Paragraph(text, styles["VCHeaderCell"])


def _build_answers_table(answers: list[dict[str, str]], styles) -> Table:
    rows: list[list[object]] = [
        [
            _header_cell("Actividad especifica", styles),
            _header_cell("Conforme", styles),
            _header_cell("No conforme", styles),
            _header_cell("No aplica", styles),
            _header_cell("Observaciones", styles),
        ]
    ]

    if not answers:
        rows.append(["Sin reactivos capturados", "", "", "", "--"])
    else:
        for answer in answers:
            result = str(answer.get("result", "")).strip().lower()
            rows.append(
                [
                    Paragraph(_safe(answer.get("activity"), "Actividad sin descripcion"), styles["VCBodySmall"]),
                    "X" if result == "conforme" else "",
                    "X" if result == "no_conforme" else "",
                    "X" if result == "no_aplica" else "",
                    Paragraph(_safe(answer.get("observations"), "--"), styles["VCBodySmall"]),
                ]
            )

    table = Table(rows, colWidths=[198, 56, 78, 66, 142], repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), ACCENT),
                ("TEXTCOLOR", (0, 0), (-1, 0), DARK),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT]),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D9DDE1")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                ("ALIGN", (1, 1), (3, -1), "CENTER"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, 0), 8),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
                ("TOPPADDING", (0, 1), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 1), (-1, -1), 6),
            ]
        )
    )
    return table


def _build_technical_table(rows_data: list[dict[str, str]], styles) -> Table:
    rows: list[list[object]] = [
        [
            _header_cell("SKU/ITEM/CODIGO/UPC INSPECCIONADO", styles),
            _header_cell("NOM APLICABLE", styles),
            _header_cell("C/NC", styles),
            _header_cell("OBSERVACIONES", styles),
        ]
    ]

    if not rows_data:
        rows.append(["Sin filas registradas", "--", "--", "--"])
    else:
        for row in rows_data:
            rows.append(
                [
                    Paragraph(_safe(row.get("sku"), "--"), styles["VCBodySmall"]),
                    Paragraph(_safe(row.get("applicable_norm"), "Sin norma"), styles["VCBodySmall"]),
                    _safe(row.get("c_nc"), "--"),
                    Paragraph(_safe(row.get("observations"), "--"), styles["VCBodySmall"]),
                ]
            )

    table = Table(rows, colWidths=[182, 150, 52, 156], repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), ACCENT),
                ("TEXTCOLOR", (0, 0), (-1, 0), DARK),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT]),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D9DDE1")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                ("ALIGN", (2, 1), (2, -1), "CENTER"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, 0), 8),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
                ("TOPPADDING", (0, 1), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 1), (-1, -1), 6),
            ]
        )
    )
    return table


def _calculate_average_score(score_by_norm: dict[str, float], fallback: float = 0.0) -> float:
    if score_by_norm:
        return round(mean(score_by_norm.values()), 1)
    return round(fallback, 1)


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
            self._draw_page_counter(total_pages)
            super().showPage()
        super().save()

    def _draw_page_counter(self, total_pages: int) -> None:
        page_width, page_height = LETTER
        self.saveState()
        self.setFillColor(DARK)
        self.setFont("Helvetica", 9)
        self.drawRightString(page_width - 36, page_height - 28, f"Paginas {self._pageNumber} de {total_pages}")
        self.restoreState()


def _decorate_page(pdf, _doc) -> None:
    page_width, page_height = LETTER
    pdf.saveState()
    if TEMPLATE_PATH.exists():
        pdf.drawImage(str(TEMPLATE_PATH), 0, 0, width=page_width, height=page_height, mask="auto")
    pdf.setFillColor(DARK)
    pdf.setFont("Helvetica-Bold", 14)
    pdf.drawCentredString(page_width / 2, page_height - 28, "Formato de Supervision")
    pdf.restoreState()


def build_formato_supervision_pdf(output_path: str | Path, payload: dict) -> Path:
    destination = Path(output_path)
    destination.parent.mkdir(parents=True, exist_ok=True)

    raw_score = _to_float(payload.get("score"), 0.0)
    inspector_supervised = _safe(payload.get("inspector_supervised") or payload.get("inspector_name"))
    supervisor_name = _safe(payload.get("evaluator"))
    protocol_answers = _normalize_answers(payload.get("protocol_answers", []))
    process_answers = _normalize_answers(payload.get("process_answers", []))
    technical_rows = _normalize_technical_rows(payload.get("technical_normative_rows", []))
    score_by_norm = _normalize_score_by_norm(payload.get("score_by_norm", {}))
    average_score = _calculate_average_score(score_by_norm, raw_score)

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="VCHeading", parent=styles["Heading1"], fontName="Helvetica-Bold", fontSize=20, textColor=DARK, spaceAfter=12))
    styles.add(ParagraphStyle(name="VCSubheading", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=12, textColor=DARK, spaceAfter=8))
    styles.add(ParagraphStyle(name="VCBody", parent=styles["BodyText"], fontName="Helvetica", fontSize=10, leading=14, textColor=DARK))
    styles.add(ParagraphStyle(name="VCBodySmall", parent=styles["BodyText"], fontName="Helvetica", fontSize=9, leading=12, textColor=DARK))
    styles.add(ParagraphStyle(name="VCHeaderCell", parent=styles["BodyText"], fontName="Helvetica-Bold", fontSize=8, leading=9, alignment=1, textColor=DARK))

    document = SimpleDocTemplate(
        str(destination),
        pagesize=LETTER,
        leftMargin=36,
        rightMargin=36,
        topMargin=72,
        bottomMargin=36,
    )

    story = [Spacer(1, 6)]

    summary_data = [
        ["Fecha de supervision", _safe(payload.get("visit_date"))],
        ["Cliente / Almacen", _safe(payload.get("client"))],
        ["Ejecutivo supervisado", inspector_supervised],
        ["Nombre del supervisor", supervisor_name],
        ["Norma evaluada", _safe(payload.get("selected_norm"))],
        ["Calificacion final (promedio)", f"{average_score:.1f}%"],
    ]

    summary_table = Table(summary_data, colWidths=[170, 340])
    summary_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, -1), ACCENT),
                ("TEXTCOLOR", (0, 0), (-1, -1), DARK),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTNAME", (1, 0), (1, -1), "Helvetica"),
                ("ROWBACKGROUNDS", (1, 0), (1, -1), [LIGHT, colors.white]),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D9DDE1")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )

    story.append(summary_table)
    story.append(Spacer(1, 14))
    story.append(Paragraph("Supervision de protocolo y habilidades blandas", styles["VCSubheading"]))
    story.append(_build_answers_table(protocol_answers, styles))
    story.append(Spacer(1, 12))
    story.append(Paragraph("Supervision de procesos", styles["VCSubheading"]))
    story.append(_build_answers_table(process_answers, styles))
    story.append(Spacer(1, 12))
    story.append(Paragraph("Supervision tecnico normativa", styles["VCSubheading"]))
    story.append(_build_technical_table(technical_rows, styles))
    document.build(
        story,
        onFirstPage=_decorate_page,
        onLaterPages=_decorate_page,
        canvasmaker=_NumberedCanvas,
    )
    return destination
