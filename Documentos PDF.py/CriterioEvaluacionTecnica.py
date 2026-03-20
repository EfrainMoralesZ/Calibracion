from __future__ import annotations

from pathlib import Path
from typing import Any

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


def _text(value: Any, fallback: str = "--") -> str:
    raw = str(value or "").strip()
    return raw or fallback


def _collect_image_paths(folder: str) -> list[Path]:
    base = Path(folder)
    if not folder or not base.exists() or not base.is_dir():
        return []

    allowed = {".png", ".jpg", ".jpeg", ".webp"}
    files: list[Path] = []
    for item in sorted(base.iterdir()):
        if item.is_file() and item.suffix.lower() in allowed:
            files.append(item)
    return files


def build_criterio_evaluacion_pdf(output_path: str | Path, payload: dict[str, Any]) -> None:
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "CriteriaTitle",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=16,
        leading=20,
        textColor=colors.HexColor("#1F2937"),
        spaceAfter=8,
    )
    subtitle_style = ParagraphStyle(
        "CriteriaSubtitle",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=10,
        leading=14,
        textColor=colors.HexColor("#4B5563"),
    )
    body_style = ParagraphStyle(
        "CriteriaBody",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#111827"),
    )

    story = []
    story.append(Paragraph("Criterios Evaluados Por Cliente", title_style))
    story.append(
        Paragraph(
            "Reporte generado desde la pestaña Criterios. Incluye respuestas de protocolo, procesos, técnico normativa y evidencia fotográfica.",
            subtitle_style,
        )
    )
    story.append(Spacer(1, 0.35 * cm))

    summary_rows = [
        ["Cliente", _text(payload.get("client"))],
        ["Ejecutivo supervisado", _text(payload.get("inspector_supervised") or payload.get("inspector_name"))],
        ["Supervisor", _text(payload.get("evaluator"))],
        ["Fecha", _text(payload.get("visit_date") or payload.get("saved_at"))],
        ["NOM evaluada", _text(payload.get("selected_norm"))],
        ["Calificación", f"{_text(payload.get('score'))}%"],
        ["Estatus", _text(payload.get("status"))],
    ]

    summary_table = Table(summary_rows, colWidths=[5.0 * cm, 12.0 * cm])
    summary_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F3F4F6")),
                ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#E2E8F0")),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTNAME", (1, 0), (1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.append(summary_table)
    story.append(Spacer(1, 0.35 * cm))

    def _build_answers_table(title: str, answers: list[dict[str, Any]]) -> None:
        story.append(Paragraph(title, styles["Heading3"]))
        if not answers:
            story.append(Paragraph("Sin criterios capturados.", body_style))
            story.append(Spacer(1, 0.25 * cm))
            return

        rows: list[list[Any]] = [["Actividad", "Resultado", "Observaciones"]]
        for row in answers:
            rows.append(
                [
                    Paragraph(_text(row.get("activity")), body_style),
                    Paragraph(_text(row.get("result")), body_style),
                    Paragraph(_text(row.get("observations")), body_style),
                ]
            )

        table = Table(rows, colWidths=[8.0 * cm, 3.0 * cm, 6.0 * cm], repeatRows=1)
        table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#111827")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                    ("FONTSIZE", (0, 0), (-1, -1), 8),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CBD5E1")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 5),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )
        story.append(table)
        story.append(Spacer(1, 0.25 * cm))

    protocol_answers = payload.get("protocol_answers", []) if isinstance(payload.get("protocol_answers"), list) else []
    process_answers = payload.get("process_answers", []) if isinstance(payload.get("process_answers"), list) else []
    technical_rows = payload.get("technical_normative_rows", []) if isinstance(payload.get("technical_normative_rows"), list) else []

    _build_answers_table("Criterios de Protocolo y Habilidades", protocol_answers)
    _build_answers_table("Criterios de Procesos", process_answers)

    story.append(Paragraph("Criterios Técnico Normativa", styles["Heading3"]))
    if technical_rows:
        tech_data: list[list[Any]] = [["SKU/ITEM/CODIGO", "NOM Aplicable", "C/NC", "Observaciones"]]
        for row in technical_rows:
            tech_data.append(
                [
                    Paragraph(_text(row.get("sku")), body_style),
                    Paragraph(_text(row.get("applicable_norm")), body_style),
                    Paragraph(_text(row.get("c_nc") or row.get("result")), body_style),
                    Paragraph(_text(row.get("observations")), body_style),
                ]
            )

        tech_table = Table(tech_data, colWidths=[4.8 * cm, 4.0 * cm, 2.0 * cm, 6.2 * cm], repeatRows=1)
        tech_table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#111827")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                    ("FONTSIZE", (0, 0), (-1, -1), 8),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CBD5E1")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 5),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )
        story.append(tech_table)
    else:
        story.append(Paragraph("Sin filas técnicas capturadas.", body_style))
    story.append(Spacer(1, 0.35 * cm))

    images = _collect_image_paths(_text(payload.get("image_folder"), ""))
    story.append(Paragraph("Evidencia fotográfica", styles["Heading3"]))
    if not images:
        story.append(Paragraph("No se encontraron imágenes para adjuntar.", body_style))
    else:
        story.append(Paragraph(f"Total de imágenes detectadas: {len(images)}", body_style))
        story.append(Spacer(1, 0.15 * cm))
        for img_path in images[:12]:
            try:
                img = Image(str(img_path))
                img._restrictSize(16.5 * cm, 9.0 * cm)
                story.append(img)
                story.append(Paragraph(_text(img_path.name), subtitle_style))
                story.append(Spacer(1, 0.2 * cm))
            except OSError:
                continue

    doc = SimpleDocTemplate(
        str(output),
        pagesize=letter,
        leftMargin=1.4 * cm,
        rightMargin=1.4 * cm,
        topMargin=1.2 * cm,
        bottomMargin=1.2 * cm,
        title="Criterios por cliente",
        author="Sistema de Calibracion V&C",
    )
    doc.build(story)
