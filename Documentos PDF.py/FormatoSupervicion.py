from __future__ import annotations

from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


ACCENT = colors.HexColor("#ECD925")
DARK = colors.HexColor("#282828")
LIGHT = colors.HexColor("#F8F9FA")
SUCCESS = colors.HexColor("#008D53")


def _safe(value: object, fallback: str = "No capturado") -> str:
	text = str(value).strip() if value not in (None, "") else ""
	return text or fallback


def build_formato_supervision_pdf(output_path: str | Path, payload: dict) -> Path:
	destination = Path(output_path)
	destination.parent.mkdir(parents=True, exist_ok=True)

	styles = getSampleStyleSheet()
	styles.add(
		ParagraphStyle(
			name="VCHeading",
			parent=styles["Heading1"],
			fontName="Helvetica-Bold",
			fontSize=20,
			textColor=DARK,
			spaceAfter=12,
		)
	)
	styles.add(
		ParagraphStyle(
			name="VCSubheading",
			parent=styles["Heading2"],
			fontName="Helvetica-Bold",
			fontSize=12,
			textColor=DARK,
			spaceAfter=8,
		)
	)
	styles.add(
		ParagraphStyle(
			name="VCBody",
			parent=styles["BodyText"],
			fontName="Helvetica",
			fontSize=10,
			leading=14,
			textColor=DARK,
		)
	)

	document = SimpleDocTemplate(
		str(destination),
		pagesize=LETTER,
		leftMargin=36,
		rightMargin=36,
		topMargin=36,
		bottomMargin=36,
	)

	story = [
		Paragraph("Formato de Supervision", styles["VCHeading"]),
		Paragraph(
			"Seguimiento operativo para el sistema de calibracion V&C. Este archivo se genera una vez que el formulario del inspector fue completado.",
			styles["VCBody"],
		),
		Spacer(1, 12),
	]

	summary_data = [
		["Inspector", _safe(payload.get("inspector_name"))],
		["Norma evaluada", _safe(payload.get("selected_norm"))],
		["Cliente / sede", _safe(payload.get("client"))],
		["Fecha de visita", _safe(payload.get("visit_date"))],
		["Desempeno", f"{_safe(payload.get('score'), '0')}%"],
		["Estado", _safe(payload.get("status"))],
		["Evaluador", _safe(payload.get("evaluator"))],
		[
			"Normas acreditadas",
			_safe(", ".join(payload.get("accredited_norms", [])), "Sin acreditaciones registradas"),
		],
	]

	summary_table = Table(summary_data, colWidths=[150, 360])
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
	story.append(Spacer(1, 16))
	story.append(Paragraph("Observaciones", styles["VCSubheading"]))
	story.append(Paragraph(_safe(payload.get("observations")), styles["VCBody"]))
	story.append(Spacer(1, 12))
	story.append(Paragraph("Acciones correctivas", styles["VCSubheading"]))
	story.append(Paragraph(_safe(payload.get("corrective_actions")), styles["VCBody"]))
	story.append(Spacer(1, 12))

	closing_data = [
		["Semaforo de seguimiento", "Atencion prioritaria" if float(payload.get("score", 0)) < 90 else "Operacion estable"],
		["Estatus de liberacion", "Documento listo para descarga"],
	]
	closing_table = Table(closing_data, colWidths=[190, 320])
	closing_table.setStyle(
		TableStyle(
			[
				("BACKGROUND", (0, 0), (1, 0), SUCCESS if float(payload.get("score", 0)) >= 90 else colors.HexColor("#ff1500")),
				("TEXTCOLOR", (0, 0), (1, 0), colors.white),
				("FONTNAME", (0, 0), (1, 0), "Helvetica-Bold"),
				("BACKGROUND", (0, 1), (1, 1), LIGHT),
				("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D9DDE1")),
				("LEFTPADDING", (0, 0), (-1, -1), 8),
				("RIGHTPADDING", (0, 0), (-1, -1), 8),
				("TOPPADDING", (0, 0), (-1, -1), 8),
				("BOTTOMPADDING", (0, 0), (-1, -1), 8),
			]
		)
	)
	story.append(closing_table)
	document.build(story)
	return destination
