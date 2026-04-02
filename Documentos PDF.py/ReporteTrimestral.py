from __future__ import annotations

from datetime import datetime
from pathlib import Path
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from runtime_paths import resource_path, writable_path


ACCENT_HEX = "#1B2A4A"
DARK_HEX = "#282828"
LIGHT_HEX = "#F8F9FA"
SUCCESS_HEX = "#008D53"
WARNING_HEX = "#ff1500"

ACCENT = colors.HexColor(ACCENT_HEX)
DARK = colors.HexColor(DARK_HEX)
LIGHT = colors.HexColor(LIGHT_HEX)
SUCCESS = colors.HexColor(SUCCESS_HEX)
WARNING = colors.HexColor(WARNING_HEX)


def _template_path() -> Path | None:
	for candidate in (resource_path("img", "plantilla.png"), writable_path("img", "plantilla.png")):
		if candidate.exists():
			return candidate
	return None


def _status_color(status_text: str) -> str:
	"""Return a readable hex color for a status label."""
	tl = status_text.lower()
	if "mayor" in tl:
		return "#006B3C"
	if "menor" in tl:
		return "#B94A2C"
	return DARK_HEX


def _safe(value: object, fallback: str = "--") -> str:
	text = str(value).strip() if value not in (None, "") else ""
	return text or fallback


def _as_int(value: object, fallback: int = 0) -> int:
	try:
		return int(value)
	except (TypeError, ValueError):
		return fallback


def _as_float(value: object, fallback: float = 0.0) -> float:
	try:
		return float(value)
	except (TypeError, ValueError):
		return fallback


def _paragraph(text: object, style) -> Paragraph:
	return Paragraph(escape(_safe(text)), style)


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


def build_trimestral_dashboard_pdf(output_path: str | Path, payload: dict) -> Path:
	destination = Path(output_path)
	destination.parent.mkdir(parents=True, exist_ok=True)

	can_edit = bool(payload.get("can_edit"))
	report_title = _safe(payload.get("report_title"), "Reporte trimestral")
	scope_label = _safe(payload.get("scope_label"), "Global")
	viewer_name = _safe(payload.get("viewer_name"), "Sistema")
	exported_at = _safe(payload.get("exported_at"), datetime.now().strftime("%Y-%m-%d %H:%M"))
	summary_text = _safe(payload.get("summary_text"), "Sin resumen disponible.")
	metrics = payload.get("metrics") if isinstance(payload.get("metrics"), dict) else {}
	rows = [item for item in payload.get("bars", []) if isinstance(item, dict)]

	styles = getSampleStyleSheet()
	styles.add(ParagraphStyle(name="VCBody", parent=styles["BodyText"], fontName="Helvetica", fontSize=10, leading=13, textColor=DARK))
	styles.add(ParagraphStyle(name="VCBodySmall", parent=styles["BodyText"], fontName="Helvetica", fontSize=9, leading=11, textColor=DARK))
	styles.add(ParagraphStyle(name="VCBodySmallBold", parent=styles["BodyText"], fontName="Helvetica-Bold", fontSize=9, leading=11, textColor=DARK))
	styles.add(ParagraphStyle(name="VCSubheading", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=12, textColor=DARK, spaceAfter=8))
	styles.add(ParagraphStyle(name="VCHeaderCell", parent=styles["BodyText"], fontName="Helvetica-Bold", fontSize=8, leading=9, alignment=1, textColor=DARK))

	document = SimpleDocTemplate(
		str(destination),
		pagesize=LETTER,
		leftMargin=36,
		rightMargin=36,
		topMargin=72,
		bottomMargin=36,
	)

	def _decorate_page(pdf, _doc) -> None:
		page_width, page_height = LETTER
		pdf.saveState()
		tpl = _template_path()
		if tpl is not None:
			pdf.drawImage(str(tpl), 0, 0, width=page_width, height=page_height, mask="auto")
		pdf.setFillColor(DARK)
		pdf.setFont("Helvetica-Bold", 14)
		pdf.drawCentredString(page_width / 2, page_height - 28, report_title)
		pdf.setFont("Helvetica", 9)
		pdf.drawRightString(page_width - 36, page_height - 42, exported_at)
		pdf.restoreState()

	story: list[object] = [Spacer(1, 8)]

	summary_data: list[list[object]] = [
		[Paragraph("Tipo de vista", styles["VCBodySmallBold"]), _paragraph("Global" if can_edit else "Ejecutivo tecnico", styles["VCBodySmall"])],
		[Paragraph("Alcance", styles["VCBodySmallBold"]), _paragraph(scope_label, styles["VCBodySmall"])],
		[Paragraph("Generado por", styles["VCBodySmallBold"]), _paragraph(viewer_name, styles["VCBodySmall"])],
		[Paragraph("Fecha de exportacion", styles["VCBodySmallBold"]), _paragraph(exported_at, styles["VCBodySmall"])],
		[Paragraph("Resumen", styles["VCBodySmallBold"]), _paragraph(summary_text, styles["VCBodySmall"])],
	]

	if not can_edit:
		summary_data.extend(
			[
				[Paragraph("Normas acreditadas", styles["VCBodySmallBold"]), _paragraph(metrics.get("total_norms", 0), styles["VCBodySmall"])],
				[Paragraph("Usos totales", styles["VCBodySmallBold"]), _paragraph(metrics.get("total_uses", 0), styles["VCBodySmall"])],
				[Paragraph("Mayor demanda", styles["VCBodySmallBold"]), _paragraph(f'{metrics.get("highest_usage", 0)} usos', styles["VCBodySmall"])],
				[Paragraph("Menor demanda", styles["VCBodySmallBold"]), _paragraph(f'{metrics.get("lowest_usage", 0)} usos', styles["VCBodySmall"])],
			]
		)

	summary_table = Table(summary_data, colWidths=[150, 390])
	summary_table.setStyle(
		TableStyle(
			[
				("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#F0F2F5")),
				("BACKGROUND", (1, 0), (1, -1), colors.white),
				("TEXTCOLOR", (0, 0), (-1, -1), DARK),
				("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
				("BOX", (0, 0), (-1, -1), 0.75, colors.HexColor("#C8CDD2")),
				("LINEBELOW", (0, 0), (-1, -2), 0.5, colors.HexColor("#D9DDE1")),
				("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
				("LEFTPADDING", (0, 0), (-1, -1), 10),
				("RIGHTPADDING", (0, 0), (-1, -1), 10),
				("TOPPADDING", (0, 0), (-1, -1), 7),
				("BOTTOMPADDING", (0, 0), (-1, -1), 7),
			]
		)
	)

	story.append(summary_table)
	story.append(Spacer(1, 14))
	story.append(Paragraph("Detalle de la grafica exportada", styles["VCSubheading"]))

	if not rows:
		story.append(Paragraph("No hay registros disponibles para exportar.", styles["VCBody"]))
	else:
		table_rows: list[list[object]] = []
		if can_edit:
			table_rows.append(
				[
					Paragraph("#", styles["VCHeaderCell"]),
					Paragraph("Ejecutivo tecnico", styles["VCHeaderCell"]),
					Paragraph("Norma", styles["VCHeaderCell"]),
					Paragraph("Usos", styles["VCHeaderCell"]),
					Paragraph("Frecuencia", styles["VCHeaderCell"]),
					Paragraph("Estado", styles["VCHeaderCell"]),
				]
			)
			col_widths = [28, 150, 120, 50, 70, 108]
			usage_col = 3
			state_col = 5
		else:
			table_rows.append(
				[
					Paragraph("#", styles["VCHeaderCell"]),
					Paragraph("Norma", styles["VCHeaderCell"]),
					Paragraph("Usos", styles["VCHeaderCell"]),
					Paragraph("Frecuencia", styles["VCHeaderCell"]),
					Paragraph("Estado", styles["VCHeaderCell"]),
				]
			)
			col_widths = [28, 200, 50, 100, 140]
			usage_col = 2
			state_col = 4

		for index, item in enumerate(rows, start=1):
			usage_count = _as_int(item.get("usage_count"), 0)
			status_text = _safe(item.get("status"), "--")
			frequency_text = "1 uso" if usage_count == 1 else f"{usage_count} usos"
			usage_paragraph = Paragraph(f"<b>{usage_count}</b>", styles["VCBodySmall"])
			status_color_hex = _status_color(status_text)
			status_paragraph = Paragraph(f'<font color="{status_color_hex}"><b>{escape(status_text)}</b></font>', styles["VCBodySmall"])

			if can_edit:
				table_rows.append(
					[
						Paragraph(str(index), styles["VCBodySmall"]),
						_paragraph(item.get("inspector"), styles["VCBodySmall"]),
						_paragraph(item.get("norm"), styles["VCBodySmall"]),
						usage_paragraph,
						_paragraph(frequency_text, styles["VCBodySmall"]),
						status_paragraph,
					]
				)
			else:
				table_rows.append(
					[
						Paragraph(str(index), styles["VCBodySmall"]),
						_paragraph(item.get("norm"), styles["VCBodySmall"]),
						usage_paragraph,
						_paragraph(frequency_text, styles["VCBodySmall"]),
						status_paragraph,
					]
				)

		detail_table = Table(table_rows, colWidths=col_widths, repeatRows=1)
		style_commands: list[tuple[object, ...]] = [
			("TEXTCOLOR", (0, 0), (-1, 0), DARK),
			("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
			("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT]),
			("BOX", (0, 0), (-1, -1), 0.75, colors.HexColor("#C8CDD2")),
			("LINEBELOW", (0, 0), (-1, 0), 1.5, colors.HexColor("#282828")),
			("LINEBELOW", (0, 1), (-1, -1), 0.5, colors.HexColor("#D9DDE1")),
			("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
			("LEFTPADDING", (0, 0), (-1, -1), 8),
			("RIGHTPADDING", (0, 0), (-1, -1), 8),
			("TOPPADDING", (0, 0), (-1, 0), 9),
			("BOTTOMPADDING", (0, 0), (-1, 0), 9),
			("TOPPADDING", (0, 1), (-1, -1), 7),
			("BOTTOMPADDING", (0, 1), (-1, -1), 7),
			("ALIGN", (0, 0), (0, -1), "CENTER"),
			("ALIGN", (usage_col, 1), (usage_col, -1), "CENTER"),
			("ALIGN", (state_col, 1), (state_col, -1), "CENTER"),
		]
		detail_table.setStyle(TableStyle(style_commands))
		story.append(detail_table)

	document.build(
		story,
		onFirstPage=_decorate_page,
		onLaterPages=_decorate_page,
		canvasmaker=_NumberedCanvas,
	)
	return destination