"""Native PDF rendering for the hosted Transfer News application.

This renderer intentionally avoids Microsoft Word, LibreOffice, and Docker.
It consumes the same verified workbook as the DOCX renderer and produces a
portable landscape PDF with the same cover, zone tables, role palette,
watermark, confidentiality footer, and assignment key.
"""
from __future__ import annotations

from datetime import datetime
from html import escape, unescape
from pathlib import Path

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import (
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from generate_transfer_document import (
    ACCENT_GOLD,
    ASSIGNMENT_KEY_ENTRIES,
    BANNER_GREEN,
    BLUE,
    BLUE_TINT,
    DEFAULT_MISSION_NAME,
    DEFAULT_PREPARED_BY,
    DEFAULT_PRESIDENT,
    GOLD_TEXT,
    GOLD_TINT,
    GREEN_TINT,
    MISSION_WATERMARK_PATH,
    MUTED_GRAY,
    NEW_MISSIONARY_FILL,
    NEW_MISSIONARY_TEXT,
    PANEL_GRAY,
    TEXT_DARK,
    TEXT_MID,
    WATERMARK_COVER_CENTER_Y_IN,
    WATERMARK_SIZE_IN,
    WATERMARK_ZONE_CENTER_Y_IN,
    ZEBRA_TINT,
    badge_colors_for,
    compute_stats,
    load_news_data,
    read_meta_values,
    read_new_missionary_count,
)


PAGE_WIDTH, PAGE_HEIGHT = landscape(letter)


def _color(value: str) -> colors.Color:
    return colors.HexColor(f"#{value}")


def _font_names() -> tuple[str, str, str]:
    """Register a portable sans family, with real bold and italic faces."""
    base_dir = Path(__file__).resolve().parent
    candidates = [
        (
            base_dir / "assets" / "Aptos.ttf",
            base_dir / "assets" / "Aptos-Bold.ttf",
            base_dir / "assets" / "Aptos-Italic.ttf",
        ),
        (
            Path("C:/Windows/Fonts/arial.ttf"),
            Path("C:/Windows/Fonts/arialbd.ttf"),
            Path("C:/Windows/Fonts/ariali.ttf"),
        ),
        (
            Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
            Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
            Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf"),
        ),
    ]
    for regular, bold, italic in candidates:
        if not all(path.exists() for path in (regular, bold, italic)):
            continue
        try:
            pdfmetrics.registerFont(TTFont("TransferSans", str(regular)))
            pdfmetrics.registerFont(TTFont("TransferSans-Bold", str(bold)))
            pdfmetrics.registerFont(TTFont("TransferSans-Italic", str(italic)))
            return "TransferSans", "TransferSans-Bold", "TransferSans-Italic"
        except Exception:
            continue
    return "Helvetica", "Helvetica-Bold", "Helvetica-Oblique"


FONT, FONT_BOLD, FONT_ITALIC = _font_names()


def _p(text: object, style: ParagraphStyle, *, markup: bool = False) -> Paragraph:
    value = str(text or "")
    if not markup:
        # Workbook values may already contain HTML entities. Decode them once
        # before escaping for ReportLab so "&amp;" is printed as "&".
        value = unescape(value)
        value = escape(value).replace("\n", "<br/>")
    return Paragraph(value, style)


def _styles() -> dict[str, ParagraphStyle]:
    styles = getSampleStyleSheet()
    return {
        "cover_mission": ParagraphStyle(
            "CoverMission", parent=styles["Title"], fontName=FONT_BOLD,
            fontSize=25, leading=29, textColor=_color(ACCENT_GOLD), alignment=TA_CENTER,
            spaceAfter=0,
        ),
        "cover_church": ParagraphStyle(
            "CoverChurch", parent=styles["Normal"], fontName=FONT,
            fontSize=7.5, leading=10, textColor=colors.white, alignment=TA_CENTER,
        ),
        "title": ParagraphStyle(
            "TransferTitle", parent=styles["Title"], fontName=FONT_BOLD,
            fontSize=25, leading=29, textColor=_color(BANNER_GREEN), alignment=TA_CENTER,
            spaceAfter=4,
        ),
        "subtitle": ParagraphStyle(
            "TransferSubtitle", parent=styles["Normal"], fontName=FONT,
            fontSize=8, leading=10, textColor=_color(MUTED_GRAY), alignment=TA_CENTER,
        ),
        "zone": ParagraphStyle(
            "ZoneTitle", parent=styles["Heading1"], fontName=FONT_BOLD,
            fontSize=14, leading=18, textColor=_color(BANNER_GREEN), alignment=TA_LEFT,
        ),
        "cell": ParagraphStyle(
            "Cell", parent=styles["Normal"], fontName=FONT,
            fontSize=13.5, leading=16.5, textColor=_color(TEXT_MID), alignment=TA_LEFT,
        ),
        "cell_center": ParagraphStyle(
            "CellCenter", parent=styles["Normal"], fontName=FONT,
            fontSize=13.5, leading=16.5, textColor=_color(TEXT_MID), alignment=TA_CENTER,
        ),
        "cell_bold": ParagraphStyle(
            "CellBold", parent=styles["Normal"], fontName=FONT_BOLD,
            fontSize=13.5, leading=16.5, textColor=_color(TEXT_DARK), alignment=TA_LEFT,
        ),
        "cell_header": ParagraphStyle(
            "CellHeader", parent=styles["Normal"], fontName=FONT_BOLD,
            fontSize=13.5, leading=16.5, textColor=colors.white, alignment=TA_CENTER,
        ),
        "small": ParagraphStyle(
            "Small", parent=styles["Normal"], fontName=FONT,
            fontSize=6.7, leading=8, textColor=_color(TEXT_MID), alignment=TA_LEFT,
        ),
        "small_center": ParagraphStyle(
            "SmallCenter", parent=styles["Normal"], fontName=FONT,
            fontSize=6.7, leading=8, textColor=_color(TEXT_MID), alignment=TA_CENTER,
        ),
    }


def _draw_watermark(canvas, center_y: float) -> None:
    if not MISSION_WATERMARK_PATH.exists():
        return
    size = WATERMARK_SIZE_IN * inch
    canvas.saveState()
    canvas.drawImage(
        str(MISSION_WATERMARK_PATH),
        (PAGE_WIDTH - size) / 2,
        center_y - size / 2,
        width=size,
        height=size,
        preserveAspectRatio=True,
        mask="auto",
    )
    canvas.restoreState()


def _font_ascent(font_name: str, size: float) -> float:
    ascent, _descent = pdfmetrics.getAscentDescent(font_name, size)
    return ascent


def _baseline_for_top(font_name: str, size: float, top: float) -> float:
    return PAGE_HEIGHT - top - _font_ascent(font_name, size)


def _draw_centered_text(
    canvas,
    text: str,
    *,
    top: float,
    font_name: str,
    size: float,
    color: str | colors.Color,
    char_space: float = 0,
    center_x: float = PAGE_WIDTH / 2,
) -> None:
    value = str(text or "")
    text_width = pdfmetrics.stringWidth(value, font_name, size)
    if value:
        text_width += max(len(value) - 1, 0) * char_space
    text_object = canvas.beginText()
    text_object.setTextOrigin(center_x - text_width / 2, _baseline_for_top(font_name, size, top))
    text_object.setFont(font_name, size)
    text_object.setCharSpace(char_space)
    text_object.setFillColor(_color(color) if isinstance(color, str) else color)
    text_object.textLine(value)
    canvas.drawText(text_object)


def _draw_segments(canvas, x: float, baseline: float, segments: list[tuple[str, str, float, str]]) -> None:
    cursor = x
    for value, font_name, size, color in segments:
        canvas.setFont(font_name, size)
        canvas.setFillColor(_color(color))
        canvas.drawString(cursor, baseline, value)
        cursor += pdfmetrics.stringWidth(value, font_name, size)


def _wrap_text(text: str, font_name: str, size: float, width: float) -> list[str]:
    words = str(text or "").split()
    if not words:
        return [""]
    lines: list[str] = []
    current = words[0]
    for word in words[1:]:
        candidate = f"{current} {word}"
        if pdfmetrics.stringWidth(candidate, font_name, size) <= width:
            current = candidate
        else:
            lines.append(current)
            current = word
    lines.append(current)
    return lines


def _zone_label(zone: object) -> str:
    value = str(zone or "").strip().upper()
    return value if value.endswith(" ZONE") else f"{value} ZONE"


def _draw_publication_cell(
    canvas,
    *,
    x: float,
    top: float,
    width: float,
    height: float,
    label: str,
    value: str,
    note: str,
) -> None:
    inner_width = width - 12
    size = 14.0
    leading = 18.6
    value_lines = _wrap_text(value.upper(), FONT_BOLD, size, inner_width)
    note_size = 13.0 if note.upper() == "ASSISTANT TO THE PRESIDENT" else size
    note_lines = _wrap_text(note.upper(), FONT_ITALIC, note_size, inner_width)
    lines = [
        (label.upper(), FONT_BOLD, MUTED_GRAY, size),
        *[(line, FONT_BOLD, BANNER_GREEN, size) for line in value_lines],
        *[(line, FONT_ITALIC, MUTED_GRAY, note_size) for line in note_lines],
    ]
    if len(lines) * leading > height - 10:
        size = 12.0
        leading = 16.0
        value_lines = _wrap_text(value.upper(), FONT_BOLD, size, inner_width)
        note_size = size
        note_lines = _wrap_text(note.upper(), FONT_ITALIC, note_size, inner_width)
        lines = [
            (label.upper(), FONT_BOLD, MUTED_GRAY, size),
            *[(line, FONT_BOLD, BANNER_GREEN, size) for line in value_lines],
            *[(line, FONT_ITALIC, MUTED_GRAY, note_size) for line in note_lines],
        ]

    block_height = len(lines) * leading
    line_top = top + (height - block_height) / 2 + 6.7
    for text, font_name, color, line_size in lines:
        baseline = _baseline_for_top(font_name, line_size, line_top)
        canvas.setFont(font_name, line_size)
        canvas.setFillColor(_color(color))
        canvas.drawString(x + 8, baseline, text)
        line_top += leading


def _draw_reference_cover(
    canvas,
    mission_name: str,
    transfer_title: str,
    stats: dict[str, int],
    grouped_zones: list[tuple[str, pd.DataFrame]],
    president: str,
    prepared_by: str,
) -> None:
    """Draw the measured publication cover from the supplied mission PDF."""
    canvas.saveState()
    _draw_watermark(canvas, PAGE_HEIGHT - WATERMARK_COVER_CENTER_Y_IN * inch)

    # Full-width mission masthead, measured from the supplied publication.
    canvas.setFillColor(_color(BANNER_GREEN))
    canvas.rect(0, PAGE_HEIGHT - 89.4, PAGE_WIDTH, 18.6, fill=1, stroke=0)
    canvas.rect(0, PAGE_HEIGHT - 110.0, PAGE_WIDTH, 18.6, fill=1, stroke=0)
    canvas.setFillColor(_color(ACCENT_GOLD))
    canvas.rect(0, PAGE_HEIGHT - 112.2, PAGE_WIDTH, 2.2, fill=1, stroke=0)
    _draw_centered_text(
        canvas, mission_name.upper(), top=77.8, font_name=FONT_BOLD, size=14,
        color=ACCENT_GOLD, char_space=2.8,
    )
    _draw_centered_text(
        canvas, "THE CHURCH OF JESUS CHRIST OF LATTER-DAY SAINTS", top=98.2,
        font_name=FONT, size=14, color=colors.white, char_space=0.08,
    )

    cover_title = transfer_title.upper()
    title_size = 14.0
    while title_size > 11 and pdfmetrics.stringWidth(cover_title, FONT_BOLD, title_size) > 560:
        title_size -= 0.5
    _draw_centered_text(
        canvas, cover_title, top=135.3, font_name=FONT_BOLD, size=title_size,
        color=BANNER_GREEN, char_space=0.35,
    )
    _draw_centered_text(
        canvas, "COMPLETE ZONE-BY-ZONE TRANSFER ASSIGNMENTS - MISSION LEADERSHIP COUNCIL",
        top=154.1, font_name=FONT_ITALIC, size=12.7, color=MUTED_GRAY,
    )

    # Compact, single-strip statistics cards.
    tiles = [
        ("ZONES", stats.get("Zones", 0), GREEN_TINT, BANNER_GREEN),
        ("MISSIONARIES", stats.get("Total Missionaries in Mission", 0), BLUE_TINT, BLUE),
        ("LEADERSHIP ROLES", stats.get("Leadership Roles", 0), GOLD_TINT, GOLD_TEXT),
        ("NEW MISSIONARIES", stats.get("New Missionaries", 0), NEW_MISSIONARY_FILL, NEW_MISSIONARY_TEXT),
    ]
    tile_x, tile_top, tile_width, tile_height = 60.9, 183.4, 164.925, 44.5
    for index, (label, value, background, foreground) in enumerate(tiles):
        x = tile_x + index * tile_width
        y = PAGE_HEIGHT - tile_top - tile_height
        canvas.setFillColor(_color(background))
        canvas.setStrokeColor(_color("DDDDDD"))
        canvas.setLineWidth(0.35)
        canvas.rect(x, y, tile_width, tile_height, fill=1, stroke=1)
        canvas.setFillColor(_color(foreground))
        canvas.rect(x, PAGE_HEIGHT - tile_top - 2.2, tile_width, 2.2, fill=1, stroke=0)
        _draw_centered_text(
            canvas, str(value), top=193.6, font_name=FONT_BOLD, size=14,
            color=foreground, center_x=x + tile_width / 2,
        )
        _draw_centered_text(
            canvas, label, top=213.2, font_name=FONT_BOLD, size=14,
            color=TEXT_MID, center_x=x + tile_width / 2,
        )

    # Zone index: three columns and four publication rows for the mission's 12 zones.
    zone_x, zone_top, zone_width = 40.5, 264.0, 700.1
    header_height, body_height = 26.9, 103.5
    row_count = max((len(grouped_zones) + 2) // 3, 1)
    row_height = body_height / row_count
    canvas.setFillColor(_color(BANNER_GREEN))
    canvas.rect(zone_x, PAGE_HEIGHT - zone_top - header_height, zone_width, header_height, fill=1, stroke=0)
    _draw_centered_text(canvas, "ZONE INDEX", top=275.2, font_name=FONT_BOLD, size=14, color=colors.white)
    cell_width = zone_width / 3
    for index in range(row_count * 3):
        row, column = divmod(index, 3)
        cell_top = zone_top + header_height + row * row_height
        cell_x = zone_x + column * cell_width
        canvas.setFillColor(_color(PANEL_GRAY))
        canvas.setStrokeColor(_color("DEDEDE"))
        canvas.setLineWidth(0.35)
        canvas.rect(cell_x, PAGE_HEIGHT - cell_top - row_height, cell_width, row_height, fill=1, stroke=1)
        if index >= len(grouped_zones):
            continue
        zone, rows = grouped_zones[index]
        baseline = PAGE_HEIGHT - (cell_top + row_height / 2 + 5.4)
        _draw_segments(canvas, cell_x + 6.5, baseline, [
            (f"{index + 1:02d}", FONT_BOLD, 14, GOLD_TEXT),
            (f"  {_zone_label(zone)}", FONT_BOLD, 14, BANNER_GREEN),
            (f"  ({len(rows)})", FONT_ITALIC, 14, MUTED_GRAY),
        ])

    # The publication details panel is intentionally tall and vertically centered.
    panel_x, panel_top, panel_width, panel_height = 66.0, 402.3, 660.0, 108.2
    panel_cell_width = panel_width / 3
    for index in range(3):
        x = panel_x + index * panel_cell_width
        canvas.setFillColor(_color(PANEL_GRAY))
        canvas.setStrokeColor(_color("DEDEDE"))
        canvas.setLineWidth(0.35)
        canvas.rect(x, PAGE_HEIGHT - panel_top - panel_height, panel_cell_width, panel_height, fill=1, stroke=1)
    effective = transfer_title.upper().replace(" TRANSFER NEWS", " TRANSFER")
    date_text = f"DOCUMENT DATE: {datetime.now():%d %B, %Y}."
    panel_defs = [
        ("PREPARED BY", prepared_by, "ASSISTANT TO THE PRESIDENT"),
        ("APPROVED BY", president, f"{mission_name} PRESIDENT"),
        ("EFFECTIVE", effective, date_text),
    ]
    for index, (label, value, note) in enumerate(panel_defs):
        _draw_publication_cell(
            canvas,
            x=panel_x + index * panel_cell_width,
            top=panel_top,
            width=panel_cell_width,
            height=panel_height,
            label=label,
            value=value,
            note=note,
        )

    _draw_centered_text(
        canvas, "STRICTLY CONFIDENTIAL - FOR MISSION USE ONLY", top=527.7,
        font_name=FONT_ITALIC, size=14, color=MUTED_GRAY,
    )
    canvas.setFillColor(_color(ACCENT_GOLD))
    canvas.rect(0, PAGE_HEIGHT - 569.2, PAGE_WIDTH, 18.7, fill=1, stroke=0)
    canvas.setFillColor(_color(BANNER_GREEN))
    canvas.rect(0, PAGE_HEIGHT - 587.8, PAGE_WIDTH, 18.6, fill=1, stroke=0)
    canvas.restoreState()


def _cover_page(
    mission_name: str,
    transfer_title: str,
    stats: dict[str, int],
    grouped_zones: list[tuple[str, pd.DataFrame]],
    president: str,
    prepared_by: str,
):
    def draw(canvas, document) -> None:
        _draw_reference_cover(
            canvas, mission_name, transfer_title, stats, grouped_zones, president, prepared_by
        )

    return draw


def _body_page(mission_name: str, transfer_title: str, president: str):
    def draw(canvas, document) -> None:
        canvas.saveState()
        _draw_watermark(canvas, PAGE_HEIGHT - WATERMARK_ZONE_CENTER_Y_IN * inch)

        canvas.setFillColor(_color(BANNER_GREEN))
        canvas.rect(0.5 * inch, PAGE_HEIGHT - 0.43 * inch, PAGE_WIDTH - inch, 0.25 * inch, fill=1, stroke=0)
        canvas.setFont(FONT_BOLD, 7)
        canvas.setFillColor(colors.white)
        canvas.drawString(0.62 * inch, PAGE_HEIGHT - 0.34 * inch, mission_name.upper())
        canvas.setFillColor(_color(ACCENT_GOLD))
        canvas.drawCentredString(PAGE_WIDTH / 2, PAGE_HEIGHT - 0.34 * inch, transfer_title.upper())
        canvas.setFillColor(colors.white)
        canvas.drawRightString(PAGE_WIDTH - 0.62 * inch, PAGE_HEIGHT - 0.34 * inch, "ALL ZONES")

        canvas.setStrokeColor(_color(BANNER_GREEN))
        canvas.setLineWidth(0.8)
        canvas.line(0.5 * inch, 0.43 * inch, PAGE_WIDTH - 0.5 * inch, 0.43 * inch)
        canvas.setFont(FONT, 6)
        canvas.setFillColor(_color(MUTED_GRAY))
        canvas.drawString(0.5 * inch, 0.25 * inch, "STRICTLY CONFIDENTIAL - MISSION USE ONLY")
        canvas.setFillColor(_color(BANNER_GREEN))
        canvas.drawCentredString(PAGE_WIDTH / 2, 0.25 * inch, f"{president.title()} - {mission_name.title()}")
        canvas.setFillColor(_color(MUTED_GRAY))
        canvas.drawRightString(PAGE_WIDTH - 0.5 * inch, 0.25 * inch, f"Page {document.page}")
        canvas.restoreState()

    return draw


def _zone_story(styles: dict[str, ParagraphStyle], number: int, zone: str, rows: pd.DataFrame) -> list:
    zone_name = _zone_label(zone)
    divider = Table(
        [[_p(f"<font color='white'><b>{number:02d}</b></font>", styles["cell_center"], markup=True),
          _p(
              f"<b>{escape(zone_name)}</b>   <font color='#{MUTED_GRAY}'><i>-   {len(rows)} missionaries</i></font>",
              styles["zone"], markup=True,
          )]],
        colWidths=[0.5 * inch, 9.361 * inch],
    )
    divider.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), _color(BANNER_GREEN)),
        ("LINEBELOW", (1, 0), (1, 0), 1.5, _color(ACCENT_GOLD)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))

    headers = ["#", "MISSIONARY", "ROLE", "ZONE", "AREA", "COMPANION(S)"]
    data = [[_p(header, styles["cell_header"]) for header in headers]]
    for position, (_, row) in enumerate(rows.iterrows(), start=1):
        data.append([
            _p(position, styles["cell_center"]),
            _p(row["Name of Missionary"], styles["cell_bold"]),
            _p(row["Assignment"], styles["cell_center"]),
            _p(row["New/Existing Zone"], styles["cell_center"]),
            _p(row["New/Existing Area"], styles["cell"]),
            _p(row["New/Existing Companion(s)"], styles["cell"]),
        ])

    table = Table(
        data,
        colWidths=[0.306 * inch, 1.563 * inch, 0.764 * inch, 1.25 * inch, 2.431 * inch, 3.547 * inch],
        repeatRows=1,
        hAlign="LEFT",
        splitByRow=1,
    )
    commands = [
        ("BACKGROUND", (0, 0), (-1, 0), _color(BANNER_GREEN)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), FONT_BOLD),
        ("GRID", (0, 0), (-1, -1), 0.35, _color("E0E0E0")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (0, -1), 2),
        ("RIGHTPADDING", (0, 0), (0, -1), 2),
    ]
    for table_row, (_, row) in enumerate(rows.iterrows(), start=1):
        if table_row % 2 == 0:
            commands.append(("BACKGROUND", (0, table_row), (-1, table_row), _color(ZEBRA_TINT)))
        bg, fg = badge_colors_for(str(row["Assignment"]))
        commands.extend([
            ("BACKGROUND", (2, table_row), (2, table_row), _color(bg)),
            ("TEXTCOLOR", (2, table_row), (2, table_row), _color(fg)),
        ])
        if "NEW MISSIONARY" in str(row["New/Existing Companion(s)"]).upper():
            commands.extend([
                ("BACKGROUND", (5, table_row), (5, table_row), _color(NEW_MISSIONARY_FILL)),
                ("BOX", (5, table_row), (5, table_row), 1, _color(NEW_MISSIONARY_TEXT)),
                ("TEXTCOLOR", (5, table_row), (5, table_row), _color(NEW_MISSIONARY_TEXT)),
            ])
    table.setStyle(TableStyle(commands))
    return [KeepTogether([divider, Spacer(1, 0.08 * inch)]), table]


def _assignment_key(styles: dict[str, ParagraphStyle]) -> list:
    entries = list(ASSIGNMENT_KEY_ENTRIES) + [("+", "NEW MISSIONARY (INCOMING)")]
    cells = []
    for code, label in entries:
        bg, fg = badge_colors_for(code)
        if code == "+":
            bg, fg = NEW_MISSIONARY_FILL, NEW_MISSIONARY_TEXT
        cells.append((_p(f"<font color='#{fg}'><b>{escape(code)}</b></font>  {escape(label.title())}", styles["cell"], markup=True), bg))
    key_columns = 6
    while len(cells) % key_columns:
        cells.append(("", PANEL_GRAY))
    data = [[_p("ASSIGNMENT KEY", styles["cell_header"])] * key_columns]
    data += [
        [cell[0] for cell in cells[index:index + key_columns]]
        for index in range(0, len(cells), key_columns)
    ]
    table = Table(data, colWidths=[9.861 * inch / key_columns] * key_columns)
    commands = [
        ("SPAN", (0, 0), (-1, 0)),
        ("BACKGROUND", (0, 0), (-1, 0), _color(BANNER_GREEN)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.lightgrey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]
    for index, (_content, bg) in enumerate(cells):
        row, col = divmod(index, key_columns)
        commands.append(("BACKGROUND", (col, row + 1), (col, row + 1), _color(bg)))
    table.setStyle(TableStyle(commands))
    return [Spacer(1, 0.12 * inch), table]


def _group_rows(df: pd.DataFrame, meta: dict[str, str], roster_mode: bool) -> list[tuple[str, pd.DataFrame]]:
    if roster_mode:
        group_column = "New/Existing Zone"
        zone_order = [zone for zone in meta.get("Zone Order", "").split("|") if zone]
        rank = {zone.upper(): index for index, zone in enumerate(zone_order)}
        df = df.copy()
        df["_zone_rank"] = df[group_column].astype(str).str.upper().map(lambda zone: rank.get(zone, len(rank)))
        df = df.sort_values(["_zone_rank"], kind="stable").drop(columns=["_zone_rank"])
    else:
        group_column = "Previous Zone" if "Previous Zone" in df.columns else "New/Existing Zone"
    return [(str(zone), rows.copy()) for zone, rows in df.groupby(group_column, sort=False)]


def build_transfer_pdf(
    workbook_path: Path,
    output_path: Path,
    mission_name: str = DEFAULT_MISSION_NAME,
    transfer_title: str | None = None,
    president: str = DEFAULT_PRESIDENT,
    prepared_by: str = DEFAULT_PREPARED_BY,
    roster_mode: bool = False,
) -> Path:
    df = load_news_data(Path(workbook_path))
    try:
        transfer_df = pd.read_excel(workbook_path, sheet_name="Transfer Sheet")
    except Exception:
        transfer_df = None
    meta = read_meta_values(Path(workbook_path))
    title = transfer_title or meta.get("Transfer Title") or f"{datetime.now():%B %Y}".upper() + " TRANSFER NEWS"
    if roster_mode:
        title = f"{title} - UPDATED ROSTER"
    grouped = _group_rows(df, meta, roster_mode)
    stats = compute_stats(df, transfer_df, read_new_missionary_count(Path(workbook_path)))
    styles = _styles()

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document = SimpleDocTemplate(
        str(output_path), pagesize=landscape(letter),
        leftMargin=0.44 * inch, rightMargin=0.44 * inch,
        topMargin=0.55 * inch, bottomMargin=0.55 * inch,
        title=title, author=mission_name,
    )
    # The first page is drawn directly on the canvas so its measured bands,
    # tables, watermark, and publication panel remain stable on every host.
    story = [PageBreak()]
    for number, (zone, rows) in enumerate(grouped, start=1):
        if number > 1:
            story.append(PageBreak())
        story.extend(_zone_story(styles, number, zone, rows))
    story.append(PageBreak())
    story.extend(_assignment_key(styles))
    document.build(
        story,
        onFirstPage=_cover_page(mission_name, title, stats, grouped, president, prepared_by),
        onLaterPages=_body_page(mission_name, title, president),
    )
    return output_path


def build_statistics_pdf(workbook_path: Path, output_path: Path) -> Path:
    df = load_news_data(Path(workbook_path))
    try:
        transfer_df = pd.read_excel(workbook_path, sheet_name="Transfer Sheet")
    except Exception:
        transfer_df = None
    stats = compute_stats(df, transfer_df, read_new_missionary_count(Path(workbook_path)))
    styles = _styles()
    data = [[_p("ITEM", styles["cell_center"]), _p("COUNT", styles["cell_center"])]]
    data.extend([[_p(label, styles["cell"]), _p(value, styles["cell_center"])] for label, value in stats.items()])
    table = Table(data, colWidths=[5 * inch, 1.5 * inch])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), _color(BANNER_GREEN)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.lightgrey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document = SimpleDocTemplate(
        str(output_path), pagesize=landscape(letter),
        leftMargin=2 * inch, rightMargin=2 * inch,
        topMargin=0.7 * inch, bottomMargin=0.7 * inch,
        title="Transfer Statistics",
    )
    title = Paragraph(
        "TRANSFER STATISTICS",
        ParagraphStyle("StatsTitle", fontName=FONT_BOLD, fontSize=18, leading=22,
                       alignment=TA_CENTER, textColor=_color(BANNER_GREEN), spaceAfter=16),
    )
    document.build([title, table])
    return output_path
