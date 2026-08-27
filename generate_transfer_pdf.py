"""Native PDF rendering for the hosted Transfer News application.

This renderer intentionally avoids Microsoft Word, LibreOffice, and Docker.
It consumes the same verified workbook as the DOCX renderer and produces a
portable landscape PDF with the same cover, zone tables, role palette,
watermark, confidentiality footer, and assignment key.
"""
from __future__ import annotations

from datetime import datetime
from html import escape
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


def _font_name() -> str:
    """Use an available Microsoft-compatible font when one is bundled."""
    candidates = [
        Path("assets/Aptos.ttf"),
        Path("assets/Calibri.ttf"),
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
    ]
    for path in candidates:
        if not path.exists():
            continue
        try:
            pdfmetrics.registerFont(TTFont("TransferSans", str(path)))
            return "TransferSans"
        except Exception:
            continue
    return "Helvetica"


FONT = _font_name()
FONT_BOLD = "Helvetica-Bold" if FONT == "Helvetica" else FONT


def _p(text: object, style: ParagraphStyle, *, markup: bool = False) -> Paragraph:
    value = str(text or "")
    if not markup:
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
            fontSize=12, leading=14, textColor=_color(BANNER_GREEN), alignment=TA_LEFT,
        ),
        "cell": ParagraphStyle(
            "Cell", parent=styles["Normal"], fontName=FONT,
            fontSize=7.2, leading=8.4, textColor=_color(TEXT_MID), alignment=TA_LEFT,
            wordWrap="CJK",
        ),
        "cell_center": ParagraphStyle(
            "CellCenter", parent=styles["Normal"], fontName=FONT,
            fontSize=7.2, leading=8.4, textColor=_color(TEXT_MID), alignment=TA_CENTER,
            wordWrap="CJK",
        ),
        "cell_bold": ParagraphStyle(
            "CellBold", parent=styles["Normal"], fontName=FONT_BOLD,
            fontSize=7.2, leading=8.4, textColor=_color(TEXT_DARK), alignment=TA_LEFT,
            wordWrap="CJK",
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


def _draw_watermark(canvas, center_y: float = PAGE_HEIGHT / 2) -> None:
    if not MISSION_WATERMARK_PATH.exists():
        return
    size = 5.5 * inch
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


def _first_page(canvas, document) -> None:
    canvas.saveState()
    _draw_watermark(canvas, PAGE_HEIGHT / 2 - 3)
    canvas.restoreState()


def _body_page(mission_name: str, transfer_title: str, president: str):
    def draw(canvas, document) -> None:
        canvas.saveState()
        _draw_watermark(canvas, PAGE_HEIGHT / 2 - 4)

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


def _cover_story(
    styles: dict[str, ParagraphStyle],
    mission_name: str,
    transfer_title: str,
    stats: dict[str, int],
    grouped_zones: list[tuple[str, pd.DataFrame]],
    president: str,
    prepared_by: str,
) -> list:
    story: list = []
    mission = Table(
        [[_p(mission_name.upper(), styles["cover_mission"])],
         [_p("THE CHURCH OF JESUS CHRIST OF LATTER-DAY SAINTS", styles["cover_church"])]],
        colWidths=[10 * inch],
    )
    mission.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), _color(BANNER_GREEN)),
        ("LINEBELOW", (0, -1), (-1, -1), 3, _color(ACCENT_GOLD)),
        ("TOPPADDING", (0, 0), (-1, 0), 8),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 5),
        ("TOPPADDING", (0, 1), (-1, 1), 2),
        ("BOTTOMPADDING", (0, 1), (-1, 1), 4),
    ]))
    cover_title = escape(transfer_title.upper())
    if cover_title.endswith(" - UPDATED ROSTER"):
        cover_title = cover_title.removesuffix(" - UPDATED ROSTER") + "<br/><font size='15'>UPDATED ROSTER</font>"
    story.extend([mission, Spacer(1, 0.18 * inch), _p(cover_title, styles["title"], markup=True)])
    story.append(_p("Complete Zone-by-Zone Transfer Assignments - Mission Leadership Council", styles["subtitle"]))
    story.append(Spacer(1, 0.16 * inch))

    tiles = [
        ("ZONES", stats.get("Zones", 0), GREEN_TINT, BANNER_GREEN),
        ("MISSIONARIES", stats.get("Total Missionaries in Mission", 0), BLUE_TINT, BLUE),
        ("LEADERSHIP ROLES", stats.get("Leadership Roles", 0), GOLD_TINT, GOLD_TEXT),
        ("NEW MISSIONARIES", stats.get("New Missionaries", 0), NEW_MISSIONARY_FILL, NEW_MISSIONARY_TEXT),
    ]
    tile_data = [[
        Paragraph(f"<font size='22'><b>{value}</b></font><br/><font size='6'>{escape(label)}</font>",
                  ParagraphStyle(f"tile-{index}", fontName=FONT, alignment=TA_CENTER, textColor=_color(fg), leading=20))
        for index, (label, value, _bg, fg) in enumerate(tiles)
    ]]
    tile_table = Table(tile_data, colWidths=[2.5 * inch] * 4, rowHeights=[0.75 * inch])
    tile_commands = [("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("BOX", (0, 0), (-1, -1), 0.4, colors.lightgrey)]
    for index, (_label, _value, bg, fg) in enumerate(tiles):
        tile_commands.extend([
            ("BACKGROUND", (index, 0), (index, 0), _color(bg)),
            ("LINEABOVE", (index, 0), (index, 0), 3, _color(fg)),
        ])
    tile_table.setStyle(TableStyle(tile_commands))
    story.extend([tile_table, Spacer(1, 0.15 * inch)])

    zone_cells = []
    for number, (zone, rows) in enumerate(grouped_zones, start=1):
        zone_cells.append(_p(f"<b>{number:02d}</b>  {escape(zone.upper())} ZONE  ({len(rows)})", styles["small"], markup=True))
    while len(zone_cells) % 3:
        zone_cells.append("")
    zone_data = [[_p("ZONE INDEX", styles["small_center"])] * 3]
    zone_data += [zone_cells[index:index + 3] for index in range(0, len(zone_cells), 3)]
    zone_table = Table(zone_data, colWidths=[10 * inch / 3] * 3)
    zone_table.setStyle(TableStyle([
        ("SPAN", (0, 0), (-1, 0)),
        ("BACKGROUND", (0, 0), (-1, 0), _color(BANNER_GREEN)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, 1), (-1, -1), _color(PANEL_GRAY)),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.lightgrey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.extend([zone_table, Spacer(1, 0.15 * inch)])

    date_text = datetime.now().strftime("%d %B, %Y").upper()
    panel_data = [[
        _p(f"<font color='#{MUTED_GRAY}' size='6'>PREPARED BY</font><br/><b>{escape(prepared_by.upper())}</b>", styles["small"], markup=True),
        _p(f"<font color='#{MUTED_GRAY}' size='6'>APPROVED BY</font><br/><b>{escape(president.upper())}</b>", styles["small"], markup=True),
        _p(f"<font color='#{MUTED_GRAY}' size='6'>DOCUMENT DATE</font><br/><b>{date_text}</b>", styles["small"], markup=True),
    ]]
    panel = Table(panel_data, colWidths=[10 * inch / 3] * 3)
    panel.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), _color(PANEL_GRAY)),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.lightgrey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.extend([panel, Spacer(1, 0.08 * inch), _p("STRICTLY CONFIDENTIAL - FOR MISSION USE ONLY", styles["small_center"])])
    return story


def _zone_story(styles: dict[str, ParagraphStyle], number: int, zone: str, rows: pd.DataFrame) -> list:
    divider = Table(
        [[_p(f"<font color='white'><b>{number:02d}</b></font>", styles["cell_center"], markup=True),
          _p(f"<b>{escape(zone.upper())} ZONE</b>   -   {len(rows)} missionaries", styles["zone"], markup=True)]],
        colWidths=[0.5 * inch, 9.5 * inch],
    )
    divider.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), _color(BANNER_GREEN)),
        ("LINEBELOW", (1, 0), (1, 0), 1.5, _color(ACCENT_GOLD)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))

    headers = ["#", "MISSIONARY", "ASSIGNMENT", "ZONE", "NEW/EXISTING AREA", "COMPANION(S)"]
    data = [[_p(header, styles["cell_center"]) for header in headers]]
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
        colWidths=[0.35 * inch, 1.45 * inch, 0.82 * inch, 1.15 * inch, 2.2 * inch, 4.03 * inch],
        repeatRows=1,
        hAlign="LEFT",
    )
    commands = [
        ("BACKGROUND", (0, 0), (-1, 0), _color(BANNER_GREEN)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), FONT_BOLD),
        ("GRID", (0, 0), (-1, -1), 0.35, _color("E0E0E0")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
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
    while len(cells) % 5:
        cells.append(("", PANEL_GRAY))
    data = [[_p("ASSIGNMENT KEY", styles["cell_center"])] * 5]
    data += [[cell[0] for cell in cells[index:index + 5]] for index in range(0, len(cells), 5)]
    table = Table(data, colWidths=[2 * inch] * 5)
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
        row, col = divmod(index, 5)
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
        leftMargin=0.5 * inch, rightMargin=0.5 * inch,
        topMargin=0.55 * inch, bottomMargin=0.55 * inch,
        title=title, author=mission_name,
    )
    story = _cover_story(styles, mission_name, title, stats, grouped, president, prepared_by)
    story.append(PageBreak())
    for number, (zone, rows) in enumerate(grouped, start=1):
        if number > 1:
            story.append(PageBreak())
        story.extend(_zone_story(styles, number, zone, rows))
    story.append(PageBreak())
    story.extend(_assignment_key(styles))
    document.build(
        story,
        onFirstPage=_first_page,
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
