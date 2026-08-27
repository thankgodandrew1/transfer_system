from __future__ import annotations

import argparse
import re
from copy import deepcopy
from datetime import datetime
from pathlib import Path

import pandas as pd
from docx import Document
from docx.table import Table
from docx.enum.section import WD_ORIENT, WD_SECTION_START
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


EXPECTED_COLUMNS = [
    "Name of Missionary",
    "Assignment",
    "New/Existing Zone",
    "New/Existing Area",
    "New/Existing Companion(s)",
]

# Palette and badge colors ported 1:1 from the mission's original design
# source (reference/JuneJuly2026_TransferNews_Generator.js).
ASSIGNMENT_BADGE_COLORS: dict[str, tuple[str, str]] = {
    "AP": ("F5EEF8", "6C3483"),
    "ZL": ("E4F2EB", "1A5C38"),
    "STL": ("EAF0FB", "1A4D8F"),
    "DL": ("E2F4EF", "0D6B55"),
    "DT": ("E2F4EF", "0D6B55"),
    "TR": ("FEF0E3", "B7560E"),
    "SA": ("FAE2D5", "C00000"),
    "JC": ("E8E8E8", "888888"),
    "SC": ("E8E8E8", "888888"),
}
DEFAULT_BADGE_COLORS = ("E8E8E8", "888888")

# The printed Transfer News is set in Aptos (the mission's Word documents);
# override with --font on the CLI if the design ever changes.
FONT_NAME = "Aptos"

# The running header/footer are set in Calibri, not FONT_NAME -- confirmed by
# measuring the mission's real sample (measure_reference_design.py): the
# header/footer text embeds as literal Calibri while every other run on the
# same page embeds as Segoe UI (the font substituted for Aptos on that
# machine). Calibri ships with Windows/Office, so requesting it directly
# avoids relying on whatever substitution a given machine happens to apply.
HEADER_FOOTER_FONT = "Calibri"

BANNER_GREEN = "1A5C38"
GREEN_TINT = "E4F2EB"
ACCENT_GOLD = "C9A227"
GOLD_TEXT = "7A6010"
GOLD_TINT = "FBF5DF"
BLUE = "1A4D8F"
BLUE_TINT = "EAF0FB"
ZEBRA_TINT = "F0F7F3"
NEW_MISSIONARY_FILL = "FDECEA"
NEW_MISSIONARY_TEXT = "A93226"
TEXT_DARK = "111111"
TEXT_MID = "3D3D3D"
MUTED_GRAY = "888888"
PANEL_GRAY = "F6F6F6"
CELL_BORDER = "E0E0E0"
NEW_MISSIONARY_MARK = "⊕"  # ⊕ as printed in the mission's Transfer News

DEFAULT_MISSION_NAME = "NIGERIA UYO MISSION"
DEFAULT_PRESIDENT = "PRESIDENT RICHARD PAAPA DADZIE"
DEFAULT_PREPARED_BY = "ELDER EDISON-MUDOR & ELDER ANDREW"
DEFAULT_PREPARED_ROLE = "ASSISTANT TO THE PRESIDENT"

# Mission seal watermark, printed faint on the cover and every zone/key page.
# assets/mission_watermark.jpg is not a recolored copy of the full-color logo
# (assets/mission_logo.png) -- it is the literal washout image extracted
# byte-for-byte from the mission's real Transfer News sample via
# measure_reference_design.py, so no runtime opacity/fade processing is
# needed; it is embedded at face value. Size and position (inches, page-
# relative) were measured the same way, not estimated: both the cover and
# the zone/key pages use an identical 6in square image, just centered a
# little higher on the cover (whose section has a shorter top margin).
MISSION_WATERMARK_PATH = Path(__file__).resolve().parent / "assets" / "mission_watermark.jpg"
WATERMARK_SIZE_IN = 6.0
WATERMARK_COVER_CENTER_X_IN = 5.5
WATERMARK_COVER_CENTER_Y_IN = 4.568
WATERMARK_ZONE_CENTER_X_IN = 5.5
WATERMARK_ZONE_CENTER_Y_IN = 4.272


def badge_colors_for(assignment: str) -> tuple[str, str]:
    code = str(assignment or "").strip().upper()
    for prefix, colors in ASSIGNMENT_BADGE_COLORS.items():
        if code.startswith(prefix):
            return colors
    return DEFAULT_BADGE_COLORS

ORDER_COLUMNS = [
    "Previous Zone Order",
    "Previous Row Order",
]

REQUIRED_COLUMNS = [
    "Name of Missionary",
    "Assignment",
    "New/Existing Zone",
    "New/Existing Area",
    "New/Existing Companion(s)",
]

def find_latest_transfer_sheet(output_dir: Path = Path("output")) -> Path:
    candidates = sorted(
        output_dir.glob("transfer_sheet_*.xlsx"),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    candidates = [path for path in candidates if not path.name.startswith("~$")]

    if not candidates:
        raise FileNotFoundError("No transfer_sheet_*.xlsx workbook found in output/.")

    return candidates[0]


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill)
    tc_pr.append(shading)


def set_cell_text(cell, text: str, bold: bool = False, size: float = 10, align=WD_ALIGN_PARAGRAPH.CENTER) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = align
    run = paragraph.add_run(str(text or "").upper())
    run.bold = bold
    run.font.name = FONT_NAME
    run.font.size = Pt(size)


def _set_char_spacing(run, twentieths: int) -> None:
    """Letter spacing (w:spacing, twentieths of a point)."""
    r_pr = run._element.get_or_add_rPr()
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:val"), str(twentieths))
    r_pr.append(spacing)


def style_cell_runs(cell, bold: bool | None = None, color: str | None = None) -> None:
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            if bold is not None:
                run.bold = bold
            if color is not None:
                run.font.color.rgb = RGBColor.from_string(color)


def set_cell_left_border(cell, color: str, size: int = 24) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), str(size))
    left.set(qn("w:space"), "0")
    left.set(qn("w:color"), color)
    borders.append(left)


def set_cell_border(cell, color: str, size: int = 8) -> None:
    """Border on all four edges of a single cell (a bordered badge/pill look,
    as close as a rectangular Word table cell can get to a rounded chip)."""
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(size))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)


def _shade_run(run, fill: str) -> None:
    """Character shading (colored chip behind a run)."""
    r_pr = run._element.get_or_add_rPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:fill"), fill)
    r_pr.append(shading)


def set_paragraph_shading(paragraph, fill: str) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill)
    p_pr.append(shading)


def set_paragraph_bottom_border(paragraph, color: str, size: int = 8) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), color)
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def set_table_borders(table, color: str = CELL_BORDER, size: int = 4) -> None:
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)

    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = f"w:{edge}"
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), str(size))
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)


def remove_element(element) -> None:
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)


def insert_body_element(document: Document, element) -> None:
    body = document._body._element
    sect_pr = body.sectPr
    if sect_pr is None:
        body.append(element)
    else:
        body.insert(body.index(sect_pr), element)


def clear_body(document: Document) -> None:
    body = document._body._element
    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)


def clear_headers_and_footers(document: Document) -> None:
    for section in document.sections:
        for part in (
            section.header,
            section.first_page_header,
            section.even_page_header,
            section.footer,
            section.first_page_footer,
            section.even_page_footer,
        ):
            for paragraph in part.paragraphs:
                set_paragraph_text(paragraph, "")


def replace_text_in_paragraph(paragraph, replacements: dict[str, str]) -> None:
    full_text = paragraph.text
    if any(placeholder in full_text for placeholder in replacements):
        for placeholder, value in replacements.items():
            full_text = full_text.replace(placeholder, str(value))

        for run in paragraph.runs:
            run.text = ""

        if paragraph.runs:
            paragraph.runs[0].text = full_text
        else:
            paragraph.add_run(full_text)
        return

    for run in paragraph.runs:
        for placeholder, value in replacements.items():
            if placeholder in run.text:
                run.text = run.text.replace(placeholder, str(value))


def set_paragraph_text(paragraph, text: str) -> None:
    for run in paragraph.runs:
        run.text = ""
    if paragraph.runs:
        paragraph.runs[0].text = text
    else:
        paragraph.add_run(text)


def replace_text_in_table(table: Table, replacements: dict[str, str]) -> None:
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                replace_text_in_paragraph(paragraph, replacements)


def replace_text_in_row(row, replacements: dict[str, str]) -> None:
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            replace_text_in_paragraph(paragraph, replacements)


def clean_text(value: str) -> str:
    return re.sub(r"\s+", " ", str(value)).strip()


def normalize_zone_name(value: str) -> str:
    value = clean_text(str(value)).upper()
    value = re.sub(r"\s+ZONE\b", "", value)
    return re.sub(r"[^A-Z0-9]", "", value)


def paragraph_text_from_element(element) -> str:
    return clean_text("".join(node.text for node in element.iter(qn("w:t")) if node.text))


def validate_zone_tables(document: Document) -> None:
    current_zone = ""
    table_index = 0

    for child in document.element.body.iterchildren():
        if child.tag == qn("w:p"):
            text = paragraph_text_from_element(child)
            if text:
                current_zone = text
            continue

        if child.tag != qn("w:tbl"):
            continue

        table = document.tables[table_index]
        table_index += 1

        if not current_zone:
            continue

        expected_zone = normalize_zone_name(current_zone)
        if not expected_zone:
            continue

        for row in list(table.rows)[1:]:
            cells = row.cells
            if len(cells) >= 6:
                row_zone = cells[3].text
            elif len(cells) >= 5:
                row_zone = cells[2].text
            else:
                continue

            actual_zone = normalize_zone_name(row_zone)
            if actual_zone and actual_zone != expected_zone:
                row_name = clean_text(cells[0].text) if cells else ""
                raise ValueError(
                    f"Zone table mismatch under {current_zone}: "
                    f"{row_name or 'row'} is assigned to {row_zone}."
                )


def find_template_parts(document: Document):
    zone_paragraph = next(
        (p for p in document.paragraphs if "{{ZONE_NAME}}" in p.text or "{{ZONE}}" in p.text),
        None,
    )
    template_table = next(
        (
            table
            for table in document.tables
            if any("{{MISSIONARY_NAME}}" in cell.text for row in table.rows for cell in row.cells)
        ),
        None,
    )
    return zone_paragraph, template_table


def set_cell_width(cell, width_inches: float) -> None:
    cell.width = Inches(width_inches)
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_w = tc_pr.first_child_found_in("w:tcW")
    if tc_w is None:
        tc_w = OxmlElement("w:tcW")
        tc_pr.append(tc_w)
    tc_w.set(qn("w:w"), str(int(width_inches * 1440)))
    tc_w.set(qn("w:type"), "dxa")


def add_zone_table(document: Document, zone_number: int, zone: str, rows: pd.DataFrame, first_zone: bool) -> None:
    if not first_zone:
        document.add_page_break()

    # Zone divider (zoneDivider in the design source): a 2-cell table — the
    # zone number in a solid green block, then the letter-spaced zone name
    # with a gold rule underneath.
    divider = document.add_table(rows=1, cols=2)
    number_cell, title_cell = divider.rows[0].cells

    _clear_cell_borders(number_cell)
    set_cell_shading(number_cell, BANNER_GREEN)
    number_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    number_para = number_cell.paragraphs[0]
    number_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    number_run = number_para.add_run(f"{zone_number:02d}")
    number_run.bold = True
    number_run.font.name = FONT_NAME
    number_run.font.size = Pt(14)
    number_run.font.color.rgb = RGBColor.from_string("FFFFFF")

    _clear_cell_borders(title_cell)
    _set_cell_edge(title_cell, "bottom", ACCENT_GOLD, size=10)
    title_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    title_para = title_cell.paragraphs[0]
    name_run = title_para.add_run(f"{zone.upper()} ZONE")
    name_run.bold = True
    name_run.font.name = FONT_NAME
    name_run.font.size = Pt(11)
    name_run.font.color.rgb = RGBColor.from_string(BANNER_GREEN)
    _set_char_spacing(name_run, 20)
    tail_run = title_para.add_run(f"   ·   {len(rows)} missionaries")
    tail_run.italic = True
    tail_run.font.name = FONT_NAME
    tail_run.font.size = Pt(7)
    tail_run.font.color.rgb = RGBColor.from_string(MUTED_GRAY)

    _set_fixed_table(divider, [700, 13500], margins=(95, 100))

    gap = document.add_paragraph()
    gap.paragraph_format.space_before = Pt(5)
    gap.paragraph_format.space_after = Pt(4)
    gap.add_run("").font.size = Pt(2)

    # Column widths from the design source (DXA): [440, 2250, 1100, 1800,
    # 3500, 5110]. "ASSIGNMENT" is wider than the design source's "ROLE" and
    # wraps badly at the original 1100 DXA, so it borrows width from
    # COMPANION(S), which has room to spare.
    headers = ["#", "MISSIONARY", "ASSIGNMENT", "ZONE", "NEW/EXISTING AREA", "COMPANION(S)"]
    widths_dxa = [440, 2250, 2350, 1800, 3500, 3860]
    left = WD_ALIGN_PARAGRAPH.LEFT
    aligns = [None, left, None, None, left, left]

    table = document.add_table(rows=1, cols=len(headers))
    set_table_borders(table)

    for index, header in enumerate(headers):
        cell = table.rows[0].cells[index]
        set_cell_text(cell, header, bold=True, size=14)
        set_cell_shading(cell, BANNER_GREEN)
        style_cell_runs(cell, color="FFFFFF")
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    _repeat_header_row(table.rows[0])

    # Measured directly off the reference design (measure_reference_design.py
    # against reference/transfer_news.pdf): every zone-table column — row
    # number included — is set at one uniform ~14pt, not the smaller mixed
    # sizes this used to have.
    sizes = [14, 14, 14, 14, 14, 14]
    for position, (_, row) in enumerate(rows.iterrows()):
        cells = table.add_row().cells
        companion = str(row["New/Existing Companion(s)"])
        is_new_missionary = "NEW MISSIONARY" in companion.upper()
        values = [
            str(position + 1),
            row["Name of Missionary"],
            row["Assignment"],
            row["New/Existing Zone"],
            row["New/Existing Area"],
            companion,
        ]

        for index, value in enumerate(values):
            set_cell_text(
                cells[index], value, size=sizes[index],
                align=aligns[index] or WD_ALIGN_PARAGRAPH.CENTER,
            )
            cells[index].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        if position % 2 == 1:
            for cell in cells:
                set_cell_shading(cell, ZEBRA_TINT)

        style_cell_runs(cells[0], color=MUTED_GRAY)
        style_cell_runs(cells[1], bold=True, color=TEXT_DARK)

        bg, fg = badge_colors_for(row["Assignment"])
        set_cell_shading(cells[2], bg)
        style_cell_runs(cells[2], bold=True, color=fg)
        style_cell_runs(cells[3], bold=True, color=BANNER_GREEN)
        style_cell_runs(cells[4], color=TEXT_MID)
        style_cell_runs(cells[5], color=TEXT_MID)

        if is_new_missionary:
            set_cell_shading(cells[5], NEW_MISSIONARY_FILL)
            set_cell_border(cells[5], NEW_MISSIONARY_TEXT, size=8)
            # Badge layout: mark inline before the text, both on one line,
            # inside a fully bordered cell (the closest a rectangular Word
            # table cell can get to a rounded pill/chip).
            cells[5].text = ""
            paragraph = cells[5].paragraphs[0]
            paragraph.alignment = aligns[5] or WD_ALIGN_PARAGRAPH.CENTER
            for text in (f"{NEW_MISSIONARY_MARK} ", "NEW MISSIONARY"):
                run = paragraph.add_run(text)
                run.bold = True
                run.font.name = FONT_NAME
                run.font.size = Pt(sizes[5])
                run.font.color.rgb = RGBColor.from_string(NEW_MISSIONARY_TEXT)

    _set_fixed_table(table, widths_dxa, margins=(80, 110))


def append_template_zone(
    document: Document,
    zone_number: int,
    zone: str,
    rows: pd.DataFrame,
    template_heading_xml,
    template_table_xml,
    first_zone: bool,
) -> None:
    if not first_zone:
        document.add_page_break()

    heading_xml = deepcopy(template_heading_xml)
    insert_body_element(document, heading_xml)
    heading = document.paragraphs[-1]
    heading_text = f"{zone_number:02d}  {zone.upper()} ZONE  ·  {len(rows)} MISSIONARIES"
    replace_text_in_paragraph(
        heading,
        {
            "{{ZONE_NAME}}": heading_text,
            "{{ZONE}}": heading_text,
        },
    )
    for run in heading.runs:
        run.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor.from_string(BANNER_GREEN)
    set_paragraph_bottom_border(heading, BANNER_GREEN)

    table_xml = deepcopy(template_table_xml)
    insert_body_element(document, table_xml)
    table = Table(table_xml, document)

    header_row = table.rows[0]
    for cell in header_row.cells:
        set_cell_shading(cell, BANNER_GREEN)
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.color.rgb = RGBColor.from_string("FFFFFF")
                run.bold = True

    placeholder_index = next(
        (
            index
            for index, row in enumerate(table.rows)
            if any("{{MISSIONARY_NAME}}" in cell.text for cell in row.cells)
        ),
        None,
    )
    if placeholder_index is None:
        return

    placeholder_row = table.rows[placeholder_index]
    placeholder_cells: dict[str, int] = {}
    for index, cell in enumerate(placeholder_row.cells):
        for placeholder in ("{{MISSIONARY_NAME}}", "{{ASSIGNMENT}}", "{{ZONE}}", "{{AREA}}", "{{COMPANION}}", "{{STATUS}}"):
            if placeholder in cell.text:
                placeholder_cells[placeholder] = index

    placeholder_row_xml = deepcopy(placeholder_row._tr)
    remove_element(placeholder_row._tr)

    name_idx = placeholder_cells.get("{{MISSIONARY_NAME}}")
    assignment_idx = placeholder_cells.get("{{ASSIGNMENT}}")
    companion_idx = placeholder_cells.get("{{COMPANION}}")

    for position, (_, row) in enumerate(rows.iterrows()):
        new_row_xml = deepcopy(placeholder_row_xml)
        table._tbl.append(new_row_xml)
        new_row = table.rows[-1]
        replacements = {
            "{{MISSIONARY_NAME}}": row["Name of Missionary"],
            "{{ASSIGNMENT}}": row["Assignment"],
            "{{ZONE}}": row["New/Existing Zone"],
            "{{AREA}}": row["New/Existing Area"],
            "{{COMPANION}}": row["New/Existing Companion(s)"],
            "{{STATUS}}": row.get("Status", ""),
        }
        replace_text_in_row(new_row, replacements)

        if position % 2 == 1:
            for cell in new_row.cells:
                set_cell_shading(cell, ZEBRA_TINT)

        if name_idx is not None:
            style_cell_runs(new_row.cells[name_idx], bold=True)

        if assignment_idx is not None:
            bg, fg = badge_colors_for(row["Assignment"])
            set_cell_shading(new_row.cells[assignment_idx], bg)
            style_cell_runs(new_row.cells[assignment_idx], bold=True, color=fg)

        if companion_idx is not None and "NEW MISSIONARY" in str(row["New/Existing Companion(s)"]).upper():
            companion_cell = new_row.cells[companion_idx]
            set_cell_shading(companion_cell, NEW_MISSIONARY_FILL)
            set_cell_border(companion_cell, NEW_MISSIONARY_TEXT, size=8)
            style_cell_runs(companion_cell, bold=True, color=NEW_MISSIONARY_TEXT)


def append_reference_zone(
    document: Document,
    zone: str,
    rows: pd.DataFrame,
    template_heading_xml,
    template_table_xml,
    first_zone: bool,
) -> None:
    if not first_zone:
        document.add_page_break()

    heading_xml = deepcopy(template_heading_xml)
    insert_body_element(document, heading_xml)
    heading = document.paragraphs[-1]
    set_paragraph_text(heading, f"{zone.upper()} ZONE")

    table_xml = deepcopy(template_table_xml)
    insert_body_element(document, table_xml)
    table = Table(table_xml, document)

    if len(table.rows) > 1:
        row_template_xml = deepcopy(table.rows[1]._tr)
        for row in list(table.rows)[1:]:
            remove_element(row._tr)
    else:
        row_template_xml = deepcopy(table.rows[0]._tr)

    for _, row in rows.iterrows():
        new_row_xml = deepcopy(row_template_xml)
        table._tbl.append(new_row_xml)
        new_row = table.rows[-1]
        values = [
            row["Name of Missionary"],
            row["Assignment"],
            row["New/Existing Zone"],
            row["New/Existing Area"],
            row["New/Existing Companion(s)"],
        ]
        for index, value in enumerate(values[: len(new_row.cells)]):
            set_cell_text(new_row.cells[index], value, size=9)


OPTIONAL_COLUMNS = ["Status", "Previous Zone", "Previous Area", "Previous Assignment"]


def load_news_data(workbook_path: Path) -> pd.DataFrame:
    df = pd.read_excel(workbook_path, sheet_name="News Format")
    missing = [column for column in REQUIRED_COLUMNS if column not in df.columns]

    if missing:
        raise ValueError(f"Workbook is missing required News Format columns: {', '.join(missing)}")

    # Rows already come pre-sorted by previous zone / role from generate_news_format.py.
    columns = EXPECTED_COLUMNS + [c for c in OPTIONAL_COLUMNS if c in df.columns]
    df = df[columns].fillna("")
    return df


def count_assignment_codes(df: pd.DataFrame, prefix: str | None = None, exact: str | None = None) -> int:
    count = 0
    for value in df.get("Assignment", []):
        codes = [code.strip().upper() for code in str(value).split("/") if code.strip()]
        if exact:
            count += sum(code == exact for code in codes)
        elif prefix:
            count += sum(code.startswith(prefix) for code in codes)
    return count


def count_leadership_roles(df: pd.DataFrame) -> int:
    """Leadership Count = APs + Zone Leaders + Sister Training Leaders +
    District Leaders / District Trainers (DL/DT). Trainers (TR) and special
    assignments are not counted."""
    count = 0
    for value in df.get("Assignment", []):
        codes = [code.strip().upper() for code in str(value).split("/") if code.strip()]
        count += sum(
            code.startswith(("AP", "ZL", "STL")) or code in ("DL", "DT")
            for code in codes
        )
    return count


def compute_stats(
    df: pd.DataFrame,
    transfer_df: pd.DataFrame | None = None,
    new_missionary_count: int | None = None,
) -> dict[str, int]:
    # The News Format table has exactly one row per missionary (new arrivals
    # are masked inside their trainer's companion cell), so it is the source
    # of truth for every count.
    names = df["Name of Missionary"].astype(str).str.upper()
    masked_rows = df["New/Existing Companion(s)"].astype(str).str.contains("NEW MISSIONARY")
    if new_missionary_count is not None:
        hidden_new_missionaries = int(new_missionary_count)
    else:
        hidden_new_missionaries = int(masked_rows.sum())
    total_missionaries = len(df) + hidden_new_missionaries
    hidden_elders = int((masked_rows & names.str.startswith("ELDER ")).sum())
    hidden_sisters = int((masked_rows & names.str.startswith("SISTER ")).sum())
    if transfer_df is not None and not transfer_df.empty:
        area_columns = [column for column in ("Zone", "District", "Area") if column in transfer_df.columns]
        companionships = int(transfer_df[area_columns].astype(str).drop_duplicates().shape[0]) if area_columns else 0
    else:
        companionships = int(
            df[["New/Existing Zone", "New/Existing Area"]]
            .astype(str)
            .drop_duplicates()
            .shape[0]
        )

    return {
        "Total Missionaries in Mission": total_missionaries,
        "Elders": int(names.str.startswith("ELDER ").sum()) + hidden_elders,
        "Sisters": int(names.str.startswith("SISTER ").sum()) + hidden_sisters,
        "New Missionaries": hidden_new_missionaries,
        "APs": count_assignment_codes(df, prefix="AP"),
        "ZLs": count_assignment_codes(df, prefix="ZL"),
        "STLs": count_assignment_codes(df, prefix="STL"),
        "DLs": count_assignment_codes(df, exact="DL"),
        "DTs": count_assignment_codes(df, exact="DT"),
        "Trainers": count_assignment_codes(df, exact="TR"),
        "Leadership Roles": count_leadership_roles(df),
        # One current zone/area group is one companionship, including real
        # trios and foursomes. Dividing the headcount by two overstates the
        # count whenever a group has more than two missionaries.
        "Companionships": companionships,
        "Zones": int(df["New/Existing Zone"].nunique()),
    }


def add_stats_section(
    document: Document,
    df: pd.DataFrame,
    transfer_df: pd.DataFrame | None = None,
    new_missionary_count: int | None = None,
) -> None:
    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("TRANSFER STATISTICS")
    run.bold = True
    run.font.name = FONT_NAME
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(0x00, 0x80, 0x00)

    stats = compute_stats(df, transfer_df, new_missionary_count)
    table = document.add_table(rows=1, cols=2)
    table.autofit = False
    set_table_borders(table)
    set_cell_text(table.rows[0].cells[0], "ITEM", bold=True)
    set_cell_text(table.rows[0].cells[1], "COUNT", bold=True)

    for label, count in stats.items():
        cells = table.add_row().cells
        set_cell_text(cells[0], label, size=10)
        set_cell_text(cells[1], str(count), size=10)


def build_statistics_docx(workbook_path: Path, output_path: Path) -> Path:
    df = load_news_data(workbook_path)
    try:
        transfer_df = pd.read_excel(workbook_path, sheet_name="Transfer Sheet")
    except Exception:
        transfer_df = None
    new_missionary_count = read_new_missionary_count(workbook_path)

    document = Document()
    set_section_margins(document.sections[0])
    add_stats_section(document, df, transfer_df, new_missionary_count)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        document.save(output_path)
        return output_path
    except PermissionError:
        fallback_path = output_path.with_name(f"{output_path.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{output_path.suffix}")
        document.save(fallback_path)
        return fallback_path


def set_section_margins(section, top: float = 0.7, bottom: float = 0.65) -> None:
    """Landscape layout matching the printed Transfer News."""
    section.orientation = WD_ORIENT.LANDSCAPE
    if section.page_height > section.page_width:
        section.page_width, section.page_height = section.page_height, section.page_width
    section.top_margin = Inches(top)
    section.bottom_margin = Inches(bottom)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    section.header_distance = Inches(0.3)
    section.footer_distance = Inches(0.3)


def _usable_width_inches(section) -> float:
    return (section.page_width - section.left_margin - section.right_margin) / Inches(1)


def _set_three_part_tabs(paragraph, section) -> None:
    width = _usable_width_inches(section)
    tab_stops = paragraph.paragraph_format.tab_stops
    tab_stops.add_tab_stop(Inches(width / 2), WD_TAB_ALIGNMENT.CENTER)
    tab_stops.add_tab_stop(Inches(width), WD_TAB_ALIGNMENT.RIGHT)


def _small_run(paragraph, text: str, color: str, bold: bool = False, italic: bool = False, size: float = 7.5, font: str | None = None):
    run = paragraph.add_run(text)
    run.font.name = font or FONT_NAME
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = RGBColor.from_string(color)
    return run


def _set_fixed_table(table, widths_dxa: list[int], margins: tuple[int, int] = (80, 110)) -> None:
    """Force a fixed-layout table with the exact column widths (DXA) from the
    design source. Without rebuilding w:tblGrid, Word ignores per-cell widths
    and renders equal columns."""
    tbl = table._tbl
    tbl_pr = tbl.tblPr

    table.autofit = False  # writes <w:tblLayout w:type="fixed"/>

    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(sum(widths_dxa)))
    tbl_w.set(qn("w:type"), "dxa")

    top_bottom, left_right = margins
    cell_mar = tbl_pr.find(qn("w:tblCellMar"))
    if cell_mar is None:
        cell_mar = OxmlElement("w:tblCellMar")
        tbl_pr.append(cell_mar)
    for tag, value in (("top", top_bottom), ("bottom", top_bottom), ("left", left_right), ("right", left_right)):
        element = cell_mar.find(qn(f"w:{tag}"))
        if element is None:
            element = OxmlElement(f"w:{tag}")
            cell_mar.append(element)
        element.set(qn("w:w"), str(value))
        element.set(qn("w:type"), "dxa")

    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        tbl.remove(grid)
    grid = OxmlElement("w:tblGrid")
    for width in widths_dxa:
        grid_col = OxmlElement("w:gridCol")
        grid_col.set(qn("w:w"), str(width))
        grid.append(grid_col)
    tbl.insert(list(tbl).index(tbl_pr) + 1, grid)

    for row in table.rows:
        seen: set[int] = set()
        column = 0
        for cell in row.cells:
            if id(cell._tc) in seen:
                column += 1
                continue
            seen.add(id(cell._tc))
            tc_pr = cell._tc.get_or_add_tcPr()
            span_el = tc_pr.find(qn("w:gridSpan"))
            span = int(span_el.get(qn("w:val"))) if span_el is not None else 1
            width = sum(widths_dxa[column:column + span]) if column < len(widths_dxa) else widths_dxa[-1]
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(width))
            tc_w.set(qn("w:type"), "dxa")
            column += span


def _repeat_header_row(row) -> None:
    """Repeat this row at the top of every page the table spans
    (tableHeader: true in the design source)."""
    tr_pr = row._tr.get_or_add_trPr()
    header = OxmlElement("w:tblHeader")
    header.set(qn("w:val"), "true")
    tr_pr.append(header)


def _gap(document: Document, points: float = 4) -> None:
    """Small fixed-height spacer paragraph."""
    para = document.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(points)
    para.add_run("").font.size = Pt(2)


def _set_cell_edge(cell, edge: str, color: str, size: int = 4, style: str = "single") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    element = borders.find(qn(f"w:{edge}"))
    if element is None:
        element = OxmlElement(f"w:{edge}")
        borders.append(element)
    element.set(qn("w:val"), style)
    element.set(qn("w:sz"), str(size))
    element.set(qn("w:space"), "0")
    element.set(qn("w:color"), color)


def _clear_cell_borders(cell) -> None:
    for edge in ("top", "bottom", "left", "right"):
        _set_cell_edge(cell, edge, "FFFFFF", 0, style="nil")


def _set_cell_top_border(cell, color: str, size: int = 20) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), str(size))
    top.set(qn("w:space"), "0")
    top.set(qn("w:color"), color)
    borders.append(top)


def _set_paragraph_top_border(paragraph, color: str, size: int = 8) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), str(size))
    top.set(qn("w:space"), "4")
    top.set(qn("w:color"), color)
    p_bdr.append(top)
    p_pr.append(p_bdr)


def _add_page_number(paragraph, color: str = "3D3D3D") -> None:
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), HEADER_FOOTER_FONT)
    fonts.set(qn("w:hAnsi"), HEADER_FOOTER_FONT)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "15")
    clr = OxmlElement("w:color")
    clr.set(qn("w:val"), color)
    rpr.append(fonts)
    rpr.append(sz)
    rpr.append(clr)
    run.append(rpr)
    text = OxmlElement("w:t")
    text.text = "1"
    run.append(text)
    fld.append(run)
    paragraph._p.append(fld)


def _anchor_picture(shape, x_emu: int, y_emu: int) -> None:
    """Convert an inline picture (as returned by Run.add_picture) into a
    floating watermark: fixed absolute page position, behind all text, no
    text wrap. Anchoring it to a header paragraph makes it repeat unchanged
    on every page of that section, exactly like Word's own watermark feature;
    text/tables painted after it (zebra rows, cell shading, etc.) still cover
    it wherever they're opaque, which is why it only shows through in the
    blank space around a page's content."""
    inline = shape._inline
    extent = inline.find(qn("wp:extent"))
    effect_extent = inline.find(qn("wp:effectExtent"))
    doc_pr = inline.find(qn("wp:docPr"))
    cnv_pr = inline.find(qn("wp:cNvGraphicFramePr"))
    graphic = inline.find(qn("a:graphic"))

    anchor = OxmlElement("wp:anchor")
    for attr, value in (
        ("distT", "0"), ("distB", "0"), ("distL", "0"), ("distR", "0"),
        ("simplePos", "0"), ("relativeHeight", "1"), ("behindDoc", "1"),
        ("locked", "0"), ("layoutInCell", "1"), ("allowOverlap", "1"),
    ):
        anchor.set(attr, value)

    simple_pos = OxmlElement("wp:simplePos")
    simple_pos.set("x", "0")
    simple_pos.set("y", "0")
    anchor.append(simple_pos)

    position_h = OxmlElement("wp:positionH")
    position_h.set("relativeFrom", "page")
    offset_h = OxmlElement("wp:posOffset")
    offset_h.text = str(int(x_emu))
    position_h.append(offset_h)
    anchor.append(position_h)

    position_v = OxmlElement("wp:positionV")
    position_v.set("relativeFrom", "page")
    offset_v = OxmlElement("wp:posOffset")
    offset_v.text = str(int(y_emu))
    position_v.append(offset_v)
    anchor.append(position_v)

    if extent is not None:
        anchor.append(extent)
    if effect_extent is None:
        effect_extent = OxmlElement("wp:effectExtent")
        for side in ("l", "t", "r", "b"):
            effect_extent.set(side, "0")
    anchor.append(effect_extent)

    anchor.append(OxmlElement("wp:wrapNone"))

    if doc_pr is not None:
        anchor.append(doc_pr)
    if cnv_pr is not None:
        anchor.append(cnv_pr)
    if graphic is not None:
        anchor.append(graphic)

    inline.getparent().replace(inline, anchor)


def add_watermark(paragraph, size_in: float, center_x_in: float, center_y_in: float) -> None:
    """Attach the mission-seal watermark to an existing paragraph. Since the
    picture is converted to a floating anchor, it does not add any visible
    line height to whatever paragraph carries it."""
    if not MISSION_WATERMARK_PATH.exists():
        return
    run = paragraph.add_run()
    size = Inches(size_in)
    shape = run.add_picture(str(MISSION_WATERMARK_PATH), width=size, height=size)
    x_emu = int(Inches(center_x_in - size_in / 2))
    y_emu = int(Inches(center_y_in - size_in / 2))
    _anchor_picture(shape, x_emu, y_emu)


def build_running_header_footer(
    document: Document,
    mission_name: str,
    transfer_title: str,
    president: str,
    section=None,
) -> None:
    """Per-page header and footer copied from the printed layout:
    header: mission name | title | 'All Zones', gold rule underneath;
    footer: confidential notice | president | page number."""
    if section is None:
        section = document.sections[-1]
    section.header.is_linked_to_previous = False
    section.footer.is_linked_to_previous = False

    header_para = section.header.paragraphs[0]
    set_paragraph_text(header_para, "")
    for run in list(header_para.runs):
        run.text = ""
    _set_three_part_tabs(header_para, section)
    set_paragraph_shading(header_para, BANNER_GREEN)
    _small_run(header_para, mission_name.upper(), "FFFFFF", bold=True, size=8.5, font=HEADER_FOOTER_FONT)
    header_para.add_run("\t")
    _small_run(header_para, transfer_title.replace(" TRANSFER NEWS", "  ·  TRANSFER NEWS"), ACCENT_GOLD, bold=True, size=8.5, font=HEADER_FOOTER_FONT)
    header_para.add_run("\t")
    _small_run(header_para, "All Zones", "FFFFFF", italic=True, size=6.5, font=HEADER_FOOTER_FONT)
    add_watermark(
        header_para, WATERMARK_SIZE_IN,
        WATERMARK_ZONE_CENTER_X_IN, WATERMARK_ZONE_CENTER_Y_IN,
    )
    set_paragraph_bottom_border(header_para, ACCENT_GOLD, size=10)

    footer_para = section.footer.paragraphs[0]
    set_paragraph_text(footer_para, "")
    for run in list(footer_para.runs):
        run.text = ""
    _set_three_part_tabs(footer_para, section)
    _set_paragraph_top_border(footer_para, BANNER_GREEN, size=8)
    _small_run(footer_para, "STRICTLY CONFIDENTIAL — Mission Use Only", MUTED_GRAY, italic=True, size=6, font=HEADER_FOOTER_FONT)
    footer_para.add_run("\t")
    _small_run(footer_para, f"{president.title()}  ·  {mission_name.title()}", BANNER_GREEN, bold=True, size=6, font=HEADER_FOOTER_FONT)
    footer_para.add_run("\t")
    _small_run(footer_para, "Page ", MUTED_GRAY, size=6, font=HEADER_FOOTER_FONT)
    _add_page_number(footer_para)


def read_meta_values(workbook_path: Path) -> dict[str, str]:
    try:
        meta = pd.read_excel(workbook_path, sheet_name="Meta")
        return {
            str(row["Metric"]): ("" if pd.isna(row["Value"]) else str(row["Value"]))
            for _, row in meta.iterrows()
        }
    except Exception:
        return {}


ASSIGNMENT_KEY_ENTRIES = [
    ("AP", "ASSISTANT TO THE PRESIDENT"),
    ("ZL", "ZONE LEADER"),
    ("STL", "SISTER TRAINING LEADER"),
    ("DL", "DISTRICT LEADER"),
    ("DT", "DISTRICT TRAINER"),
    ("TR", "TRAINER"),
    ("SA", "SPECIAL ASSIGNMENT"),
    ("JC", "JUNIOR COMPANION"),
    ("SC", "SENIOR COMPANION"),
]


def read_new_missionary_count(workbook_path: Path) -> int | None:
    try:
        meta = pd.read_excel(workbook_path, sheet_name="Meta")
        row = meta[meta["Metric"] == "New Missionaries (Incoming)"]
        if row.empty:
            return None
        return int(row.iloc[0]["Value"])
    except Exception:
        return None


def build_cover_page(
    document: Document,
    mission_name: str,
    transfer_title: str,
    stats: dict[str, int],
    zone_rows: list[tuple[int, str, int]],
    president: str = DEFAULT_PRESIDENT,
    prepared_by: str = DEFAULT_PREPARED_BY,
) -> None:
    banner = document.add_paragraph()
    banner.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = banner.add_run(mission_name.upper())
    run.bold = True
    run.font.name = FONT_NAME
    run.font.size = Pt(26)
    run.font.color.rgb = RGBColor.from_string(ACCENT_GOLD)
    _set_char_spacing(run, 60)
    set_paragraph_shading(banner, BANNER_GREEN)

    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run("THE CHURCH OF JESUS CHRIST OF LATTER-DAY SAINTS")
    run.font.name = FONT_NAME
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor.from_string("FFFFFF")
    _set_char_spacing(run, 20)
    set_paragraph_shading(subtitle, BANNER_GREEN)
    set_paragraph_bottom_border(subtitle, ACCENT_GOLD, size=18)

    _gap(document, 3)

    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run(transfer_title.upper())
    run.bold = True
    run.font.name = FONT_NAME
    run.font.size = Pt(29)
    run.font.color.rgb = RGBColor.from_string(BANNER_GREEN)
    _set_char_spacing(run, 20)

    subtitle2 = document.add_paragraph()
    subtitle2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle2.add_run("Complete Zone-by-Zone Transfer Assignments  ·  Mission Leadership Council")
    run.italic = True
    run.font.name = FONT_NAME
    run.font.size = Pt(8.5)
    run.font.color.rgb = RGBColor.from_string(MUTED_GRAY)
    add_watermark(
        subtitle2, WATERMARK_SIZE_IN,
        WATERMARK_COVER_CENTER_X_IN, WATERMARK_COVER_CENTER_Y_IN,
    )

    _gap(document, 3)

    # Stat cards (statCard in the design source): thick accent top border,
    # tinted background, big accent number, small caps label.
    tile_defs = [
        ("ZONES", stats.get("Zones", 0), GREEN_TINT, BANNER_GREEN),
        ("MISSIONARIES", stats.get("Total Missionaries in Mission", 0), BLUE_TINT, BLUE),
        ("LEADERSHIP ROLES", stats.get("Leadership Roles", 0), GOLD_TINT, GOLD_TEXT),
        ("NEW MISSIONARIES", stats.get("New Missionaries", 0), NEW_MISSIONARY_FILL, NEW_MISSIONARY_TEXT),
    ]
    tile_table = document.add_table(rows=2, cols=len(tile_defs))
    set_table_borders(tile_table, color="DDDDDD", size=4)
    for index, (label, value, bg, fg) in enumerate(tile_defs):
        number_cell = tile_table.rows[0].cells[index]
        set_cell_shading(number_cell, bg)
        set_cell_text(number_cell, str(value), bold=True, size=26)
        style_cell_runs(number_cell, color=fg)
        _set_cell_top_border(number_cell, fg, size=20)
        label_cell = tile_table.rows[1].cells[index]
        set_cell_shading(label_cell, bg)
        set_cell_text(label_cell, label, bold=True, size=6.5)
        style_cell_runs(label_cell, color=TEXT_MID)
    _set_fixed_table(tile_table, [3300] * 4, margins=(30, 100))

    _gap(document, 3)

    # ZONE INDEX table with the merged green header row (buildZoneIndex in
    # the design source).
    cols = 3
    index_widths = [4666, 4667, 4667]
    row_count = max((len(zone_rows) + cols - 1) // cols, 1)
    zone_table = document.add_table(rows=row_count + 1, cols=cols)
    set_table_borders(zone_table, color="DEDEDE", size=3)

    header_cell = zone_table.rows[0].cells[0]
    header_cell = header_cell.merge(zone_table.rows[0].cells[1]).merge(zone_table.rows[0].cells[2])
    set_cell_shading(header_cell, BANNER_GREEN)
    set_cell_text(header_cell, "ZONE INDEX", bold=True, size=7.5)
    style_cell_runs(header_cell, color="FFFFFF")

    for i in range(row_count * cols):
        r, c = divmod(i, cols)
        cell = zone_table.rows[r + 1].cells[c]
        cell.text = ""
        set_cell_shading(cell, PANEL_GRAY)
        if i < len(zone_rows):
            number, name, count = zone_rows[i]
            para = cell.paragraphs[0]
            _small_run(para, f"{number:02d}", GOLD_TEXT, bold=True, size=7)
            _small_run(para, f"  {name.upper()} ZONE", BANNER_GREEN, bold=True, size=7)
            _small_run(para, f"  ({count})", MUTED_GRAY, italic=True, size=6.5)
    _set_fixed_table(zone_table, index_widths, margins=(50, 130))

    _gap(document, 3)

    # PREPARED BY / APPROVED BY / EFFECTIVE panel
    effective_date = f"{datetime.now():%d %B, %Y}".upper()
    panel_defs = [
        ("PREPARED BY", prepared_by.upper(), DEFAULT_PREPARED_ROLE),
        ("APPROVED BY", president.upper(), f"{mission_name.upper()} PRESIDENT"),
        ("EFFECTIVE", transfer_title.upper().replace(" TRANSFER NEWS", " TRANSFER"), f"DOCUMENT DATE: {effective_date}"),
    ]
    panel = document.add_table(rows=1, cols=len(panel_defs))
    for index, (label, value, note) in enumerate(panel_defs):
        cell = panel.rows[0].cells[index]
        set_cell_shading(cell, PANEL_GRAY)
        cell.text = ""
        p1 = cell.paragraphs[0]
        _small_run(p1, label, MUTED_GRAY, bold=True, size=6)
        p2 = cell.add_paragraph()
        _small_run(p2, value, BANNER_GREEN, bold=True, size=8.5)
        p3 = cell.add_paragraph()
        _small_run(p3, note, MUTED_GRAY, italic=True, size=6.5)
    _set_fixed_table(panel, [4400, 4400, 4400], margins=(70, 160))

    confidential = document.add_paragraph()
    confidential.alignment = WD_ALIGN_PARAGRAPH.CENTER
    confidential.paragraph_format.space_before = Pt(2)
    confidential.paragraph_format.space_after = Pt(2)
    _small_run(confidential, "STRICTLY CONFIDENTIAL — FOR MISSION USE ONLY", MUTED_GRAY, italic=True, size=6.5)

    # Closing gold + green bars at the foot of the cover (design source):
    # a single line — gold fill with a thick green rule beneath it.
    bar = document.add_paragraph()
    set_paragraph_shading(bar, ACCENT_GOLD)
    set_paragraph_bottom_border(bar, BANNER_GREEN, size=28)
    bar.paragraph_format.space_before = Pt(0)
    bar.paragraph_format.space_after = Pt(0)
    bar_run = bar.add_run(" ")
    bar_run.font.size = Pt(3)


def build_assignment_key_page(document: Document) -> None:
    document.add_page_break()

    # 6-column key grid with a merged green title row, exactly like the
    # printed layout: AP ZL STL DL DT TR / SA JC SC ⊕.
    entries = [
        ("AP", "ASSISTANT TO THE PRESIDENT"),
        ("ZL", "ZONE LEADER"),
        ("STL", "SISTER TRAINING LEADER"),
        ("DL", "DISTRICT LEADER"),
        ("DT", "DISTRICT TRAINER"),
        ("TR", "TRAINER"),
        ("SA", "SPECIAL ASSIGNMENT"),
        ("JC", "JUNIOR COMPANION"),
        ("SC", "SENIOR COMPANION"),
        (NEW_MISSIONARY_MARK, "NEW MISSIONARY (INCOMING)"),
        ("", ""),
        ("", ""),
    ]
    cols = 6
    widths = [2366, 2367, 2367, 2367, 2367, 2366]
    row_count = (len(entries) + cols - 1) // cols
    table = document.add_table(rows=row_count + 1, cols=cols)
    set_table_borders(table, color="DEDEDE", size=3)

    header_cell = table.rows[0].cells[0]
    for extra in table.rows[0].cells[1:]:
        header_cell = header_cell.merge(extra)
    set_cell_shading(header_cell, BANNER_GREEN)
    set_cell_text(header_cell, "ASSIGNMENT KEY", bold=True, size=14)
    style_cell_runs(header_cell, color="FFFFFF")

    for i, (code, label) in enumerate(entries):
        r, c = divmod(i, cols)
        cell = table.rows[r + 1].cells[c]
        cell.text = ""
        if not code:
            set_cell_shading(cell, PANEL_GRAY)
            continue
        if code == NEW_MISSIONARY_MARK:
            bg, fg = NEW_MISSIONARY_FILL, NEW_MISSIONARY_TEXT
        elif code in ("JC", "SC"):
            bg, fg = PANEL_GRAY, MUTED_GRAY
        else:
            bg, fg = badge_colors_for(code)
        set_cell_shading(cell, bg)
        para = cell.paragraphs[0]
        if code == NEW_MISSIONARY_MARK:
            # Mark inline before the label, both bold in the badge color —
            # matches the reference design's key page and the zone-table cell.
            _small_run(para, f"{code} ", fg, bold=True, size=14)
            _small_run(para, label.title(), fg, bold=True, size=14)
        else:
            _small_run(para, code, fg, bold=True, size=14)
            _small_run(para, f"  {label.title()}", TEXT_MID, size=14)
    _set_fixed_table(table, widths, margins=(80, 120))


def build_transfer_docx(
    workbook_path: Path,
    output_path: Path,
    template_path: Path | None = None,
    mission_name: str = DEFAULT_MISSION_NAME,
    transfer_title: str | None = None,
    president: str = DEFAULT_PRESIDENT,
    prepared_by: str = DEFAULT_PREPARED_BY,
    roster_mode: bool = False,
) -> Path:
    """Build the Transfer News document.

    roster_mode=False: the printed Transfer News — zone sections grouped by
    each missionary's PREVIOUS zone, in the previous news' zone order.
    roster_mode=True: the 'updated previous transfer news' — the exact same
    missionaries grouped by their CURRENT zone, so the document can be fed
    straight back in as the previous-news input of the next transfer cycle.
    """
    df = load_news_data(workbook_path)
    try:
        transfer_df = pd.read_excel(workbook_path, sheet_name="Transfer Sheet")
    except Exception:
        transfer_df = None
    meta = read_meta_values(workbook_path)
    new_missionary_count = read_new_missionary_count(workbook_path)
    if transfer_title is None:
        transfer_title = meta.get("Transfer Title") or None

    use_template = template_path is not None and template_path.exists()
    document = Document(str(template_path)) if use_template else Document()
    template_heading_xml = None
    template_table_xml = None
    reference_heading_xml = None
    reference_table_xml = None

    if use_template:
        zone_paragraph, template_table = find_template_parts(document)
        if zone_paragraph is not None and template_table is not None:
            template_heading_xml = deepcopy(zone_paragraph._p)
            template_table_xml = deepcopy(template_table._tbl)
        elif document.tables:
            reference_heading = next((p for p in document.paragraphs if p.text.strip()), None)
            reference_heading_xml = deepcopy(reference_heading._p) if reference_heading is not None else None
            reference_table_xml = deepcopy(document.tables[0]._tbl)
        else:
            use_template = False
            document = Document()

    clear_headers_and_footers(document)
    clear_body(document)

    set_section_margins(document.sections[0])

    for style_name in ("Normal",):
        style = document.styles[style_name]
        style.font.name = FONT_NAME
        style.font.size = Pt(9)

    # Zone sections follow the mission's printed layout: missionaries are
    # grouped by the zone they served in LAST transfer (Previous Zone), while
    # each row's zone column shows where they serve now. In roster mode the
    # same rows are grouped by their CURRENT zone instead, in the canonical
    # zone order, so the output round-trips as next transfer's input.
    if roster_mode and "New/Existing Zone" in df.columns:
        group_column = "New/Existing Zone"
        zone_order = [z for z in meta.get("Zone Order", "").split("|") if z]
        rank = {z.upper(): i for i, z in enumerate(zone_order)}
        df = df.copy()
        df["_zone_rank"] = df[group_column].astype(str).str.upper().map(
            lambda z: rank.get(z, len(rank))
        )
        df = df.sort_values(["_zone_rank"], kind="stable").drop(columns=["_zone_rank"])
    else:
        group_column = "Previous Zone" if "Previous Zone" in df.columns else "New/Existing Zone"

    grouped_zones = []
    for zone, rows in df.groupby(group_column, sort=False):
        grouped_zones.append((str(zone), zone, rows))

    stats = compute_stats(df, transfer_df, new_missionary_count)
    zone_rows = [(number, str(zone), len(rows)) for number, (_, zone, rows) in enumerate(grouped_zones, start=1)]
    resolved_title = transfer_title or f"{datetime.now():%B %Y}".upper() + " TRANSFER NEWS"
    if roster_mode:
        resolved_title = f"{resolved_title} — UPDATED ROSTER"

    # Like the design source, the cover lives in its own section: tighter
    # margins and no running header/footer; the zone pages follow in a second
    # section that carries them.
    set_section_margins(document.sections[0], top=0.4, bottom=0.35)
    build_cover_page(document, mission_name, resolved_title, stats, zone_rows, president, prepared_by)
    body_section = document.add_section(WD_SECTION_START.NEW_PAGE)
    # The paragraph carrying the section break still occupies a line on the
    # cover; make it effectively zero-height so it cannot push a blank page.
    break_para = document.paragraphs[-1]
    break_para.paragraph_format.space_before = Pt(0)
    break_para.paragraph_format.space_after = Pt(0)
    break_run = break_para.add_run("")
    break_run.font.size = Pt(1)
    set_section_margins(body_section)
    build_running_header_footer(document, mission_name, resolved_title, president, section=body_section)

    first_zone = True
    for zone_number, (_, zone, rows) in enumerate(grouped_zones, start=1):
        if template_heading_xml is not None and template_table_xml is not None:
            append_template_zone(document, zone_number, str(zone), rows, template_heading_xml, template_table_xml, first_zone)
        elif reference_heading_xml is not None and reference_table_xml is not None:
            append_reference_zone(document, str(zone), rows, reference_heading_xml, reference_table_xml, first_zone)
        else:
            add_zone_table(document, zone_number, str(zone), rows, first_zone)
        first_zone = False

    build_assignment_key_page(document)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        document.save(output_path)
        return output_path
    except PermissionError:
        fallback_path = output_path.with_name(f"{output_path.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{output_path.suffix}")
        document.save(fallback_path)
        return fallback_path


def title_slug(title: str) -> str:
    """'JUNE / JULY 2026 TRANSFER NEWS' -> 'June_July_2026'."""
    cleaned = re.sub(r"\s*TRANSFER NEWS.*$", "", str(title or "").upper()).strip()
    cleaned = re.sub(r"[^A-Z0-9]+", "_", cleaned).strip("_")
    return cleaned.title() if cleaned else datetime.now().strftime("%B_%Y")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Create Word/PDF transfer news from a generated transfer workbook.")
    parser.add_argument("--workbook", help="Path to transfer_sheet_*.xlsx. Defaults to newest file in output/.")
    parser.add_argument("--docx", help="DOCX output path. Defaults to output/Transfer_News_<months>_<timestamp>.docx.")
    parser.add_argument("--template", default=None, help="Optional DOCX template with placeholders. Omit to use the built-in printed design.")
    parser.add_argument("--pdf", action="store_true", help="Also create native PDFs without Word or LibreOffice.")
    parser.add_argument("--no-roster", action="store_true", help="Skip the updated-previous-transfer roster document.")
    parser.add_argument("--stats-docx", help="Statistics DOCX output path. Defaults beside the transfer news document.")
    parser.add_argument("--mission-name", default=DEFAULT_MISSION_NAME, help="Mission name shown on the cover page banner.")
    parser.add_argument("--president", default=DEFAULT_PRESIDENT, help="Mission president shown on the cover and footer.")
    parser.add_argument("--prepared-by", default=DEFAULT_PREPARED_BY, help="'Prepared by' line on the cover page.")
    parser.add_argument("--transfer-title", help="Cover page title, e.g. 'JUNE / JULY 2026 TRANSFER NEWS'. Defaults to the months detected from the Transfer Management PDF (via the workbook Meta sheet).")
    parser.add_argument("--font", default=None, help="Typeface for the whole document (default: Aptos, as printed).")
    return parser.parse_args()


def main() -> None:
    global FONT_NAME
    args = parse_args()
    if args.font:
        FONT_NAME = args.font
    workbook_path = Path(args.workbook) if args.workbook else find_latest_transfer_sheet()
    meta = read_meta_values(workbook_path)
    title = args.transfer_title or meta.get("Transfer Title") or None
    slug = title_slug(title or "")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    docx_path = Path(args.docx) if args.docx else Path("output") / f"Transfer_News_{slug}_{timestamp}.docx"

    template_path = Path(args.template) if args.template else None
    docx_path = build_transfer_docx(
        workbook_path, docx_path, template_path, args.mission_name, title,
        president=args.president, prepared_by=args.prepared_by,
    )
    print(f"Word document created: {docx_path}")

    roster_path = None
    if not args.no_roster:
        roster_path = docx_path.with_name(f"Updated_Previous_Transfer_News_{slug}_{timestamp}.docx")
        roster_path = build_transfer_docx(
            workbook_path, roster_path, None, args.mission_name, title,
            president=args.president, prepared_by=args.prepared_by, roster_mode=True,
        )
        print(f"Updated previous-transfer roster created: {roster_path}")
        print("Use this file (or its PDF) as the 'previous transfer news' input next transfer.")

    stats_path = Path(args.stats_docx) if args.stats_docx else docx_path.with_name("Transfer Statistics.docx")
    stats_path = build_statistics_docx(workbook_path, stats_path)
    print(f"Statistics document created: {stats_path}")

    if args.pdf:
        from generate_transfer_pdf import build_statistics_pdf, build_transfer_pdf

        pdf_path = build_transfer_pdf(
            workbook_path, docx_path.with_suffix(".pdf"), args.mission_name, title,
            president=args.president, prepared_by=args.prepared_by,
        )
        print(f"PDF created: {pdf_path}")
        if roster_path:
            roster_pdf = build_transfer_pdf(
                workbook_path, roster_path.with_suffix(".pdf"), args.mission_name, title,
                president=args.president, prepared_by=args.prepared_by, roster_mode=True,
            )
            print(f"Updated roster PDF created: {roster_pdf}")
        stats_pdf = build_statistics_pdf(workbook_path, stats_path.with_suffix(".pdf"))
        print(f"Statistics PDF created: {stats_pdf}")


if __name__ == "__main__":
    main()
