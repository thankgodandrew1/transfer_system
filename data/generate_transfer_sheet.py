"""
data/generate_transfer_sheet.py

Wraps the root Transfer Management PDF parser (generate_transfer_sheet.py) and
the News Format cross-referencer (generate_news_format.py) into a single
workbook with three sheets: News Format, Transfer Sheet, Summary.

This is the reusable verified-pipeline entry point used by the web app and CLI.
"""
from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from generate_transfer_sheet import (  # noqa: E402
    extract_pdf_lines,
    parse_area_assignments,
    split_page_into_blocks,
)
from generate_news_format import generate_news_format  # noqa: E402
from verify_news import Issue, VerificationResult, verify  # noqa: E402


def _verification_sheet(vresult: VerificationResult) -> pd.DataFrame:
    rows: list[dict[str, str]] = []
    for c in vresult.corrections:
        rows.append({"Type": "CORRECTED", "Missionary": c.row_name, "Field": c.field,
                     "Old Value": c.old_value, "New Value": c.new_value, "Detail": c.reason})
    for e in vresult.blocking_errors:
        rows.append({"Type": "BLOCKING", "Missionary": e.row_name, "Field": "", "Old Value": "", "New Value": "", "Detail": e.detail})
    for n in vresult.notes:
        rows.append({"Type": "NOTE", "Missionary": n.row_name, "Field": "", "Old Value": "", "New Value": "", "Detail": n.detail})
    return pd.DataFrame(rows, columns=["Type", "Missionary", "Field", "Old Value", "New Value", "Detail"])


def generate_transfer_sheet(
    pdf_path: Path,
    output_dir: Path,
    previous_news_pdf: Path | None = None,
    current_report_path: Path | None = None,
    old_report_path: Path | None = None,
    manual_corrections_path: Path | None = None,
    transfer_title_override: str | None = None,
    raise_on_blocking: bool = True,
) -> tuple[Path, Path, Path]:
    pdf_path = Path(pdf_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    extracted_dir = output_dir.parent / "extracted"
    logs_dir = output_dir.parent / "logs"
    extracted_dir.mkdir(exist_ok=True)
    logs_dir.mkdir(exist_ok=True)

    all_records: list[dict[str, str]] = []
    page_lines = extract_pdf_lines(pdf_path)

    for lines in page_lines:
        zone, blocks = split_page_into_blocks(lines)
        for missionary_pool, district, detail_lines in blocks:
            all_records.extend(
                parse_area_assignments(detail_lines, zone, district, missionary_pool)
            )

    if not all_records:
        raise RuntimeError("No transfer rows were found. Check that the PDF is a current Transfer Management export.")

    df = pd.DataFrame(all_records)
    df = df.drop_duplicates().sort_values(["Zone", "District", "Area"]).reset_index(drop=True)

    summary = (
        df.groupby("Zone", as_index=False)
        .agg(Districts=("District", "nunique"), Areas=("Area", "count"))
        .sort_values("Zone")
    )

    news_df = pd.DataFrame()
    changes_df = pd.DataFrame()
    news_status = "not generated (no previous news PDF supplied)"
    new_missionary_count = 0
    transfer_title = ""
    zone_order: list[str] = []
    if previous_news_pdf is not None:
        previous_news_pdf = Path(previous_news_pdf)
        if previous_news_pdf.exists():
            result = generate_news_format(pdf_path, previous_news_pdf)
            news_df = result.news
            changes_df = result.changes
            transfer_title = transfer_title_override or result.transfer_title
            zone_order = result.zone_order
            # A missionary not found in the previous Transfer News gets no row
            # of their own; they only surface as "NEW MISSIONARY" in whoever
            # they're companioned with. That count is the reliable source of
            # truth for how many new missionaries arrived this transfer.
            if not news_df.empty:
                new_missionary_count = int(
                    news_df["New/Existing Companion(s)"].eq("NEW MISSIONARY").sum()
                )
            news_status = f"{len(news_df)} rows"
        else:
            news_status = f"not generated (previous news PDF not found: {previous_news_pdf})"

    # Verify against the mission's own Transfer Management reports (more
    # reliable than the PDF-derived data above) and auto-correct whatever
    # they state unambiguously. Reports default to sitting next to the
    # current Transfer Management PDF.
    vresult: VerificationResult | None = None
    verification_df = pd.DataFrame(columns=["Type", "Missionary", "Field", "Old Value", "New Value", "Detail"])
    if not news_df.empty:
        current_report_path = current_report_path or pdf_path.parent / "current_transfer_report.xlsx"
        old_report_path = old_report_path or pdf_path.parent / "old_transfer_report.xlsx"
        vresult = verify(
            news_df,
            current_report_path,
            old_report_path,
            manual_corrections_path=manual_corrections_path,
        )
        news_df = vresult.news
        if vresult.new_missionary_count is not None and vresult.new_missionary_count != new_missionary_count:
            vresult.notes.insert(0, Issue(
                "", f"New Missionaries (Incoming): News Format's masked-row count was "
                f"{new_missionary_count}; using the ground-truth roster diff ({vresult.new_missionary_count}) instead.",
                "note",
            ))
            new_missionary_count = vresult.new_missionary_count
        verification_df = _verification_sheet(vresult)

        # Verification can remove a falsely named new missionary row and add
        # the corresponding NEW MISSIONARY placeholder. Rebuild this change
        # category from the verified table so Changes and Meta agree.
        if not changes_df.empty:
            changes_df = changes_df[changes_df["Change"] != "NEW MISSIONARY (MASKED)"].copy()
        masked_rows: list[dict[str, str]] = []
        for _, row in news_df.iterrows():
            mentions = str(row.get("New/Existing Companion(s)", "")).upper().count("NEW MISSIONARY")
            for _ in range(mentions):
                masked_rows.append({
                    "Change": "NEW MISSIONARY (MASKED)",
                    "Missionary": "",
                    "Detail": f"{row.get('New/Existing Zone', '')} / {row.get('New/Existing Area', '')}",
                })
        if masked_rows:
            changes_df = pd.concat([changes_df, pd.DataFrame(masked_rows)], ignore_index=True)

        # The log and GUI should report the verified row count, not the
        # parser's provisional count from before ground-truth corrections.
        news_status = f"{len(news_df)} verified rows"

    meta = pd.DataFrame(
        [
            {"Metric": "New Missionaries (Incoming)", "Value": new_missionary_count},
            {"Metric": "Transfer Title", "Value": transfer_title},
            {"Metric": "Zone Order", "Value": "|".join(zone_order)},
        ]
    )

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_path = output_dir / f"transfer_sheet_{timestamp}.xlsx"
    csv_path = extracted_dir / f"transfer_sheet_{timestamp}.csv"
    log_path = logs_dir / f"transfer_sheet_{timestamp}.txt"

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        if not news_df.empty:
            news_export = news_df[[column for column in news_df.columns if not column.startswith("_")]]
            news_export.to_excel(writer, sheet_name="News Format", index=False)
        df.to_excel(writer, sheet_name="Transfer Sheet", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)
        if not changes_df.empty:
            changes_df.to_excel(writer, sheet_name="Changes", index=False)
        meta.to_excel(writer, sheet_name="Meta", index=False)
        if not verification_df.empty:
            verification_df.to_excel(writer, sheet_name="Verification", index=False)

        header_fill = PatternFill("solid", fgColor="1A5C38")
        zebra_fill = PatternFill("solid", fgColor="F0F7F3")
        warning_fill = PatternFill("solid", fgColor="FFF4E5")
        blocking_fill = PatternFill("solid", fgColor="FDECEA")
        corrected_fill = PatternFill("solid", fgColor="EAF0FB")
        bottom_border = Border(bottom=Side(style="thin", color="D8E2DC"))
        for worksheet in writer.sheets.values():
            worksheet.freeze_panes = "A2"
            worksheet.sheet_view.showGridLines = False
            worksheet.row_dimensions[1].height = 24
            if worksheet.max_row > 1 and worksheet.max_column > 0:
                worksheet.auto_filter.ref = worksheet.dimensions
            for cell in worksheet[1]:
                cell.fill = header_fill
                cell.font = Font(name="Aptos", size=10, bold=True, color="FFFFFF")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            for row_number in range(2, worksheet.max_row + 1):
                worksheet.row_dimensions[row_number].height = 20
                for cell in worksheet[row_number]:
                    cell.font = Font(name="Aptos", size=9, color="222222")
                    cell.alignment = Alignment(vertical="center", wrap_text=False)
                    cell.border = bottom_border
                    if row_number % 2 == 0:
                        cell.fill = zebra_fill

            if worksheet.title == "Verification" and worksheet.max_row > 1:
                type_column = next((cell.column for cell in worksheet[1] if cell.value == "Type"), None)
                if type_column:
                    for row_number in range(2, worksheet.max_row + 1):
                        issue_type = str(worksheet.cell(row_number, type_column).value or "").upper()
                        fill = blocking_fill if issue_type == "BLOCKING" else corrected_fill if issue_type == "CORRECTED" else warning_fill
                        for cell in worksheet[row_number]:
                            cell.fill = fill

            for column_cells in worksheet.columns:
                max_length = max(len(str(cell.value or "")) for cell in column_cells)
                worksheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_length + 2, 12), 60)

    df.to_csv(csv_path, index=False)

    log_path.write_text(
        "\n".join(
            [
                "TRANSFER SHEET GENERATION LOG",
                "=============================",
                f"Source PDF: {pdf_path}",
                f"Previous News PDF: {previous_news_pdf if previous_news_pdf else '(none supplied)'}",
                f"Transfer Title: {transfer_title or '(not detected)'}",
                f"News Format: {news_status}",
                f"Changes Detected: {len(changes_df)}",
                f"New Missionaries (Incoming): {new_missionary_count}",
                f"Verification: {len(vresult.corrections)} corrected, {len(vresult.blocking_errors)} blocking, "
                f"{len(vresult.notes)} notes" if vresult is not None else "Verification: not run (no reports supplied)",
                f"Rows: {len(df)}",
                f"Zones: {df['Zone'].nunique()}",
                f"Districts: {df['District'].nunique()}",
                f"Areas: {df['Area'].nunique()}",
                f"Excel: {excel_path}",
                f"CSV: {csv_path}",
            ]
        ),
        encoding="utf-8",
    )

    if raise_on_blocking and vresult is not None and vresult.blocking_errors:
        noun = "discrepancy" if len(vresult.blocking_errors) == 1 else "discrepancies"
        details = "; ".join(f"{e.row_name}: {e.detail}" for e in vresult.blocking_errors)
        raise RuntimeError(
            f"Verification found {len(vresult.blocking_errors)} unresolved {noun} the ground-truth "
            f"reports couldn't settle — see the Verification sheet in {excel_path}. {details}"
        )

    return excel_path, csv_path, log_path
