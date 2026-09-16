from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

import pandas as pd

from generate_news_format import parse_tm_zones


def extract_transfer_records(pdf_path: Path) -> list[dict[str, str]]:
    """One row per staffed area card in the Transfer Management PDF.

    Uses the same layout-aware parser as the News Format, so both sheets
    always describe the same set of areas. Cards with nobody assigned are
    closing areas and are left out."""
    records: list[dict[str, str]] = []
    for area in parse_tm_zones(Path(pdf_path)):
        if not area.members:
            continue
        records.append(
            {
                "Zone": area.zone,
                "District": area.district,
                "Area": area.area,
                "Assignment": " / ".join(m.badge for m in area.members if m.badge),
                "Missionary Pool": " & ".join(m.name for m in area.members),
            }
        )
    return records


def generate_transfer_sheet(pdf_path: Path, output_dir: Path) -> tuple[Path, Path, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    extracted_dir = output_dir.parent / "extracted"
    logs_dir = output_dir.parent / "logs"
    extracted_dir.mkdir(exist_ok=True)
    logs_dir.mkdir(exist_ok=True)

    all_records = extract_transfer_records(pdf_path)

    if not all_records:
        raise RuntimeError("No transfer rows were found. Check that the PDF is a current Transfer Management export.")

    df = pd.DataFrame(all_records)
    df = df.drop_duplicates().sort_values(["Zone", "District", "Area"]).reset_index(drop=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_path = output_dir / f"transfer_sheet_{timestamp}.xlsx"
    csv_path = extracted_dir / f"transfer_sheet_{timestamp}.csv"
    log_path = logs_dir / f"transfer_sheet_{timestamp}.txt"

    summary = (
        df.groupby("Zone", as_index=False)
        .agg(Districts=("District", "nunique"), Areas=("Area", "count"))
        .sort_values("Zone")
    )

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Transfer Sheet", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)

        workbook = writer.book
        for worksheet in writer.sheets.values():
            worksheet.freeze_panes = "A2"
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

    return excel_path, csv_path, log_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate a transfer sheet from a Transfer Management PDF.")
    parser.add_argument(
        "--pdf",
        default="data/current_transfer_management.pdf",
        help="Path to the current Transfer Management PDF.",
    )
    parser.add_argument(
        "--output-dir",
        default="output",
        help="Folder where the Excel workbook should be saved.",
    )
    parser.add_argument(
        "--previous-news",
        default=None,
        help="Path to the previous Transfer News PDF. If given, also builds the News Format sheet.",
    )
    parser.add_argument("--current-report", help="Path to the authoritative current roster XLSX.")
    parser.add_argument("--old-report", help="Path to the authoritative previous roster XLSX.")
    parser.add_argument("--manual-corrections", help="Optional CSV with Name, Field, Value columns.")
    parser.add_argument("--transfer-title", help="Optional transfer-title override.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    pdf_path = Path(args.pdf)
    output_dir = Path(args.output_dir)

    if args.previous_news:
        from data.generate_transfer_sheet import generate_transfer_sheet as generate_full_transfer_sheet

        excel_path, csv_path, log_path = generate_full_transfer_sheet(
            pdf_path,
            output_dir,
            Path(args.previous_news),
            current_report_path=Path(args.current_report) if args.current_report else None,
            old_report_path=Path(args.old_report) if args.old_report else None,
            manual_corrections_path=Path(args.manual_corrections) if args.manual_corrections else None,
            transfer_title_override=args.transfer_title,
        )
    else:
        excel_path, csv_path, log_path = generate_transfer_sheet(pdf_path, output_dir)

    print("Transfer sheet generated successfully.")
    print(f"Excel: {excel_path}")
    print(f"CSV: {csv_path}")
    print(f"Log: {log_path}")


if __name__ == "__main__":
    main()
