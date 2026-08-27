from __future__ import annotations

import argparse
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
from pypdf import PdfReader # type: ignore


ASSIGNMENT_CODES = {
    "AP",
    "AP1",
    "AP2",
    "DL",
    "DT",
    "JC",
    "SA",
    "SC",
    "STL1",
    "STL2",
    "TR",
    "ZL1",
    "ZL2",
}

CODE_PATTERN = re.compile(
    r"(STL1|STL2|ZL1|ZL2|AP1|AP2|DL|DT|JC|SA|SC|TR|AP)"
)


def clean_text(value: str) -> str:
    value = value.replace("&amp;", "&")
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def is_footer(line: str) -> bool:
    return (
        "Transfer Management" in line
        or "imos.churchofjesuschrist.org" in line
        or bool(re.match(r"^\d{1,2}/\d{1,2}/\d{2},", line))
    )


def has_assignment_code(line: str) -> bool:
    return bool(CODE_PATTERN.search(line))


def looks_like_name_pool(line: str) -> bool:
    if has_assignment_code(line):
        return False
    if any(char.isdigit() for char in line):
        return False
    if "/" in line:
        return False
    return len(line) >= 18


def split_codes_and_text(line: str) -> list[tuple[str, str]]:
    """Split a mixed PDF line into ordered TEXT and CODE tokens."""
    if clean_text(line) == "T":
        return []

    tokens: list[tuple[str, str]] = []
    index = 0

    for match in CODE_PATTERN.finditer(line):
        start, end = match.span()
        before = clean_text(line[index:start])
        if before:
            tokens.append(("TEXT", before))

        code = match.group(1)
        tokens.append(("CODE", code))

        index = end

    after = clean_text(line[index:])
    if after:
        tokens.append(("TEXT", after))

    if not tokens and line:
        tokens.append(("TEXT", clean_text(line)))

    return tokens


def finalize_record(
    records: list[dict[str, str]],
    zone: str,
    district: str,
    missionary_pool: str,
    area: str,
    codes: list[str],
) -> None:
    area = clean_text(area)
    codes = [code for code in codes if code in ASSIGNMENT_CODES]

    if not area or not codes:
        return

    records.append(
        {
            "Zone": zone,
            "District": district,
            "Area": area,
            "Assignment": " / ".join(codes),
            "Missionary Pool": missionary_pool,
        }
    )


def parse_area_assignments(
    lines: list[str],
    zone: str,
    district: str,
    missionary_pool: str,
) -> list[dict[str, str]]:
    records: list[dict[str, str]] = []
    current_area = ""
    current_codes: list[str] = []

    for line in lines:
        for token_type, value in split_codes_and_text(line):
            if token_type == "TEXT":
                if current_area and current_codes:
                    finalize_record(
                        records,
                        zone,
                        district,
                        missionary_pool,
                        current_area,
                        current_codes,
                    )
                    current_area = value
                    current_codes = []
                elif current_area:
                    current_area = clean_text(f"{current_area} {value}")
                else:
                    current_area = value
            else:
                current_codes.append(value)

    finalize_record(records, zone, district, missionary_pool, current_area, current_codes)
    return records


def split_page_into_blocks(lines: list[str]) -> tuple[str, list[tuple[str, str, list[str]]]]:
    zone = lines[0] if lines else "UNKNOWN ZONE"
    blocks: list[tuple[str, str, list[str]]] = []

    if len(lines) < 3:
        return zone, blocks

    index = 1
    while index + 1 < len(lines):
        missionary_pool = lines[index]
        district = lines[index + 1]
        index += 2

        detail_lines: list[str] = []
        while index < len(lines):
            line = lines[index]
            next_line = lines[index + 1] if index + 1 < len(lines) else ""

            if detail_lines and looks_like_name_pool(line) and next_line and not has_assignment_code(next_line):
                break

            detail_lines.append(line)
            index += 1

        blocks.append((missionary_pool, district, detail_lines))

    return zone, blocks


def extract_pdf_lines(pdf_path: Path) -> list[list[str]]:
    reader = PdfReader(str(pdf_path))
    pages: list[list[str]] = []

    for page in reader.pages:
        text = page.extract_text() or ""
        lines = [
            clean_text(line)
            for line in text.splitlines()
            if clean_text(line) and not is_footer(clean_text(line))
        ]
        pages.append(lines)

    return pages


def generate_transfer_sheet(pdf_path: Path, output_dir: Path) -> tuple[Path, Path, Path]:
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
