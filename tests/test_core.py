from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd
from pypdf import PdfReader

import generate_news_format
from generate_transfer_document import build_transfer_docx, compute_stats
from generate_transfer_pdf import build_statistics_pdf, build_transfer_pdf
from verify_news import verify


def test_month_only_title_uses_source_file_year(monkeypatch, tmp_path: Path) -> None:
    source = tmp_path / "management.pdf"
    source.write_bytes(b"%PDF-placeholder")
    stamp = datetime(2026, 8, 15, 12, 0).timestamp()
    os.utime(source, (stamp, stamp))

    class Page:
        @staticmethod
        def extract_text() -> str:
            return "Nigeria Uyo Mission\nAugust Transfer"

    class FakePdf:
        pages = [Page()]

        def __enter__(self):
            return self

        def __exit__(self, *_args):
            return False

    monkeypatch.setattr(generate_news_format.pdfplumber, "open", lambda _path: FakePdf())
    assert generate_news_format.parse_transfer_title(source) == "AUGUST 2026 TRANSFER NEWS"


def test_unmatched_news_row_is_blocking(tmp_path: Path) -> None:
    news = pd.DataFrame(
        [["ELDER MISSING", "JC", "UYO", "CENTRAL", "ELDER KNOWN", "SAME", "UYO", "CENTRAL", "JC"]],
        columns=[
            "Name of Missionary", "Assignment", "New/Existing Zone", "New/Existing Area",
            "New/Existing Companion(s)", "Status", "Previous Zone", "Previous Area", "Previous Assignment",
        ],
    )
    report = tmp_path / "current.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["Zone", "District", "Area", "Missionary Name", "Type", "Position", "Companion"])
    sheet.append(["Uyo Zone", "Central", "Central", "Known, Elder", "Elder", "JC", "Missing, Elder (JC)"])
    workbook.save(report)

    result = verify(news, report)
    assert len(result.blocking_errors) == 1
    assert "cannot safely verify" in result.blocking_errors[0].detail


def test_stats_count_a_trio_as_one_companionship(sample_news: pd.DataFrame) -> None:
    transfer = pd.DataFrame(
        [{"Zone": "UYO", "District": "CENTRAL", "Area": "CENTRAL"}]
    )
    stats = compute_stats(sample_news, transfer, new_missionary_count=1)
    assert stats["Total Missionaries in Mission"] == 4
    assert stats["Companionships"] == 1
    assert stats["Zones"] == 1


def test_native_pdf_and_docx_outputs_are_structurally_valid(sample_workbook: Path, tmp_path: Path) -> None:
    news_pdf = build_transfer_pdf(sample_workbook, tmp_path / "news.pdf")
    stats_pdf = build_statistics_pdf(sample_workbook, tmp_path / "stats.pdf")
    news_docx = build_transfer_docx(sample_workbook, tmp_path / "news.docx", None)

    reader = PdfReader(news_pdf)
    assert len(reader.pages) >= 3
    pdf_text = "\n".join(page.extract_text() or "" for page in reader.pages)
    assert "AUGUST 2026 TRANSFER NEWS" in pdf_text
    assert "ELDER ALPHA" in pdf_text
    assert len(PdfReader(stats_pdf).pages) == 1

    document = openpyxl.load_workbook(sample_workbook, read_only=True)
    assert {"News Format", "Transfer Sheet", "Meta"}.issubset(document.sheetnames)
    document.close()
    assert news_docx.exists() and news_docx.stat().st_size > 10_000
