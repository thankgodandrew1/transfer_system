from pathlib import Path

import openpyxl
import pandas as pd
import pytest

from app import _parse_current_role_overrides
from generate_transfer_document import assignment_key_entries, badge_colors_for, build_transfer_docx
from generate_transfer_pdf import build_transfer_pdf
from roles import normalize_role
from verify_news import _apply_manual_corrections, _normalize_position, load_report


def test_name_keys_support_accents_and_non_latin_names():
    from text_utils import comparison_key
    assert comparison_key("Élder D’Ángelo") == comparison_key("Elder Dangelo")
    assert comparison_key("王") != comparison_key("李")
    assert _parse_current_role_overrides("王 | TR") == [("王", "TR")]


def test_no_corrections_are_inherited_from_the_working_directory(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    (tmp_path / "data").mkdir()
    (tmp_path / "data" / "manual_corrections.csv").write_text("Name,Field,Value\nElder Alpha,Assignment,SA\n")
    news = pd.DataFrame({"Name of Missionary": ["ELDER ALPHA"], "Assignment": ["SC"]})
    _apply_manual_corrections(news, None, [])
    assert news.iloc[0]["Assignment"] == "SC"


def test_missing_verification_never_approves_publication(sample_workbook):
    from app import _workbook_summary
    with pytest.raises(ValueError, match="verification results could not be read"):
        _workbook_summary(sample_workbook)


@pytest.mark.parametrize("role", ["STL/TR", "AP2/TR", "WELFARE", "OFFICE SUPPORT", "MEDICAL+SA", "ÉQUIPE"])
def test_roles_are_open_ended(role):
    assert _parse_current_role_overrides(f"Sister Gabriel | {role}") == [("Sister Gabriel", role)]
    assert badge_colors_for(role) == badge_colors_for(role)
    assert all(len(color) == 6 for color in badge_colors_for(role))


@pytest.mark.parametrize("role", ["", "STL//TR", "STL/", "A" * 41, "<script>", "=1+1", "TR|red"])
def test_roles_reject_malformed_labels(role):
    with pytest.raises(ValueError):
        normalize_role(role)


def test_colors_do_not_confuse_new_roles_with_prefixes():
    from generate_transfer_document import count_assignment_codes, count_leadership_roles
    assert badge_colors_for("AP1") == badge_colors_for("AP")
    assert badge_colors_for("STL/TR") != badge_colors_for("STL")
    assert badge_colors_for("APARTMENT") != badge_colors_for("AP")
    assert ("STL/TR", "SISTER TRAINING LEADER / TRAINER") in assignment_key_entries(["STL/TR"])
    frame = pd.DataFrame({"Assignment": ["APARTMENT", "STL/TR", "AP2"]})
    assert count_assignment_codes(frame, prefix="AP") == 1
    assert count_leadership_roles(frame) == 2


def test_corrections_are_atomic_and_names_must_be_unique(tmp_path):
    news = pd.DataFrame({"Name of Missionary": ["ELDER SMITH", "SISTER SMITH"], "Assignment": ["JC", "JC"]})
    path = tmp_path / "corrections.csv"
    path.write_text("Name,Field,Value\nELDER SMITH,Assignment,SA\nSMITH,Assignment,TR\n", encoding="utf-8-sig")
    with pytest.raises(ValueError, match="more than one"):
        _apply_manual_corrections(news, path, [])
    assert news["Assignment"].tolist() == ["JC", "JC"]
    path.write_text("Name,Field,Value\nSister Smith,Assignment,stl / tr\n", encoding="utf-8-sig")
    _apply_manual_corrections(news, path, [])
    assert news.iloc[1]["Assignment"] == "STL/TR"


@pytest.mark.parametrize("name,field,message", [("Missing", "Assignment", "does not match"), ("Elder Smith", "Typo", "unknown field")])
def test_bad_corrections_explain_recovery(tmp_path, name, field, message):
    news = pd.DataFrame({"Name of Missionary": ["ELDER SMITH"], "Assignment": ["JC"]})
    path = tmp_path / "corrections.csv"
    path.write_text(f"Name,Field,Value\n{name},{field},SA\n", encoding="utf-8")
    with pytest.raises(ValueError, match=message):
        _apply_manual_corrections(news, path, [])


def test_report_finds_roster_after_cover_and_title_rows(tmp_path):
    path = tmp_path / "report.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.create_sheet("Roster")
    ws.append(["Exported roster"])
    ws.append([" missionary   NAME ", "ZONE", "Area", "Position", "Companion"])
    ws.append(["Gabriel, Example", "New Zone", "New Area", "(STL, TR)", ""])
    wb.save(path)
    records = load_report(path)
    assert records[0].position == "STL/TR"
    assert records[0].district == ""
    ws.append(["Gabriel, Example", "Other Zone", "Other Area", "SC", ""])
    wb.save(path)
    with pytest.raises(ValueError, match="more than once"):
        load_report(path)


def test_position_preserves_future_combinations():
    assert _normalize_position("(WELFARE, SA)") == "WELFARE/SA"
    assert _normalize_position("(DL, TR, WELFARE)") == "DL/TR/WELFARE"
    assert _normalize_position("(DL, TR)") == "DT"
    assert _normalize_position("(ZL, JC)") == "ZL"


def test_custom_roles_survive_both_publications(sample_workbook, tmp_path):
    from docx import Document
    from pypdf import PdfReader
    with pd.ExcelFile(sample_workbook) as excel:
        sheets = {sheet: pd.read_excel(excel, sheet_name=sheet) for sheet in excel.sheet_names}
    sheets["News Format"].loc[0, "Assignment"] = "STL/TR"
    sheets["News Format"].loc[1, "Assignment"] = "WELFARE"
    with pd.ExcelWriter(sample_workbook) as writer:
        for sheet, frame in sheets.items():
            frame.to_excel(writer, sheet_name=sheet, index=False)
    word = tmp_path / "news.docx"
    pdf = tmp_path / "news.pdf"
    build_transfer_docx(sample_workbook, word)
    build_transfer_pdf(sample_workbook, pdf)
    from movement_plan import parse_transfer_pdf
    parsed = parse_transfer_pdf(pdf)
    assert {item.name for item in parsed.assignments} == {"ELDER ALPHA", "ELDER BETA", "ELDER GAMMA"}
    document = Document(word)
    word_text = " ".join(cell.text for table in document.tables for row in table.rows for cell in row.cells)
    pdf_text = " ".join(page.extract_text() for page in PdfReader(pdf).pages)
    for text in (word_text, pdf_text):
        compact = "".join(text.split())
        assert compact.count("STL/TR") >= 2
        assert compact.count("WELFARE") >= 2


def test_movement_fallback_does_not_split_surname_at_role_substring():
    from movement_plan import ApartmentDirectory, DirectoryEntry, parse_transfer_text
    directory = ApartmentDirectory([DirectoryEntry("UYO", "CENTRAL APT", "CENTRAL")])
    parsed = parse_transfer_text("1 ELDER STAPLETON SC UYO CENTRAL COMPANION", directory)
    assert parsed.assignments[0].name.upper() == "ELDER STAPLETON"
