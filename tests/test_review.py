from dataclasses import asdict
import io
import json
from pathlib import Path

import openpyxl
import pandas as pd
import pytest
from werkzeug.datastructures import MultiDict
from pypdf import PdfWriter

import app as webapp
from review_workflow import (baseline_workbook, edit_assignment, export_report, initialize_review,
    link_identity, load_review, review_summary, save_review, undo_last)
from verify_news import ReportRecord, load_report


def record(name, area="OLD AREA", role="SC", companions=()):
    return asdict(ReportRecord("UYO", "CENTRAL", area, name, name.split(",")[0], "ELDER", role, list(companions)))


@pytest.fixture
def state():
    return {"revision": 0, "edits": [], "previous": [record("Alpha, Adam", companions=["Beta, Ben"]),
        record("Beta, Ben", role="JC", companions=["Alpha, Adam"])],
        "current": [record("Alpha, Adam", "NEW AREA", companions=["Beta, Ben"]),
        record("Beta, Ben", "NEW AREA", "JC", ["Alpha, Adam"]), record("Gamma, George", "NEW AREA")],
        "values": {"mission_name": "TEST MISSION", "president": "PRESIDENT TEST", "prepared_by": "OFFICE",
            "transfer_title": "NOVEMBER 2026 TRANSFER NEWS", "current_role_overrides": "", "include_roster": True, "include_stats": False}}


def edit_form(**changes):
    return MultiDict(dict(source="previous", missionary="0", zone="UYO", area="EMERGENCY AREA", role="SA",
        reason="Emergency transfer before this cycle", **changes))


def test_emergency_assignment_updates_both_companions_and_audit(state):
    form = edit_form()
    form.setlist("companions", ["1"])
    edit_assignment(state, form)
    assert state["revision"] == 1
    assert all(row["area"] == "EMERGENCY AREA" for row in state["previous"])
    assert state["previous"][0]["companions"] == ["Beta, Ben"]
    assert state["previous"][1]["companions"] == ["Alpha, Adam"]
    assert state["previous"][1]["position"] == "JC"
    assert len(state["edits"][0]["changes"]) == 2
    assert state["current"][0]["area"] == "NEW AREA"


def test_solo_emergency_removes_stale_reciprocal_links(state):
    edit_assignment(state, edit_form())
    assert all(not row["companions"] for row in state["previous"])


def test_undo_restores_complete_group_and_keeps_history(state):
    original = json.loads(json.dumps(state["previous"]))
    edit_assignment(state, edit_form())
    undo_last(state)
    assert state["previous"] == original
    assert state["revision"] == 2
    assert not state["edits"]
    assert state["undone"][0]["correction"]["reason"]


@pytest.mark.parametrize("field,value", [("missionary", "-1"), ("missionary", "999"), ("source", "inputs"),
    ("role", ""), ("reason", ""), ("zone", ""), ("area", "=BAD\nFORMULA")])
def test_invalid_edits_leave_state_unchanged(state, field, value):
    before = json.dumps(state)
    form = edit_form()
    form[field] = value
    with pytest.raises(ValueError):
        edit_assignment(state, form)
    assert json.dumps(state) == before


def test_identity_link_is_explicit_and_preserves_history(state):
    state["previous"][0].update(full_name="Alfa, Adam", last="Alfa")
    state["previous"][1]["companions"] = ["Alfa, Adam"]
    form = MultiDict(dict(current_identity="0", previous_identity="0", reason="Office confirmed alternate spelling"))
    link_identity(state, form)
    assert len(review_summary(state)["arrivals"]) == 1
    assert not review_summary(state)["departures"]
    assert state["previous"][1]["companions"] == ["Alpha, Adam"]
    assert state["edits"][0]["name"] == "Alfa, Adam"
    with pytest.raises(ValueError, match="Only link"):
        link_identity(state, form)
    undo_last(state)
    assert state["previous"][0]["full_name"] == "Alfa, Adam"
    assert state["previous"][1]["companions"] == ["Alfa, Adam"]


def test_exports_preserve_custom_roles_and_literal_excel_text(state, tmp_path):
    state["previous"][0]["position"] = "STL/TR"
    state["previous"][0]["area"] = "=LITERAL AREA"
    path = tmp_path / "roster.xlsx"
    export_report(state["previous"], path)
    rows = load_report(path)
    assert rows[0].position == "STL/TR"
    assert rows[0].area == "=LITERAL AREA"
    assert rows[0].companions == ["Beta, Ben"]


def test_same_surname_and_initial_stays_distinct_without_ids(state, tmp_path):
    from review_workflow import printable_names
    from verify_news import verify, _last_name
    previous = [record("Smith, John", companions=["Smith, James"]), record("Smith, James", role="JC", companions=["Smith, John"])]
    state["previous"] = previous
    state["current"] = json.loads(json.dumps(previous))
    names = printable_names(previous)
    assert names == ["ELDER SMITH JO.", "ELDER SMITH JA."]
    from generate_news_format import _base_name
    assert _base_name("SMITH JO.") == "SMITH"
    assert _base_name("COLLINS JR.") == "COLLINS JR"
    assert _last_name(names[0]) == "SMITH"
    assert _last_name("ELDER COLLINS JR.") == "COLLINS JR"
    old, current = tmp_path / "old.xlsx", tmp_path / "current.xlsx"
    export_report(previous, old)
    export_report(state["current"], current)
    workbook = tmp_path / "baseline.xlsx"
    baseline_workbook(state, workbook)
    frame = pd.read_excel(workbook, sheet_name="News Format").fillna("")
    result = verify(frame, current, old)
    assert not result.blocking_errors
    assert list(result.news["Name of Missionary"]) == names


def test_quick_roles_are_visible_editable_and_undoable(state, tmp_path):
    inputs = tmp_path / "inputs"
    inputs.mkdir()
    export_report(state["current"], inputs / "current_transfer_report.xlsx")
    export_report(state["previous"], inputs / "old_transfer_report.xlsx")
    original = (inputs / "current_transfer_report.xlsx").read_bytes()
    snapshot = initialize_review(tmp_path, state["values"], [("Elder Alpha", "WELFARE/TR")])
    assert snapshot["current"][0]["position"] == "WELFARE/TR"
    assert snapshot["values"]["current_role_overrides"] == ""
    undo_last(snapshot)
    assert snapshot["current"][0]["position"] == "SC"
    assert (inputs / "current_transfer_report.xlsx").read_bytes() == original


def test_uninitialled_arrival_with_same_surname_does_not_match_continuing_person(state, tmp_path):
    from verify_news import verify
    previous = [record("Smith, John", "OLD AREA")]
    current = previous + [record("Smith, James", "NEW AREA", "JC")]
    old, now = tmp_path / "old.xlsx", tmp_path / "now.xlsx"
    export_report(previous, old)
    export_report(current, now)
    frame = pd.DataFrame([{"Name of Missionary": "ELDER SMITH", "Assignment": row["position"],
        "New/Existing Zone": "UYO", "New/Existing Area": row["area"], "New/Existing Companion(s)": "",
        "Previous Zone": "UYO", "Previous Area": row["area"], "Previous Assignment": row["position"], "Status": "STAYING"} for row in current])
    result = verify(frame, now, old)
    assert not result.blocking_errors
    assert len(result.news) == 1
    assert result.news.iloc[0]["New/Existing Area"] == "OLD AREA"


def test_baseline_uses_emergency_locations_and_current_name_disambiguation(state, tmp_path):
    state["previous"].append(record("Alpha, Aaron", "RELEASED AREA"))
    edit_assignment(state, edit_form())
    path = tmp_path / "baseline.xlsx"
    baseline_workbook(state, path)
    rows = pd.read_excel(path, sheet_name="News Format")
    assert rows.iloc[0]["Name of Missionary"] == "ELDER ALPHA"
    assert rows.iloc[0]["New/Existing Area"] == "EMERGENCY AREA"
    assert len(set(rows["Name of Missionary"])) == 3


def _client(monkeypatch, tmp_path):
    root = tmp_path / "jobs"
    root.mkdir()
    monkeypatch.setattr(webapp, "JOB_ROOT", root)
    monkeypatch.setattr(webapp, "ACCESS_KEY", "")
    webapp.app.config.update(TESTING=True, SECRET_KEY="review-test", SESSION_COOKIE_SECURE=False)
    client = webapp.app.test_client()
    client.get("/")
    with client.session_transaction() as session:
        token = session["csrf_token"]
    return client, root, token


def _draft(root, state):
    job_id = "a" * 32
    directory = root / job_id
    directory.mkdir()
    (directory / "inputs").mkdir()
    save_review(directory, state)
    webapp._write_manifest(directory, dict(job_id=job_id, kind="transfer_news", status="review", has_review=True,
        summary={}, files=[], warnings=[], expires_at="2099-01-01T00:00:00Z"))
    return job_id, directory


def test_review_save_stale_tab_csrf_and_safe_render(monkeypatch, tmp_path, state):
    client, root, token = _client(monkeypatch, tmp_path)
    state["previous"][0]["area"] = "<script>alert(1)</script>"
    job_id, directory = _draft(root, state)
    url = f"/jobs/{job_id}/review"
    page = client.get(url)
    assert page.status_code == 200
    assert b"&lt;script&gt;" in page.data
    assert b"<script>alert(1)</script>" not in page.data
    assert client.post(url, data={"action": "edit"}).status_code == 400
    form = edit_form()
    form.update(dict(csrf_token=token, revision="0", action="edit"))
    response = client.post(url, data=form, follow_redirects=True)
    assert b"Correction saved" in response.data
    response = client.post(url, data=form, follow_redirects=True)
    assert b"another tab" in response.data
    assert load_review(directory)["revision"] == 1


def test_publish_requires_confirmation_and_blocked_job_can_be_corrected(monkeypatch, tmp_path, state):
    client, root, token = _client(monkeypatch, tmp_path)
    job_id, directory = _draft(root, state)
    url = f"/jobs/{job_id}/review"
    response = client.post(url, data=dict(csrf_token=token, revision="0", action="publish"), follow_redirects=True)
    assert b"Confirm that" in response.data
    def blocked(*args):
        return [], {"blocking_details": [{"Missionary": "ELDER ALPHA", "Detail": "Check identity"}]}, [], "Needs review"
    monkeypatch.setattr(webapp, "_run_generation", blocked)
    response = client.post(url, data=dict(csrf_token=token, revision="0", action="publish", confirmed="on"), follow_redirects=True)
    assert b"Check identity" in response.data
    assert json.loads((directory / "manifest.json").read_text())["status"] == "blocked"
    form = edit_form()
    form.update(dict(csrf_token=token, revision="0", action="edit"))
    client.post(url, data=form)
    assert json.loads((directory / "manifest.json").read_text())["status"] == "review"


def test_completed_package_is_immutable_and_revision_gets_new_job(monkeypatch, tmp_path, state):
    client, root, token = _client(monkeypatch, tmp_path)
    job_id, directory = _draft(root, state)
    def complete(*args):
        return [], {"blocking": 0}, [], None
    monkeypatch.setattr(webapp, "_run_generation", complete)
    url = f"/jobs/{job_id}/review"
    client.post(url, data=dict(csrf_token=token, revision="0", action="publish", confirmed="on"))
    form = edit_form()
    form.update(dict(csrf_token=token, revision="0", action="edit"))
    response = client.post(url, data=form, follow_redirects=True)
    assert b"already generated" in response.data
    assert load_review(directory)["revision"] == 0
    response = client.post(f"/jobs/{job_id}/revise", data=dict(csrf_token=token))
    assert response.status_code == 302
    assert job_id not in response.location
    assert len(list(root.glob("*/review.json"))) == 2


def test_upload_goes_to_review_without_generating_documents(monkeypatch, tmp_path, state):
    client, root, token = _client(monkeypatch, tmp_path)
    def must_not_generate(*args):
        pytest.fail("Publication must wait for staff confirmation")
    monkeypatch.setattr(webapp, "_run_generation", must_not_generate)
    current, previous = tmp_path / "current.xlsx", tmp_path / "previous.xlsx"
    export_report(state["current"], current)
    export_report(state["previous"], previous)
    writer = PdfWriter()
    writer.add_blank_page(612, 792)
    pdf = io.BytesIO()
    writer.write(pdf)
    response = client.post("/generate", data={"csrf_token": token, "review_before_publish": "on",
        "management_pdf": (io.BytesIO(pdf.getvalue()), "current.pdf"),
        "previous_news_pdf": (io.BytesIO(pdf.getvalue()), "previous.pdf"),
        "current_report": (io.BytesIO(current.read_bytes()), "current.xlsx"),
        "old_report": (io.BytesIO(previous.read_bytes()), "previous.xlsx")}, follow_redirects=True)
    assert response.status_code == 200
    assert b"Roster preview" in response.data
    assert not list(root.glob("*/outputs/*.pdf"))


@pytest.mark.parametrize("replacement", ["none", "pdf", "text"])
def test_movement_handoff_preserves_explicit_replacement_priority(monkeypatch, tmp_path, state, replacement):
    from movement_plan import Assignment, ParseResult
    client, root, token = _client(monkeypatch, tmp_path)
    job_id, directory = _draft(root, state)
    outputs = directory / "outputs"
    outputs.mkdir()
    for filename, contents in (("baseline.pdf", b"saved previous"), ("news.pdf", b"saved current")):
        (outputs / filename).write_bytes(contents)
    manifest = json.loads((directory / "manifest.json").read_text())
    manifest.update(status="complete", files=[{"label": "Previous baseline for movements - PDF", "relative_path": "outputs/baseline.pdf"},
        {"label": "Transfer News - PDF", "relative_path": "outputs/news.pdf"}])
    webapp._write_manifest(directory, manifest)
    seen = {}
    def parse(pdf_path, text, apartment_directory, label):
        seen[label] = (pdf_path.read_bytes() if pdf_path else None, text)
        area = "OLD AREA" if label.startswith("previous") else "NEW AREA"
        return ParseResult([Assignment("ELDER ALPHA", "UYO", area, 1)])
    monkeypatch.setattr(webapp, "_parse_movement_source", parse)
    data = {"csrf_token": token, "news_job": job_id, "effective_date": "2026-11-17",
        "apartment_directory": (io.BytesIO(b"Zone,Apartment,Area\nUYO,OLD APARTMENT,OLD AREA\nUYO,NEW APARTMENT,NEW AREA\n"), "apartments.csv")}
    if replacement == "pdf":
        writer = PdfWriter()
        writer.add_blank_page(612, 792)
        pdf = io.BytesIO()
        writer.write(pdf)
        data["movement_previous_pdf"] = (io.BytesIO(pdf.getvalue()), "replacement.pdf")
    elif replacement == "text":
        data["previous_transfer_text"] = "replacement table"
    response = client.post("/movement/generate", data=data, follow_redirects=True)
    assert response.status_code == 200
    assert seen["current Transfer News"][0] == b"saved current"
    if replacement == "none":
        assert seen["previous Transfer News"][0] == b"saved previous"
    elif replacement == "pdf":
        assert seen["previous Transfer News"][0].startswith(b"%PDF")
    else:
        assert seen["previous Transfer News"] == (None, "replacement table")
