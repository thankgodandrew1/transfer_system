"""Opt-in end-to-end test against the office's confidential local exports."""
import csv
import io
import json
import os
from pathlib import Path

import pandas as pd
import pytest

import app as webapp
from movement_plan import parse_transfer_pdf, canonical
from review_workflow import load_review
from text_utils import comparison_key


@pytest.mark.skipif(os.getenv("RUN_REAL_DATA_TESTS") != "1", reason="confidential fixtures are opt-in")
def test_real_review_emergency_correction_publication_and_movement(monkeypatch, tmp_path):
    data = Path(__file__).resolve().parents[1] / "data"
    inputs = {"management_pdf": data / "current_transfer_management.pdf", "previous_news_pdf": data / "old_transfer_news.pdf",
        "current_report": data / "current_transfer_report.xlsx", "old_report": data / "old_transfer_report.xlsx"}
    if not all(path.exists() for path in inputs.values()):
        pytest.skip("local exports are unavailable")
    root = tmp_path / "jobs"
    root.mkdir()
    monkeypatch.setattr(webapp, "JOB_ROOT", root)
    monkeypatch.setattr(webapp, "ACCESS_KEY", "")
    webapp.app.config.update(TESTING=True, SECRET_KEY="review-real", SESSION_COOKIE_SECURE=False)
    client = webapp.app.test_client()
    client.get("/")
    with client.session_transaction() as session:
        token = session["csrf_token"]
    response = client.post("/generate", data={"csrf_token": token, "review_before_publish": "on", "include_roster": "on",
        **{key: (io.BytesIO(path.read_bytes()), path.name) for key, path in inputs.items()}}, follow_redirects=True)
    assert b"Roster preview" in response.data
    job = next(root.iterdir())
    state = load_review(job)
    current_names = {comparison_key(row["full_name"]) for row in state["current"]}
    index = next(i for i, row in enumerate(state["previous"]) if comparison_key(row["full_name"]) in current_names
        and sum(comparison_key(other["last"]) == comparison_key(row["last"]) for other in state["previous"]) == 1)
    row = state["previous"][index]
    name = row["full_name"]
    original_previous = (job / "inputs" / "old_transfer_report.xlsx").read_bytes()
    url = f"/jobs/{job.name}/review"
    response = client.post(url, data={"csrf_token": token, "revision": "0", "action": "edit", "source": "previous",
        "missionary": str(index), "zone": row["zone"], "area": "EMERGENCY REVIEW AREA", "role": row["position"] or "SC",
        "reason": "Test confirmed pre-transfer emergency relocation",
        "companions": [str(i) for i, other in enumerate(state["previous"]) if other["full_name"] in row["companions"]]}, follow_redirects=True)
    assert b"Correction saved" in response.data
    response = client.post(url, data={"csrf_token": token, "revision": "1", "action": "publish", "confirmed": "on"}, follow_redirects=True)
    manifest = json.loads((job / "manifest.json").read_text())
    assert manifest["status"] == "complete", manifest.get("error")
    assert manifest["summary"]["blocking"] == 0
    assert (job / "inputs" / "old_transfer_report.xlsx").read_bytes() == original_previous
    def file_path(label):
        return job / next(item["relative_path"] for item in manifest["files"] if item["label"] == label)
    baseline = parse_transfer_pdf(file_path("Previous baseline for movements - PDF"))
    current = parse_transfer_pdf(file_path("Transfer News - PDF"))
    workbook = pd.read_excel(file_path("Review workbook"), sheet_name="News Format")
    surname = name.split(",")[0].upper()
    target = next(row for row in baseline.assignments if row.name.upper().endswith(" " + surname))
    assert target.area == "EMERGENCY REVIEW AREA"
    assert "EMERGENCY REVIEW AREA" in set(workbook["Previous Area"])
    assert b"Prepare movements using these assignments" in response.data
    page = client.get(f"/movement?news_job={job.name}")
    assert b"reviewed previous baseline" in page.data
    directory = io.StringIO()
    writer = csv.writer(directory)
    writer.writerow(["Zone", "Apartment", "Area"])
    seen = set()
    for assignment in baseline.assignments + current.assignments:
        if canonical(assignment.area) not in seen:
            seen.add(canonical(assignment.area))
            writer.writerow([assignment.zone, assignment.area + " APARTMENT", assignment.area])
    response = client.post("/movement/generate", data={"csrf_token": token, "news_job": job.name,
        "effective_date": "2026-11-17", "apartment_directory": (io.BytesIO(directory.getvalue().encode()), "apartments.csv")}, follow_redirects=True)
    movement_job = next(path for path in root.iterdir() if path != job)
    movement = json.loads((movement_job / "manifest.json").read_text())
    assert movement["status"] == "complete", movement.get("error")
    csv_file = next(item for item in movement["files"] if item["kind"] == "csv" and "Movement" in item["label"])
    movements = pd.read_csv(movement_job / csv_file["relative_path"])
    target_movement = movements[movements["Missionary"].str.upper() == target.name.upper()].iloc[0]
    assert target_movement["From Apartment"] == "EMERGENCY REVIEW AREA APARTMENT"
