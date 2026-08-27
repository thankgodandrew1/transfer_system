from __future__ import annotations

import io
from pathlib import Path

import app as webapp


def _client(monkeypatch, tmp_path: Path):
    job_root = tmp_path / "jobs"
    job_root.mkdir()
    monkeypatch.setattr(webapp, "JOB_ROOT", job_root)
    monkeypatch.setattr(webapp, "ACCESS_KEY", "")
    webapp.app.config.update(TESTING=True, SECRET_KEY="test-secret", SESSION_COOKIE_SECURE=False)
    return webapp.app.test_client()


def _csrf(client) -> str:
    client.get("/")
    with client.session_transaction() as session:
        return session["csrf_token"]


def test_health_and_security_headers(monkeypatch, tmp_path: Path) -> None:
    client = _client(monkeypatch, tmp_path)
    response = client.get("/health")
    assert response.status_code == 200
    assert response.json["status"] == "ok"
    assert response.headers["X-Content-Type-Options"] == "nosniff"
    assert response.headers["X-Frame-Options"] == "DENY"
    assert response.headers["Cache-Control"] == "no-store"
    assert "camera=()" in response.headers["Permissions-Policy"]
    assert client.get("/static/styles.css").headers["Cache-Control"] == "public, max-age=86400"


def test_generation_rejects_expired_form(monkeypatch, tmp_path: Path) -> None:
    client = _client(monkeypatch, tmp_path)
    response = client.post("/generate", data={})
    assert response.status_code == 400
    assert b"form expired" in response.data


def test_access_key_is_csrf_protected_and_redirect_stays_local(monkeypatch, tmp_path: Path) -> None:
    client = _client(monkeypatch, tmp_path)
    monkeypatch.setattr(webapp, "ACCESS_KEY", "a-strong-shared-key")

    page = client.get("/access?next=//example.invalid")
    assert page.status_code == 200
    with client.session_transaction() as session:
        token = session["csrf_token"]

    response = client.post(
        "/access?next=//example.invalid",
        data={"csrf_token": token, "access_key": "a-strong-shared-key"},
        follow_redirects=False,
    )
    assert response.status_code == 302
    assert response.headers["Location"] == "/"


def test_invalid_pdf_upload_is_rejected(monkeypatch, tmp_path: Path) -> None:
    client = _client(monkeypatch, tmp_path)
    token = _csrf(client)
    response = client.post(
        "/generate",
        data={
            "csrf_token": token,
            "management_pdf": (io.BytesIO(b"not a pdf"), "input.pdf"),
        },
        content_type="multipart/form-data",
        follow_redirects=True,
    )
    assert response.status_code == 200
    assert b"Expected a valid PDF file" in response.data


def test_complete_job_can_be_downloaded(monkeypatch, tmp_path: Path) -> None:
    client = _client(monkeypatch, tmp_path)

    def fake_run(job_dir: Path, _values):
        output = job_dir / "outputs" / "package.zip"
        output.parent.mkdir(parents=True)
        output.write_bytes(b"PK-test-package")
        files = [{
            "name": output.name,
            "label": "Download everything",
            "kind": "zip",
            "size": output.stat().st_size,
            "relative_path": "outputs/package.zip",
        }]
        summary = {
            "missionaries": 4, "zones": 1, "new_missionaries": 1,
            "corrections": 2, "blocking": 0, "notes": 0, "changes": 2,
        }
        return files, summary, [], None

    monkeypatch.setattr(webapp, "_run_generation", fake_run)
    response = client.post(
        "/generate",
        data={"csrf_token": _csrf(client)},
        follow_redirects=True,
    )
    assert response.status_code == 200
    assert b"package is ready" in response.data
    assert b"Download everything" in response.data

    with client.session_transaction():
        pass
    manifest_paths = list((tmp_path / "jobs").glob("*/manifest.json"))
    assert len(manifest_paths) == 1
    job_id = manifest_paths[0].parent.name
    download = client.get(f"/jobs/{job_id}/download/package.zip")
    assert download.status_code == 200
    assert download.data == b"PK-test-package"
    assert "attachment" in download.headers["Content-Disposition"]


def test_blocked_job_keeps_review_downloads(monkeypatch, tmp_path: Path) -> None:
    client = _client(monkeypatch, tmp_path)

    def fake_run(job_dir: Path, _values):
        output = job_dir / "outputs" / "review.xlsx"
        output.parent.mkdir(parents=True)
        output.write_bytes(b"review")
        files = [{
            "name": output.name,
            "label": "Review workbook",
            "kind": "xlsx",
            "size": output.stat().st_size,
            "relative_path": "outputs/review.xlsx",
        }]
        summary = {
            "missionaries": 3, "zones": 1, "new_missionaries": 0,
            "corrections": 0, "blocking": 1, "notes": 0, "changes": 0,
        }
        return files, summary, [], "Publication stopped because verification found 1 unresolved discrepancy."

    monkeypatch.setattr(webapp, "_run_generation", fake_run)
    response = client.post("/generate", data={"csrf_token": _csrf(client)}, follow_redirects=True)
    assert response.status_code == 200
    assert b"Publication paused safely" in response.data
    assert b"Review workbook" in response.data
