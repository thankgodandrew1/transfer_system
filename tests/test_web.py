from __future__ import annotations

import io
from html.parser import HTMLParser
from pathlib import Path

import app as webapp


def test_guide_has_two_independent_lazy_video_players(monkeypatch, tmp_path: Path) -> None:
    class Videos(HTMLParser):
        def __init__(self):
            super().__init__()
            self.players = []

        def handle_starttag(self, tag, attrs):
            attrs = dict(attrs)
            if tag == "video":
                self.players.append(dict(attrs, sources=[], captions=[]))
            elif tag == "source":
                self.players[-1]["sources"].append(attrs["src"])
            elif tag == "track":
                self.players[-1]["captions"].append(attrs["src"])

    client = _client(monkeypatch, tmp_path)
    page = client.get("/guide")
    assert page.status_code == 200
    parsed = Videos()
    parsed.feed(page.data.decode())
    assert len(parsed.players) == 2
    assert all(player["preload"] == "none" and "controls" in player and "playsinline" in player for player in parsed.players)
    assert parsed.players[0]["sources"] == ["/static/walkthrough.mp4"]
    assert parsed.players[0]["captions"] == ["/static/walkthrough.vtt"]
    assert parsed.players[1]["sources"] == ["/static/walkthrough-detailed.mp4"]
    assert not parsed.players[1]["captions"]  # Do not attach the first video's unrelated captions.
    for player in parsed.players:
        assert client.get(player["poster"]).status_code == 200


def test_detailed_walkthrough_supports_range_requests(monkeypatch, tmp_path: Path) -> None:
    client = _client(monkeypatch, tmp_path)
    path = webapp.BASE_DIR / "static" / "walkthrough-detailed.mp4"
    assert 0 < path.stat().st_size < 100 * 1024 * 1024
    response = client.get("/static/walkthrough-detailed.mp4", headers={"Range": "bytes=0-1023"})
    try:
        assert response.status_code == 206
        assert response.mimetype == "video/mp4"
        assert len(response.data) == 1024
        assert response.headers["Content-Range"] == f"bytes 0-1023/{path.stat().st_size}"
    finally:
        response.close()


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


def test_branding_links_creator_repository_and_favicon(monkeypatch, tmp_path: Path) -> None:
    client = _client(monkeypatch, tmp_path)
    response = client.get("/")
    assert response.status_code == 200
    assert b"ThankGod Andrew" in response.data
    assert b"https://github.com/thankgodandrew1/transfer_system" in response.data
    assert b"favicon.svg" in response.data
    assert b"Current assignment overrides" in response.data


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
