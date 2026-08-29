from __future__ import annotations

import csv
import hashlib
import hmac
import json
import logging
import os
import re
import secrets
import shutil
import tempfile
import threading
import time
import zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

import openpyxl
import pandas as pd
from flask import (
    Flask,
    abort,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from pypdf import PdfReader
from dotenv import load_dotenv

from data.generate_transfer_sheet import generate_transfer_sheet
from generate_transfer_document import (
    DEFAULT_MISSION_NAME,
    DEFAULT_PREPARED_BY,
    DEFAULT_PRESIDENT,
    build_statistics_docx,
    build_transfer_docx,
    read_meta_values,
    title_slug,
)
from generate_transfer_pdf import build_statistics_pdf, build_transfer_pdf


BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")
JOB_ROOT = Path(os.getenv("TRANSFER_JOB_ROOT", Path(tempfile.gettempdir()) / "transfer-news-jobs")).resolve()
JOB_ROOT.mkdir(parents=True, exist_ok=True)

APP_NAME = os.getenv("APP_NAME", "Transfer News Generator")
GITHUB_URL = os.getenv(
    "GITHUB_URL",
    "https://github.com/thankgodandrew1/transfer_system",
).strip()
CREATOR_NAME = os.getenv("CREATOR_NAME", "ThankGod Andrew").strip() or "ThankGod Andrew"
ACCESS_KEY = os.getenv("APP_ACCESS_KEY", "")
REQUIRE_ACCESS_KEY = os.getenv("REQUIRE_ACCESS_KEY", "0") == "1"
JOB_TTL_MINUTES = max(int(os.getenv("JOB_TTL_MINUTES", "60")), 10)
MAX_FILE_MB = max(int(os.getenv("MAX_FILE_MB", "20")), 1)
MAX_FILE_BYTES = MAX_FILE_MB * 1024 * 1024

if REQUIRE_ACCESS_KEY and len(ACCESS_KEY) < 12:
    raise RuntimeError("APP_ACCESS_KEY must be set to at least 12 characters when REQUIRE_ACCESS_KEY=1.")

app = Flask(__name__)
app.config.update(
    SECRET_KEY=os.getenv("FLASK_SECRET_KEY") or secrets.token_hex(32),
    MAX_CONTENT_LENGTH=MAX_FILE_BYTES * 6,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    # Local development runs over HTTP. Production hosting sets this to 1.
    SESSION_COOKIE_SECURE=os.getenv("COOKIE_SECURE", "0") == "1",
)

generation_lock = threading.BoundedSemaphore(value=1)


class GenerationWarningHandler(logging.Handler):
    def __init__(self) -> None:
        super().__init__(level=logging.WARNING)
        self.messages: list[str] = []

    def emit(self, record: logging.LogRecord) -> None:
        self.messages.append(self.format(record))


def utc_now() -> datetime:
    return datetime.now(timezone.utc)


def _csrf_token() -> str:
    token = session.get("csrf_token")
    if not token:
        token = secrets.token_urlsafe(32)
        session["csrf_token"] = token
    return token


def _require_csrf() -> None:
    supplied = request.form.get("csrf_token", "")
    expected = session.get("csrf_token", "")
    if not supplied or not expected or not hmac.compare_digest(supplied, expected):
        abort(400, "The form expired. Refresh the page and try again.")


def _form_text(field: str, fallback: str, limit: int) -> str:
    value = request.form.get(field, "").strip()
    value = re.sub(r"[\x00-\x1f\x7f]", " ", value)
    return (value or fallback)[:limit]


def _safe_next_path(value: str | None) -> str:
    if value and value.startswith("/") and not value.startswith("//"):
        return value
    return url_for("index")


def _job_dir(job_id: str) -> Path:
    if not re.fullmatch(r"[a-f0-9]{32}", job_id):
        abort(404)
    target = (JOB_ROOT / job_id).resolve()
    if target.parent != JOB_ROOT:
        abort(404)
    return target


def _manifest_path(job_id: str) -> Path:
    return _job_dir(job_id) / "manifest.json"


def _load_manifest(job_id: str) -> dict[str, Any]:
    path = _manifest_path(job_id)
    if not path.exists():
        abort(404)
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        abort(404)


def _write_manifest(job_dir: Path, manifest: dict[str, Any]) -> None:
    temp_path = job_dir / "manifest.tmp"
    temp_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")
    temp_path.replace(job_dir / "manifest.json")


def _cleanup_expired_jobs() -> None:
    cutoff = time.time() - JOB_TTL_MINUTES * 60
    try:
        children = list(JOB_ROOT.iterdir())
    except OSError:
        return
    for child in children:
        try:
            resolved = child.resolve()
            if not child.is_dir() or resolved.parent != JOB_ROOT or child.is_symlink():
                continue
            if child.stat().st_mtime < cutoff:
                shutil.rmtree(resolved)
        except OSError:
            continue


def _validate_zip_office(path: Path, expected_suffix: str) -> None:
    if path.suffix.lower() != expected_suffix:
        raise ValueError(f"Expected a {expected_suffix} file.")
    try:
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
            if "[Content_Types].xml" not in names:
                raise ValueError("The Office file is missing its content manifest.")
            bad_member = archive.testzip()
            if bad_member:
                raise ValueError("The Office file is damaged.")
    except zipfile.BadZipFile as exc:
        raise ValueError("The Office file is damaged or has the wrong extension.") from exc


def _validate_upload(path: Path, kind: str) -> None:
    size = path.stat().st_size
    if size == 0:
        raise ValueError("An uploaded file is empty.")
    if size > MAX_FILE_BYTES:
        raise ValueError(f"Each upload must be {MAX_FILE_MB} MB or smaller.")
    if kind == "pdf":
        if path.suffix.lower() != ".pdf" or path.read_bytes()[:5] != b"%PDF-":
            raise ValueError("Expected a valid PDF file.")
        try:
            if not PdfReader(path).pages:
                raise ValueError("The PDF has no pages.")
        except Exception as exc:
            raise ValueError("The PDF could not be read.") from exc
    elif kind == "xlsx":
        _validate_zip_office(path, ".xlsx")
        try:
            workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
            if not workbook.sheetnames:
                raise ValueError("The workbook has no worksheets.")
            workbook.close()
        except Exception as exc:
            raise ValueError("The Excel report could not be read.") from exc
    elif kind == "docx":
        _validate_zip_office(path, ".docx")
    elif kind == "csv":
        if path.suffix.lower() != ".csv":
            raise ValueError("Expected a CSV file.")
        with path.open("r", encoding="utf-8-sig", newline="") as stream:
            headers = set(next(csv.reader(stream), []))
        required = {"Name", "Field", "Value"}
        if not required.issubset(headers):
            raise ValueError("Corrections CSV must contain Name, Field, and Value columns.")


def _save_upload(field: str, destination: Path, kind: str, required: bool = True) -> Path | None:
    uploaded = request.files.get(field)
    if uploaded is None or not uploaded.filename:
        if required:
            raise ValueError(f"Missing required upload: {field.replace('_', ' ')}.")
        return None
    destination.parent.mkdir(parents=True, exist_ok=True)
    uploaded.save(destination)
    try:
        _validate_upload(destination, kind)
    except Exception:
        destination.unlink(missing_ok=True)
        raise
    return destination


def _file_record(path: Path, label: str, kind: str) -> dict[str, Any]:
    return {
        "name": path.name,
        "label": label,
        "kind": kind,
        "size": path.stat().st_size,
    }


def _workbook_summary(workbook_path: Path) -> dict[str, Any]:
    news = pd.read_excel(workbook_path, sheet_name="News Format")
    meta = read_meta_values(workbook_path)
    summary: dict[str, Any] = {
        "missionaries": len(news) + int(float(meta.get("New Missionaries (Incoming)", 0) or 0)),
        "named_rows": len(news),
        "new_missionaries": int(float(meta.get("New Missionaries (Incoming)", 0) or 0)),
        "zones": int(news["New/Existing Zone"].nunique()),
        "transfer_title": meta.get("Transfer Title", ""),
        "corrections": 0,
        "blocking": 0,
        "notes": 0,
        "changes": 0,
    }
    try:
        verification = pd.read_excel(workbook_path, sheet_name="Verification")
        counts = verification["Type"].astype(str).str.upper().value_counts().to_dict()
        summary.update(
            corrections=int(counts.get("CORRECTED", 0)),
            blocking=int(counts.get("BLOCKING", 0)),
            notes=int(counts.get("NOTE", 0)),
        )
    except Exception:
        pass
    try:
        changes = pd.read_excel(workbook_path, sheet_name="Changes")
        summary["changes"] = len(changes)
    except Exception:
        pass
    return summary


def _build_download_zip(job_dir: Path, files: list[dict[str, Any]], zip_path: Path) -> None:
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for record in files:
            path = job_dir / record["relative_path"]
            if path.exists() and path.is_file():
                archive.write(path, arcname=record["name"])


def _run_generation(
    job_dir: Path,
    values: dict[str, Any],
) -> tuple[list[dict[str, Any]], dict[str, Any], list[str], str | None]:
    input_dir = job_dir / "inputs"
    output_dir = job_dir / "outputs"
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    management_pdf = _save_upload("management_pdf", input_dir / "current_transfer_management.pdf", "pdf")
    previous_news = _save_upload("previous_news_pdf", input_dir / "previous_transfer_news.pdf", "pdf")
    current_report = _save_upload("current_report", input_dir / "current_transfer_report.xlsx", "xlsx")
    old_report = _save_upload("old_report", input_dir / "old_transfer_report.xlsx", "xlsx")
    template = _save_upload("word_template", input_dir / "template.docx", "docx", required=False)
    corrections = _save_upload("manual_corrections", input_dir / "manual_corrections.csv", "csv", required=False)

    warning_handler = GenerationWarningHandler()
    warning_handler.setFormatter(logging.Formatter("%(message)s"))
    news_logger = logging.getLogger("generate_news_format")
    previous_level = news_logger.level
    news_logger.addHandler(warning_handler)
    news_logger.setLevel(logging.WARNING)
    try:
        workbook, csv_path, log_path = generate_transfer_sheet(
            management_pdf,
            output_dir,
            previous_news,
            current_report_path=current_report,
            old_report_path=old_report,
            manual_corrections_path=corrections,
            transfer_title_override=values["transfer_title"] or None,
            raise_on_blocking=False,
        )
    finally:
        news_logger.removeHandler(warning_handler)
        news_logger.setLevel(previous_level)

    if warning_handler.messages:
        with Path(log_path).open("a", encoding="utf-8") as log_stream:
            log_stream.write("\n\nPARSER WARNINGS\n===============\n")
            log_stream.write("\n".join(warning_handler.messages))
            log_stream.write("\n")

    summary = _workbook_summary(Path(workbook))
    if summary["blocking"]:
        file_specs = [
            (Path(workbook), "Review workbook", "xlsx"),
            (Path(csv_path), "Transfer-sheet CSV", "csv"),
            (Path(log_path), "Generation log", "log"),
        ]
        records: list[dict[str, Any]] = []
        for path, label, kind in file_specs:
            path = path.resolve()
            record = _file_record(path, label, kind)
            record["relative_path"] = str(path.relative_to(job_dir)).replace("\\", "/")
            records.append(record)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_path = output_dir / f"Transfer_Review_Package_{stamp}.zip"
        _build_download_zip(job_dir, records, zip_path)
        zip_record = _file_record(zip_path, "Download review package", "zip")
        zip_record["relative_path"] = str(zip_path.relative_to(job_dir)).replace("\\", "/")
        records.insert(0, zip_record)
        noun = "discrepancy" if summary["blocking"] == 1 else "discrepancies"
        message = (
            f"Publication stopped because verification found {summary['blocking']} unresolved {noun}. "
            "Download the review workbook, resolve each BLOCKING row, then run the generator again."
        )
        return records, summary, warning_handler.messages, message

    meta = read_meta_values(workbook)
    transfer_title = values["transfer_title"] or meta.get("Transfer Title") or f"{datetime.now():%B %Y}".upper() + " TRANSFER NEWS"
    slug = title_slug(transfer_title)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    news_docx = output_dir / f"Transfer_News_{slug}_{stamp}.docx"
    news_pdf = output_dir / f"Transfer_News_{slug}_{stamp}.pdf"
    roster_docx = output_dir / f"Updated_Previous_Transfer_News_{slug}_{stamp}.docx"
    roster_pdf = output_dir / f"Updated_Previous_Transfer_News_{slug}_{stamp}.pdf"
    stats_docx = output_dir / "Transfer_Statistics.docx"
    stats_pdf = output_dir / "Transfer_Statistics.pdf"

    build_transfer_docx(
        workbook, news_docx, template,
        mission_name=values["mission_name"], transfer_title=transfer_title,
        president=values["president"], prepared_by=values["prepared_by"],
    )
    build_transfer_pdf(
        workbook, news_pdf,
        mission_name=values["mission_name"], transfer_title=transfer_title,
        president=values["president"], prepared_by=values["prepared_by"],
    )

    if values["include_roster"]:
        build_transfer_docx(
            workbook, roster_docx, None,
            mission_name=values["mission_name"], transfer_title=transfer_title,
            president=values["president"], prepared_by=values["prepared_by"], roster_mode=True,
        )
        build_transfer_pdf(
            workbook, roster_pdf,
            mission_name=values["mission_name"], transfer_title=transfer_title,
            president=values["president"], prepared_by=values["prepared_by"], roster_mode=True,
        )

    if values["include_stats"]:
        build_statistics_docx(workbook, stats_docx)
        build_statistics_pdf(workbook, stats_pdf)

    file_specs = [
        (Path(workbook), "Review workbook", "xlsx"),
        (Path(csv_path), "Transfer-sheet CSV", "csv"),
        (Path(log_path), "Generation log", "log"),
        (news_docx, "Transfer News - Word", "docx"),
        (news_pdf, "Transfer News - PDF", "pdf"),
    ]
    if values["include_roster"]:
        file_specs.extend([
            (roster_docx, "Updated roster - Word", "docx"),
            (roster_pdf, "Updated roster - PDF", "pdf"),
        ])
    if values["include_stats"]:
        file_specs.extend([
            (stats_docx, "Transfer statistics - Word", "docx"),
            (stats_pdf, "Transfer statistics - PDF", "pdf"),
        ])

    records: list[dict[str, Any]] = []
    for path, label, kind in file_specs:
        path = Path(path).resolve()
        if not path.exists() or job_dir.resolve() not in path.parents:
            continue
        record = _file_record(path, label, kind)
        record["relative_path"] = str(path.relative_to(job_dir)).replace("\\", "/")
        records.append(record)

    zip_path = output_dir / f"Transfer_News_Package_{slug}_{stamp}.zip"
    _build_download_zip(job_dir, records, zip_path)
    zip_record = _file_record(zip_path, "Download everything", "zip")
    zip_record["relative_path"] = str(zip_path.relative_to(job_dir)).replace("\\", "/")
    records.insert(0, zip_record)
    return records, summary, warning_handler.messages, None


@app.before_request
def before_request() -> Any:
    _cleanup_expired_jobs()
    if not ACCESS_KEY:
        return None
    if request.endpoint in {"access", "health", "static"}:
        return None
    if not session.get("authorized"):
        return redirect(url_for("access", next=request.path))
    return None


@app.after_request
def security_headers(response):
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("X-Frame-Options", "DENY")
    response.headers.setdefault("Referrer-Policy", "no-referrer")
    response.headers.setdefault("Permissions-Policy", "camera=(), microphone=(), geolocation=()")
    response.headers.setdefault("Cross-Origin-Opener-Policy", "same-origin")
    response.headers.setdefault("Cross-Origin-Resource-Policy", "same-origin")
    response.headers.setdefault(
        "Content-Security-Policy",
        "default-src 'self'; img-src 'self' data:; media-src 'self'; style-src 'self'; script-src 'self'",
    )
    response.headers["Cache-Control"] = "public, max-age=86400" if request.endpoint == "static" else "no-store"
    return response


@app.context_processor
def common_context() -> dict[str, Any]:
    return {
        "app_name": APP_NAME,
        "github_url": GITHUB_URL,
        "creator_name": CREATOR_NAME,
        "csrf_token": _csrf_token(),
        "job_ttl_minutes": JOB_TTL_MINUTES,
        "max_file_mb": MAX_FILE_MB,
    }


@app.route("/access", methods=["GET", "POST"])
def access():
    if not ACCESS_KEY:
        return redirect(url_for("index"))
    if request.method == "POST":
        _require_csrf()
        supplied = request.form.get("access_key", "")
        if hmac.compare_digest(supplied, ACCESS_KEY):
            session.clear()
            session["authorized"] = True
            session["csrf_token"] = secrets.token_urlsafe(32)
            return redirect(_safe_next_path(request.args.get("next")))
        flash("That access key is not correct.", "error")
    return render_template("access.html")


@app.post("/logout")
def logout():
    _require_csrf()
    session.clear()
    return redirect(url_for("access"))


@app.get("/")
def index():
    return render_template(
        "index.html",
        defaults={
            "mission_name": DEFAULT_MISSION_NAME,
            "president": DEFAULT_PRESIDENT,
            "prepared_by": DEFAULT_PREPARED_BY,
        },
    )


@app.post("/generate")
def generate():
    _require_csrf()
    if not generation_lock.acquire(blocking=False):
        flash("Another transfer is being generated. Please wait a moment and try again.", "error")
        return redirect(url_for("index")), 429

    job_id = secrets.token_hex(16)
    job_dir = _job_dir(job_id)
    job_dir.mkdir(parents=True, exist_ok=False)
    created = utc_now()
    values = {
        "mission_name": _form_text("mission_name", DEFAULT_MISSION_NAME, 120),
        "president": _form_text("president", DEFAULT_PRESIDENT, 120),
        "prepared_by": _form_text("prepared_by", DEFAULT_PREPARED_BY, 160),
        "transfer_title": _form_text("transfer_title", "", 120).upper(),
        "include_roster": request.form.get("include_roster") == "on",
        "include_stats": request.form.get("include_stats") == "on",
    }
    manifest: dict[str, Any] = {
        "job_id": job_id,
        "status": "processing",
        "created_at": created.isoformat(),
        "expires_at": (created + timedelta(minutes=JOB_TTL_MINUTES)).isoformat(),
        "files": [],
        "summary": {},
        "warnings": [],
    }
    _write_manifest(job_dir, manifest)
    try:
        files, summary, warnings, blocked_message = _run_generation(job_dir, values)
        manifest.update(
            status="blocked" if blocked_message else "complete",
            files=files,
            summary=summary,
            warnings=warnings,
        )
        if blocked_message:
            manifest["error"] = blocked_message
    except ValueError as exc:
        manifest.update(status="error", error=str(exc))
    except Exception as exc:
        # Return a useful but compact error to the owner without logging
        # uploaded names or document contents to the hosting provider.
        manifest.update(status="error", error=f"Generation stopped: {exc}")
    finally:
        _write_manifest(job_dir, manifest)
        generation_lock.release()
    return redirect(url_for("job_result", job_id=job_id))


@app.get("/jobs/<job_id>")
def job_result(job_id: str):
    manifest = _load_manifest(job_id)
    return render_template("result.html", job=manifest)


@app.get("/jobs/<job_id>/download/<path:filename>")
def download(job_id: str, filename: str):
    manifest = _load_manifest(job_id)
    record = next((item for item in manifest.get("files", []) if item.get("name") == filename), None)
    if record is None:
        abort(404)
    path = (_job_dir(job_id) / record["relative_path"]).resolve()
    if _job_dir(job_id) not in path.parents or not path.is_file():
        abort(404)
    return send_file(path, as_attachment=True, download_name=record["name"], conditional=True)


@app.get("/guide")
def guide():
    return render_template("guide.html")


@app.get("/health")
def health():
    return {"status": "ok", "service": APP_NAME}


@app.errorhandler(413)
def too_large(_error):
    return render_template("error.html", message=f"The upload is too large. Each file must be {MAX_FILE_MB} MB or smaller."), 413


@app.errorhandler(400)
def bad_request(error):
    return render_template("error.html", message=getattr(error, "description", "The request could not be processed.")), 400


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=int(os.getenv("PORT", "5000")), debug=os.getenv("FLASK_DEBUG") == "1")
