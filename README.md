# Transfer News Generator

A small Flask web application that turns mission transfer exports into a verified, downloadable Transfer News package. It runs as a native Python service—there is no Docker requirement and no Microsoft Word or LibreOffice dependency for PDF generation.

The repository link shown in the web header is controlled by `GITHUB_URL`. Set it to your GitHub profile or repository URL after you create the repository.

## What the system does

The generator:

1. Parses the current Transfer Management PDF.
2. Compares it with the previous Transfer News PDF.
3. Verifies current and previous facts against the two authoritative roster reports.
4. Masks confirmed incoming missionaries as `NEW MISSIONARY`.
5. Stops publication when a row cannot be verified safely.
6. Creates Word, native PDF, Excel, CSV, log, statistics, updated-roster, and ZIP downloads.

Each browser run uses an isolated temporary folder. Files expire automatically and are not stored in a database.

## Inputs

| Input | Format | Purpose |
|---|---|---|
| Current Transfer Management | PDF | Current areas, zones, assignments, and companionships |
| Previous Transfer News | PDF | Previous printed assignments and zone order |
| Current Transfer Report | XLSX | Authoritative current roster verification |
| Previous Transfer Report | XLSX | Authoritative previous-state and incoming-missionary verification |
| Word template | DOCX, optional | Approved custom document layout |
| Manual corrections | CSV, optional | Approved overrides with `Name,Field,Value` columns |

## Outputs

- Complete ZIP package
- Verified review workbook with `News Format`, `Transfer Sheet`, `Summary`, `Changes`, `Meta`, and `Verification` sheets as applicable
- Transfer News DOCX and native PDF
- Updated-roster DOCX and native PDF
- Statistics DOCX and native PDF
- Transfer-sheet CSV and generation log

If verification finds an unresolved discrepancy, publishing documents are not created. The results page still provides a review workbook, CSV, log, and review ZIP so the issue can be resolved safely.

## Run locally

### Windows PowerShell

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements-dev.txt
Copy-Item .env.example .env
python app.py
```

Open <http://127.0.0.1:5000>.

For local HTTP, keep `COOKIE_SECURE=0`. Choose a long `FLASK_SECRET_KEY` and replace the example access key before using real records.

### macOS or Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements-dev.txt
cp .env.example .env
python app.py
```

## Configuration

| Variable | Purpose | Local default |
|---|---|---|
| `APP_NAME` | Name shown in the interface | `Transfer News Generator` |
| `APP_ACCESS_KEY` | Shared key required before opening the generator | blank |
| `REQUIRE_ACCESS_KEY` | Refuse startup if a strong access key is missing | `0` |
| `FLASK_SECRET_KEY` | Signs browser sessions and CSRF tokens | random per restart |
| `GITHUB_URL` | GitHub profile or repository link shown in the site | hidden when blank |
| `JOB_TTL_MINUTES` | Temporary-file lifetime | `60` |
| `MAX_FILE_MB` | Maximum size of each upload | `20` |
| `COOKIE_SECURE` | Send session cookies over HTTPS only | `0` locally, `1` in production |
| `TRANSFER_JOB_ROOT` | Optional temporary job directory | operating-system temp folder |

## Tests

```powershell
python -m pytest -q
```

The confidential local integration fixture is opt-in:

```powershell
$env:RUN_REAL_DATA_TESTS='1'
python -m pytest -q
```

The normal suite uses synthetic records. Never commit real input or generated output files.

## Walkthrough video

The **How to use** page embeds a 49-second, captioned, sanitized walkthrough. Its presenter notes are in [docs/walkthrough-script.md](docs/walkthrough-script.md). Rebuild the video after a major UI change with:

```powershell
python scripts/build_walkthrough.py
```

## Free deployment

[Render](https://render.com/docs/free) is the recommended free host for this project because it supports a native Python web service directly from GitHub and can use the included `render.yaml`. The free service sleeps after inactivity, has ephemeral storage, and is intended for hobby/testing use; those limits fit the app’s temporary-job design.

Follow [docs/deployment.md](docs/deployment.md) for the full GitHub and Render walkthrough. No Dockerfile is used.

## Privacy and operational rules

- Use a private GitHub repository unless you intentionally want the source public.
- Never commit real PDFs, spreadsheets, corrections, logs, or generated documents.
- Set `APP_ACCESS_KEY` on any internet-facing deployment.
- Download packages promptly; free-host restarts can remove temporary jobs before their stated expiry.
- Review every correction, note, warning, and `BLOCKING` row before publication.

## Main project files

- `app.py` — Flask routes, upload security, isolated jobs, and downloads
- `data/generate_transfer_sheet.py` — verified end-to-end workbook pipeline
- `generate_news_format.py` — PDF comparison and News Format construction
- `verify_news.py` — authoritative roster verification and masking
- `generate_transfer_document.py` — Word renderer and shared design/statistics logic
- `generate_transfer_pdf.py` — native ReportLab PDF renderer
- `templates/` and `static/` — responsive web interface and walkthrough assets
- `tests/` — synthetic unit/web tests and opt-in local integration test

