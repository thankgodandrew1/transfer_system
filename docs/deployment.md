# Free deployment: GitHub + Render

This deployment uses Render’s native Python runtime and the repository’s `render.yaml`. It does not use Docker.

## 1. Verify the project locally

From PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
python -m pytest -q
python app.py
```

Open <http://127.0.0.1:5000>, confirm the home page and **How to use** video, then stop the server with `Ctrl+C`.

## 2. Create a GitHub repository

1. Sign in to GitHub and create a new repository, for example `transfer-news-generator`.
2. Choose **Private** unless you intentionally want the source public.
3. Do not initialize it with a README because this project already contains one.
4. Copy the repository URL, such as `https://github.com/YOUR-ID/transfer-news-generator`.

Before committing, confirm confidential files are ignored:

```powershell
git status --short --ignored
git check-ignore data/current_transfer_management.pdf data/current_transfer_report.xlsx data/old_transfer_report.xlsx
```

Then commit and push:

```powershell
git add .
git status --short
git commit -m "Convert Transfer News generator to web app"
git branch -M main
git remote add origin https://github.com/YOUR-ID/transfer-news-generator.git
git push -u origin main
```

Inspect `git status --short` before committing. No real `.pdf`, `.xlsx`, `.csv`, output, log, or job file should be staged. The safe `templates/template.docx`, captions, poster, and walkthrough MP4 are expected.

## 3. Create the free Render service

1. Create or sign in to a [Render account](https://dashboard.render.com/).
2. Select **New → Blueprint**.
3. Connect GitHub and select the new repository.
4. Render reads `render.yaml` and creates `transfer-news-generator` as a free native Python web service.
5. When prompted for environment values, set:
   - `APP_ACCESS_KEY` — a long shared key of at least 12 characters; use a passphrase, not a common password.
   - `GITHUB_URL` — the repository or GitHub profile URL you want displayed in the site.
6. Apply the Blueprint and wait for the first build.

The Blueprint already configures:

- `pip install -r requirements.txt`
- `gunicorn app:app --workers 1 --threads 2 --timeout 300`
- `/health` health checks
- HTTPS-only session cookies
- mandatory access-key enforcement
- one small free web service

## 4. Verify the live deployment

1. Open the `onrender.com` URL.
2. Confirm the access-key page appears. If the service will not start, check that `APP_ACCESS_KEY` is at least 12 characters.
3. Enter the access key and open **How to use**.
4. Confirm the video plays and the GitHub link opens the expected URL.
5. Run a sanitized test package before uploading real records.
6. Open `/health`; it should return JSON with `"status": "ok"`.

## 5. Understand the free-tier limits

Render documents that free web services spin down after 15 minutes without inbound traffic, use ephemeral storage, and share a monthly free instance-hour allowance. The first request after sleep can therefore take longer. See [Render’s free-service documentation](https://render.com/docs/free) and [instance types](https://render.com/docs/compute-plans) for the current limits.

Because storage is ephemeral:

- Download output immediately after generation.
- Do not treat the app as an archive.
- A restart or redeploy can remove jobs before their normal 60-minute expiry.

## 6. Deploy updates

After testing a change locally:

```powershell
git add .
git commit -m "Describe the change"
git push
```

Render automatically deploys the pushed commit. Watch the deploy log, then repeat the live health and sanitized-generation checks.

## Troubleshooting

- **Build fails while installing packages:** open the Render build log and confirm the service uses Python 3.12 and `requirements.txt`.
- **Service refuses to start:** set `APP_ACCESS_KEY` to at least 12 characters. Production intentionally fails closed.
- **CSRF/form expired message:** refresh the page and resubmit. Do not reuse an old form after a restart.
- **Generation is busy:** the free service runs one generation at a time to avoid cross-run interference; wait for the active run to finish.
- **Files disappeared:** free storage is temporary. Generate again and download promptly.
- **First page is slow:** the free service is waking from its idle state.

