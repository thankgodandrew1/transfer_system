# Free deployment: GitHub + Render

This project deploys through Render's native Python runtime and the included `render.yaml`. It does not use Docker.

## 1. Verify the project locally

From PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
python -m pytest -q
python app.py
```

Open <http://127.0.0.1:5000>. Confirm that the generator loads, the **How to use** video plays, and the GitHub links open the correct repository. Stop the server with `Ctrl+C`.

## 2. Push to the existing GitHub repository

The project repository is [thankgodandrew1/transfer_system](https://github.com/thankgodandrew1/transfer_system). Keep it private unless you intentionally want the source public.

Before committing, confirm confidential files are ignored:

```powershell
git status --short --ignored
git check-ignore data/current_transfer_management.pdf data/current_transfer_report.xlsx data/old_transfer_report.xlsx
```

Then commit and push:

```powershell
git add .
git status --short
git commit -m "Match generated Transfer News to publication design"
git push -u origin main
```

Inspect `git status --short` before committing. No real `.pdf`, `.xlsx`, `.csv`, generated output, log, or job file should be staged. The safe template, captions, poster, and walkthrough MP4 are expected.

## 3. Create the free Render service

1. Create or sign in to a [Render account](https://dashboard.render.com/).
2. Select **New -> Blueprint**.
3. Connect GitHub and select `thankgodandrew1/transfer_system`.
4. Render reads `render.yaml` and creates the free native Python web service.
5. When prompted, set `APP_ACCESS_KEY` to a long shared key of at least 12 characters. Use a passphrase, not a common password.
6. Apply the Blueprint and wait for the first build.

The repository URL and the `ThankGod Andrew` creator credit are already configured in `render.yaml`. The Blueprint also configures:

- `pip install -r requirements.txt`
- `gunicorn app:app --workers 1 --threads 2 --timeout 300`
- `/health` health checks
- HTTPS-only session cookies
- mandatory access-key enforcement
- one small free web service

## 4. Verify the live deployment

1. Open the new `onrender.com` URL.
2. Confirm the access-key page appears. If the service will not start, check that `APP_ACCESS_KEY` is at least 12 characters.
3. Enter the key and open **How to use**.
4. Confirm the video plays, the favicon appears, and the GitHub links open `thankgodandrew1/transfer_system`.
5. Run one sanitized test package before uploading real records.
6. Open `/health`; it should return JSON with `"status": "ok"`.

## 5. Understand the free-tier limits

Render's free web services can sleep after inactivity, use ephemeral storage, and have free-tier usage limits. The first request after sleep can take longer. Check [Render's free-service documentation](https://render.com/docs/free) for the current limits.

Because storage is ephemeral:

- Download output immediately after generation.
- Do not treat the app as an archive.
- A restart or redeploy can remove jobs before their normal 60-minute expiry.

## 6. Deploy updates

After testing a change locally:

```powershell
git add .
git status --short
git commit -m "Describe the change"
git push
```

Render automatically deploys the pushed commit. Watch the deploy log, then repeat the live health and sanitized-generation checks.

## Troubleshooting

- **Build fails while installing packages:** check the Render build log and confirm the service uses Python 3.12 and `requirements.txt`.
- **Service refuses to start:** set `APP_ACCESS_KEY` to at least 12 characters. Production intentionally fails closed.
- **CSRF/form expired:** refresh the page and resubmit; do not reuse an old form after a restart.
- **Generation is busy:** the free service runs one generation at a time; wait for the active run to finish.
- **Files disappeared:** free storage is temporary. Generate again and download promptly.
- **First request is slow:** the free service is waking from its idle state.
