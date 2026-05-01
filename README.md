# Calendar Sheet Manager

Single-page web app to manage your Google Calendar and export results to Google Sheets with the same flow you already use:

- Weekly incremental export (current month logic)
- Full export for current month
- Full export for a custom month (`MM/YYYY`)

Data remains in Google services:

- Source of truth for schedule: Google Calendar
- Output/report: Google Sheets

## 1) Prerequisites

- Node.js 20+
- A Google Cloud project with:
  - Calendar API enabled
  - Sheets API enabled
- A service account JSON key

## 2) Permissions

Share your resources with the service account email:

- Google Calendar: at least `See all event details`
- Google Spreadsheet: `Editor`

## 3) Setup

```bash
npm install
cp .env.example .env
```

Fill `.env` with:

- `GOOGLE_SERVICE_ACCOUNT_JSON`
- `GOOGLE_CALENDAR_ID`
- `GOOGLE_SPREADSHEET_ID`
- `GOOGLE_TIMEZONE`
- optional `APP_TOKEN`

## 4) Run locally

```bash
npm run dev
```

Open:

- `http://localhost:8080`
- If `APP_TOKEN` is set, use: `http://localhost:8080/?token=YOUR_TOKEN`

## 5) Deploy always-on

Recommended:

- Backend + frontend together on Railway/Render/Fly.io
- Keep minimum instance > 0 (no sleep)

### Railway quick steps

1. Push this folder to GitHub
2. New Railway project from repo
3. Add environment variables from `.env`
4. Start command: `npm start`
5. Keep service always on

### GitHub Actions -> VPS (auto deploy)

This project includes:

- Workflow: [.github/workflows/deploy-vps.yml](.github/workflows/deploy-vps.yml)
- Remote deploy script: [scripts/deploy_remote.sh](scripts/deploy_remote.sh)

On every push to `main` (or manual run), GitHub Actions will:

1. Sync source code to your VPS path
2. Install production dependencies
3. Restart app (`pm2` if available, otherwise `nohup node`)
4. Run local health check on VPS

Required GitHub repository `Secrets`:

- `VPS_HOST`: example `103.179.173.39`
- `VPS_PORT`: example `22`
- `VPS_USER`: example `root`
- `VPS_APP_DIR`: example `/root/russweb/calendar-sheet`
- `VPS_SSH_KEY`: private key content (recommended deploy key)

Optional GitHub repository `Variables`:

- `APP_NAME`: default `calendar-sheet-manager`
- `APP_PORT`: default `8080`

Important:

- Keep your runtime `.env` only on VPS. Do not commit `.env`.
- If you currently use password SSH only, switch to SSH key auth first for CI/CD.

## 6) API endpoints

- `GET /api/health`
- `GET /api/config`
- `GET /api/status`
- `POST /api/export/weekly-current`
- `POST /api/export/month-current`
- `POST /api/export/month-custom` body: `{ "month": "04/2026" }`

## 7) Notes

- Sheet naming follows: `TongKet_<year>_<month>`.
- Name matching uses Unicode normalization to avoid Win/Mac accent-encoding mismatches.
- Existing fee column is preserved; export writes summary columns only.

## 8) Notes

- Publish version v0.0.1