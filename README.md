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

- Required for current `/app` dashboard:
- `GOOGLE_SERVICE_ACCOUNT_JSON`
- `GOOGLE_CALENDAR_ID`
- `GOOGLE_TIMEZONE` (optional, default `Asia/Ho_Chi_Minh`)
- `SESSION_SECRET`
- `APP_USERS_JSON`

- Optional (only if you still use legacy export-to-sheet endpoints):
- `GOOGLE_SPREADSHEET_ID`
- `SHEET_GID`

- Optional:
- `APP_TOKEN`
- `SESSION_SECRET` (bat buoc khi deploy production)
- `SESSION_TTL_HOURS` (mac dinh `12`)
- `DEFAULT_SESSION_RATE` (mac dinh `300000`)
- `APP_USERS_JSON` (danh sach tai khoan login)
- `STUDENT_RATE_JSON` (don gia theo hoc vien)
- `DATABASE_URL` (strongly recommended on Render to avoid data loss after restart)

Example:

```env
SESSION_SECRET=replace_with_strong_secret
DEFAULT_SESSION_RATE=300000
APP_USERS_JSON=[{"username":"teacher","password":"yourStrongPwd","role":"teacher","displayName":"Co Trang"},{"username":"an.nguyen","password":"abc12345","role":"student","displayName":"An Nguyen","studentKey":"An Nguyen"}]
STUDENT_RATE_JSON={"An Nguyen":300000,"Binh Tran":280000}
```

## 4) Run locally

```bash
npm run dev
```

Open:

- `http://localhost:8080` (landing page)
- `http://localhost:8080/app` (dashboard login + quan ly hoc vien)
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

### Heroku quick steps

Note:

- Heroku no longer offers classic free dynos. You may need a paid dyno type (Eco/Basic) and billing enabled.

Steps:

1. Install Heroku CLI and login

```bash
heroku login
```

2. Create app

```bash
heroku create your-app-name
```

3. Set Node buildpack (usually auto-detected, but explicit is safer)

```bash
heroku buildpacks:set heroku/nodejs -a your-app-name
```

4. Add Config Vars in Heroku Dashboard (Settings > Config Vars)

- `GOOGLE_SERVICE_ACCOUNT_JSON`
- `GOOGLE_CALENDAR_ID`
- `GOOGLE_SPREADSHEET_ID`
- `GOOGLE_TIMEZONE`
- `SHEET_GID`
- `SESSION_SECRET`
- `SESSION_TTL_HOURS`
- `DEFAULT_SESSION_RATE`
- `APP_USERS_JSON`
- `STUDENT_RATE_JSON`
- optional `APP_TOKEN`

5. Deploy from git

```bash
git add .
git commit -m "prepare heroku deployment"
git push heroku main
```

6. Open app

```bash
heroku open -a your-app-name
```

7. Check logs if needed

```bash
heroku logs --tail -a your-app-name
```

Important for this project:

- `data/payments.json` is on local disk. On Heroku dyno restart/redeploy, local disk is ephemeral.
- For production-grade persistence, move payment states to a real DB (Postgres) or another persistent store.

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
- `POST /api/auth/login`
- `POST /api/auth/logout`
- `GET /api/auth/me`
- `GET /api/dashboard/month?month=MM/YYYY`
- `POST /api/payments/weekly` body: `{ "month": "08/2026", "studentKey": "an nguyen", "weekIndex": 2, "paid": true }`
- `POST /api/payments/monthly` body: `{ "month": "08/2026", "studentKey": "an nguyen", "paid": true }`
- `GET /api/payments/requests?month=MM/YYYY`
- `POST /api/payments/requests` body: `{ "month": "08/2026", "amount": 600000, "method": "bank_transfer", "note": "ck dot 1" }`
- `POST /api/payments/requests/:requestId/review` body: `{ "action": "approve" }` (teacher only)
- `POST /api/rates/update` body: `{ "studentKey": "an nguyen", "rate": 320000 }` (teacher only)
- `GET /api/config`
- `GET /api/status`
- `POST /api/export/weekly-current`
- `POST /api/export/month-current`
- `POST /api/export/month-custom` body: `{ "month": "04/2026" }`

## 7) Notes
## 7) Notes & Database Persistence

- Sheet naming follows: `TongKet_<year>_<month>`.
- Name matching uses Unicode normalization to avoid Win/Mac accent-encoding mismatches.
- Existing fee column is preserved; export writes summary columns only.
- Payment states are stored in `data/payments.json` (auto-created).
- If `DATABASE_URL` is set, payment data is stored in Postgres (`app_kv_store`) and survives restarts.
- Teacher can mark payment by week or by full month directly on `/app`.
- Student role is read-only and only sees their own row.

## 8) Notes

- Publish version v0.0.1
- Source of truth for schedule: Google Calendar.
- Payment states and custom rates are stored in PostgreSQL (`app_kv_store` và `student_rates`) when `DATABASE_URL` is set.
- A local file store (`data/payments.json`) acts as an immediate fallback & backup store if PostgreSQL is temporarily unavailable.
- Khi PostgreSQL kết nối lại, hệ thống tự động đồng bộ fallback rates vào database để đảm bảo không bị nhảy về đơn giá mặc định (300k).
- **Lưu ý triển khai trên Render**:
  - Nên dùng **Internal Database URL** (thay vì External URL) và đảm bảo Web Service và PostgreSQL cùng chung một Region (ví dụ `Singapore`).
  - Internal Database URL có độ trễ cực thấp, kết nối ổn định không qua Internet công cộng và không tốn chi phí băng thông ngoại mạng.
- Giáo viên có thể tick thu học phí theo từng tuần (V/X) hoặc chốt thu cả tháng trực tiếp trên `/app`.
- Học viên đăng nhập chỉ xem được thông tin và trạng thái của chính mình.