# TAIGA Center & CRM (calendar-sheet-manager)

Hệ thống tích hợp Website trung tâm tiếng Nga **TAIGA Center** và nền tảng quản lý học viên 1-1 (**TAIGA CRM**) tự động hóa dữ liệu từ Google Calendar.

---

## 📌 Tổng quan dự án (Project Overview)

Dự án giải quyết bài toán quản lý lớp học kèm 1-1, tính số buổi/thời lượng học và tính học phí cho giáo viên dạy kèm (cô Trang), thay thế hoàn toàn việc nhập liệu thủ công mỗi ngày vào bảng tính:

- **Nguồn dữ liệu gốc (Source of Truth):** Lịch dạy trên **Google Calendar**. Giáo viên chỉ cần lên lịch các buổi dạy (tiêu đề chứa tên học viên và thời lượng như `1h30`, `90p`...).
- **Tự động phân tích & Tính toán:** Backend tự động đọc sự kiện từ Google Calendar API qua Service Account, phân tích thời lượng, nhóm theo thứ trong tuần (T2 - CN), tổng hợp số giờ học và nhân với đơn giá (hourly rate) riêng của từng học viên.
- **Theo dõi học phí & Thu tiền:** Giáo viên có thể đánh dấu đã thanh toán theo từng tuần hoặc chốt thu cả tháng trực tiếp trên giao diện web.
- **Phân quyền người dùng (Role-based Auth):**
  - **Giáo viên (`teacher`):** Xem toàn bộ học viên, điều chỉnh đơn giá, quản lý trạng thái thanh toán, cấp tài khoản đăng nhập cho học sinh, ẩn/hiện cột đơn giá khi cần chụp màn hình đối soát.
  - **Học viên (`student`):** Đăng nhập bằng tài khoản được cấp và chỉ xem được lịch sử học, thời lượng, học phí và trạng thái thanh toán của chính mình.
- **Lưu trữ bền vững 2 lớp:** PostgreSQL kết hợp cơ chế tự động Fallback sang file JSON cục bộ (`data/payments.json`) nếu database mất kết nối, và tự đồng bộ ngược lại khi kết nối phục hồi.
- **Website quảng bá trung tâm:** Cung cấp trang Landing page giới thiệu trung tâm TAIGA Center, các khóa học tiếng Nga, bảng giá và form đăng ký tư vấn.

---

## 🌐 Các phân hệ & Đường dẫn chính (Routing)

| Đường dẫn | Phân hệ / Tính năng | Mô tả |
| :--- | :--- | :--- |
| `/home` | **Landing Page TAIGA Center** | Giới thiệu trung tâm, đội ngũ giáo viên, các khóa học tiếng Nga, biểu phí và form đăng ký. |
| `/app` | **TAIGA CRM Dashboard** | Cổng quản lý học viên: Đăng nhập (Teacher/Student), chọn tháng xem thống kê, đối soát buổi học, đổi đơn giá, cấp tài khoản học sinh. |
| `/learn` | **Tài liệu & Lộ trình** | Trang phụ giới thiệu tài liệu, lộ trình học tiếng Nga từ cơ bản đến nâng cao. |
| `/contact` | **Trang Liên hệ** | Thông tin liên hệ và tư vấn trực tiếp của trung tâm. |
| `/articles` | **Góc chia sẻ & Bài viết** | Các bài viết kinh nghiệm học tiếng Nga và văn hóa Nga. |
| `/` | **Control Room (Legacy)** | Công cụ xuất dữ liệu từ Google Calendar sang Google Sheets theo tuần/tháng. |

---

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

## 6) Danh sách API Endpoints

### Xác thực & Phiên làm việc (Auth)
- `POST /api/auth/login` - Đăng nhập hệ thống (set cookie `taiga_session`)
- `POST /api/auth/logout` - Đăng xuất và hủy phiên
- `GET /api/auth/me` - Lấy thông tin user hiện tại (Role, displayName, studentKey)

### Thống kê & Quản lý buổi học (CRM)
- `GET /api/dashboard/month?month=MM/YYYY` - Lấy dữ liệu buổi học, tổng giờ và học phí (Teacher xem tất cả, Student xem của chính mình)
- `POST /api/payments/weekly` - Đánh dấu đóng học phí theo tuần `{ month, studentKey, weekIndex, paid }` *(Teacher only)*
- `POST /api/payments/monthly` - Chốt trạng thái đóng học phí cả tháng `{ month, studentKey, paid }` *(Teacher only)*
- `POST /api/rates/update` - Đổi đơn giá theo giờ của học viên `{ studentKey, rate }` *(Teacher only)*

### Quản lý tài khoản học sinh (Student Accounts)
- `GET /api/student-accounts` - Xem danh sách tài khoản học sinh *(Teacher only)*
- `POST /api/student-accounts/upsert` - Cấp mới / đổi mật khẩu học sinh `{ username, password, studentKey, displayName }` *(Teacher only)*
- `DELETE /api/student-accounts/:studentKey` - Xóa tài khoản đăng nhập của học sinh *(Teacher only)*

### Hệ thống & Chẩn đoán
- `GET /api/health` - Kiểm tra tình trạng server
- `GET /api/db/diagnostics` - Kiểm tra kết nối PostgreSQL và tình trạng bảng dữ liệu
- `GET /api/config` - Lấy cấu hình runtime công khai (yêu cầu Token)
- `GET /api/status` - Xem trạng thái lần xuất dữ liệu gần nhất (yêu cầu Token)

### Thao tác sự kiện Calendar & Export Sheets (Legacy)
- `POST /api/export/weekly-current` - Xuất tuần hiện tại ra Google Sheet (yêu cầu Token)
- `POST /api/export/month-current` - Xuất cả tháng hiện tại ra Google Sheet (yêu cầu Token)
- `POST /api/export/month-custom` - Xuất tháng tùy chỉnh (`MM/YYYY`) ra Sheet (yêu cầu Token)
- `GET /api/events` - Đọc danh sách sự kiện từ Google Calendar
- `POST /api/events/create` - Tạo sự kiện mới lên Google Calendar
- `PUT /api/events/:eventId` - Chỉnh sửa sự kiện trên Calendar
- `DELETE /api/events/:eventId` - Xóa sự kiện trên Calendar

---

## 7) Cơ sở dữ liệu & Cơ chế Dự phòng (Database & Fallback Persistence)

- **Source of truth cho lịch học:** Google Calendar.
- **Cơ chế lưu trữ kép (Dual-layer persistence):**
  - **PostgreSQL:** Khi có `DATABASE_URL`, hệ thống tự tạo và quản lý:
    - `student_rates`: Lưu đơn giá tùy chỉnh theo từng học viên (`student_key`, `rate`, `updated_at`).
    - `app_kv_store`: Lưu trạng thái thanh toán tuần/tháng theo dạng key-value.
    - `student_accounts`: Lưu tài khoản đăng nhập của học sinh.
  - **Fallback JSON (`data/payments.json`):** Nếu PostgreSQL tạm thời bị ngắt kết nối (hoặc chưa cấu hình), hệ thống tự động lưu vào file JSON cục bộ để không gián đoạn công việc của giáo viên.
  - **Auto-Sync:** Khi PostgreSQL kết nối lại, hệ thống tự động đồng bộ fallback rates lên database để bảo toàn dữ liệu và không bị reset về đơn giá mặc định (300k).
- **Lưu ý triển khai trên Render:**
  - Nên dùng **Internal Database URL** (thay vì External URL) và đảm bảo Web Service và PostgreSQL cùng chung một Region (ví dụ `Singapore`).
  - Internal Database URL có độ trễ cực thấp (< 5ms), bảo mật nội bộ và không tính băng thông ra ngoài.