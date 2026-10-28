# Quy Tắc Dự Án & Chuẩn Phát Triển - Taiga Talk

## 1. PWA & Background Resume (QUY TẮC BẤT BIẾN)
- **Tuyệt đối KHÔNG màn xanh / màn che:** Không dùng bất kỳ overlay, full-screen color veil (`#003838`, v.v.) nào khi nạp hoặc mở app.
- **Mở lại từ chạy ngầm (Mobile Resume):** 100% in-memory sync (streak, XP, UI tree). Tuyệt đối KHÔNG reload trang, KHÔNG giật nháy UI.
- **Tự động cập nhật khi deploy:** Chỉ reload khi Service Worker phát hiện bản deploy mới và kích hoạt (`controllerchange`).

## 2. Header & Tab Navigation UI Rules
- **Header trên Mobile:** Logo cố định bên trái, Gamification Stats (`🔥`, `⚡`) bên phải góc trên; thanh Navigation nằm ngang bên dưới.
- **Active Tab / Selector Styling (Quy tắc thống nhất toàn app):**
  - Áp dụng đồng bộ cho: Header tabs (`.ff-nav-tab`), Trình độ (`.ff-level-btn`, `.level-btn`), Kỹ năng Nghe - Nói - Đọc - Viết (`.ff-skill-btn`, `.skill-btn`), Chuyên ngành (`.spec-btn`), Chặng 1/2 (`.stage-tab-btn`), và Danh mục bài viết (`.filter-chip`).
  - **KHÔNG** dùng viền bo tròn (pill border / outline).
  - **KHÔNG** dùng khối hộp nền màu đặc dày cộp (no heavy solid background boxes).
  - **ĐẶC TRƯNG ACTIVE:** Chữ in đậm (`font-weight: 800`) + màu xanh thương hiệu (`var(--ff-primary)`) + vầng sáng nhẹ Luminous Glow (`text-shadow: 0 0 12px rgba(39, 196, 181, 0.65), 0 0 22px rgba(39, 196, 181, 0.35)`).
- **Trang `/app`:** Header kéo full 100% màn hình; nút Đăng xuất ẩn hoàn toàn khi chưa đăng nhập.

## 3. Git Workflow
- Nhánh phát triển hiện tại: `feature/dictionary-widget`.
- Chỉ merge sang `dev_update-main-photo-home-page` khi có yêu cầu cụ thể từ người dùng; không merge trực tiếp vào `main`.
