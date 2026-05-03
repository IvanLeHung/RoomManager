# Quản lý phòng trọ

Web app cá nhân để quản lý phòng, người trọ, phiếu thu theo tháng và QR chuyển khoản VietQR.

## Chạy local

```bash
npm install
npm run dev
```

Mở link Vite hiện ra, thường là `http://localhost:5173`.

PIN mặc định: `1234`

## Build deploy

```bash
npm run build
npm run preview
```

Thư mục deploy: `dist`

## Lưu ý dữ liệu

Bản này lưu dữ liệu bằng `localStorage`, nên dữ liệu nằm trong trình duyệt của máy đang dùng. Nếu muốn đồng bộ nhiều thiết bị, nên nâng cấp thêm Supabase/Firebase.
