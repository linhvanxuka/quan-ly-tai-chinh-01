# SỔ QUẢN LÝ TÀI CHÍNH GIA ĐÌNH • HÒA & QUỲNH (MỤC TIÊU 27.000k)

Ứng dụng web quản lý tài chính cá nhân đồng bộ 2 chiều với Google Sheets.
- **Đơn vị tính:** nghìn đồng (`k VNĐ`). Ví dụ: Bát phở nhập `45` (45k), Tiền điện `2500` (2,5 triệu).
- **Hạn mức kế hoạch:** 27.000k (Vợ tự chủ: 19.440k | Chồng & ăn ngoài: 7.560k).

---

## CẤU TRÚC THƯ MỤC TRIỂN KHAI VERCEL:

```text
├── api/
│   └── index.py            # Vercel Serverless Function (Python Flask)
├── templates/
│   └── index.html          # Giao diện Web Responsive PWA (Mobile & Desktop)
├── google_apps_script.js   # Mã script dán vào Google Sheets để Vercel kết nối
├── requirements.txt        # Các thư viện Python cần thiết
├── vercel.json             # File cấu hình định tuyến Vercel
├── start_app.sh            # File chạy nhanh trên máy Mac ở nhà
└── README.md               # Tài liệu hướng dẫn
```

---

## 3 BƯỚC ĐẨY LÊN VERCEL DÙNG VĨNH VIỄN TRÊN ĐIỆN THOẠI:

### BƯỚC 1: ĐẨY TOÀN BỘ CODE LÊN GITHUB
1. Mở Terminal tại thư mục `app` này.
2. Khởi tạo Git và đẩy lên Repository mới trên GitHub của anh:
   ```bash
   git init
   git add .
   git commit -m "Khoi tao App Quan Ly Tai Chinh Hoa Quynh"
   git branch -M main
   git remote add origin https://github.com/TÊN_GITHUB_CỦA_ANH/TÊN_REPO.git
   git push -u origin main
   ```

### BƯỚC 2: TẠO WEBHOOK GOOGLE APPS SCRIPT
1. Mở trang Google Sheets: [Chi tiêu tháng 9, 10, 11](https://docs.google.com/spreadsheets/d/1OdHwBFJ5L5JO5_Q_pk-_F_acBHUbpMsdG0fd7pyb_HA/edit).
2. Chọn menu: **Tiện ích mở rộng (Extensions)** > **Apps Script**.
3. Copy toàn bộ nội dung file `google_apps_script.js` dán vào và bấm **Lưu 💾**.
4. Bấm nút màu xanh **Triển khai (Deploy)** góc phải trên > **Tùy chọn triển khai mới (New deployment)**.
5. Chọn loại **Ứng dụng web (Web app)**:
   * Ai có quyền truy cập (Who has access): Chọn **Bất kỳ ai (Anyone)**.
6. Bấm **Triển khai (Deploy)** và copy đường link URL Web App (dạng `https://script.google.com/macros/s/.../exec`).

### BƯỚC 3: DEPLOY TRÊN VERCEL
1. Truy cập [vercel.com](https://vercel.com) > Bấm **Add New...** > **Project**.
2. Chọn Repository GitHub vừa tạo ở Bước 1 > Bấm **Import**.
3. Ở mục **Environment Variables** (Biến môi trường), thêm 1 biến:
   * **Key:** `GSHEETS_URL`
   * **Value:** *(Dán đường link Web App vừa copy ở Bước 2)*
4. Bấm **Deploy**!
Chỉ sau 30 giây, Vercel sẽ cấp cho anh link vĩnh viễn (ví dụ `https://tai-chinh-hoa-quynh.vercel.app`) để lưu vào màn hình chính điện thoại dùng mọi lúc mọi nơi!
