# PLT Process

Đây là hệ thống Web tích hợp các công cụ hỗ trợ xử lý công việc chuyên dụng của tôi. Hiện tại, hệ thống bao gồm 3 công cụ chính:

## 1. 📊 Xử lý file .KEW (KEW Analyzer)
Công cụ phân tích dữ liệu điện năng từ máy đo Kyoritsu (file `.KEW` hoặc tệp `.ZIP`).
- Phân tích thông số: Apparent Power, Voltage, Current, THD...
- Phát hiện sự kiện PQ: Dip, Swell, Inrush, Transient.
- Tự động nội suy (Fix) dữ liệu cho các pha bị khuyết.
- Đánh giá chất lượng điện theo tiêu chuẩn IEEE 519.

## 2. 📸 Xử lý ảnh đo (Image Editor)
Công cụ chỉnh sửa thông số hiển thị trên ảnh chụp màn hình máy đo (file `.BMP`).
- Thay thế chỉ số (Pixel Replacement) trực tiếp trên trình duyệt bằng HTML5 Canvas.
- Hỗ trợ nhiều mẫu đồng hồ (Template) khác nhau như Kyoritsu KEW 6315, 6305, Hioki...
- Xử lý hàng loạt và đóng gói file ZIP sau khi sửa.

## 3. 📝 Xử lý Excel số điện (Excel Handler)
Công cụ tự động hóa việc nhập liệu và xử lý tệp Excel báo cáo số điện.
- Đọc dữ liệu từ text thô (String mode) hoặc nhập thủ công.
- Tự động tìm kiếm và ghi đè giá trị vào đúng dòng, cột trong file Excel báo cáo.
- Hỗ trợ quản lý lịch sử nhập liệu trong phiên làm việc.

## 4. 📝 Báo cáo KEW Word (Synopex)
Công cụ tạo báo cáo Word từ bộ ảnh đo KEW đóng gói trong file `.zip`.
- Luồng xử lý: upload ZIP -> server giải nén và đọc ảnh -> server sinh file `.docx` -> tải file về.
- Không cần upload file mẫu Word hay cấu hình `Tesseract path` từ phía client.
- Nếu server không có file mẫu cấu hình sẵn, hệ thống sẽ tự dùng template Word tối thiểu tích hợp.

---

## 🚀 Cài Đặt & Khởi Chạy

### Yêu cầu hệ thống
- Python 3.8 trở lên.

### Các bước cài đặt
1. Cài đặt các thư viện cần thiết:
   ```bash
   pip install -r requirements.txt
   ```
2. Khởi chạy Server:
   ```bash
   python3 app.py
   ```
3. Truy cập Dashboard tại: `http://localhost:5525`

### Biến môi trường hữu ích khi deploy
- `HOST`: host bind của Flask app. Mặc định `0.0.0.0`.
- `PORT`: cổng chạy app. Mặc định `5525`.
- `FLASK_DEBUG`: bật debug khi cần (`1`, `true`, `yes`, `on`).
- `MAX_UPLOAD_MB`: giới hạn dung lượng upload. Mặc định `256`.
- `SYNOPEX_TEMPLATE_FILE`: đường dẫn tới file mẫu `.docx` nếu muốn dùng mẫu server-side riêng.
- `TESSERACT_CMD`: đường dẫn Tesseract server-side nếu muốn bật fallback OCR bằng Tesseract.

## ⚙️ Cấu Trúc Dự Án
- `modules/`: Chứa các Blueprint xử lý Backend (KEW, Excel, Image).
- `static/js/`: Các module xử lý logic Frontend (KEW charts, Image Editor, Excel Handler).
- `templates/dashboard.html`: Layout tổng (shell) của dashboard.
- `templates/components/layout/`: Thành phần layout dùng chung (sidebar, ...).
- `templates/components/workspaces/`: Mỗi tool là một component/template riêng (`kew.html`, `image.html`, `excel.html`).
- `templates/components/modals/`: Các modal dùng chung.
- `templates/components/scripts/`: Script điều hướng workspace và bundle script.
- `utils/`: Các hàm tiện ích dùng chung.
