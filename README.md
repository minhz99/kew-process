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
- Thay thế chỉ số (Pixel Replacement) và dán thời gian theo toạ độ bằng xử lý server-side cho ảnh `.BMP`.
- Hỗ trợ cộng thời gian giữa các ảnh theo khoảng ngẫu nhiên `m-n` giây.
- Hỗ trợ nhiều mẫu đồng hồ (Template) khác nhau như Kyoritsu KEW 6315, 6305, Hioki...
- Xử lý hàng loạt và đóng gói file ZIP sau khi sửa.

## 3. 📝 Xử lý Excel số điện (Excel Handler)
Công cụ tự động hóa việc nhập liệu và xử lý tệp Excel báo cáo số điện.
- Đọc dữ liệu từ text thô (String mode) hoặc nhập thủ công.
- Tự động tìm kiếm và ghi đè giá trị vào đúng dòng, cột trong file Excel báo cáo.
- Nếu chưa tải file mẫu, hệ thống tự dùng file mặc định `static/excel-template/excel-so-dien.xlsx`.
- Hỗ trợ quản lý lịch sử nhập liệu trong phiên làm việc.



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

## ⚙️ Cấu Trúc Dự Án
- `app.py`: Flask app entry, đăng ký Blueprint.
- `config.json`: Tham số phân tích dùng chung (Ornstein-Uhlenbeck, ngưỡng đánh giá).
- `modules/`: Backend (mỗi tool một sub-package).
  - `modules/kew/`: phân tích / hiệu chỉnh / nội suy KEW, tổ chức hồ sơ hiện trường.
  - `modules/image/`: chỉnh sửa ảnh BMP máy đo (`image_api.py`, `kew6315_layout.py`).
  - `modules/excel/`: cập nhật file Excel số điện.
  - `modules/report/`: sinh báo cáo Word (`gen_word.py` + `context_keys.json` mô tả khoá template).
- `static/js/`: Frontend logic (KEW charts, Image Editor, Excel Handler).
- `static/{css,digits,time-digits,excel-template,word-template}/`: tài nguyên tĩnh.
- `templates/dashboard.html`: Layout tổng (shell) của dashboard.
- `templates/components/{layout,modals,scripts,workspaces}/`: Thành phần dùng chung & mỗi tool là một workspace riêng (`kew.html`, `image.html`, `excel.html`).
- `utils/`: Các hàm tiện ích dùng chung (file/zip handling).
