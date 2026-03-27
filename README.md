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

## ⚙️ Cấu Trúc Dự Án
- `modules/`: Chứa các Blueprint xử lý Backend (KEW, Excel).
- `static/js/`: Các module xử lý logic Frontend (Charts, Image Editor, Excel).
- `templates/`: Giao diện Dashboard (Bootstrap 5).
- `utils/`: Các hàm tiện ích dùng chung.
