# PLT Process

Hệ thống Web tích hợp các công cụ hỗ trợ xử lý dữ liệu kỹ thuật và tự động hóa báo cáo chuyên dụng.

## Tính năng chính
- **Phân tích KEW**: Xử lý dữ liệu điện năng từ máy đo Kyoritsu, phát hiện sự cố PQ và nội suy dữ liệu.
- **Chỉnh sửa Ảnh**: Thay đổi thông số trực tiếp trên ảnh `.BMP`, xử lý hàng loạt và đồng bộ thời gian.
- **Tự động hóa Excel**: Đọc dữ liệu thô và tự động điền vào các biểu mẫu báo cáo Excel.
- **Xuất Báo cáo Word**: Tự động tổng hợp dữ liệu và sinh báo cáo Word dựa trên template.

## Cài đặt & Khởi chạy
1. **Cài đặt:**
   ```bash
   pip install -r requirements.txt
   ```
2. **Chạy Server:**
   ```bash
   python app.py
   ```
3. **Truy cập:** `http://localhost:5525`

## Cấu trúc dự án
- `app.py`: Flask entry point.
- `modules/`: Chứa các module xử lý chính (Excel, Image, KEW, Report).
- `static/` & `templates/`: Giao diện Dashboard và tài liệu mẫu (Word, Excel).
- `config.json`: Tham số cấu hình phân tích.

---
*Mặc định chạy trên cổng `5525`. Có thể tùy chỉnh qua biến môi trường `PORT`.*
