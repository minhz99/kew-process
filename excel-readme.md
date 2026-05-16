# Hướng dẫn File Excel Hiện trường (Field Excel)

File Excel hiện trường là "linh hồn" của hệ thống, dùng để tổ chức hồ sơ từ máy đo và cung cấp dữ liệu đầu vào để tự động sinh nhận xét trong báo cáo Word.

## 📋 Cấu trúc các cột (Bắt buộc)
File Excel cần có đúng **18 cột** sau (tên cột không phân biệt hoa thường):

| STT | Tên cột | Mô tả | Ghi chú |
|:---:|:---|:---|:---|
| 1 | `stt` | Số thứ tự | Thứ tự xuất hiện của thiết bị trong báo cáo Word. |
| 2 | `name` | Tên thiết bị | Tên hiển thị (VD: MBA T1, Tủ MSB...). |
| 3 | `file` | Mã thư mục | Mã `Sxxxx` tương ứng trên máy đo (VD: 0001 hoặc S0001). |
| 4 | `img` | Ảnh đầu | Số hiệu ảnh BMP bắt đầu (VD: 641). |
| 5 | `imgend` | Ảnh cuối | Số hiệu ảnh BMP kết thúc (VD: 646). |
| 6 | `imgomit` | Ảnh bỏ qua | Các ảnh lỗi trong dải cần bỏ qua (VD: `642, 644`). |
| 7 | `type` | Loại | `MBA` (Máy biến áp), `4` (Thiết bị Chương 4), hoặc `device` (mặc định). |
| 8 | `pdm` | P định mức | Công suất định mức (kVA) của MBA hoặc thiết bị. |
| 9 | `current_char` | Đặc tính tải | "Ổn định", "Dao động nhẹ", "Biến đổi liên tục", hoặc "Chu kỳ Load-Unload". |
| 10 | `u_min` | U min (V) | Điện áp thấp nhất đo được. |
| 11 | `u_max` | U max (V) | Điện áp cao nhất đo được. |
| 12 | `i_max` | I max (A) | Dòng điện lớn nhất đo được. |
| 13 | `delta_u` | ΔU (%) | Độ lệch pha (mất cân bằng) điện áp lớn nhất. |
| 14 | `delta_i` | ΔI (%) | Độ lệch pha (mất cân bằng) dòng điện lớn nhất. |
| 15 | `p` | P (kW) | Công suất tác dụng trung bình. |
| 16 | `cos_phi` | cosφ | Hệ số công suất trung bình. |
| 17 | `thd` | THD (%) | Tổng biến dạng sóng hài điện áp lớn nhất. |
| 18 | `tdd` | TDD (%) | Tổng biến dạng sóng hài dòng điện lớn nhất. |

---

## 💡 Quy tắc nhập liệu quan trọng

### 1. Đặc tính dòng điện (`current_char`)
Từ khoá này quyết định cụm từ mô tả biểu đồ trong báo cáo Word:
- **Ổn định**: "Biểu đồ dòng điện tiêu thụ tại thời điểm đo kiểm ổn định."
- **Chu kỳ Load-Unload**: "Biểu đồ dòng điện tiêu thụ biến đổi theo chu kỳ Load/Unload."
- **Biến đổi liên tục**: "Biểu đồ dòng điện tiêu thụ biến đổi liên tục với biên độ nhỏ."

### 2. Thuật toán Đánh giá tự động (Loi_Dem)
Hệ thống tự động chấm điểm lỗi dựa trên các cột thông số:
- **ΔU > 5.0%**: Lỗi nghiêm trọng (+2 điểm).
- **Cosφ < 0.9**: +1 điểm.
- **ΔI > 10.0%**: +1 điểm.
- **THD > 8.0%**: +1 điểm.
- **TDD > 12.0%** (MBA) hoặc **20.0%** (Tải): +1 điểm.

**Kết quả đánh giá:**
- 0 điểm: **Tốt**
- 1 điểm: **Tương đối tốt**
- ≥ 2 điểm: **Chưa tốt** (hoặc "Chưa thực sự tốt" đối với MBA).

### 3. Nhận xét thủ công
Nếu bạn muốn tự viết nhận xét cho một thiết bị mà không muốn hệ thống sinh tự động:
- Tại cột bất kỳ (thường là cột ghi chú thêm), hãy bắt đầu bằng cụm từ **"Nhận xét:"**. 
- Hệ thống sẽ nhận diện và ưu tiên dùng nguyên văn đoạn văn đó vào báo cáo.

---

## 📁 Tổ chức hình ảnh
Hệ thống yêu cầu các file ảnh trong mỗi thư mục thiết bị sau khi xử lý:
- `a.png`: Ảnh tổng quan (Overview).
- `PS-SDxxx.BMP`: Các ảnh thông số (lấy theo dải `img` đến `imgend` trong Excel).

*Lưu ý: Nếu thiếu file `a.png` hoặc các file `.BMP` không đúng dải, hệ thống sẽ báo lỗi khi tạo báo cáo Word.*
