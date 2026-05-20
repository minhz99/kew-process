# Hướng Dẫn File Excel Hiện Trường Cho Xử Lý KEW

Tài liệu này mô tả riêng file Excel hiện trường đang được code đọc trong phần **Xử lý KEW**. Hướng dẫn đầy đủ theo từng nút giao diện nằm tại `docs/huong-dan-xu-ly-kew.md`.

## Định Dạng File

- Nên dùng `.xlsx` hoặc `.xlsm`.
- Dữ liệu đặt ở sheet đầu tiên/active sheet.
- Dòng 1 là header.
- Tên cột không phân biệt hoa thường, nhưng nên viết đúng như bên dưới.

## Cột Bắt Buộc

| Cột | Mô tả | Ví dụ |
| --- | --- | --- |
| `stt` | Thứ tự thiết bị trong báo cáo/bảng tổng hợp | `1` |
| `name` | Tên thiết bị, cũng dùng để đặt tên thư mục sau xử lý | `MBA T1` |
| `file` | Mã thư mục record KEW; chấp nhận `1`, `0001`, `S0001` | `S0001` |
| `img` | Số ảnh BMP bắt đầu | `641` |
| `imgend` | Số ảnh BMP kết thúc | `646` |
| `type` | Loại section: `MBA`, `4`/`device4`, hoặc `device` | `MBA` |
| `pdm` | Công suất định mức kVA | `1600` |
| `current_char` | Đặc tính dòng điện | `ổn định`, `dao động nhẹ`, `biến đổi liên tục`, `load/unload` |

## Cột Tùy Chọn

| Cột | Mô tả | Ví dụ |
| --- | --- | --- |
| `imgomit` | Ảnh cần loại khỏi dải `img`-`imgend` | `944,945` |
| `imglu` | Ảnh load/unload riêng, được copy thành `load-unload-xxx.BMP` | `647` |

## Cột OCR/Thông Số Đo

Các cột này có thể để trống. Khi chạy **Xử lý file sơ bộ**, OCR sẽ tạo và điền nếu đọc được ảnh; mặc định không ghi đè ô đã có dữ liệu.

| Cột | Mô tả |
| --- | --- |
| `p` | Công suất tác dụng kW |
| `cos_phi` | Hệ số công suất; nên nhập dạng `0.987` |
| `i_max` | Dòng điện lớn nhất A |
| `u_min` | Điện áp nhỏ nhất V |
| `u_max` | Điện áp lớn nhất V |
| `delta_u` | Độ lệch/mất cân bằng điện áp % |
| `delta_i` | Độ lệch/mất cân bằng dòng điện % |
| `thd` | THD điện áp lớn nhất % |
| `tdd` | TDD dòng điện lớn nhất % |

## Lưu Ý Quan Trọng

- Schema hiện tại dùng `cos_phi`, `i_max`, `delta_i`; không dùng các tên cũ như `pf`, `i1`, `i2`, `i3`, `di` cho luồng chính.
- `file` phải trỏ tới một thư mục record tồn tại trong ZIP, ví dụ `S0001`.
- `img` đến `imgend` là dải ảnh bao gồm cả hai đầu.
- Dải ảnh giữa các thiết bị không được chồng lấn.
- `imgomit` chỉ nên chứa số ảnh nằm trong dải của chính thiết bị đó.
- Mỗi thư mục thiết bị cần có `a.png` để tạo báo cáo Word (nếu không có, hệ thống sẽ tự động dùng ảnh mặc định tại `static/word-template/a.png`).
