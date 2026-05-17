# Hướng dẫn và cấu trúc phần Xử lý KEW

Tài liệu này tổng hợp theo code hiện tại của dự án `plt-process`, tập trung vào tab **Xử lý KEW** trên giao diện.

## 1. Cấu trúc chức năng trong code

### Giao diện

- `templates/components/workspaces/kew.html`
  - Subtab **Xử lý file sơ bộ**.
  - Subtab **Tạo báo cáo Word** gồm 4 nút: **Tạo chương 4**, **Tạo chương 5**, **Excel MBA**, **Tạo Bảng 6.3**.

### API Flask

- `POST /api/kew/organize-field-zip`: sắp xếp hồ sơ hiện trường và trả về `KEW_HoSoDaXuLy.zip`.
- `POST /api/kew/generate-chapter4`: sinh file Word Chương 4.
- `POST /api/kew/generate-chapter5`: sinh file Word Chương 5.
- `POST /api/kew/generate-excel-mba`: sinh file Excel MBA.
- `POST /api/kew/generate-table6`: sinh Word Bảng 6.3.
- `POST /api/kew/generate-word-report`: endpoint tổng hợp cũ, hiện không có nút trên giao diện.

### Module xử lý

- `modules/kew/organize_field_zip.py`: đọc Excel hiện trường, map thư mục `Sxxxx`, copy ảnh BMP, chạy OCR, tạo `Project_Output`.
- `modules/image/ocr_kew.py`: OCR các ảnh màn hình KEW6315 để điền các cột thông số vào Excel.
- `modules/kew/analyse_kew.py`: parse file `INPSxxxx.KEW` và tổng hợp thông số.
- `modules/report/gen_word.py`: đọc ZIP đã sắp xếp, lấy metadata Excel, chọn ảnh, render/merge Word.
- `modules/report/gen_excel_mba.py`: đọc `INPSxxxx.KEW`, điền vào template `MBA.xlsm`.

### Template và tài nguyên bắt buộc của dự án

- Word templates:
  - `static/word-template/mba.docx`
  - `static/word-template/device.docx`
  - `static/word-template/device4.docx`
  - `static/word-template/totalmba.docx`
  - `static/word-template/table6.docx`
- Excel template:
  - `static/excel-template/MBA.xlsm`
  - Template này cần có sheet `MBA1`; nếu có sheet `Tổn thất MBA` thì hệ thống sẽ điền bảng tổng hợp MBA vào sheet đó.
- OCR digits:
  - `static/digits/*.bmp`

## 2. Excel hiện trường: cấu trúc đúng với code hiện tại

File Excel nên là `.xlsx` hoặc `.xlsm`, dữ liệu đặt ở sheet đầu tiên/active sheet, dòng đầu tiên là header.

Tên cột được so khớp không phân biệt hoa thường và bỏ khoảng trắng đầu/cuối, nhưng nên dùng đúng tên sau đây để tránh lệch schema.

### Cột bắt buộc

| Cột | Ý nghĩa | Ví dụ |
| --- | --- | --- |
| `stt` | Thứ tự thiết bị trong báo cáo/bảng tổng hợp | `1` |
| `name` | Tên hiển thị của thiết bị; cũng là tên thư mục sau khi sắp xếp | `MBA T1` |
| `file` | Mã thư mục record trên máy đo; chấp nhận `1`, `0001`, `S0001` | `S0001` |
| `img` | Số ảnh BMP đầu tiên trong dải ảnh của thiết bị | `641` |
| `imgend` | Số ảnh BMP cuối cùng trong dải ảnh của thiết bị | `646` |
| `type` | Loại section: `MBA`, `4`/`device4`, hoặc `device` | `MBA` |
| `pdm` | Công suất định mức kVA, dùng cho MBA và bảng tổng hợp MBA | `1600` |
| `current_char` | Đặc tính dòng điện để sinh nhận xét | `ổn định`, `dao động nhẹ`, `biến đổi liên tục`, `load/unload` |

### Cột tùy chọn

| Cột | Ý nghĩa | Ví dụ |
| --- | --- | --- |
| `imgomit` | Danh sách ảnh cần bỏ qua trong dải `img` - `imgend` | `944,945` hoặc `PS-SD944.BMP` |
| `imglu` | Ảnh riêng cho load/unload; được copy thành `load-unload-xxx.BMP` | `647` |

### Cột thông số đo/OCR

Nếu các cột này chưa có, bước **Xử lý file sơ bộ** sẽ tự tạo và điền bằng OCR khi đọc được ảnh. Nếu cột đã có giá trị, mặc định không ghi đè.

| Cột | Cách dùng |
| --- | --- |
| `p` | Công suất tác dụng kW |
| `cos_phi` | Hệ số công suất; nên nhập dạng `0.987` |
| `i_max` | Dòng điện lớn nhất A |
| `u_min` | Điện áp nhỏ nhất V |
| `u_max` | Điện áp lớn nhất V |
| `delta_u` | Mất cân bằng/lệch pha điện áp % |
| `delta_i` | Mất cân bằng/lệch pha dòng điện % |
| `thd` | THD điện áp lớn nhất % |
| `tdd` | TDD dòng điện lớn nhất % |

Lưu ý: Một số mô tả cũ trong giao diện/readme nhắc đến `pf`, `i1`, `i2`, `i3`, `di`. Code hiện tại không đọc các cột đó cho nhận xét chính; schema đúng là `cos_phi`, `i_max`, `delta_i`.

## 3. Yêu cầu ZIP đầu vào cho từng tính năng

### 3.1. Xử lý file sơ bộ

Nút trên UI: **Xử lý và tải ZIP**.

API: `POST /api/kew/organize-field-zip`.

ZIP đầu vào cần có:

```text
HoSoGoc.zip
├── ke-hoach-hien-truong.xlsx
├── S0001/
│   ├── INPS0001.KEW
│   ├── a.png
│   └── các file khác trong record nếu có
├── S0002/
│   ├── INPS0002.KEW
│   ├── a.png
│   └── ...
├── PS-SD641.BMP
├── PS-SD642.BMP
├── PS-SD643.BMP
└── ...
```

Quy tắc:

- Phải có ít nhất 1 file Excel `.xlsx`/`.xlsm` trong ZIP.
- Phải có các thư mục record tên đúng dạng `S0001`, `S0002`, ...
- Giá trị cột `file` trong Excel được chuẩn hóa về `Sxxxx` và phải tìm thấy thư mục tương ứng.
- File ảnh phải có tên dạng `PS-SDxxx.BMP`; dải ảnh `img` đến `imgend` là inclusive.
- Các dải ảnh giữa các dòng Excel không được chồng lấn, kể cả khi có `imgomit`.
- `imgomit` chỉ loại ảnh khỏi việc copy/OCR nếu nằm trong dải.
- `a.png` cần có sẵn trong thư mục `Sxxxx`, vì bước này chỉ copy nguyên thư mục record, không tạo `a.png`.

Thứ tự ảnh để OCR:

- Ảnh hợp lệ đầu tiên sau khi bỏ `imgomit` được xem là màn hình `SD140`.
- Ảnh thứ 2 là `SD141`.
- Ảnh thứ 5 là `SD144`.
- Ảnh thứ 6 là `SD145`.
- OCR điền các cột: `p`, `cos_phi`, `i_max`, `u_min`, `u_max`, `delta_u`, `delta_i`, `thd`, `tdd`.

Output:

```text
KEW_HoSoDaXuLy.zip
└── Project_Output/
    ├── ke-hoach-hien-truong.xlsx
    ├── MBA T1/
    │   ├── INPS0001.KEW
    │   ├── a.png
    │   ├── PS-SD641.BMP
    │   ├── PS-SD642.BMP
    │   └── ...
    └── Tên thiết bị khác/
        ├── INPS0002.KEW
        ├── a.png
        └── ...
```

### 3.2. Tạo chương 4

Nút trên UI: **Tạo chương 4**.

API: `POST /api/kew/generate-chapter4`.

Input nên là `KEW_HoSoDaXuLy.zip` từ bước sơ bộ. Hệ thống tìm `Project_Output`; nếu không có thì dùng root ZIP.

Chỉ các dòng Excel có `type = 4` hoặc `type = device4` mới được đưa vào Chương 4. Mỗi thư mục thiết bị cần có:

- `a.png`
- Các file `PS-SDxxx.BMP` theo thứ tự số. Template device cần 6 ảnh; nếu ít hơn, code sẽ lặp lại ảnh đầu tiên cho slot thiếu, nhưng nên chuẩn bị đủ 6 ảnh để báo cáo đúng.

Template sử dụng: `static/word-template/device4.docx`.

### 3.3. Tạo chương 5

Nút trên UI: **Tạo chương 5**.

API: `POST /api/kew/generate-chapter5`.

Chương 5 gồm:

- Tất cả MBA.
- Các thiết bị không phải `type = 4`/`device4`.
- Bảng tổng hợp MBA chèn trước phần thiết bị nếu có MBA.

Quy tắc phân loại:

- Nếu ZIP có Excel metadata: nên điền rõ `type = MBA` cho MBA, `type = 4` cho Chương 4, `type = device` hoặc để thiết bị thường.
- Nếu ZIP không có Excel: code đoán MBA theo tên thư mục, ví dụ bắt đầu bằng `MBA`, `TR`, `TBA`, `T1`, `MBT`, hoặc có chữ "máy biến áp"/"biến áp".
- Nếu có Excel nhưng cột `type` trong dòng bị trống/không nhận diện, phần Word sẽ xem là `device`, không đoán theo tên.

Mỗi thư mục MBA/thiết bị cần có:

- `a.png`
- `PS-SDxxx.BMP`
- Với MBA, nên có `INPSxxxx.KEW` để bảng thông số trong template lấy dữ liệu thật. Nếu thiếu, một số bảng sẽ hiện dấu `—`.

Templates sử dụng:

- MBA: `static/word-template/mba.docx`
- Thiết bị: `static/word-template/device.docx`
- Bảng tổng hợp MBA: `static/word-template/totalmba.docx`

### 3.4. Excel MBA

Nút trên UI: **Excel MBA**.

API: `POST /api/kew/generate-excel-mba`.

Output mặc định trên UI: `NX MBA <Tên dự án>.xlsm`.

Input:

- ZIP đã sắp xếp có Excel metadata và các thư mục thiết bị.
- Nên điền `type = MBA` cho tất cả MBA.
- Mỗi thư mục MBA nên có file `INPSxxxx.KEW`.

Xử lý:

- Lọc danh sách MBA.
- Mở template `static/excel-template/MBA.xlsm`.
- Dùng sheet `MBA1` làm sheet mẫu, copy thành nhiều sheet nếu có nhiều MBA.
- Đọc `INPSxxxx.KEW`, bỏ dòng đầu, map các cột `AVG_A1[A]`, `AVG_P[W]`, `AVG_VL1[V]`, `AVG_THDVR1[%]`, ...
- Ghi bảng dữ liệu vào từng sheet MBA và điền đánh giá vào cột `AE`.
- Nếu không đọc được `INPSxxxx.KEW`, chỉ fallback một số ô đánh giá từ Excel metadata: `u_min`, `u_max`, `delta_u`, `cos_phi`, `thd`, `tdd`.
- Nếu có sheet `Tổn thất MBA`, hệ thống điền STT, tên MBA, `pdm` và công thức link về sheet từng MBA.

### 3.5. Tạo Bảng 6.3

Nút trên UI: **Tạo Bảng 6.3**.

API: `POST /api/kew/generate-table6`.

Input cần có Excel metadata trong ZIP. Nếu không có Excel, file Word vẫn được tạo nhưng bảng sẽ rỗng.

Cột Excel được đọc:

- `stt`: sắp xếp dòng.
- `name`: tên thiết bị.
- `i_max`: cột I trong bảng.
- `delta_i`: cột delta I.
- `cos_phi`: cột cos phi.
- `p`: cột P.
- `tdd`: cột TDD.

Code hiện tại đưa tất cả dòng Excel vào Bảng 6.3, không lọc theo `type`.

Nhận xét tự động:

- `cos_phi < 0.75`: "Hệ số Cosφ còn thấp".
- `delta_i >= 10`: "Độ lệch pha dòng điện còn cao".
- `tdd >= 12`: "Tổng biến dạng sóng hài dòng điện còn cao".
- Nếu không có vi phạm: "Thiết bị vận hành ổn định".

Template sử dụng: `static/word-template/table6.docx`.

## 4. Lỗi thường gặp

- **Không tìm thấy Excel**: ZIP không có `.xlsx`/`.xlsm`, hoặc Excel nằm trong file ẩn/thư mục bị bỏ qua.
- **Thiếu cột bắt buộc**: Excel cần đủ `stt`, `name`, `file`, `img`, `imgend`, `type`, `pdm`, `current_char`.
- **Sai tên cột**: `pf` không thay cho `cos_phi`; `di` không thay cho `delta_i`.
- **Không có thư mục Sxxxx**: cột `file` trỏ tới record không có trong ZIP.
- **Thiếu ảnh BMP**: dải `img` - `imgend` có ảnh không tồn tại.
- **Trùng dải ảnh**: hai thiết bị có dải ảnh giao nhau.
- **Thiếu `a.png`**: bước sơ bộ có thể chạy, nhưng bước Word sẽ lỗi khi render section.
- **Chương 4 rỗng**: không có dòng Excel nào có `type = 4`/`device4`.
- **Excel MBA báo không tìm thấy MBA**: cột `type` chưa điền `MBA`, hoặc tên thư mục không được đoán là MBA.

## 5. Đề xuất cải tiến và tinh gọn

1. Đồng bộ schema trong UI, `excel-readme.md` và code: thay `pf`, `i1`, `i2`, `i3`, `di` bằng `cos_phi`, `i_max`, `delta_i`; bỏ mô tả `pdm` là điện áp định mức.
2. Tạo file Excel template hiện trường mẫu để người dùng tải về, gồm đầy đủ 19 cột hiện tại và validation dropdown cho `type`, `current_char`.
3. Tách JavaScript trong `kew.html` thành file riêng, gom logic upload/download dùng chung cho Chương 4, Chương 5, Table 6 và Excel MBA.
4. Hiển thị warnings sau khi xử lý thành công. Backend đã trả warning qua header, nhưng UI hiện chưa đọc các header `X-KEW-*`.
5. Expose tùy chọn OCR trên UI: bật/tắt `run_ocr`, cho phép/không cho phép `ocr_overwrite`.
6. Chuẩn hóa lại phân loại MBA/device: Word và Excel MBA đang có hành vi khác nhau khi Excel có `type` trống/không hợp lệ.
7. Hỗ trợ alias cột để tương thích dữ liệu cũ: `pf -> cos_phi`, `di -> delta_i`, `i1/i2/i3 -> i_max` nếu cần.
8. Validate `a.png` ngay ở bước sơ bộ để báo lỗi sớm, vì hiện tại lỗi chỉ xuất hiện khi tạo Word.
9. Làm rõ hỗ trợ định dạng Excel: bước sơ bộ thông báo `.xls` nhưng code chỉ tìm `.xlsx/.xlsm`; nếu cần `.xls` thì thêm engine đọc phù hợp.
10. Bảo vệ giải nén ZIP: dùng safe extract để chặn đường dẫn nguy hiểm, giới hạn kích thước ZIP và số lượng file.
11. Gom các hằng số ngưỡng đánh giá (`5%`, `0.9`, `8%`, `12%`, `20%`) vào `config.json` thay vì rải trong `gen_word.py`.
12. Thêm test tự động cho: map cột Excel, validate dải ảnh, OCR write-back, sinh Chương 4/5, Table 6 và Excel MBA với ZIP mẫu nhỏ.
13. Quyết định số phận của endpoint `/generate-word-report`: nếu cần trên UI thì thêm nút "Tạo báo cáo tổng hợp"; nếu không thì xóa/gom để tránh API thừa.
14. Làm rõ vai trò `imglu`: hiện được copy thành `load-unload-xxx.BMP` nhưng chưa được template Word tự động sử dụng.
