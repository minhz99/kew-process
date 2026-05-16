import os
import io
import shutil
import tempfile
import traceback
import zipfile
import re
import urllib.parse
from flask import Blueprint, request, jsonify, send_file


try:
    import pandas as pd
    from openpyxl import load_workbook
    from typing import Mapping
    _MBA_DEPS_OK = True
except ImportError:
    _MBA_DEPS_OK = False

# ─── Cấu hình MBA export ─────────────────────────────────────────────────────
_MBA_SKIP_ROWS = 1
_MBA_START_ROW = 2
_MBA_START_COL = 2
_MBA_TEMPLATE_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    '..', '..', 'static', 'excel-template', 'MBA.xlsm'
)
_MBA_PREBUILT_COUNT = 10  # Số sheet có sẵn trong template (MBA1 … MBA10)
_MBA_COLUMN_MAPPING: Mapping[str, str] = {
    "AVG_A1[A]": "AVG_A1[A]",
    "AVG_A2[A]": "AVG_A2[A]",
    "AVG_A3[A]": "AVG_A3[A]",
    "AVG_P[W]": "AVG_P[W]",
    "AVG_Q[var]": "AVG_Q[var]",
    "AVG_S[VA]": "AVG_S[VA]",
    "AVG_PF[_]": "AVG_PF",
    "AVG_VL1[V]": "AVG_VL1[V]",
    "AVG_VL2[V]": "AVG_VL2[V]",
    "AVG_VL3[V]": "AVG_VL3[V]",
    "AVG_THDVR1[%]": "AVG_Vthd1[%]",
    "AVG_THDVR2[%]": "AVG_Vthd2[%]",
    "AVG_THDVR3[%]": "AVG_Vthd3[%]",
    "AVG_THDAR1[%]": "AVG_Athd1[%]",
    "AVG_THDAR2[%]": "AVG_Athd2[%]",
    "AVG_THDAR3[%]": "AVG_Athd3[%]",
    "AVG_UV[%]": "AVG_Vunb[%]",
    "AVG_UA[%]": "AVG_Aunb[%]",
}
_MBA_SCALE_DIV_1000 = ("AVG_P[W]", "AVG_Q[var]", "AVG_S[VA]")
_MBA_ROUND: Mapping[str, int] = {
    "AVG_A1[A]": 2, "AVG_A2[A]": 2, "AVG_A3[A]": 2,
    "AVG_P[W]": 2, "AVG_Q[var]": 2, "AVG_S[VA]": 2,
    "AVG_PF": 4,
    "AVG_VL1[V]": 1, "AVG_VL2[V]": 1, "AVG_VL3[V]": 1,
    "AVG_Vthd1[%]": 3, "AVG_Vthd2[%]": 3, "AVG_Vthd3[%]": 3,
    "AVG_Athd1[%]": 2, "AVG_Athd2[%]": 2, "AVG_Athd3[%]": 2,
    "AVG_Vunb[%]": 4, "AVG_Aunb[%]": 3,
}
_MBA_FMT: Mapping[str, str] = {
    "AVG_A1[A]": "0.00", "AVG_A2[A]": "0.00", "AVG_A3[A]": "0.00",
    "AVG_P[W]": "0.00", "AVG_Q[var]": "0.00", "AVG_S[VA]": "0.00",
    "AVG_PF": "0.0000",
    "AVG_VL1[V]": "0.0", "AVG_VL2[V]": "0.0", "AVG_VL3[V]": "0.0",
    "AVG_Vthd1[%]": "0.000", "AVG_Vthd2[%]": "0.000", "AVG_Vthd3[%]": "0.000",
    "AVG_Athd1[%]": "0.00", "AVG_Athd2[%]": "0.00", "AVG_Athd3[%]": "0.00",
    "AVG_Vunb[%]": "0.0000", "AVG_Aunb[%]": "0.000",
}
_KM_RE = re.compile(r"^\s*([+-]?\d+(?:\.\d+)?)\s*([kKmM])\s*$")


def _mba_to_number(v):
    """
    Chuyển đổi một giá trị từ file KEW sang số thực (float).
    Xử lý các đơn vị k (kilo), m (mega) và loại bỏ các ký hiệu đơn vị khác.
    
    Args:
        v: Giá trị cần chuyển đổi.
        
    Returns:
        float or pd.NA: Giá trị số sau khi chuyển đổi hoặc pd.NA nếu không hợp lệ.
    """
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return pd.NA
    s = str(v).strip()
    
    # Xoá dấu âm để lấy giá trị tuyệt đối
    s = s.replace('-', '')
    
    if not s or s.lower() == "nan":
        return pd.NA
        
    # Regex tìm số thập phân và ký tự k/m (chấp nhận dính chữ như kvar, kW)
    m = re.search(r"(\d+(?:\.\d+)?)\s*([kKmM])?", s)
    if m:
        base = float(m.group(1))
        unit = (m.group(2) or "").lower()
        if unit == "k":
            base *= 1_000.0
        elif unit == "m":
            base *= 1_000_000.0
        return base
        
    return pd.to_numeric(s, errors="coerce")


def _mba_extract(df: "pd.DataFrame") -> "tuple[pd.DataFrame, list[str]]":
    """
    Trích xuất và chuẩn hóa dữ liệu từ DataFrame gốc của file KEW.
    Áp dụng mapping cột, chuyển đổi đơn vị và làm tròn số.
    
    Args:
        df (pd.DataFrame): DataFrame dữ liệu gốc.
        
    Returns:
        tuple: (DataFrame đã chuẩn hóa, danh sách các cảnh báo warnings).
    """
    orig_cols = list(_MBA_COLUMN_MAPPING.keys())
    available_cols = [c for c in orig_cols if c in df.columns]
    missing = [c for c in orig_cols if c not in df.columns]
    
    out = df[available_cols].copy()
    rename_mapping = {k: v for k, v in _MBA_COLUMN_MAPPING.items() if k in available_cols}
    out = out.rename(columns=rename_mapping)
    
    for m in missing:
        out[_MBA_COLUMN_MAPPING[m]] = pd.NA

    target_ordered = list(_MBA_COLUMN_MAPPING.values())
    out = out[target_ordered]

    for col in out.columns:
        out[col] = out[col].map(_mba_to_number)
        
    cols_to_check = [
        "AVG_A1[A]", "AVG_A2[A]", "AVG_A3[A]",
        "AVG_P[W]", "AVG_Q[var]", "AVG_S[VA]"
    ]
    for col in cols_to_check:
        renamed = dict(_MBA_COLUMN_MAPPING).get(col, col)
        if renamed in out.columns:
            out[renamed] = out[renamed].apply(lambda x: x * 1000.0 if pd.notna(x) and abs(x) < 10.0 else x)

    for col in _MBA_SCALE_DIV_1000:
        renamed = dict(_MBA_COLUMN_MAPPING).get(col, col)
        if renamed in out.columns:
            out[renamed] = out[renamed] / 1000.0
    for col, nd in _MBA_ROUND.items():
        if col in out.columns:
            out[col] = pd.to_numeric(out[col], errors='coerce').round(nd)
            
    warnings = []
    if missing:
        warnings.append(f"Dữ liệu gốc khuyết {len(missing)} cột: {missing}")
        
    return out, warnings


def _evaluate_for_excel(df: "pd.DataFrame") -> dict[int, str]:
    """
    Đánh giá các chỉ số kỹ thuật (điện áp, PF, THD, unbalance) để ghi nhận xét vào file Excel.
    Sử dụng các hàm đánh giá dùng chung với module report.
    
    Args:
        df (pd.DataFrame): DataFrame dữ liệu đã chuẩn hóa.
        
    Returns:
        dict: Một dictionary mapping dòng (row index) với chuỗi nhận xét (ví dụ {8: "Đạt"}).
    """
    from modules.report.gen_word import (
        _eval_voltage, _eval_pf, _eval_thd, _eval_unbalance,
        _V_DEV_LIMIT_PCT, _THDV_LIMIT_PCT, _TDD_LIMIT_PCT, _MBA_NOMINAL_VOLTAGE_V
    )
    
    res = {}
    
    # AE8: Điện áp
    if "AVG_VL1[V]" in df.columns:
        u_vals = df["AVG_VL1[V]"].dropna()
        if not u_vals.empty:
            eval_str, _, _, _ = _eval_voltage(u_vals.max(), u_vals.min(), u_vals.mean(), _MBA_NOMINAL_VOLTAGE_V)
            res[8] = eval_str
            
    # AE11: Dòng điện (Word không có logic đánh giá)
    res[11] = "—"
    
    # AE14: Pha áp
    if "AVG_Vunb[%]" in df.columns:
        vu_vals = df["AVG_Vunb[%]"].dropna()
        if not vu_vals.empty:
            res[14] = _eval_unbalance(vu_vals.max(), vu_vals.mean(), _V_DEV_LIMIT_PCT)
            
    # AE15: Pha dòng
    if "AVG_Aunb[%]" in df.columns:
        au_vals = df["AVG_Aunb[%]"].dropna()
        if not au_vals.empty:
            res[15] = _eval_unbalance(au_vals.max(), au_vals.mean(), 10.0)
            
    # AE16: PF
    if "AVG_PF" in df.columns:
        pf_vals = df["AVG_PF"].dropna()
        if not pf_vals.empty:
            res[16] = _eval_pf(pf_vals.max(), pf_vals.min(), pf_vals.mean())
            
    # AE20: THD
    thd_cols = [c for c in ["AVG_Vthd1[%]", "AVG_Vthd2[%]", "AVG_Vthd3[%]"] if c in df.columns]
    if thd_cols:
        max_vals = [df[c].max() for c in thd_cols if not df[c].dropna().empty]
        avg_vals = [df[c].mean() for c in thd_cols if not df[c].dropna().empty]
        res[20] = _eval_thd(max_vals, avg_vals, _THDV_LIMIT_PCT)
        
    # AE23: TDD
    tdd_cols = [c for c in ["AVG_Athd1[%]", "AVG_Athd2[%]", "AVG_Athd3[%]"] if c in df.columns]
    if tdd_cols:
        max_vals = [df[c].max() for c in tdd_cols if not df[c].dropna().empty]
        avg_vals = [df[c].mean() for c in tdd_cols if not df[c].dropna().empty]
        res[23] = _eval_thd(max_vals, avg_vals, _TDD_LIMIT_PCT)
        
    return res


def _mba_write(ws, df: "pd.DataFrame") -> None:
    """
    Ghi dữ liệu từ DataFrame vào worksheet Excel và áp dụng định dạng.
    Cũng ghi các nhận xét đánh giá vào cột AE (cột 31).
    
    Args:
        ws: Worksheet của openpyxl.
        df (pd.DataFrame): DataFrame dữ liệu đã chuẩn hóa.
    """
    sr, sc = _MBA_START_ROW, _MBA_START_COL
    
    # Xoá vùng dữ liệu mẫu cũ (nếu có)
    last_row = ws.max_row
    last_col = sc + len(df.columns) - 1
    for r in range(sr + 1, last_row + 1):
        for c in range(sc, last_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.value = None
            cell.number_format = "General"
            
    for ci, cname in enumerate(df.columns):
        ws.cell(row=sr, column=sc + ci, value=cname)
    for ri, row in enumerate(df.values):
        for ci, val in enumerate(row):
            if pd.notna(val):
                cname = df.columns[ci]
                cell = ws.cell(row=sr + 1 + ri, column=sc + ci, value=float(val))
                fmt = _MBA_FMT.get(cname)
                if fmt:
                    cell.number_format = fmt

    # Ghi nhận xét vào các ô AE (col 31) theo logic của Word
    evals = _evaluate_for_excel(df)
    for r_idx, ev_str in evals.items():
        cell = ws.cell(row=r_idx, column=31)
        cell.value = ev_str

kew_bp = Blueprint('kew_bp', __name__)


@kew_bp.route('/export-mba', methods=['POST'])
def export_mba():
    """
    API endpoint để xuất báo cáo MBA sang file Excel (.xlsm).
    
    Nhận các file KEW hoặc ZIP từ người dùng, trích xuất dữ liệu và điền vào 
    template Excel có sẵn. Mỗi file KEW sẽ tương ứng với một sheet MBA trong file kết quả.
    
    Returns:
        Response: File Excel kết quả hoặc lỗi JSON.
    """
    if not _MBA_DEPS_OK:
        return jsonify({'error': 'Thiếu thư viện pandas hoặc openpyxl.'}), 500

    out_filename = request.form.get('filename', '').strip() or 'NX-MBA.xlsm'
    # Đảm bảo đuôi .xlsm để bảo toàn macro của template
    if not out_filename.lower().endswith(('.xlsx', '.xlsm')):
        out_filename += '.xlsm'

    # ── Lấy tất cả file KEW bytes từ request ─────────────────────────────────
    def _extract_kew_bytes(f) -> list:
        """Từ 1 FileStorage trả về list [(name, bytes)]."""
        raw = f.read()
        if f.filename.upper().endswith('.ZIP'):
            with zipfile.ZipFile(io.BytesIO(raw), 'r') as zf:
                entries = [n for n in zf.namelist() if n.upper().endswith('.KEW')]
                return [(os.path.basename(n), zf.read(n)) for n in entries]
        return [(os.path.basename(f.filename), raw)]

    kew_list = []  # list of (name, bytes)
    if 'files' in request.files:
        for f in request.files.getlist('files'):
            kew_list.extend(_extract_kew_bytes(f))
    elif 'zip' in request.files:
        kew_list.extend(_extract_kew_bytes(request.files['zip']))

    if not kew_list:
        return jsonify({'error': 'Cần upload ít nhất 1 file .KEW hoặc .ZIP.'}), 400


    # ── Load template ─────────────────────────────────────────────────────────
    template_path = os.path.normpath(_MBA_TEMPLATE_PATH)
    if not os.path.isfile(template_path):
        return jsonify({'error': f'Không tìm thấy template MBA.xlsm tại {template_path}'}), 500

    try:
        wb = load_workbook(template_path, keep_vba=True)
    except Exception as e:
        return jsonify({'error': f'Không mở được template: {e}'}), 500

    # ── Danh sách sheet có sẵn trong template (MBA1 … MBA10) ─────────────────
    # Template đã có sẵn _MBA_PREBUILT_COUNT sheet; chỉ cần trỏ vào và đổi tên.
    # Nếu số lượng MBA vượt quá số sheet có sẵn → copy_worksheet từ sheet cuối.
    prebuilt_sheets = wb.sheetnames[:_MBA_PREBUILT_COUNT]

    errors_list = []
    try:
        for idx, (kew_name, kew_bytes) in enumerate(kew_list):
            # ── Parse KEW ────────────────────────────────────────────────────────
            try:
                df_raw = pd.read_csv(
                    io.BytesIO(kew_bytes),
                    skiprows=_MBA_SKIP_ROWS,
                    low_memory=False,
                )
                df_raw.columns = df_raw.columns.str.strip()
            except Exception as e:
                errors_list.append(f'{kew_name}: không đọc được ({e})')
                continue

            df, warnings = _mba_extract(df_raw)
            if warnings:
                errors_list.extend([f"{kew_name}: {w}" for w in warnings])

            # ── Lấy sheet đích ───────────────────────────────────────────────────
            try:
                if idx < len(prebuilt_sheets):
                    # Tái sử dụng sheet có sẵn, KHÔNG đổi tên
                    ws = wb[prebuilt_sheets[idx]]
                else:
                    # Vượt quá số sheet template → copy từ sheet cuối cùng có sẵn
                    ws_ref = wb[wb.sheetnames[len(prebuilt_sheets) - 1]]
                    ws = wb.copy_worksheet(ws_ref)
                    ws.title = f'MBA{idx + 1}'
            except Exception as e:
                errors_list.append(f'{kew_name}: không thể lấy sheet ({e})')
                continue

            try:
                _mba_write(ws, df)
            except Exception as e:
                errors_list.append(f'{kew_name}: lỗi ghi dữ liệu ({e})')
                continue

        if errors_list and len(errors_list) == len(kew_list):
            # Tất cả đều lỗi
            return jsonify({'error': 'Tất cả file thất bại: ' + '; '.join(errors_list)}), 400

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Lỗi hệ thống khi xử lý: {str(e)}'}), 500

    # Giữ đuôi .xlsm để bảo toàn macro
    if not out_filename.lower().endswith('.xlsm'):
        out_filename = os.path.splitext(out_filename)[0] + '.xlsm'

    resp = send_file(
        output,
        as_attachment=True,
        download_name=out_filename,
        mimetype='application/vnd.ms-excel.sheet.macroEnabled.12',
    )
    if errors_list:
        resp.headers['X-MBA-Warnings'] = urllib.parse.quote('; '.join(errors_list))
    return resp


@kew_bp.route("/organize-field-zip", methods=["POST"])
def organize_field_zip():
    """
    API endpoint để tổ chức lại file ZIP hồ sơ hiện trường.
    
    Thực hiện:
    1. Đọc file Excel kế hoạch trong ZIP.
    2. Đổi tên các thư mục Sxxxx thành tên thiết bị tương ứng.
    3. Di chuyển các file ảnh PS-SDxxx.BMP vào đúng thư mục thiết bị.
    4. (Tùy chọn) Chạy OCR tự động đọc các thông số đo từ ảnh BMP và ghi vào Excel.
    5. Nén lại thành file ZIP kết quả.

    Form params:
        zip / file: File ZIP đầu vào.
        run_ocr: '1' hoặc 'true' để chạy OCR sau khi sắp xếp (mặc định: true).
        ocr_overwrite: '1' hoặc 'true' để ghi đè các ô Excel đã có giá trị (mặc định: false).
    
    Returns:
        Response: File ZIP đã tổ chức lại hoặc lỗi JSON.
    """
    from modules.kew import organize_field_zip as organize_mod

    zf = request.files.get("zip") or request.files.get("file")
    if zf is None or not getattr(zf, "filename", None):
        return jsonify({"error": "Cần upload file ZIP (form field zip hoặc file)."}), 400
    if not str(zf.filename).lower().endswith(".zip"):
        return jsonify({"error": "Chỉ chấp nhận file .zip."}), 400

    zip_bytes = zf.read()
    if not zip_bytes:
        return jsonify({"error": "File ZIP rỗng."}), 400

    # Tham số OCR từ form
    run_ocr_raw = request.form.get("run_ocr", "true").strip().lower()
    run_ocr = run_ocr_raw not in ("0", "false", "no", "")
    ocr_overwrite_raw = request.form.get("ocr_overwrite", "false").strip().lower()
    ocr_overwrite = ocr_overwrite_raw in ("1", "true", "yes")

    work = tempfile.mkdtemp(prefix="kew_field_org_")
    try:
        out_path, warnings, fatal = organize_mod.process_field_zip_bytes(
            zip_bytes, work, run_ocr=run_ocr, ocr_overwrite=ocr_overwrite
        )
        if fatal:
            return jsonify({"errors": fatal, "warnings": warnings}), 400
        if not out_path or not os.path.isfile(out_path):
            return jsonify({"error": "Không tạo được file kết quả.", "warnings": warnings}), 500
        with open(out_path, "rb") as fh:
            buf = io.BytesIO(fh.read())
        buf.seek(0)
        resp = send_file(
            buf,
            as_attachment=True,
            download_name="KEW_HoSoDaXuLy.zip",
            mimetype="application/zip",
        )
        if warnings:
            resp.headers["X-KEW-Field-Warnings"] = urllib.parse.quote("; ".join(warnings))
        return resp
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"Lỗi xử lý hồ sơ KEW: {e}"}), 500
    finally:
        shutil.rmtree(work, ignore_errors=True)


@kew_bp.route("/generate-word-report", methods=["POST"])
def generate_word_report():
    """
    API endpoint để sinh báo cáo Word tổng hợp từ hồ sơ thiết bị (file ZIP).
    
    Quy trình xử lý:
    1. Tiếp nhận file ZIP từ client và giải nén vào bộ nhớ tạm.
    2. Duyệt cấu trúc thư mục, nhận diện Máy biến áp (MBA) và thiết bị phụ tải.
    3. Tạo và trộn các section theo thứ tự: Các MBA -> Bảng tổng kết MBA -> Các thiết bị phụ tải.
    4. Trả về file Word hoàn chỉnh cho người dùng.
    
    Returns:
        Response: File .docx báo cáo (application/vnd.openxmlformats-officedocument.wordprocessingml.document) hoặc lỗi JSON.
    """
    from modules.report.gen_word import build_word_report_from_zip

    zf = request.files.get("zip") or request.files.get("file")
    if zf is None or not getattr(zf, "filename", None):
        return jsonify({"error": "Cần upload file ZIP (form field zip hoặc file)."}), 400
    if not str(zf.filename).lower().endswith(".zip"):
        return jsonify({"error": "Chỉ chấp nhận file .zip."}), 400

    zip_bytes = zf.read()
    if not zip_bytes:
        return jsonify({"error": "File ZIP rỗng."}), 400

    out_name = (request.form.get("filename", "") or "").strip() or "BaoCao_KEW.docx"
    if not out_name.lower().endswith(".docx"):
        out_name += ".docx"

    work = tempfile.mkdtemp(prefix="kew_word_")
    try:
        out_path = os.path.join(work, out_name)
        try:
            _, warnings = build_word_report_from_zip(zip_bytes, out_path)
        except (FileNotFoundError, ValueError) as e:
            return jsonify({"error": str(e)}), 400
        except RuntimeError as e:
            return jsonify({"error": str(e)}), 400
        if not os.path.isfile(out_path):
            return jsonify({"error": "Không tạo được file Word."}), 500

        with open(out_path, "rb") as fh:
            buf = io.BytesIO(fh.read())
        buf.seek(0)
        resp = send_file(
            buf,
            as_attachment=True,
            download_name=out_name,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        if warnings:
            resp.headers["X-KEW-Word-Warnings"] = urllib.parse.quote("; ".join(warnings))
        return resp
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"Lỗi sinh báo cáo Word: {e}"}), 500
    finally:
        shutil.rmtree(work, ignore_errors=True)


@kew_bp.route("/generate-chapter4", methods=["POST"])
def generate_chapter4():
    """
    API endpoint để sinh Chương 4 Word: chỉ các thiết bị có ``type=\"4\"``.

    Sử dụng template ``device4.docx`` (cấu trúc giống ``device.docx``).

    Returns:
        Response: File .docx Chương 4 hoặc lỗi JSON.
    """
    from modules.report.gen_word import build_chapter4_from_zip

    zf = request.files.get("zip") or request.files.get("file")
    if zf is None or not getattr(zf, "filename", None):
        return jsonify({"error": "Cần upload file ZIP (form field zip hoặc file)."}), 400
    if not str(zf.filename).lower().endswith(".zip"):
        return jsonify({"error": "Chỉ chấp nhận file .zip."}), 400

    zip_bytes = zf.read()
    if not zip_bytes:
        return jsonify({"error": "File ZIP rỗng."}), 400

    out_name = (request.form.get("filename", "") or "").strip() or "Chương 4.docx"
    if not out_name.lower().endswith(".docx"):
        out_name += ".docx"

    work = tempfile.mkdtemp(prefix="kew_chap4_")
    try:
        out_path = os.path.join(work, out_name)
        try:
            _, warnings = build_chapter4_from_zip(zip_bytes, out_path)
        except (FileNotFoundError, ValueError) as e:
            return jsonify({"error": str(e)}), 400
        except RuntimeError as e:
            return jsonify({"error": str(e)}), 400
        if not os.path.isfile(out_path):
            return jsonify({"error": "Không tạo được file Word Chương 4."}), 500

        with open(out_path, "rb") as fh:
            buf = io.BytesIO(fh.read())
        buf.seek(0)
        resp = send_file(
            buf,
            as_attachment=True,
            download_name=out_name,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        if warnings:
            resp.headers["X-KEW-Word-Warnings"] = urllib.parse.quote("; ".join(warnings))
        return resp
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"Lỗi sinh báo cáo Chương 4: {e}"}), 500
    finally:
        shutil.rmtree(work, ignore_errors=True)


@kew_bp.route("/generate-chapter5", methods=["POST"])
def generate_chapter5():
    """
    API endpoint để sinh Chương 5 Word: MBA + các thiết bị không có ``type=\"4\"``.

    Returns:
        Response: File .docx Chương 5 hoặc lỗi JSON.
    """
    from modules.report.gen_word import build_chapter5_from_zip

    zf = request.files.get("zip") or request.files.get("file")
    if zf is None or not getattr(zf, "filename", None):
        return jsonify({"error": "Cần upload file ZIP (form field zip hoặc file)."}), 400
    if not str(zf.filename).lower().endswith(".zip"):
        return jsonify({"error": "Chỉ chấp nhận file .zip."}), 400

    zip_bytes = zf.read()
    if not zip_bytes:
        return jsonify({"error": "File ZIP rỗng."}), 400

    out_name = (request.form.get("filename", "") or "").strip() or "Chương 5.docx"
    if not out_name.lower().endswith(".docx"):
        out_name += ".docx"

    work = tempfile.mkdtemp(prefix="kew_chap5_")
    try:
        out_path = os.path.join(work, out_name)
        try:
            _, warnings = build_chapter5_from_zip(zip_bytes, out_path)
        except (FileNotFoundError, ValueError) as e:
            return jsonify({"error": str(e)}), 400
        except RuntimeError as e:
            return jsonify({"error": str(e)}), 400
        if not os.path.isfile(out_path):
            return jsonify({"error": "Không tạo được file Word Chương 5."}), 500

        with open(out_path, "rb") as fh:
            buf = io.BytesIO(fh.read())
        buf.seek(0)
        resp = send_file(
            buf,
            as_attachment=True,
            download_name=out_name,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        if warnings:
            resp.headers["X-KEW-Word-Warnings"] = urllib.parse.quote("; ".join(warnings))
        return resp
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"Lỗi sinh báo cáo Chương 5: {e}"}), 500
    finally:
        shutil.rmtree(work, ignore_errors=True)


@kew_bp.route("/generate-table6", methods=["POST"])
def generate_table6():
    """
    API endpoint để sinh bảng tổng hợp kết quả đo kiểm (Table 6).
    
    Trích xuất các chỉ số I, P, PF, THD, TDD từ Excel hiện trường và 
    tạo bảng tổng hợp trong file Word.
    
    Returns:
        Response: File Word chứa Table 6 hoặc lỗi JSON.
    """
    from modules.report.gen_word import generate_table6_from_zip

    zf = request.files.get("zip") or request.files.get("file")
    if zf is None or not getattr(zf, "filename", None):
        return jsonify({"error": "Cần upload file ZIP (form field zip hoặc file)."}), 400
    if not str(zf.filename).lower().endswith(".zip"):
        return jsonify({"error": "Chỉ chấp nhận file .zip."}), 400

    zip_bytes = zf.read()
    if not zip_bytes:
        return jsonify({"error": "File ZIP rỗng."}), 400

    out_name = (request.form.get("filename", "") or "").strip() or "Bang_TongHop_Table6.docx"
    if not out_name.lower().endswith(".docx"):
        out_name += ".docx"
    work = tempfile.mkdtemp(prefix="kew_table6_")
    try:
        out_path = os.path.join(work, out_name)
        try:
            _, warnings = generate_table6_from_zip(zip_bytes, out_path)
        except (FileNotFoundError, ValueError) as e:
            return jsonify({"error": str(e)}), 400
        except RuntimeError as e:
            return jsonify({"error": str(e)}), 400
        if not os.path.isfile(out_path):
            return jsonify({"error": "Không tạo được bảng tổng hợp."}), 500

        with open(out_path, "rb") as fh:
            buf = io.BytesIO(fh.read())
        buf.seek(0)
        resp = send_file(
            buf,
            as_attachment=True,
            download_name=out_name,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        if warnings:
            resp.headers["X-KEW-Table6-Warnings"] = urllib.parse.quote("; ".join(warnings))
        return resp
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"Lỗi sinh bảng tổng hợp: {e}"}), 500
    finally:
        shutil.rmtree(work, ignore_errors=True)
