import os
import io
import json
import shutil
import tempfile
import traceback
import zipfile
import re
import urllib.parse
from typing import Mapping, Optional
from copy import copy
from flask import Blueprint, request, jsonify, send_file, current_app
from utils.file_utils import process_zip, group_kew_files_by_id, analyse_folder

try:
    import pandas as pd
    from openpyxl import load_workbook
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


def _mba_write(ws, df: "pd.DataFrame") -> None:
    """Ghi header + data vào sheet, bắt đầu từ B2."""
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

kew_bp = Blueprint('kew_bp', __name__)

def _build_response(results, errors):
    response = {
        'count': len(results),
        'datasets': results,
    }
    if errors:
        response['warnings'] = errors
    # Tương thích ngược: flat keys khi chỉ có 1 bộ
    if len(results) == 1:
        response['summary'] = results[0]['summary']
        response['series'] = results[0]['series']
        response['inps_series'] = results[0].get('inps_series', {})
        response['commentary'] = results[0].get('commentary', '')
        response['device_name'] = results[0].get('device_name', '')
    return jsonify(response)


@kew_bp.route('/upload', methods=['POST'])
def upload_files():
    """
    Handle uploading of .KEW files or a ZIP archive containing KEW data.
    Parses the upload, groups files by device ID, and analyses each group to extract power quality metrics.
    Returns a JSON payload containing series data and KPI summaries.
    """
    # ── Trường hợp 1: Upload file ZIP ──
    if 'zip' in request.files:
        zip_file = request.files['zip']
        if zip_file.filename == '':
            return jsonify({'error': 'File ZIP rỗng.'}), 400

        zip_bytes = io.BytesIO(zip_file.read())
        results, errors = process_zip(zip_bytes)

        if not results:
            return jsonify({'error': '; '.join(errors) or 'ZIP không chứa dữ liệu KEW hợp lệ.'}), 400

        return _build_response(results, errors)

    # ── Trường hợp 2: Upload file KEW rời ──
    if 'files' not in request.files:
        return jsonify({'error': 'Cần upload file .KEW hoặc file .ZIP.'}), 400

    uploaded_files = request.files.getlist('files')
    if not uploaded_files or uploaded_files[0].filename == '':
        return jsonify({'error': 'Chưa chọn file nào.'}), 400

    # Nếu file đầu tiên là ZIP, xử lý như ZIP
    first = uploaded_files[0]
    if first.filename.upper().endswith('.ZIP'):
        zip_bytes = io.BytesIO(first.read())
        results, errors = process_zip(zip_bytes)
        if not results:
            return jsonify({'error': '; '.join(errors) or 'ZIP không chứa dữ liệu KEW.'}), 400
        return _build_response(results, errors)

    # Lọc file KEW và nhóm theo ID
    kew_files = [f for f in uploaded_files if f.filename.upper().endswith('.KEW')]
    if not kew_files:
        return jsonify({'error': 'Không tìm thấy file .KEW hợp lệ.'}), 400

    # Nhóm file KEW theo ID, lưu vào temp dir riêng
    id_to_files = {}
    for f in kew_files:
        fname = os.path.basename(f.filename)
        basename = os.path.splitext(fname)[0]
        match = re.match(r'^([A-Za-z]+)(.+)$', basename)
        file_id = match.group(2) if match else basename
        id_to_files.setdefault(file_id, []).append(f)

    results = []
    errors = []
    for device_id, files in id_to_files.items():
        temp_dir = tempfile.mkdtemp(prefix=f'kew_{device_id}_')
        try:
            for f in files:
                f.save(os.path.join(temp_dir, os.path.basename(f.filename)))
            r = analyse_folder(temp_dir, device_name=device_id)
            if r:
                results.append(r)
            else:
                errors.append(f"ID {device_id}: phân tích thất bại.")
        except Exception as e:
            errors.append(f"ID {device_id}: lỗi – {e}")
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    if not results:
        return jsonify({'error': '; '.join(errors) or 'Tất cả bộ file thất bại.'}), 400

    return _build_response(results, errors)


@kew_bp.route('/fix', methods=['POST'])
def fix_files():
    """
    Simulate and interpolate missing current phases (A2, A3) using Ornstein-Uhlenbeck processes 
    based on a reference phase. Returns a downloadable ZIP archive containing the patched KEW datasets.
    """
    from modules.kew import interpolate_kew
    temp_in = tempfile.mkdtemp(prefix='kew_fix_in_')
    temp_out = tempfile.mkdtemp(prefix='kew_fix_out_')
    
    try:
        if 'zip' in request.files:
            zip_file = request.files['zip']
            with zipfile.ZipFile(io.BytesIO(zip_file.read()), 'r') as zf:
                zf.extractall(temp_in)
        elif 'files' in request.files:
            uploaded_files = request.files.getlist('files')
            first = uploaded_files[0]
            if first.filename.upper().endswith('.ZIP'):
                with zipfile.ZipFile(io.BytesIO(first.read()), 'r') as zf:
                    zf.extractall(temp_in)
            else:
                for f in uploaded_files:
                    f.save(os.path.join(temp_in, os.path.basename(f.filename)))
        else:
            return jsonify({'error': 'No files provided'}), 400
            
        kew_files = []
        for root, _, files in os.walk(temp_in):
            for fname in files:
                if fname.upper().endswith('.KEW'):
                    kew_files.append(os.path.join(root, fname))
                    
        if not kew_files:
            return jsonify({'error': 'Không tìm thấy file .KEW'}), 400
            
        dirs = set(os.path.dirname(f) for f in kew_files)
        for d in dirs:
            out_d = os.path.join(temp_out, os.path.basename(d))
            interpolate_kew.process_folder(d, out_d)
            
        memory_file = io.BytesIO()
        with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root, _, files in os.walk(temp_out):
                for fname in files:
                    file_path = os.path.join(root, fname)
                    arcname = os.path.relpath(file_path, temp_out)
                    zf.write(file_path, arcname)
                    
        memory_file.seek(0)
        return send_file(memory_file, download_name='KEW_Fixed_Data.zip', as_attachment=True, mimetype='application/zip')
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': f"Lỗi nội suy: {str(e)}"}), 500
    finally:
        shutil.rmtree(temp_in, ignore_errors=True)
        shutil.rmtree(temp_out, ignore_errors=True)


@kew_bp.route('/detect', methods=['POST'])
def detect_phases():
    """Detect which current phases are missing in an uploaded KEW set, without full analysis."""
    from modules.kew import interpolate_kew
    temp_in = tempfile.mkdtemp(prefix='kew_det_')
    try:
        if 'files' in request.files:
            uploaded_files = request.files.getlist('files')
            first = uploaded_files[0]
            if first.filename.upper().endswith('.ZIP'):
                with zipfile.ZipFile(io.BytesIO(first.read()), 'r') as zf:
                    zf.extractall(temp_in)
            else:
                for f in uploaded_files:
                    fname = os.path.basename(f.filename)
                    f.save(os.path.join(temp_in, fname))
        elif 'zip' in request.files:
            zf_obj = request.files['zip']
            with zipfile.ZipFile(io.BytesIO(zf_obj.read()), 'r') as zf:
                zf.extractall(temp_in)
        else:
            return jsonify({'error': 'No files'}), 400

        # Find the folder containing KEW files
        kew_dirs = set()
        for root, _, files in os.walk(temp_in):
            if any(f.upper().endswith('.KEW') for f in files):
                kew_dirs.add(root)

        results = []
        for d in sorted(kew_dirs):
            info = interpolate_kew.detect_missing_phases(d)
            results.append(info)

        return jsonify({'results': results})
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500
    finally:
        shutil.rmtree(temp_in, ignore_errors=True)


@kew_bp.route('/correct', methods=['POST'])
def correct_files():
    """Apply per-channel multiplier/offset corrections and return corrected KEW files as ZIP."""
    from modules.kew import correct_kew
    temp_in = tempfile.mkdtemp(prefix='kew_corr_in_')
    temp_out = tempfile.mkdtemp(prefix='kew_corr_out_')
    try:
        # Parse corrections JSON
        corrections_str = request.form.get('corrections', '{}')
        try:
            corrections = json.loads(corrections_str)
        except Exception:
            return jsonify({'error': 'Định dạng corrections JSON không hợp lệ'}), 400

        if not corrections:
            return jsonify({'error': 'Chưa nhập thông số hiệu chỉnh'}), 400

        # Save uploaded files
        if 'files' in request.files:
            uploaded_files = request.files.getlist('files')
            first = uploaded_files[0]
            if first.filename.upper().endswith('.ZIP'):
                with zipfile.ZipFile(io.BytesIO(first.read()), 'r') as zf:
                    zf.extractall(temp_in)
            else:
                for f in uploaded_files:
                    f.save(os.path.join(temp_in, os.path.basename(f.filename)))
        elif 'zip' in request.files:
            with zipfile.ZipFile(io.BytesIO(request.files['zip'].read()), 'r') as zf:
                zf.extractall(temp_in)
        else:
            return jsonify({'error': 'Không có file nào được upload'}), 400

        # Find directories with KEW files
        kew_dirs = set()
        for root, _, files in os.walk(temp_in):
            if any(f.upper().endswith('.KEW') for f in files):
                kew_dirs.add(root)

        if not kew_dirs:
            return jsonify({'error': 'Không tìm thấy file .KEW'}), 400

        for d in sorted(kew_dirs):
            out_d = os.path.join(temp_out, os.path.relpath(d, temp_in))
            correct_kew.process_folder(d, out_d, corrections)

        # Package output as ZIP
        mem = io.BytesIO()
        with zipfile.ZipFile(mem, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root, _, files in os.walk(temp_out):
                for fname in files:
                    fp = os.path.join(root, fname)
                    zf.write(fp, os.path.relpath(fp, temp_out))
        mem.seek(0)
        return send_file(mem, download_name='KEW_Corrected.zip', as_attachment=True, mimetype='application/zip')
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': f'Lỗi hiệu chỉnh: {str(e)}'}), 500
    finally:
        shutil.rmtree(temp_in, ignore_errors=True)
        shutil.rmtree(temp_out, ignore_errors=True)


@kew_bp.route('/export-mba', methods=['POST'])
def export_mba():
    """
    Nhận nhiều file INPS (.KEW hoặc .ZIP), mỗi file → 1 sheet trong output Excel.
    Form fields:
        files[]  – list các file KEW / ZIP (multipart, có thể gửi nhiều)
        sheets   – JSON array tên sheet tương ứng, ví dụ '["MBA1","MBA2"]'
                   Nếu thiếu / ngắn hơn số file, các sheet còn lại tự đặt tên MBA1, MBA2, ...
        filename – tên file xuất (mặc định 'NX-MBA.xlsm')
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
