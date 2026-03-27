import os
import io
import shutil
import tempfile
import zipfile
import re
from flask import Blueprint, request, jsonify, send_file
from utils.file_utils import process_zip, group_kew_files_by_id, analyse_folder

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
        import traceback
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
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500
    finally:
        shutil.rmtree(temp_in, ignore_errors=True)


@kew_bp.route('/correct', methods=['POST'])
def correct_files():
    """Apply per-channel multiplier/offset corrections and return corrected KEW files as ZIP."""
    from modules.kew import correct_kew
    import json
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
        import traceback; traceback.print_exc()
        return jsonify({'error': f'Lỗi hiệu chỉnh: {str(e)}'}), 500
    finally:
        shutil.rmtree(temp_in, ignore_errors=True)
        shutil.rmtree(temp_out, ignore_errors=True)
