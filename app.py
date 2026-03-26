from flask import Flask, request, jsonify, send_file
import os
import re
import io
import tempfile
import shutil
import zipfile
from analyse_kew import build_analysis, sanitize, generate_commentary
import sys
from PIL import Image
sys.path.append(os.path.join(os.path.dirname(__file__), 'scripts'))
import meter_editor

app = Flask(__name__, static_folder='.', static_url_path='')

@app.route('/')
def index():
    return send_file('dashboard.html')


def group_kew_files_by_id(file_paths):
    """
    Nhóm các file KEW (theo đường dẫn) theo ID thiết bị từ tên file.
    Ví dụ: INHS9951.KEW, INPS9951.KEW → nhóm '9951'
    Trả về: { id: [filepath, ...], ... }
    """
    groups = {}
    for fp in file_paths:
        fname = os.path.basename(fp)
        if not fname.upper().endswith('.KEW'):
            continue
        basename = os.path.splitext(fname)[0]
        match = re.match(r'^([A-Za-z]+)(.+)$', basename)
        file_id = match.group(2) if match else basename
        groups.setdefault(file_id, []).append(fp)
    return groups


def analyse_folder(folder_path, device_name=''):
    """Phân tích một thư mục KEW, trả về dict kết quả đã sanitize."""
    result = build_analysis(folder_path)
    if not result:
        return None
    sanitized = sanitize(result)
    # Tạo nhận xét tự động
    name = device_name or os.path.basename(folder_path.rstrip(os.sep))
    try:
        commentary = generate_commentary(result, device_name=name)
    except Exception as e:
        commentary = f"(Lỗi tạo nhận xét: {e})"
    sanitized['commentary'] = commentary
    sanitized['device_name'] = name
    return sanitized


def process_zip(zip_file_obj):
    """
    Giải nén ZIP vào temp dir. Hỗ trợ 2 cấu trúc:
    - Flat: tất cả KEW trong gốc ZIP
    - Thư mục con: mỗi thư mục là 1 bộ đo
    Trả về list kết quả.
    """
    temp_root = tempfile.mkdtemp(prefix='kew_zip_')
    results = []
    errors = []

    try:
        with zipfile.ZipFile(zip_file_obj, 'r') as zf:
            zf.extractall(temp_root)

        # Tìm tất cả KEW files trong temp_root (bao gồm thư mục con)
        all_kew = []
        for root, dirs, files in os.walk(temp_root):
            for fname in files:
                if fname.upper().endswith('.KEW'):
                    all_kew.append(os.path.join(root, fname))

        if not all_kew:
            return [], ["Không tìm thấy file .KEW nào trong ZIP."]

        # Kiểm tra cấu trúc: flat hay thư mục con
        # Nếu tất cả KEW nằm trong cùng 1 thư mục → flat hoặc 1 bộ
        kew_dirs = set(os.path.dirname(f) for f in all_kew)

        if len(kew_dirs) == 1:
            # Flat: tất cả trong 1 thư mục, nhóm theo ID
            flat_dir = list(kew_dirs)[0]
            groups = group_kew_files_by_id(all_kew)
            for device_id, paths in groups.items():
                sub_dir = tempfile.mkdtemp(prefix=f'kew_{device_id}_', dir=temp_root)
                for p in paths:
                    shutil.copy2(p, sub_dir)
                r = analyse_folder(sub_dir, device_name=device_id)
                if r:
                    results.append(r)
                else:
                    errors.append(f"ID {device_id}: phân tích thất bại.")
        else:
            # Mỗi thư mục là 1 bộ đo khác nhau
            for kew_dir in sorted(kew_dirs):
                dir_name = os.path.basename(kew_dir)
                r = analyse_folder(kew_dir, device_name=dir_name)
                if r:
                    results.append(r)
                else:
                    # Có thể là thư mục cha không có INHS nhưng con có
                    pass

        # Nếu chưa có kết quả nào, thử phân tích từng thư mục có INHS
        if not results:
            for kew_dir in sorted(kew_dirs):
                inhs = [f for f in os.listdir(kew_dir) if f.upper().startswith('INHS')]
                if inhs:
                    r = analyse_folder(kew_dir)
                    if r:
                        results.append(r)

    except zipfile.BadZipFile:
        errors.append("File upload không phải định dạng ZIP hợp lệ.")
    except Exception as e:
        errors.append(f"Lỗi xử lý ZIP: {str(e)}")
    finally:
        shutil.rmtree(temp_root, ignore_errors=True)

    return results, errors


@app.route('/api/upload', methods=['POST'])
def upload_files():
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


@app.route('/api/fix', methods=['POST'])
def fix_files():
    import interpolate_kew
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


@app.route('/api/detect', methods=['POST'])
def detect_phases():
    """Detect which current phases are missing in an uploaded KEW set, without full analysis."""
    import interpolate_kew
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


@app.route('/api/correct', methods=['POST'])
def correct_files():
    """Apply per-channel multiplier/offset corrections and return corrected KEW files as ZIP."""
    import correct_kew
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

@app.route('/api/edit-meter-images', methods=['POST'])
def edit_meter_images():
    # ... (existing code for ZIP)
    try:
        if 'zip' not in request.files:
            return jsonify({'error': 'Không có file ZIP nào được upload'}), 400
            
        zip_file = request.files['zip']
        params = request.form.to_dict()
        if 'fluctuate' in params:
            params['fluctuate'] = params['fluctuate'] == 'true'
            
        zip_bytes = io.BytesIO(zip_file.read())
        out_zip_bytes = meter_editor.process_zip(zip_bytes, params)
        
        return send_file(out_zip_bytes, download_name='Edited_Meter_Images.zip', as_attachment=True, mimetype='application/zip')
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Lỗi chỉnh sửa ảnh: {str(e)}'}), 500

@app.route('/api/edit-meter-image', methods=['POST'])
def edit_meter_image():
    try:
        if 'image' not in request.files:
            return jsonify({'error': 'Không có file ảnh nào được upload'}), 400
            
        file = request.files['image']
        idx = int(request.form.get('idx', 0))
        params = request.form.to_dict()
        if 'fluctuate' in params:
            params['fluctuate'] = params['fluctuate'] == 'true'
            
        img = Image.open(file).convert('RGB')
        img = meter_editor.process_image(img, idx, params)
        
        img_io = io.BytesIO()
        img.save(img_io, 'BMP')
        img_io.seek(0)
        
        return send_file(img_io, mimetype='image/bmp')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/edit-meter-images-bulk', methods=['POST'])
def edit_meter_images_bulk():
    try:
        files = request.files.getlist('files')
        if not files:
            return jsonify({'error': 'Không có file nào được upload'}), 400
            
        params = request.form.to_dict()
        if 'fluctuate' in params:
            params['fluctuate'] = params['fluctuate'] == 'true'
            
        # Create an in-memory ZIP from the uploaded files
        in_zip_bytes = io.BytesIO()
        with zipfile.ZipFile(in_zip_bytes, 'w') as zf:
            for idx, f in enumerate(files):
                # Use filename or just index if missing
                fname = f.filename or f"image_{idx}.bmp"
                zf.writestr(fname, f.read())
        
        in_zip_bytes.seek(0)
        out_zip_bytes = meter_editor.process_zip(in_zip_bytes, params)
        
        return send_file(out_zip_bytes, download_name='Edited_Meter_Images.zip', as_attachment=True, mimetype='application/zip')
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500




if __name__ == '__main__':
    print("Khởi động KEW Server trên cổng 5000...")
    app.run(host='0.0.0.0', port=5000, debug=True)
