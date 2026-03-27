import os
import re
import tempfile
import shutil
import zipfile
from modules.kew.analyse_kew import build_analysis, sanitize, generate_commentary

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
