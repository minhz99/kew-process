import os
import io
import json
from flask import Blueprint, jsonify, current_app, request, send_file
from PIL import Image, ImageDraw
from modules.image.kew6315_layout import SCREENS

image_bp = Blueprint('image_bp', __name__)

# ── Toạ độ dán chữ số thời gian ────────────────────────────────────────────────
# Mỗi chữ số PNG 6×9 px; các số được dán từ trái sang phải từ điểm (x, y).
TIME_FIELDS = {
    # Ngày
    'dd': {'x': 258, 'y': 1},
    'mo': {'x': 276, 'y': 1},  # tháng (mm)
    'yyyy': {'x': 294, 'y': 1},
    # Giờ
    'hh': {'x': 264, 'y': 11},
    'mi': {'x': 282, 'y': 11},  # phút (mm)
    'ss': {'x': 300, 'y': 11},
}
DIGIT_W = 6
DIGIT_H = 9

@image_bp.route('/templates', methods=['GET'])
def get_templates():
    """Trả về danh sách các mẫu đồng hồ hỗ trợ (dùng cho mở rộng sau này)."""
    templates = [
        {"id": "kew6315", "name": "Kyoritsu KEW 6315"},
        {"id": "kew6305", "name": "Kyoritsu KEW 6305"},
        {"id": "hioki3198", "name": "Hioki PQ3198"},
        {"id": "chauvin", "name": "Chauvin Arnoux C.A 8336"}
    ]
    return jsonify(templates)

CHAR_MAP = {'.': 'dot', '-': 'minus'}
_DIGIT_TEMPLATES = {}

def get_digit_img(char, color, digits_dir):
    """
    Lấy hình ảnh của một chữ số hoặc ký tự đặc biệt (chấm, trừ) với màu sắc tương ứng.
    
    Args:
        char (str): Ký tự cần lấy (0-9, '.', '-').
        color (str): Mã màu ('w' cho trắng, 'g' cho xanh lục...).
        digits_dir (str): Thư mục chứa các file ảnh chữ số.
        
    Returns:
        PIL.Image: Ảnh chữ số đã load hoặc None nếu không tìm thấy.
    """
    s = CHAR_MAP.get(char, char)
    key = f"{s}_{color}"
    if key in _DIGIT_TEMPLATES:
        return _DIGIT_TEMPLATES[key]
        
    filename = f"{s}{color}.bmp"
    filepath = os.path.join(digits_dir, filename)
    if os.path.exists(filepath):
        img = Image.open(filepath).convert("RGBA")
        _DIGIT_TEMPLATES[key] = img
        return img
        
    fallback_color = 'g' if color == 'w' else 'w'
    key_fall = f"{s}_{fallback_color}"
    if key_fall in _DIGIT_TEMPLATES:
        return _DIGIT_TEMPLATES[key_fall]
        
    filename = f"{s}{fallback_color}.bmp"
    filepath = os.path.join(digits_dir, filename)
    if os.path.exists(filepath):
        img = Image.open(filepath).convert("RGBA")
        _DIGIT_TEMPLATES[key_fall] = img
        return img
    
    return None

def apply_text_to_image(img, img_draw, config, text, digits_dir):
    """
    Dán văn bản (số) lên ảnh dựa trên cấu hình tọa độ và thư mục chứa chữ số.
    
    Args:
        img (PIL.Image): Ảnh gốc.
        img_draw (PIL.ImageDraw): Đối tượng vẽ lên ảnh.
        config (dict): Cấu hình vị trí và thuộc tính của overlay.
        text (str): Văn bản/số cần dán.
        digits_dir (str): Thư mục chứa ảnh chữ số.
    """
    x_right = config['x']
    y_bot = config['y']
    color = config.get('bg', 'w')
    w_clear = config.get('w_clear', 50)
    h_clear = 15
    
    x_left = max(0, x_right - w_clear + 1)
    y_top = max(0, y_bot - h_clear + 1)

    pixel_color = img.getpixel((x_left, y_bot))
    img_draw.rectangle([x_left, y_top, x_left + w_clear - 1, y_top + h_clear - 1], fill=pixel_color)

    normalized_text = str(text).replace(',', '.')
    chars = list(normalized_text)[::-1]
    curr_x = x_right + 1

    for char in chars:
        c = '.' if char == '/' else char
        digit_img = get_digit_img(c, color, digits_dir)
        if digit_img:
            dw = digit_img.width
            dh = digit_img.height
            spacing = 1 if dw >= 8 else 2

            curr_x -= dw
            paste_y = y_bot - dh + 1
                
            img.paste(digit_img, (curr_x, paste_y), digit_img)
            curr_x -= spacing
        else:
            curr_x -= 6

def _safe_screen(idx):
    if not SCREENS:
        return {'overlays': []}
    return SCREENS[idx % len(SCREENS)]


def _build_bmp_name(filename):
    return filename or 'edited.bmp'


@image_bp.route('/process', methods=['POST'])
def process_image():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
        
    file = request.files['file']
    screen_idx_str = request.form.get('screenIdx', '0')
    params_str = request.form.get('parameters', '{}')
    meter_model = request.form.get('meterModel', 'kew6315')
    
    try:
        screen_idx = int(screen_idx_str)
        params = json.loads(params_str)
    except (ValueError, json.JSONDecodeError):
        return jsonify({"error": "Invalid screen_idx or parameters"}), 400
        
    try:
        original_img = Image.open(file).convert("RGB")
    except Exception:
        return jsonify({"error": "Invalid image file"}), 400
        
    if meter_model != 'kew6315':
        current_app.logger.info("meterModel=%s chưa hỗ trợ riêng, dùng layout kew6315", meter_model)

    sc = _safe_screen(screen_idx)
        
    digits_dir = os.path.join(current_app.static_folder, 'digits')
    img_draw = ImageDraw.Draw(original_img)
    
    for overlay in sc.get('overlays', []):
        val = params.get(overlay['id'])
        if val is None and 'alias' in overlay:
            val = params.get(overlay['alias'])
        
        if val is not None and str(val).strip() != "":
            apply_text_to_image(original_img, img_draw, overlay, val, digits_dir)
            
    buf = io.BytesIO()
    original_img.save(buf, format='BMP')
    buf.seek(0)
    
    return send_file(
        buf,
        mimetype='image/bmp',
        as_attachment=True,
        download_name=_build_bmp_name(getattr(file, 'filename', 'edited.bmp')),
    )


# ── Bộ cache ảnh chữ số thời gian ─────────────────────────────────────────────
_TIME_DIGIT_CACHE = {}

def _get_time_digit(digit_char, time_digits_dir):
    """Trả về RGBA PIL Image cho chữ số 0-9 từ thư mục time-digits."""
    if digit_char in _TIME_DIGIT_CACHE:
        return _TIME_DIGIT_CACHE[digit_char]
    path = os.path.join(time_digits_dir, f"{digit_char}.png")
    if not os.path.exists(path):
        return None
    img = Image.open(path).convert('RGBA')
    _TIME_DIGIT_CACHE[digit_char] = img
    return img


def _paste_digits_in_field(target_img, digits_str, field_cfg, time_digits_dir):
    """Dán từng chữ số 6×9 vào field_cfg từ trái sang phải."""
    x_cursor = field_cfg['x']
    y_top = field_cfg['y']
    for ch in digits_str:
        if not ch.isdigit():
            continue
        d_img = _get_time_digit(ch, time_digits_dir)
        if d_img is None:
            x_cursor += DIGIT_W
            continue
        # Resize nếu kích thước khác 6×9 (dự phòng)
        if d_img.width != DIGIT_W or d_img.height != DIGIT_H:
            d_img = d_img.resize((DIGIT_W, DIGIT_H), Image.NEAREST)
        mask = d_img.split()[3]
        target_img.paste(d_img, (x_cursor, y_top), mask)
        x_cursor += DIGIT_W


def _parse_timestamp_values(form):
    """
    Phân tích các giá trị timestamp từ form request.
    
    Returns:
        dict: Một dictionary chứa các chuỗi đã chuẩn hóa (zfill) cho dd, mo, yyyy, hh, mi, ss.
    """
    timestamp = form.get('timestamp', '').strip()
    if timestamp:
        try:
            date_part, time_part = timestamp.split(' ', 1)
            dd_s, mo_s, yyyy_s = date_part.split('/')
            hh_s, mi_s, ss_s = time_part.split(':')
        except ValueError as exc:
            raise ValueError('Định dạng timestamp phải là dd/mm/yyyy hh:mm:ss') from exc
    else:
        dd_s = form.get('dd', '')
        mo_s = form.get('mo', '')
        yyyy_s = form.get('yyyy', '')
        hh_s = form.get('hh', '')
        mi_s = form.get('mi', '')
        ss_s = form.get('ss', '')

    return {
        'dd': dd_s.zfill(2),
        'mo': mo_s.zfill(2),
        'yyyy': yyyy_s.zfill(4),
        'hh': hh_s.zfill(2),
        'mi': mi_s.zfill(2),
        'ss': ss_s.zfill(2),
    }


@image_bp.route('/apply-timestamp', methods=['POST'])
def apply_timestamp():
    """Dán chữ số thời gian (dd/mm/yyyy hh:mm:ss) vào ảnh BMP.

    Form fields:
        file       – ảnh BMP gốc
        timestamp  – chuỗi có dạng 'dd/mm/yyyy hh:mm:ss'
                     (có thể truyền riêng từng phần qua dd, mo, yyyy, hh, mi, ss)
    """
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']
    try:
        original_img = Image.open(file).convert('RGB')
    except Exception:
        return jsonify({'error': 'Invalid image file'}), 400

    try:
        values = _parse_timestamp_values(request.form)
    except ValueError as exc:
        return jsonify({'error': str(exc)}), 400

    time_digits_dir = os.path.join(current_app.static_folder, 'time-digits')
    if not os.path.isdir(time_digits_dir):
        return jsonify({'error': 'Thiếu thư mục static/time-digits trên server'}), 500

    rgba_img = original_img.convert('RGBA')

    for field_key, field_cfg in TIME_FIELDS.items():
        val = values.get(field_key, '')
        if val:
            _paste_digits_in_field(rgba_img, val, field_cfg, time_digits_dir)

    out_img = rgba_img.convert('RGB')
    buf = io.BytesIO()
    out_img.save(buf, format='BMP')
    buf.seek(0)

    return send_file(
        buf,
        mimetype='image/bmp',
        as_attachment=True,
        download_name=_build_bmp_name(getattr(file, 'filename', 'edited.bmp')),
    )
