import os
import base64
import io
import json
from flask import Blueprint, jsonify, current_app, request, send_file
from PIL import Image, ImageDraw
from modules.image.kew6315_layout import SCREENS

image_bp = Blueprint('image_bp', __name__)

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


@image_bp.route('/digits', methods=['GET'])
def get_digits():
    """Trả về toàn bộ digit templates dạng base64 PNG để client-side dùng cho canvas."""
    digits_dir = os.path.join(current_app.static_folder, 'digits')
    result = {}

    symbols = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'dot', 'minus']
    colors = ['w', 'g']

    for s in symbols:
        for c in colors:
            filename = f"{s}{c}.bmp"
            filepath = os.path.join(digits_dir, filename)
            if not os.path.exists(filepath):
                continue
            try:
                img = Image.open(filepath).convert('RGBA')
                buf = io.BytesIO()
                img.save(buf, format='PNG')
                b64 = base64.b64encode(buf.getvalue()).decode('utf-8')
                key = f"{s}_{c}"
                result[key] = f"data:image/png;base64,{b64}"
            except Exception as e:
                current_app.logger.warning(f"Không thể đọc digit {filename}: {e}")

    return jsonify(result)

CHAR_MAP = {'.': 'dot', '-': 'minus'}
_DIGIT_TEMPLATES = {}

def get_digit_img(char, color, digits_dir):
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
    except:
        return jsonify({"error": "Invalid screen_idx or parameters"}), 400
        
    try:
        original_img = Image.open(file).convert("RGB")
    except Exception as e:
        return jsonify({"error": "Invalid image file"}), 400
        
    if meter_model == 'kew6315':
        sc = SCREENS[screen_idx % 6] if (screen_idx % 6) < len(SCREENS) else SCREENS[0]
    else:
        sc = SCREENS[screen_idx % 6] if (screen_idx % 6) < len(SCREENS) else SCREENS[0]
        
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
    
    # Sửa filename (nếu có .bmp thì xoá hoặc đổi thành Edit_)
    fname = getattr(file, 'filename', 'edited.bmp')
    if not fname.lower().endswith('.bmp'):
        fname += '.bmp'
        
    return send_file(buf, mimetype='image/bmp', as_attachment=True, download_name=f"Edited_{fname}")
