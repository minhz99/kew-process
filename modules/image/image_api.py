import os
import base64
import io
import json
from flask import Blueprint, jsonify, current_app, request, send_file
from PIL import Image, ImageDraw

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

def make_grid(ids, x_rights, y_bot, bg, scale=0.96):
    return [{"id": id_, "x": x, "y": y_bot, "bg": bg, "scale": scale} for id_, x in zip(ids, x_rights)]

def _map_sd140(o):
    if o["id"] in ["PF1", "PF2", "PF3"]:
        o["w_clear"] = 55
    return o

SCREENS = [
    {
        "id": "SD140",
        "overlays": list(map(_map_sd140, 
            make_grid(["V1", "V2", "V3"], [94, 158, 222], 54, "w") +
            make_grid(["A1", "A2", "A3"], [94, 158, 222], 70, "g") +
            make_grid(["P1", "P2", "P3"], [94, 158, 222], 86, "w") +
            make_grid(["Q1", "Q2", "Q3"], [94, 158, 222], 102, "g") +
            make_grid(["S1", "S2", "S3"], [94, 158, 222], 118, "w") +
            make_grid(["PF1", "PF2", "PF3"], [94, 158, 222], 134, "g") +
            [
                {"id": "P", "x": 94, "y": 153, "bg": "w", "scale": 0.96},
                {"id": "freq", "alias": "f", "x": 222, "y": 153, "bg": "w", "scale": 0.96},
                {"id": "Q", "x": 94, "y": 169, "bg": "g", "scale": 0.96},
                {"id": "S", "x": 94, "y": 185, "bg": "w", "scale": 0.96},
                {"id": "PF", "x": 94, "y": 201, "bg": "g", "scale": 0.96, "w_clear": 55},
                {"id": "An", "x": 222, "y": 201, "bg": "g", "scale": 0.96}
            ]
        ))
    },
    {
        "id": "SD141",
        "overlays": [
          {"id": "V1", "x": 63, "y": 36, "bg": "w", "scale": 0.85, "w_clear": 45},
          {"id": "Vdeg1", "x": 121, "y": 36, "bg": "w", "scale": 0.85, "w_clear": 45},
          {"id": "V2", "x": 63, "y": 52, "bg": "g", "scale": 0.85, "w_clear": 45},
          {"id": "Vdeg2", "x": 121, "y": 52, "bg": "g", "scale": 0.85, "w_clear": 45},
          {"id": "V3", "x": 63, "y": 68, "bg": "w", "scale": 0.85, "w_clear": 45},
          {"id": "Vdeg3", "x": 121, "y": 68, "bg": "w", "scale": 0.85, "w_clear": 45},
          {"id": "A1", "x": 63, "y": 87, "bg": "w", "scale": 0.85, "w_clear": 45},
          {"id": "Adeg1", "x": 121, "y": 87, "bg": "w", "scale": 0.85, "w_clear": 45},
          {"id": "A2", "x": 63, "y": 103, "bg": "g", "scale": 0.85, "w_clear": 45},
          {"id": "Adeg2", "x": 121, "y": 103, "bg": "g", "scale": 0.85, "w_clear": 45},
          {"id": "A3", "x": 63, "y": 119, "bg": "w", "scale": 0.85, "w_clear": 45},
          {"id": "Adeg3", "x": 121, "y": 119, "bg": "w", "scale": 0.85, "w_clear": 45},
          {"id": "freq", "alias": "f", "x": 83, "y": 154, "bg": "w", "scale": 0.85, "w_clear": 45},
          {"id": "V_unb", "alias": "V%", "x": 83, "y": 189, "bg": "g", "scale": 0.85, "w_clear": 45},
          {"id": "A_unb", "alias": "A%", "x": 83, "y": 205, "bg": "w", "scale": 0.85, "w_clear": 45}
        ]
    },
    { 
        "id": "SD142", 
        "overlays": make_grid(["V1", "V2", "V3"], [76, 136, 196], 47, "w") + make_grid(["A1", "A2", "A3"], [76, 136, 196], 63, "g")
    },
    {
        "id": "SD143", 
        "overlays": make_grid(["V1", "V2", "V3"], [76, 136, 196], 47, "w") + make_grid(["A1", "A2", "A3"], [76, 136, 196], 63, "g")
    },
    {
        "id": "SD144", 
        "overlays": make_grid(["V1", "V2", "V3"], [76, 136, 196], 47, "w") + make_grid(["THDV1", "THDV2", "THDV3"], [76, 136, 196], 63, "g")
    },
    {
        "id": "SD145", 
        "overlays": make_grid(["A1", "A2", "A3"], [76, 136, 196], 47, "w") + make_grid(["THDA1", "THDA2", "THDA3"], [76, 136, 196], 63, "g")
    }
]

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
    scale = config.get('scale', 1.0)
    
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
            dw = int(digit_img.width * scale)
            dh = int(digit_img.height * scale)
            spacing = 1 if dw >= 8 else 2

            curr_x -= dw
            paste_y = y_bot - dh + 1
            if scale != 1.0:
                digit_resized = digit_img.resize((dw, dh), Image.NEAREST)
            else:
                digit_resized = digit_img
                
            img.paste(digit_resized, (curr_x, paste_y), digit_resized)
            curr_x -= spacing
        else:
            curr_x -= int(6 * scale)

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
