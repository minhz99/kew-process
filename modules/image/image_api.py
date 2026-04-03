import os
import base64
import io
from flask import Blueprint, jsonify, current_app
from PIL import Image

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
    # Determine absolute path to the static/digits directory safely
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
    digits_dir = os.path.join(base_dir, 'static', 'digits')
    
    result = {}
    
    if not os.path.exists(digits_dir):
        current_app.logger.error(f"Thư mục digits không tồn tại: {digits_dir}")
        return jsonify(result)

    symbols = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'dot', 'minus']
    colors = ['w', 'g']

    try:
        dir_files = os.listdir(digits_dir)
    except Exception as e:
        current_app.logger.error(f"Không thể đọc thư mục {digits_dir}: {e}")
        return jsonify(result)

    for s in symbols:
        for c in colors:
            target_lower = f"{s}{c}.bmp".lower()
            
            # Case-insensitive resolution for Linux deployment safety
            actual_filename = next((f for f in dir_files if f.lower() == target_lower), None)
            
            if not actual_filename:
                continue
                
            filepath = os.path.join(digits_dir, actual_filename)
            try:
                img = Image.open(filepath).convert('RGBA')
                buf = io.BytesIO()
                img.save(buf, format='PNG')
                b64 = base64.b64encode(buf.getvalue()).decode('utf-8')
                key = f"{s}_{c}"
                result[key] = f"data:image/png;base64,{b64}"
            except Exception as e:
                current_app.logger.warning(f"Không thể đọc digit {actual_filename}: {e}")

    return jsonify(result)
