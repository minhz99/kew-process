from flask import Blueprint, jsonify

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
