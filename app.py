import os

from flask import Flask, jsonify, render_template, request
from werkzeug.exceptions import RequestEntityTooLarge

from modules.excel.excel_api import excel_bp
from modules.kew.kew_api import kew_bp
from modules.image.image_api import image_bp

def _env_int(name, default):
    """
    Chuyển đổi biến môi trường sang kiểu integer.
    
    Args:
        name (str): Tên biến môi trường.
        default (int): Giá trị mặc định nếu biến không tồn tại hoặc lỗi chuyển đổi.
        
    Returns:
        int: Giá trị integer của biến môi trường hoặc giá trị mặc định.
    """
    try:
        return int(os.environ.get(name, default))
    except (TypeError, ValueError):
        return default


app = Flask(__name__)
app.json.ensure_ascii = False
app.config["MAX_CONTENT_LENGTH"] = _env_int("MAX_UPLOAD_MB", 256) * 1024 * 1024
app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 0

# Register Blueprints for specialized toolset
app.register_blueprint(excel_bp, url_prefix='/api/excel')
app.register_blueprint(kew_bp, url_prefix='/api/kew')
app.register_blueprint(image_bp, url_prefix='/api/image')


@app.errorhandler(RequestEntityTooLarge)
def handle_request_entity_too_large(_exc):
    """
    Xử lý lỗi khi file upload vượt quá kích thước cho phép.
    
    Args:
        _exc: Đối tượng exception.
        
    Returns:
        Response: Thông báo lỗi dạng JSON hoặc text với mã trạng thái 413.
    """
    max_upload_mb = app.config["MAX_CONTENT_LENGTH"] // (1024 * 1024)
    if request.path.startswith("/api/"):
        return jsonify({"error": f"File upload vượt quá giới hạn {max_upload_mb} MB của server."}), 413
    return f"File upload vượt quá giới hạn {max_upload_mb} MB của server.", 413

@app.route('/')
def index():
    """Render the main dashboard UI application."""
    return render_template('dashboard.html')

if __name__ == '__main__':
    host = os.environ.get("HOST", "0.0.0.0")
    port = _env_int("PORT", 5525)
    debug = os.environ.get("FLASK_DEBUG", "").strip().lower() in {"1", "true", "yes", "on"}
    print(f"Khởi động PLT Process Server trên {host}:{port}...")
    app.run(host=host, port=port, debug=debug)
