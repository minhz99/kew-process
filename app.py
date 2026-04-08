import os

from flask import Flask, jsonify, render_template, request
from werkzeug.exceptions import RequestEntityTooLarge

from modules.excel.excel_api import excel_bp
from modules.kew.kew_api import kew_bp
from modules.image.image_api import image_bp
from modules.synopex.synopex_api import synopex_bp


def _env_int(name, default):
    try:
        return int(os.environ.get(name, default))
    except (TypeError, ValueError):
        return default


app = Flask(__name__)
app.json.ensure_ascii = False
app.config["MAX_CONTENT_LENGTH"] = _env_int("MAX_UPLOAD_MB", 256) * 1024 * 1024

# Register Blueprints for specialized toolset
app.register_blueprint(excel_bp, url_prefix='/api/excel')
app.register_blueprint(kew_bp, url_prefix='/api/kew')
app.register_blueprint(image_bp, url_prefix='/api/image')
app.register_blueprint(synopex_bp, url_prefix='/api/synopex')


@app.errorhandler(RequestEntityTooLarge)
def handle_request_entity_too_large(_exc):
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
