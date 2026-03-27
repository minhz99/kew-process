from flask import Flask, render_template
from modules.excel.excel_api import excel_bp
from modules.kew.kew_api import kew_bp
from modules.image.image_api import image_bp

app = Flask(__name__)

# Register Blueprints for specialized toolset
app.register_blueprint(excel_bp, url_prefix='/api/excel')
app.register_blueprint(kew_bp, url_prefix='/api/kew')
app.register_blueprint(image_bp, url_prefix='/api/image')

@app.route('/')
def index():
    """Render the main dashboard UI application."""
    return render_template('dashboard.html')

if __name__ == '__main__':
    print("Khởi động PLT Process Server trên cổng 5525...")
    app.run(host='0.0.0.0', port=5525, debug=True)
