import os
from pathlib import Path
from flask import Flask

BASE_DIR = Path(__file__).parent.parent

def create_app():
    app = Flask(__name__, template_folder="templates", static_folder=str(BASE_DIR / "static"))
    app.secret_key = "portfolio-analyzer-local"

    os.makedirs(BASE_DIR / "runs", exist_ok=True)
    os.makedirs(BASE_DIR / "uploads", exist_ok=True)

    try:
        from app.routes.upload import bp as upload_bp
        from app.routes.results import bp as results_bp
        from app.routes.property_detail import bp as property_bp
        from app.routes.history import bp as history_bp
        from app.routes.download import bp as download_bp

        for bp in (upload_bp, results_bp, property_bp, history_bp, download_bp):
            app.register_blueprint(bp)
    except ImportError:
        pass  # Routes not yet created

    return app
