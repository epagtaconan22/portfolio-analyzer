import os
from flask import Flask

def create_app():
    app = Flask(__name__, template_folder="templates", static_folder="../static")
    app.secret_key = "portfolio-analyzer-local"

    os.makedirs("runs", exist_ok=True)
    os.makedirs("uploads", exist_ok=True)

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
