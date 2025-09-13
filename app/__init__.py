import os
from flask import Flask

_THIS_DIR = os.path.abspath(os.path.dirname(__file__))
_TEMPLATE_DIR = os.path.join(_THIS_DIR, "templates")
_STATIC_DIR = os.path.join(_THIS_DIR, "static")


def create_app():
    app = Flask(__name__, template_folder=_TEMPLATE_DIR, static_folder=_STATIC_DIR)

    app.config.update(
        SECRET_KEY="change-this-in-prod",
        MAX_CONTENT_LENGTH=512 * 1024 * 1024,  # 512MB
    )

    # 업로드 및 출력 경로 설정
    root = os.path.abspath(os.path.join(_THIS_DIR, ".."))
    uploads_root = os.path.join(root, "uploads")
    outputs_root = os.path.join(root, "outputs")

    app.config["UPLOAD_ORIGINAL_DIR"] = os.path.join(uploads_root, "original")
    app.config["UPLOAD_REVISED_DIR"] = os.path.join(uploads_root, "revised")
    app.config["OUTPUT_DIR"] = outputs_root

    for p in [uploads_root, outputs_root,
              app.config["UPLOAD_ORIGINAL_DIR"],
              app.config["UPLOAD_REVISED_DIR"]]:
        os.makedirs(p, exist_ok=True)

    from .routes import bp as main_bp
    app.register_blueprint(main_bp)

    return app