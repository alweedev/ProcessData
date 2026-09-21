import os

from flask import Blueprint, abort, send_from_directory

from backend.core.config import settings
from backend.core.logging import get_logger

logger = get_logger()

frontend_bp = Blueprint("frontend", __name__)


@frontend_bp.route("/api/<path:path>", methods=["POST", "PUT", "PATCH", "DELETE"])
def api_desconhecida(path):
    # O curinga do SPA só aceita GET: sem esta rota, um POST numa rota /api removida responderia 405, não 404.
    abort(404)


@frontend_bp.route("/", defaults={"path": "index.html"})
@frontend_bp.route("/<path:path>")
def serve_frontend(path):
    frontend_dir = settings.FRONTEND_DIR
    root = os.path.realpath(frontend_dir)
    candidate = os.path.realpath(os.path.join(root, path))
    # Só serve arquivos que resolvem para dentro de frontend_dir (bloqueia ../).
    within_root = candidate == root or candidate.startswith(root + os.sep)
    if within_root and os.path.isfile(candidate):
        return send_from_directory(root, os.path.relpath(candidate, root))
    index_path = os.path.join(root, "index.html")
    if os.path.isfile(index_path):
        return send_from_directory(root, "index.html")
    abort(404)
