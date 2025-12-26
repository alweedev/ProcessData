import os
from flask import Blueprint, send_from_directory, abort, jsonify, request
from backend.core.config import settings
from backend.core.logging import get_logger

logger = get_logger()

frontend_bp = Blueprint('frontend', __name__)


@frontend_bp.route('/', defaults={'path': 'index.html'})
@frontend_bp.route('/<path:path>')
def serve_frontend(path):
    frontend_dir = settings.FRONTEND_DIR
    full_path = os.path.join(frontend_dir, path)
    if os.path.exists(full_path) and os.path.isfile(full_path):
        return send_from_directory(frontend_dir, path)
    index_path = os.path.join(frontend_dir, 'index.html')
    if os.path.exists(index_path):
        return send_from_directory(frontend_dir, 'index.html')
    abort(404)


@frontend_bp.route('/health', methods=['GET', 'HEAD'])
def health_check():
    return ('', 204) if request.method == 'HEAD' else jsonify({'status': 'OK', 'message': 'Servidor rodando'}), 200
