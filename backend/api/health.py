from flask import Blueprint, jsonify, request

health_bp = Blueprint('health', __name__, url_prefix='/api')


@health_bp.route('/health', methods=['GET', 'HEAD'])
def api_health():
    if request.method == 'HEAD':
        return ('', 204)
    return jsonify({'status': 'OK', 'message': 'Servidor rodando'}), 200