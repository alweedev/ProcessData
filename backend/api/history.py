import hmac

from flask import Blueprint, jsonify, request

from backend.core.config import settings
from backend.core.logging import get_logger
from backend.services.audit_service import AuditService

logger = get_logger()

history_bp = Blueprint("history", __name__, url_prefix="/api/history")

_LOCAL_ADDRS = {"127.0.0.1", "::1", "localhost"}


def _delete_allowed() -> bool:
    """DELETE só de localhost ou com X-Admin-Token válido (se configurado)."""
    if (request.remote_addr or "") in _LOCAL_ADDRS:
        return True
    token = settings.HISTORY_ADMIN_TOKEN
    provided = request.headers.get("X-Admin-Token") or ""
    return bool(token) and hmac.compare_digest(provided, token)


@history_bp.route("", methods=["GET"])
def get_history():
    try:
        limit = int(request.args.get("limit", "200"))
    except Exception:
        limit = 200
    events = AuditService.list_events(limit=limit)
    return jsonify({"items": events, "total": len(events)}), 200


@history_bp.route("", methods=["DELETE"])
def clear_history():
    if not _delete_allowed():
        logger.warning("DELETE /api/history negado para %s", request.remote_addr)
        return jsonify({"error": "Não autorizado a limpar o histórico."}), 403
    AuditService.clear_events()
    return jsonify({"success": True}), 200
