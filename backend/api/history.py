from flask import Blueprint, jsonify, request

from backend.services.audit_service import AuditService

history_bp = Blueprint("history", __name__, url_prefix="/api/history")


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
    AuditService.clear_events()
    return jsonify({"success": True}), 200
