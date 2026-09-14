import os

from flask import Blueprint, jsonify, request, send_file

from backend.core.config import settings
from backend.core.logging import get_logger
from backend.services.audit_service import AuditService
from backend.services.export_service import ExportService
from backend.services.processing_service import ProcessingService
from backend.shared.upload_validation import save_and_validate_upload

logger = get_logger()

cadastro_bp = Blueprint("cadastro", __name__, url_prefix="/api")


@cadastro_bp.route("/process_cadastro", methods=["POST"])
def api_process_cadastro():
    paths = []
    try:
        uploaded = request.files.getlist("files[]") or request.files.getlist("files")
        if not uploaded:
            return jsonify({"error": "Nenhum arquivo enviado"}), 400

        for f in uploaded:
            p, err = save_and_validate_upload(f, settings.UPLOAD_FOLDER)
            if p:
                paths.append(p)
            if err:
                return jsonify({"error": err}), 400

        login_choice = request.form.get("login_choice", "CPF")
        fluxo = request.form.get("fluxo", "SELF")

        errors, df_final = ProcessingService.process_records_from_files(
            paths,
            login_choice=login_choice,
            fluxo=fluxo,
        )

        if df_final.empty:
            AuditService.record(
                event_type="cadastro",
                status="empty",
                details={"files": len(uploaded), "errors": errors},
            )
            return jsonify({"error": "Nenhum registro processado", "errors": errors}), 400

        output = ExportService.to_excel_bytes(df_final, sheet_name="Cadastro")

        AuditService.record(
            event_type="cadastro",
            status="success",
            details={
                "files": len(uploaded),
                "rows": int(df_final.shape[0]),
                "columns": int(df_final.shape[1]),
                "login_choice": login_choice,
                "fluxo": fluxo,
                "invalid_rows": len(errors),
                "errors": errors,
            },
        )
        return send_file(
            output,
            download_name="saida_cadastro.xlsx",
            as_attachment=True,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        logger.exception("Erro em /api/process_cadastro")
        AuditService.record(
            event_type="cadastro",
            status="error",
            details={"message": str(e)},
        )
        return jsonify({"error": "Erro interno ao processar a solicitação."}), 500
    finally:
        for p in paths:
            try:
                if os.path.exists(p):
                    os.remove(p)
            except Exception:
                logger.warning(f"Falha ao remover temporário {p}")
