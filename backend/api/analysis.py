import os

from flask import Blueprint, jsonify, request

from backend.core.config import settings
from backend.core.logging import get_logger
from backend.services.audit_service import AuditService
from backend.services.processing_service import ProcessingService
from backend.services.report_service import ReportService
from backend.shared.file_utils import gerar_nome_arquivo_temporario, validar_extensao_arquivo
from backend.shared.upload_validation import validar_conteudo_xlsx

logger = get_logger()

analysis_bp = Blueprint("analysis", __name__, url_prefix="/api")


@analysis_bp.route("/analysis/summary", methods=["POST"])
def analysis_summary():
    paths = []
    try:
        uploaded = request.files.getlist("files[]") or request.files.getlist("files")
        if not uploaded:
            return jsonify({"error": "Nenhum arquivo enviado"}), 400

        for file_item in uploaded:
            is_valid, error_msg = validar_extensao_arquivo(file_item.filename, {".xlsx", ".xls"})
            if not is_valid:
                return jsonify({"error": error_msg}), 400

        for file_item in uploaded:
            path = gerar_nome_arquivo_temporario(file_item.filename, settings.UPLOAD_FOLDER)
            file_item.save(path)
            paths.append(path)
            ok, msg = validar_conteudo_xlsx(path)
            if not ok:
                return jsonify({"error": msg}), 400

        login_choice = request.form.get("login_choice", "CPF")
        fluxo = request.form.get("fluxo", "SELF")

        errors, df_final = ProcessingService.process_records_from_files(
            paths,
            login_choice=login_choice,
            fluxo=fluxo,
        )

        if (df_final is None or df_final.empty) and errors:
            detail = "; ".join(f"{p}: {msg}" for p, msg in errors.items())
            # não grava `detail` no histórico: pode conter caminho de arquivo
            # do servidor, e /api/history é legível sem autenticação (GET).
            AuditService.record(
                event_type="analysis_summary",
                status="error",
                details={"files": len(uploaded), "errors_count": len(errors)},
            )
            return jsonify({"error": f"Falha ao processar arquivo(s): {detail}"}), 400

        report = ReportService.build_quality_report(df_final, errors)
        preview = []
        if df_final is not None and not df_final.empty:
            preview = df_final.head(20).to_dict(orient="records")

        AuditService.record(
            event_type="analysis_summary",
            status="success",
            details={
                "files": len(uploaded),
                "total_rows": report.get("total_rows", 0),
                "invalid_rows": report.get("invalid_rows", 0),
                "login_choice": login_choice,
                "fluxo": fluxo,
            },
        )

        return jsonify(
            {
                "report": report,
                "preview": preview,
            }
        ), 200
    except Exception:
        logger.exception("Erro em /api/analysis/summary")
        AuditService.record(event_type="analysis_summary", status="error", details={})
        return jsonify({"error": "Erro interno ao processar a solicitação."}), 500
    finally:
        for path in paths:
            try:
                if os.path.exists(path):
                    os.remove(path)
            except Exception:
                logger.warning("Falha ao remover temporário %s", path)
