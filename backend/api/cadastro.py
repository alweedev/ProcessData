import os

from flask import Blueprint, jsonify, request, send_file

from backend.core.config import settings
from backend.core.logging import get_logger
from backend.services.audit_service import AuditService
from backend.services.export_service import ExportService
from backend.services.processing_service import ProcessingService
from backend.shared.cadastro_request import (
    display_name,
    parse_login_and_fluxo,
    parse_name_overrides,
    user_facing_file_errors,
)
from backend.shared.name_splitter import ACCEPTED_WEIGHT, EDITED_WEIGHT, MAX_FIELD_LEN, get_vocabulary
from backend.shared.upload_validation import save_and_validate_upload

logger = get_logger()

cadastro_bp = Blueprint("cadastro", __name__, url_prefix="/api")


@cadastro_bp.route("/process_cadastro", methods=["POST"])
def api_process_cadastro():
    paths = []
    names = {}  # caminho temporário no servidor -> nome que o usuário enviou
    try:
        uploaded = request.files.getlist("files[]") or request.files.getlist("files")
        if not uploaded:
            return jsonify({"error": "Nenhum arquivo enviado"}), 400

        for f in uploaded:
            p, err = save_and_validate_upload(f, settings.UPLOAD_FOLDER)
            if p:
                paths.append(p)
                names[p] = display_name(f.filename)
            if err:
                return jsonify({"error": err}), 400

        login_choice, fluxo, param_error = parse_login_and_fluxo(request.form)
        if param_error:
            return jsonify({"error": param_error}), 400

        name_overrides, overrides_error = parse_name_overrides(request.form)
        if overrides_error:
            return jsonify({"error": overrides_error}), 400

        errors, df_final = ProcessingService.process_records_from_files(
            paths,
            login_choice=login_choice,
            fluxo=fluxo,
            name_overrides=name_overrides,
        )

        if df_final.empty:
            AuditService.record(
                event_type="cadastro",
                status="empty",
                # não grava o dict de erros no histórico: em caso de falha de
                # leitura ele é indexado por caminho de arquivo do servidor,
                # e /api/history é legível sem autenticação (GET).
                details={"files": len(uploaded), "errors_count": len(errors)},
            )
            # A resposta também vai por nome de arquivo: o caminho temporário é do servidor.
            file_errors = user_facing_file_errors(errors, names, settings.UPLOAD_FOLDER)
            detail = "; ".join(f"{name}: {msg}" for name, msg in file_errors.items())
            message = f"Nenhum registro processado: {detail or 'nenhuma linha com dados foi encontrada.'}"
            return jsonify({"error": message, "errors": file_errors}), 400

        # Nome ou Sobrenome acima do limite da plataforma não sobe: a tela só libera o Gerar depois do ajuste, e aqui
        # a regra vale também para quem chamar a rota direto (o nome não é cortado em silêncio).
        too_long = [
            item
            for item in df_final.attrs.get("name_review", [])
            if item["estouro"]["nome"] or item["estouro"]["sobrenome"]
        ]
        if too_long:
            where = ", ".join(item["label"] for item in too_long[:5]) + ("…" if len(too_long) > 5 else "")
            plural = len(too_long) > 1
            return (
                jsonify(
                    {
                        "error": (
                            f"{len(too_long)} nome{'s' if plural else ''} com mais de {MAX_FIELD_LEN} caracteres "
                            f"em Nome ou Sobrenome ({where}). Ajuste na conferência de nomes antes de gerar."
                        )
                    }
                ),
                400,
            )

        overrides_applied = df_final.attrs.get("name_overrides_applied", [])
        # `attrs` é copiado a cada operação do pandas na exportação (uma por coluna): não leva a lista de nomes junto.
        df_final.attrs = {}
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
                "names_confirmed": len(overrides_applied),  # só a contagem: nome de passageiro não vai para o histórico
            },
        )
        # O vocabulário aprende com o que o usuário confirmou (sucesso da geração = ele confirmou de fato). Corrigir
        # à mão pesa mais que aceitar a sugestão. Melhor esforço: não atrapalha a entrega do arquivo.
        if overrides_applied:
            get_vocabulary().learn_many(
                (o["nome"], o["sobrenome"], EDITED_WEIGHT if o["editado"] else ACCEPTED_WEIGHT)
                for o in overrides_applied
            )
        return send_file(
            output,
            download_name="saida_cadastro.xlsx",
            as_attachment=True,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception:
        # detalhe completo só no log do servidor (logger.exception); o
        # histórico é público via GET /api/history, então não grava str(exc).
        logger.exception("Erro em /api/process_cadastro")
        AuditService.record(event_type="cadastro", status="error", details={})
        return jsonify({"error": "Erro interno ao processar a solicitação."}), 500
    finally:
        for p in paths:
            try:
                if os.path.exists(p):
                    os.remove(p)
            except Exception:
                logger.warning(f"Falha ao remover temporário {p}")
