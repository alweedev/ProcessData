import os
from typing import Any

import pandas as pd
from flask import Blueprint, jsonify, request, send_file

from backend.core.config import settings
from backend.core.logging import get_logger
from backend.services.approval_service import ApprovalService
from backend.services.audit_service import AuditService
from backend.services.export_service import ExportService
from backend.shared.upload_validation import save_and_validate_upload

logger = get_logger()


aprovacao_bp = Blueprint("aprovacao", __name__, url_prefix="/api/aprovacao")


@aprovacao_bp.route("/remover/preview", methods=["POST"])
def aprovacao_remover_preview():
    users_path: str | None = None
    base_path: str | None = None
    try:
        users_file = request.files.get("users_file")
        base_file = request.files.get("base_file")
        form = request.form or {}
        raw_json = request.get_json(silent=True) if request.is_json else None
        cpf_raw = form.get("cpf") or (raw_json or {}).get("cpf")

        # Obter flag de remover segundo nível para cálculo correto
        remove_second_level_raw = form.get("remove_second_level")
        if remove_second_level_raw is None and raw_json is not None:
            remove_second_level_raw = raw_json.get("remove_second_level")
        if isinstance(remove_second_level_raw, bool):
            remove_second_level = remove_second_level_raw
        else:
            remove_second_level = str(remove_second_level_raw or "").lower() in {"1", "true", "yes", "on"}

        if not users_file or not base_file:
            return jsonify({"error": "Envie 'users_file' e 'base_file' (arquivos Excel)."}), 400

        cpf_digits, cpf_formatted = ApprovalService.normalize_cpf_input(cpf_raw)

        users_path, err = save_and_validate_upload(users_file, settings.UPLOAD_FOLDER, label="users_file")
        if err:
            return jsonify({"error": err}), 400
        base_path, err = save_and_validate_upload(base_file, settings.UPLOAD_FOLDER, label="base_file")
        if err:
            return jsonify({"error": err}), 400

        _, approver_name = ApprovalService.load_users_and_find_approver(users_path, cpf_digits)

        df_base = pd.read_excel(base_path, dtype=str).fillna("")
        cols = ApprovalService.detect_approval_columns(df_base)

        preview = ApprovalService.build_preview_for_cpf(
            df_base, cpf_digits, cols, check_empty=True, remove_second_level=remove_second_level
        )

        # Converter estruturas internas para o formato esperado pelo frontend
        raw_structures: list[dict[str, Any]] = preview.get("structures", []) or []
        structures_without_approvers = preview.get("structures_without_approvers", [])
        empty_ids = {s.get("aprovacaoId") for s in structures_without_approvers}

        por_aprovacao_por: dict[str, int] = {}
        items: list[dict[str, Any]] = []
        for rec in raw_structures:
            aprovacao_por = (rec.get("aprovacao_por") or "").strip()
            chave_tipo = aprovacao_por or "OUTRO"
            por_aprovacao_por[chave_tipo] = por_aprovacao_por.get(chave_tipo, 0) + 1

            cost_center = rec.get("cost_center") or ""
            cc_codigo: str | None = None
            cc_descricao: str | None = None
            if cost_center:
                # Dividir em "codigo - descricao" se possível
                partes = [p.strip() for p in str(cost_center).split("-", 1)]
                if len(partes) == 2:
                    cc_codigo, cc_descricao = partes[0] or None, partes[1] or None
                else:
                    cc_descricao = partes[0] or None

            positions = rec.get("positions") or []
            try:
                posicoes_norm = [int(p) for p in positions]
            except Exception:
                posicoes_norm = []

            item = {
                "aprovacaoId": rec.get("aprovacao_id"),
                "aprovacaoPor": aprovacao_por or None,
                "aprovacao": rec.get("aprovacao"),
                "tipo": rec.get("tipo"),
                "valor": rec.get("valor"),
                "viajanteNomeCompleto": rec.get("traveler_name"),
                "ccCodigo": cc_codigo,
                "ccDescricao": cc_descricao,
                "posicoes": posicoes_norm,
                "segundoNivel": bool(rec.get("in_second_level")),
                "ficaraSemAprovador": rec.get("aprovacao_id") in empty_ids,
            }
            items.append(item)

        response = {
            "approver": {
                "cpf": cpf_formatted,
                "nomeCompleto": approver_name,
            },
            "summary": {
                "estruturasAfetadas": preview.get("total_structures", 0),
                "ocorrenciasTotal": preview.get("total_occurrences", 0),
                "porAprovacaoPor": por_aprovacao_por,
                "estruturasSemAprovador": len(structures_without_approvers),
            },
            "items": items,
            "alertas": {
                "estruturasSemAprovador": structures_without_approvers,
            },
        }
        return jsonify(response), 200
    except ValueError as ve:
        logger.warning(f"Preview aprovacao remover - erro de validação: {ve}")
        return jsonify({"error": str(ve)}), 400
    except Exception:  # pragma: no cover - proteção extra
        logger.exception("Erro em /api/aprovacao/remover/preview")
        return jsonify({"error": "Erro interno ao processar a solicitação."}), 500
    finally:
        for path in [users_path, base_path]:
            try:
                if path and os.path.exists(path):
                    os.remove(path)
            except Exception as cleanup_exc:  # pragma: no cover
                logger.warning(f"Falha ao remover temporário {path}: {cleanup_exc}")


@aprovacao_bp.route("/remover/export", methods=["POST"])
def aprovacao_remover_export():
    users_path: str | None = None
    base_path: str | None = None
    try:
        users_file = request.files.get("users_file")
        base_file = request.files.get("base_file")
        form = request.form or {}

        raw_json = request.get_json(silent=True) if request.is_json else None

        cpf_raw = form.get("cpf") or (raw_json or {}).get("cpf")
        mode = (form.get("mode") or (raw_json or {}).get("mode") or "all").lower()

        # selected_aprovacao_ids[] pode vir como múltiplos campos de formulário
        selected_ids_form = form.getlist("selected_aprovacao_ids[]") or form.getlist("selected_aprovacao_ids")
        selected_ids_json = (raw_json or {}).get("selected_aprovacao_ids") or []
        selected_ids: set[str] = set(str(x) for x in (selected_ids_form or selected_ids_json or []))

        remove_second_level_raw = form.get("remove_second_level")
        if remove_second_level_raw is None and raw_json is not None:
            remove_second_level_raw = raw_json.get("remove_second_level")
        if isinstance(remove_second_level_raw, bool):
            remove_second_level = remove_second_level_raw
        else:
            remove_second_level = str(remove_second_level_raw or "").lower() in {"1", "true", "yes", "on"}

        # Flag para ignorar alerta de estruturas vazias
        ignore_empty_warning_raw = form.get("ignore_empty_warning")
        if ignore_empty_warning_raw is None and raw_json is not None:
            ignore_empty_warning_raw = raw_json.get("ignore_empty_warning")
        if isinstance(ignore_empty_warning_raw, bool):
            ignore_empty_warning = ignore_empty_warning_raw
        else:
            ignore_empty_warning = str(ignore_empty_warning_raw or "").lower() in {"1", "true", "yes", "on"}

        if not users_file or not base_file:
            return jsonify({"error": "Envie 'users_file' e 'base_file' (arquivos Excel)."}), 400

        cpf_digits, cpf_formatted = ApprovalService.normalize_cpf_input(cpf_raw)

        users_path, err = save_and_validate_upload(users_file, settings.UPLOAD_FOLDER, label="users_file")
        if err:
            return jsonify({"error": err}), 400
        base_path, err = save_and_validate_upload(base_file, settings.UPLOAD_FOLDER, label="base_file")
        if err:
            return jsonify({"error": err}), 400

        # Garante que o CPF existe na base de usuários (e obtém nome apenas para validação/coerência)
        _, _ = ApprovalService.load_users_and_find_approver(users_path, cpf_digits)

        df_base = pd.read_excel(base_path, dtype=str).fillna("")
        cols = ApprovalService.detect_approval_columns(df_base)

        preview = ApprovalService.build_preview_for_cpf(
            df_base, cpf_digits, cols, check_empty=True, remove_second_level=remove_second_level
        )
        affected_ids_all: set[str] = set(preview.get("affected_ids") or [])
        if not affected_ids_all:
            return jsonify({"error": "CPF não está presente em nenhuma estrutura de aprovação."}), 400

        if mode == "selected":
            if not selected_ids:
                return jsonify({"error": "Informe 'selected_aprovacao_ids' quando mode='selected'."}), 400
            target_ids = affected_ids_all.intersection(selected_ids)
            if not target_ids:
                return jsonify({"error": "Nenhuma AprovacaoId selecionada contém o CPF informado."}), 400
        else:
            target_ids = affected_ids_all

        # Verificar se há estruturas que ficarão sem aprovadores
        structures_without_approvers = preview.get("structures_without_approvers", [])
        affected_empty = [s for s in structures_without_approvers if s.get("aprovacaoId") in target_ids]

        if affected_empty and not ignore_empty_warning:
            return jsonify(
                {
                    "error": "Algumas estruturas ficarão sem nenhum aprovador após a remoção.",
                    "warning": True,
                    "estruturasSemAprovador": affected_empty,
                    "message": f"{len(affected_empty)} estrutura(s) ficará(ão) sem aprovadores. Deseja continuar mesmo assim?",
                }
            ), 400

        df_updated, stats = ApprovalService.remove_cpf_and_compact(
            df_base=df_base,
            cpf_digits=cpf_digits,
            cols=cols,
            target_ids=target_ids,
            remove_second_level=remove_second_level,
        )

        # Exporta todas as linhas das estruturas alvo (a Argo precisa da estrutura
        # completa), mas só carimba Operacao=UPDATE nas linhas efetivamente
        # alteradas pela remoção/compactação/promoção.
        aprovacao_id_col = cols.get("aprovacao_id")
        if aprovacao_id_col and aprovacao_id_col in df_updated.columns:
            df_export = df_updated[df_updated[aprovacao_id_col].astype(str).isin(target_ids)].copy()
        else:
            df_export = df_updated.copy()

        changed_indices = stats.get("changed_indices") or set()
        if "Operacao" not in df_export.columns:
            df_export.insert(0, "Operacao", "")
        changed_mask = df_export.index.isin(changed_indices)
        df_export.loc[changed_mask, "Operacao"] = "UPDATE"

        # Mover Operacao para a primeira posição
        cols_list = df_export.columns.tolist()
        if cols_list and cols_list[0] != "Operacao":
            cols_list.remove("Operacao")
            cols_list.insert(0, "Operacao")
            df_export = df_export[cols_list]

        # Escrita/estilo/neutralização de fórmula centralizados no ExportService.
        output = ExportService.to_excel_bytes(df_export, sheet_name="Aprovacao")

        filename = f"base_aprovacao_atualizada_{cpf_formatted.replace('-', '')}.xlsx"
        logger.info(
            "Export aprovacao remover gerado - apenas estruturas alteradas",
        )
        logger.info(
            "Estruturas atualizadas: %s | Ocorrencias removidas: %s | Linhas exportadas: %s (de %s total)",
            stats.get("structures_updated"),
            stats.get("occurrences_removed"),
            len(df_export),
            len(df_base),
        )

        AuditService.record(
            event_type="aprovacao_remocao",
            status="success",
            details={
                "structures_updated": stats.get("structures_updated"),
                "occurrences_removed": stats.get("occurrences_removed"),
            },
        )
        return send_file(
            output,
            download_name=filename,
            as_attachment=True,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except ValueError as ve:
        logger.warning(f"Export aprovacao remover - erro de validação: {ve}")
        return jsonify({"error": str(ve)}), 400
    except Exception:  # pragma: no cover - proteção extra
        logger.exception("Erro em /api/aprovacao/remover/export")
        AuditService.record(event_type="aprovacao_remocao", status="error", details={})
        return jsonify({"error": "Erro interno ao processar a solicitação."}), 500
    finally:
        for path in [users_path, base_path]:
            try:
                if path and os.path.exists(path):
                    os.remove(path)
            except Exception as cleanup_exc:  # pragma: no cover
                logger.warning(f"Falha ao remover temporário {path}: {cleanup_exc}")
