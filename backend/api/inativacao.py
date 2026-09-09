import os
import re
import pandas as pd
from flask import Blueprint, request, jsonify, send_file

# Use absolute imports to be robust to direct script execution
from backend.core.config import settings
from backend.core.logging import get_logger
from backend.services.audit_service import AuditService
from backend.services.inactivation_service import InactivationService
from backend.services.export_service import ExportService
from backend.shared.upload_validation import validar_conteudo_xlsx
from backend.utils import validar_extensao_arquivo, gerar_nome_arquivo_temporario

logger = get_logger()

inativacao_bp = Blueprint('inativacao', __name__, url_prefix='/api')


@inativacao_bp.route("/inativacao/buscar", methods=["POST"])
def api_inativacao_buscar():
    base_path = None
    lista_path = None
    try:
        base_file = request.files.get("base")
        if not base_file:
            return jsonify({"error": "Envie a base (arquivo Excel)"}), 400
        
        # Validar extensão do arquivo
        is_valid, error_msg = validar_extensao_arquivo(base_file.filename)
        if not is_valid:
            return jsonify({"error": error_msg}), 400
        
        base_path = gerar_nome_arquivo_temporario(base_file.filename, settings.UPLOAD_FOLDER)
        base_file.save(base_path)
        ok, msg = validar_conteudo_xlsx(base_path)
        if not ok:
            return jsonify({"error": msg}), 400
        df_base = pd.read_excel(base_path, dtype=str).fillna("")

        # Extrair itens (CPFs ou nomes)
        itens = []
        if request.is_json:
            try:
                payload = request.get_json(silent=True) or {}
                itens = payload.get("itens", []) or []
            except Exception:
                itens = []
        if not itens:
            itens_field = request.form.get("itens")
            if itens_field:
                try:
                    import json as _json
                    itens = _json.loads(itens_field)
                except Exception:
                    itens = []
        if not itens:
            lista_file = request.files.get("lista")
            lista_text = request.form.get("lista_text", "")
            if lista_file:
                # Validar extensão do arquivo
                is_valid, error_msg = validar_extensao_arquivo(lista_file.filename)
                if not is_valid:
                    return jsonify({"error": error_msg}), 400
                
                lista_path = gerar_nome_arquivo_temporario(lista_file.filename, settings.UPLOAD_FOLDER)
                lista_file.save(lista_path)
                ok, msg = validar_conteudo_xlsx(lista_path)
                if not ok:
                    return jsonify({"error": msg}), 400
                try:
                    df_lista = pd.read_excel(lista_path, dtype=str).fillna("")
                    df_lista = InactivationService.normalize_lista_columns(df_lista)
                    if 'CPF' in df_lista.columns:
                        itens = [str(x) for x in df_lista['CPF'].tolist() if str(x).strip()]
                except Exception:
                    itens = []
            elif lista_text.strip():
                itens = [line.strip() for line in lista_text.split('\n') if line.strip()]

        search = InactivationService.search_matches(df_base, itens)
        AuditService.record(
            event_type="inativacao_busca",
            status="success",
            details={
                "items_in": len(itens or []),
                "items_out": search.get("total", 0),
            },
        )
        return jsonify(search), 200
    except Exception as e:
        logger.exception("Erro em /api/inativacao/buscar")
        AuditService.record(
            event_type="inativacao_busca",
            status="error",
            details={"message": str(e)},
        )
        return jsonify({"error": "Erro interno ao processar a solicitação."}), 500
    finally:
        for path in [base_path, lista_path]:
            try:
                if path and os.path.exists(path):
                    os.remove(path)
            except Exception:
                pass


# NOTE: a rota POST /api/inativacao/executar foi removida (2026-09).
# Ela retornava {"success": true, "message": "N usuário(s) inativados."} sem
# ler a base, gerar arquivo ou registrar auditoria — sucesso falso, sem
# consumidor no frontend. A inativação real é feita por /api/process_inativacao.


@inativacao_bp.route("/process_inativacao", methods=["POST"])
def api_process_inativacao():
    base_path = None
    lista_path = None
    try:
        base_file = request.files.get("base")
        lista_file = request.files.get("lista")
        lista_text = request.form.get("lista_text", "").strip()

        if not base_file:
            logger.error("Nenhum arquivo 'base' enviado")
            return jsonify({"error": "Envie a base"}), 400

        if not (lista_file or lista_text):
            logger.error("Nenhum arquivo 'lista' ou texto enviado")
            return jsonify({"error": "Envie a lista ou insira os nomes/CPFs"}), 400

        # Validar extensão do arquivo base
        is_valid, error_msg = validar_extensao_arquivo(base_file.filename)
        if not is_valid:
            return jsonify({"error": error_msg}), 400

        base_path = gerar_nome_arquivo_temporario(base_file.filename, settings.UPLOAD_FOLDER)
        base_file.save(base_path)
        ok, msg = validar_conteudo_xlsx(base_path)
        if not ok:
            return jsonify({"error": msg}), 400
        logger.info(f"Arquivo base salvo em: {base_path}")

        if lista_file:
            # Validar extensão do arquivo lista
            is_valid, error_msg = validar_extensao_arquivo(lista_file.filename)
            if not is_valid:
                return jsonify({"error": error_msg}), 400

            lista_path = gerar_nome_arquivo_temporario(lista_file.filename, settings.UPLOAD_FOLDER)
            lista_file.save(lista_path)
            ok, msg = validar_conteudo_xlsx(lista_path)
            if not ok:
                return jsonify({"error": msg}), 400
            logger.info(f"Arquivo lista salvo em: {lista_path}")
            df_lista = pd.read_excel(lista_path, dtype=str).fillna("")
            df_lista = InactivationService.normalize_lista_columns(df_lista)
        else:
            logger.info("Processando lista a partir de texto")
            lista_items = [item.strip() for item in lista_text.split('\n') if item.strip()]
            if not lista_items:
                return jsonify({"error": "Texto de lista vazio ou sem CPF/Nome/E-mail válidos"}), 400
            email_pat = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", re.IGNORECASE)
            rows = []
            for it in lista_items:
                digits = re.sub(r"\D", "", it)
                if email_pat.match(it):
                    rows.append({"CPF": "", "NomeCompleto": "", "Email": it})
                elif len(digits) == 11:
                    rows.append({"CPF": it, "NomeCompleto": "", "Email": ""})
                else:
                    rows.append({"CPF": "", "NomeCompleto": it, "Email": ""})
            df_lista = pd.DataFrame(rows)
            if 'Email' not in df_lista.columns:
                df_lista['Email'] = ''

        df_base = pd.read_excel(base_path, dtype=str).fillna("")

        use_fuzzy = request.form.get('use_fuzzy', 'false').lower() in ['1', 'true', 'yes']
        try:
            fuzzy_cutoff = float(request.form.get('fuzzy_cutoff', 0.90))
        except Exception:
            fuzzy_cutoff = 0.90

        out = InactivationService.process_from_dataframes(
            df_base,
            df_lista,
            use_fuzzy=use_fuzzy,
            fuzzy_cutoff=fuzzy_cutoff,
        )
        if isinstance(out, tuple) and len(out) == 2:
            out_df, stats = out
        else:
            out_df = out
            stats = {}

        logger.info("DataFrame de inativação gerado: %s linhas x %s colunas", out_df.shape[0], out_df.shape[1])
        if out_df.empty:
            logger.warning("Nenhuma linha ativa correspondente para inativação")
            if stats and stats.get('inactive_matches'):
                AuditService.record(
                    event_type="inativacao_geracao",
                    status="empty",
                    details={"stats": stats},
                )
                return jsonify({"error": "Nenhuma linha ativa correspondeu; foram encontradas correspondências INATIVAS.", "stats": stats}), 400
            AuditService.record(
                event_type="inativacao_geracao",
                status="empty",
                details={"stats": stats},
            )
            return jsonify({"error": "Nenhum dado processado para inativação", "stats": stats}), 400

        output = ExportService.to_excel_bytes(out_df, sheet_name="Inativacao")

        logger.info("Arquivo de inativação gerado e enviado")
        AuditService.record(
            event_type="inativacao_geracao",
            status="success",
            details={
                "rows": int(out_df.shape[0]),
                "columns": int(out_df.shape[1]),
                "use_fuzzy": use_fuzzy,
                "fuzzy_cutoff": fuzzy_cutoff,
            },
        )
        return send_file(output,
                         download_name="saida_inativacao.xlsx",
                         as_attachment=True,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        logger.exception("Erro em /api/process_inativacao")
        AuditService.record(
            event_type="inativacao_geracao",
            status="error",
            details={"message": str(e)},
        )
        return jsonify({"error": "Erro interno ao processar a solicitação."}), 500
    finally:
        for path in [base_path, lista_path]:
            try:
                if path and os.path.exists(path):
                    os.remove(path)
                    logger.info(f"Arquivo temporário removido: {path}")
            except Exception as e:
                logger.warning(f"Falha ao remover {path}: {str(e)}")


@inativacao_bp.route("/preview_inativacao", methods=["POST"])
def api_preview_inativacao():
    base_path = None
    lista_path = None
    try:
        base_file = request.files.get("base")
        lista_file = request.files.get("lista")
        lista_text = request.form.get("lista_text", "").strip()

        if not base_file:
            return jsonify({"error": "Envie a base"}), 400

        # Validar extensão do arquivo base
        is_valid, error_msg = validar_extensao_arquivo(base_file.filename)
        if not is_valid:
            return jsonify({"error": error_msg}), 400

        base_path = gerar_nome_arquivo_temporario(base_file.filename, settings.UPLOAD_FOLDER)
        base_file.save(base_path)
        ok, msg = validar_conteudo_xlsx(base_path)
        if not ok:
            return jsonify({"error": msg}), 400

        if lista_file:
            # Validar extensão do arquivo lista
            is_valid, error_msg = validar_extensao_arquivo(lista_file.filename)
            if not is_valid:
                return jsonify({"error": error_msg}), 400

            lista_path = gerar_nome_arquivo_temporario(lista_file.filename, settings.UPLOAD_FOLDER)
            lista_file.save(lista_path)
            ok, msg = validar_conteudo_xlsx(lista_path)
            if not ok:
                return jsonify({"error": msg}), 400
            df_lista = pd.read_excel(lista_path, dtype=str).fillna("")
        else:
            if not lista_text:
                return jsonify({"error": "Envie a lista como arquivo ou cole nomes/CPFs no campo de texto"}), 400
            lista_items = [item.strip() for item in lista_text.split('\n') if item.strip()]
            if not lista_items:
                return jsonify({"error": "Texto de lista vazio ou sem CPF/Nome/E-mail válidos"}), 400
            email_pat = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", re.IGNORECASE)
            rows = []
            for it in lista_items:
                digits = re.sub(r"\D", "", it)
                if email_pat.match(it):
                    rows.append({"CPF": "", "NomeCompleto": "", "Email": it})
                elif len(digits) == 11:
                    rows.append({"CPF": it, "NomeCompleto": "", "Email": ""})
                else:
                    rows.append({"CPF": "", "NomeCompleto": it, "Email": ""})
            df_lista = pd.DataFrame(rows)
            if 'Email' not in df_lista.columns:
                df_lista['Email'] = ''

        df_base = pd.read_excel(base_path, dtype=str).fillna("")

        use_fuzzy = request.form.get('use_fuzzy', 'false').lower() in ['1', 'true', 'yes']
        try:
            fuzzy_cutoff = float(request.form.get('fuzzy_cutoff', 0.90))
        except Exception:
            fuzzy_cutoff = 0.90
        out = InactivationService.process_from_dataframes(
            df_base, df_lista, use_fuzzy=use_fuzzy, fuzzy_cutoff=fuzzy_cutoff
        )
        if isinstance(out, tuple) and len(out) == 2:
            out_df, stats = out
        else:
            out_df = out
            stats = {}

        try:
            count = int(stats.get('total_matches')) if stats and 'total_matches' in stats else (int(out_df.shape[0]) if out_df is not None else 0)
        except Exception:
            count = int(out_df.shape[0]) if out_df is not None else 0

        sample = []
        columns = []
        records = []
        if out_df is not None and not out_df.empty:
            try:
                columns = list(out_df.columns)
            except Exception:
                columns = []
            try:
                sample = out_df.head(10).to_dict(orient="records")
            except Exception:
                sample = []
            try:
                records = out_df.head(500).to_dict(orient='records')
            except Exception:
                records = sample[:]
        else:
            try:
                inactive = (stats or {}).get('inactive_matches') or {}
                all_rows = []
                for k, lst in (inactive.items() if isinstance(inactive, dict) else []):
                    if isinstance(lst, list):
                        for it in lst:
                            try:
                                r = dict(it)
                                r['match_type'] = k
                                all_rows.append(r)
                            except Exception:
                                pass
                if all_rows:
                    colset = set()
                    for r in all_rows[:50]:
                        try:
                            colset.update(list(r.keys()))
                        except Exception:
                            pass
                    columns = list(colset)
                    sample = all_rows[:10]
                    records = all_rows[:500]
            except Exception:
                pass

        try:
            logger.info(f"/api/preview_inativacao -> count={count} sample={len(sample)} records={len(records)} columns={len(columns)}")
            if sample and isinstance(sample, list) and len(sample) > 0:
                logger.debug(f"preview sample keys: {list(sample[0].keys())[:10]}")
        except Exception:
            pass
        return jsonify({"count": count, "sample": sample, "columns": columns, "records": records, "stats": stats}), 200
    except Exception as e:
        logger.exception("Erro em /api/preview_inativacao")
        return jsonify({"error": "Erro interno ao processar a solicitação."}), 500
    finally:
        for path in [base_path, lista_path]:
            try:
                if path and os.path.exists(path):
                    os.remove(path)
                    logger.info(f"Arquivo temporário removido: {path}")
            except Exception as e:
                logger.warning(f"Falha ao remover {path}: {str(e)}")
