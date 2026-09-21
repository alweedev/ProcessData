import json
import os

import pandas as pd
from flask import Blueprint, jsonify, request, send_file
from werkzeug.exceptions import HTTPException, RequestEntityTooLarge

from backend.core.config import settings
from backend.core.logging import get_logger
from backend.services.audit_service import AuditService
from backend.services.export_service import ExportService
from backend.services.inactivation_cascade_service import Analise, InactivationCascadeService, InativacaoError
from backend.services.inactivation_service import InactivationService
from backend.shared.cpf_mask import mascarar_cpf
from backend.shared.upload_validation import save_and_validate_upload

logger = get_logger()

inativacao_bp = Blueprint("inativacao", __name__, url_prefix="/api")

_VERDADEIRO = {"1", "true", "yes", "on"}


@inativacao_bp.before_request
def _limite_de_upload() -> None:
    # Duas planilhas grandes na mesma requisição estouram os 16 MB globais: teto próprio destas rotas.
    request.max_content_length = settings.INATIVACAO_MAX_CONTENT_LENGTH


@inativacao_bp.errorhandler(RequestEntityTooLarge)
def _arquivo_grande(_exc):
    limite = max(1, settings.INATIVACAO_MAX_CONTENT_LENGTH // (1024 * 1024))
    mensagem = f"Os arquivos enviados passam do limite de {limite} MB."
    return jsonify({"error": mensagem, "code": "ARQUIVO_GRANDE"}), 413


def _erro(exc: InativacaoError):
    return jsonify({"error": exc.message, "code": exc.code, **exc.extra}), exc.status


def _limpar(paths: list[str]) -> None:
    for path in paths:
        try:
            if path and os.path.exists(path):
                os.remove(path)
        except Exception as exc:  # pragma: no cover - limpeza nunca derruba a resposta
            logger.warning("Falha ao remover temporário %s: %s", path, exc)


def _ler_planilha(campo: str, rotulo: str, paths: list[str]) -> pd.DataFrame:
    arquivo = request.files.get(campo)
    if not arquivo:
        raise InativacaoError("BASE_AUSENTE", f"Envie a {rotulo}.")
    path, err = save_and_validate_upload(arquivo, settings.UPLOAD_FOLDER, label=rotulo)
    if path:
        paths.append(path)
    if err:
        raise InativacaoError("ARQUIVO_INVALIDO", err)
    try:
        return pd.read_excel(path, dtype=str).fillna("")
    except Exception:
        raise InativacaoError("ARQUIVO_INVALIDO", f"{rotulo}: não foi possível ler a planilha.") from None


def _json_lista(campo: str) -> list[str]:
    bruto = request.form.get(campo, "")
    if not bruto:
        return []
    try:
        dados = json.loads(bruto)
    except ValueError:
        raise InativacaoError("LISTA_VAZIA", f"O campo '{campo}' não é um JSON válido.") from None
    return [str(x).strip() for x in dados if str(x).strip()] if isinstance(dados, list) else []


def _extrair_itens(paths: list[str]) -> list[str]:
    """Itens da lista: `itens` (JSON), `lista` (planilha) ou `lista_text` (um por linha), nessa ordem."""
    if request.form.get("itens"):
        return _json_lista("itens")
    arquivo = request.files.get("lista")
    if arquivo:
        path, err = save_and_validate_upload(arquivo, settings.UPLOAD_FOLDER, label="lista")
        if path:
            paths.append(path)
        if err:
            raise InativacaoError("ARQUIVO_INVALIDO", err)
        df = InactivationService.normalize_lista_columns(pd.read_excel(path, dtype=str).fillna(""))
        itens: list[str] = []
        for _idx, row in df.iterrows():
            valores = (str(row.get(col, "")).strip() for col in ("CPF", "Email", "NomeCompleto"))
            itens.append(next((v for v in valores if v), ""))
        return [i for i in itens if i]
    return [linha.strip() for linha in request.form.get("lista_text", "").split("\n") if linha.strip()]


def _detalhes_analise(analise: Analise) -> dict:
    """Só contagens, impressão digital e CPF mascarado: nome e e-mail nunca vão para o histórico."""
    resumo = dict(analise.payload["resumo"])
    resumo["duplicados"] = [mascarar_cpf(c) for c in resumo.get("duplicados", [])]  # CPFs digitados em dobro
    return {
        "resumo": resumo,
        "impressaoDigital": analise.payload["impressaoDigital"],
        "usuarios": [{"cpf": u["cpfMascarado"], "situacao": u["situacao"]} for u in analise.payload["usuarios"]],
    }


def _interno(evento: str, rota: str):
    logger.exception("Erro em %s", rota)
    AuditService.record(event_type=evento, status="error", details={"code": "ERRO_INTERNO"})
    return jsonify({"error": "Erro interno ao processar a solicitação.", "code": "ERRO_INTERNO"}), 500


@inativacao_bp.route("/inativacao/analisar", methods=["POST"])
def api_inativacao_analisar():
    paths: list[str] = []
    try:
        df_cadastro = _ler_planilha("cadastro", "base de cadastro", paths)
        df_estruturas = _ler_planilha("estruturas", "base de estruturas", paths)
        itens = _extrair_itens(paths)
        analise = InactivationCascadeService.analisar(df_cadastro, df_estruturas, itens, _json_lista("selecionados"))
        AuditService.record(event_type="inativacao_analise", status="success", details=_detalhes_analise(analise))
        return jsonify(analise.payload), 200
    except InativacaoError as exc:
        AuditService.record(event_type="inativacao_analise", status="error", details={"code": exc.code})
        return _erro(exc)
    except HTTPException:
        raise
    except Exception:
        return _interno("inativacao_analise", "/api/inativacao/analisar")
    finally:
        _limpar(paths)


@inativacao_bp.route("/inativacao/executar", methods=["POST"])
def api_inativacao_executar():
    paths: list[str] = []
    try:
        df_cadastro = _ler_planilha("cadastro", "base de cadastro", paths)
        df_estruturas = _ler_planilha("estruturas", "base de estruturas", paths)
        cpfs = _json_lista("cpfs")
        ignorar_orfas = request.form.get("ignore_orphan_warning", "").lower() in _VERDADEIRO
        execucao = InactivationCascadeService.executar(
            df_cadastro, df_estruturas, cpfs, request.form.get("impressaoDigital", ""), ignorar_orfas
        )
        arquivos = {
            "saida_inativacao.xlsx": ExportService.to_excel_bytes(execucao.ficha, sheet_name="Inativacao"),
            "estruturas_atualizadas.xlsx": ExportService.to_excel_bytes(execucao.estruturas, sheet_name="Aprovacao"),
        }
        pacote = ExportService.to_zip_bytes(arquivos)
        AuditService.record(
            event_type="inativacao_execucao",
            status="success",
            details={"resumo": execucao.resumo, "cpfs": [mascarar_cpf(c) for c in cpfs]},
        )
        return send_file(pacote, download_name="inativacao.zip", as_attachment=True, mimetype="application/zip")
    except InativacaoError as exc:
        AuditService.record(event_type="inativacao_execucao", status="error", details={"code": exc.code})
        return _erro(exc)
    except HTTPException:
        raise
    except Exception:
        return _interno("inativacao_execucao", "/api/inativacao/executar")
    finally:
        _limpar(paths)
