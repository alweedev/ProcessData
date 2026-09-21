"""Parâmetros e mensagens de erro das rotas de cadastro (validação e geração), num lugar só."""

import json
import re

from backend.domain.rules import FLUXOS, LOGIN_CHOICES


def parse_login_and_fluxo(form) -> tuple[str, str, str | None]:
    """Lê ``login_choice`` e ``fluxo`` do formulário. Ausente ou vazio = padrão (CPF / SELF); caixa é
    ignorada. Valor desconhecido devolve a mensagem de erro em vez de gerar um arquivo com Login vazio."""
    login = (form.get("login_choice") or "").strip().upper() or "CPF"
    fluxo = (form.get("fluxo") or "").strip().upper() or "SELF"
    if login not in LOGIN_CHOICES:
        return login, fluxo, "Tipo de login inválido: use CPF ou EMAIL."
    if fluxo not in FLUXOS:
        return login, fluxo, "Fluxo inválido: use SELF ou FRONT."
    return login, fluxo, None


_MAX_OVERRIDES = 5000  # mais que o de uma planilha razoável (todas as linhas de 5 arquivos)
_MAX_TEXT = 200


def parse_name_overrides(form) -> tuple[dict, str | None]:
    """Lê ``name_overrides`` (JSON) do formulário: a decisão do usuário na conferência de nomes.

    Formato: ``{"arquivo:linha": {"nome_completo": str, "nome": str, "sobrenome": str, "editado": bool}}``.
    Ausente = sem correções. Estrutura inválida devolve a mensagem de erro (o corpo vem do cliente)."""
    raw = form.get("name_overrides")
    if not raw:
        return {}, None
    invalid = "Correções de nomes inválidas."
    try:
        data = json.loads(raw)
    except ValueError:
        return {}, invalid
    if not isinstance(data, dict) or len(data) > _MAX_OVERRIDES:
        return {}, invalid
    overrides: dict = {}
    for key, value in data.items():
        if not isinstance(key, str) or len(key) > 40 or not isinstance(value, dict):
            return {}, invalid
        fields = {name: value.get(name) for name in ("nome_completo", "nome", "sobrenome")}
        if any(not isinstance(v, str) or len(v) > _MAX_TEXT for v in fields.values()):
            return {}, invalid
        overrides[key] = {**fields, "editado": value.get("editado") is True}
    return overrides, None


def display_name(filename: str | None) -> str:
    """Nome do arquivo como o usuário o enviou, sem diretórios (o cliente controla o `filename`)."""
    name = (filename or "").replace("\\", "/").rsplit("/", 1)[-1].strip()
    return name or "arquivo"


def user_facing_file_errors(errors: dict, names: dict[str, str], upload_folder: str) -> dict[str, str]:
    """Erros de leitura por arquivo — chaveados pelo caminho temporário no servidor — como
    ``{nome enviado: mensagem}``, sem nenhum caminho do servidor (nem no texto que a biblioteca de
    leitura devolveu). Ignora as chaves que não são de arquivo (número de linha e ``__geral__``)."""
    folder = re.compile(re.escape(upload_folder.replace("\\", "/").rstrip("/")) + r"/*", re.IGNORECASE)
    result: dict[str, str] = {}
    for key, message in errors.items():
        if not isinstance(key, str) or key == "__geral__":
            continue
        name = names.get(key, "arquivo")
        text = str(message).replace("\\", "/").replace(key.replace("\\", "/"), name)
        result[name] = folder.sub("", text)
    return result
