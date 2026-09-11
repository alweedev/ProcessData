"""Utilitários de arquivo: validação de extensão e nome temporário."""

import os
import re
import uuid

_DEFAULT_ALLOWED = {".xlsx", ".xls", ".xltx"}
_UNSAFE_CHARS = re.compile(r"[^A-Za-z0-9._-]+")


def _sanitize_filename_part(name: str) -> str:
    """Reduz `name` a um componente de arquivo seguro: descarta qualquer
    diretório embutido (`/` ou `\\`, o filename de um upload multipart é
    controlado pelo cliente) e troca caracteres fora de [A-Za-z0-9._-] por
    `_`, prevenindo path traversal ao montar o caminho do temporário."""
    name = name.replace("\\", "/").rsplit("/", 1)[-1]
    name = name.lstrip(".").strip()
    name = _UNSAFE_CHARS.sub("_", name)
    return name or "file"


def validar_extensao_arquivo(filename: str, allowed_extensions: set[str] | None = None) -> tuple[bool, str]:
    """Valida se a extensão do arquivo é permitida.

    Returns:
        (bool, str): (é_válido, mensagem_erro)
    """
    if allowed_extensions is None:
        allowed_extensions = _DEFAULT_ALLOWED

    if not filename:
        return False, "Nenhum arquivo fornecido"

    _, ext = filename.rsplit(".", 1) if "." in filename else ("", "")
    ext = f".{ext.lower()}" if ext else ""

    if not ext:
        return False, "Arquivo sem extensão"

    if ext not in allowed_extensions:
        exts_str = ", ".join(sorted(allowed_extensions))
        return False, f"Extensão não permitida. Aceitos: {exts_str}"

    return True, ""


def gerar_nome_arquivo_temporario(filename: str, upload_folder: str) -> str:
    """Gera um caminho único para arquivo temporário, preservando a extensão."""
    filename = _sanitize_filename_part(filename or "file")

    if "." in filename:
        name_part, ext = filename.rsplit(".", 1)
        ext = f".{ext}"
    else:
        name_part = filename
        ext = ""

    unique_id = str(uuid.uuid4())[:8]
    temp_filename = f"{unique_id}-{name_part}{ext}"

    os.makedirs(upload_folder, exist_ok=True)
    return os.path.join(upload_folder, temp_filename)
