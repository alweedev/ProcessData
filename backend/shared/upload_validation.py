"""Validação de conteúdo de uploads (defesa contra arquivos maliciosos).

A validação de extensão (``validar_extensao_arquivo``) não garante que o
conteúdo seja realmente uma planilha. Aqui checamos a assinatura/estrutura do
arquivo já salvo em disco. ``.xls`` (BIFF binário) não é checado — fica a cargo
do pandas/xlrd na leitura.
"""

import os
import zipfile
from typing import Any

from backend.shared.file_utils import gerar_nome_arquivo_temporario, validar_extensao_arquivo

_ZIP_MAGIC = b"PK\x03\x04"
_OOXML_EXTS = {".xlsx", ".xltx"}


def save_and_validate_upload(file: Any, upload_folder: str, *, label: str | None = None) -> tuple[str | None, str | None]:
    """Salva um upload (`werkzeug.FileStorage`) em arquivo temporário, validando
    extensão e conteúdo. Repete o mesmo trio (extensão -> salvar -> conteúdo)
    que toda rota de upload do projeto precisa fazer.

    Retorna ``(path, None)`` em sucesso. Em falha, ``path`` ainda pode vir
    preenchido (o arquivo já foi salvo em disco antes da validação de
    conteúdo) — o chamador deve sempre agendar ``path`` pra limpeza no
    ``finally`` da rota, mesmo quando a mensagem de erro não é ``None``.
    """
    prefix = f"{label}: " if label else ""

    is_valid, error_msg = validar_extensao_arquivo(file.filename)
    if not is_valid:
        return None, f"{prefix}{error_msg}"

    path = gerar_nome_arquivo_temporario(file.filename, upload_folder)
    file.save(path)
    ok, msg = validar_conteudo_xlsx(path)
    if not ok:
        return path, f"{prefix}{msg}"
    return path, None


def validar_conteudo_xlsx(path: str) -> tuple[bool, str]:
    """Retorna ``(ok, mensagem_erro)`` para um arquivo já salvo em ``path``."""
    ext = os.path.splitext(path)[1].lower()
    if ext not in _OOXML_EXTS:
        return True, ""

    try:
        with open(path, "rb") as fh:
            head = fh.read(4)
        if head != _ZIP_MAGIC:
            return False, "Arquivo não é um .xlsx válido (conteúdo inesperado)."
        with zipfile.ZipFile(path) as zf:
            names = zf.namelist()
        if "[Content_Types].xml" not in names or not any(n.startswith("xl/") for n in names):
            return False, "Arquivo .xlsx sem a estrutura esperada de planilha."
    except zipfile.BadZipFile:
        return False, "Arquivo .xlsx corrompido ou inválido."
    except OSError:
        return False, "Não foi possível ler o arquivo enviado."
    return True, ""
