import os
import tempfile
from dataclasses import dataclass

# Diretorio de trabalho para uploads temporarios e trilha de auditoria.
# Fora do repositorio por padrao (evita vazar planilhas / history.jsonl para o
# git); sobrescrevivel por variavel de ambiente em producao.
_DEFAULT_UPLOAD_FOLDER = os.getenv("UPLOAD_FOLDER", os.path.join(tempfile.gettempdir(), "processdata_uploads"))
_DEFAULT_HISTORY_LOG_FILE = os.getenv("HISTORY_LOG_FILE", os.path.join(_DEFAULT_UPLOAD_FOLDER, "history.log.jsonl"))
# Vocabulário de nomes aprendido com as conferências (só palavras soltas + contagem, nunca nome completo nem CPF).
# Fora do repositório e, ao contrário dos uploads, num lugar que o sistema não limpa: é o que se aprendeu com o uso.
_DEFAULT_NAME_VOCAB_FILE = os.getenv(
    "NAME_VOCAB_FILE", os.path.join(os.path.expanduser("~"), ".processdata", "name_vocabulary.json")
)


@dataclass
class Settings:
    # Directories
    PROJECT_ROOT: str = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
    BACKEND_DIR: str = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    FRONTEND_DIR: str = os.path.abspath(os.path.join(BACKEND_DIR, "..", "frontend"))
    FRONTEND_STATIC_DIR: str = os.path.abspath(os.path.join(FRONTEND_DIR, "static"))

    # Uploads / persistência (env: UPLOAD_FOLDER, HISTORY_LOG_FILE)
    UPLOAD_FOLDER: str = _DEFAULT_UPLOAD_FOLDER
    HISTORY_LOG_FILE: str = _DEFAULT_HISTORY_LOG_FILE
    NAME_VOCAB_FILE: str = _DEFAULT_NAME_VOCAB_FILE  # env: NAME_VOCAB_FILE
    MAX_CONTENT_LENGTH: int = 16 * 1024 * 1024
    # Inativação em cascata: teto de upload próprio (duas planilhas na mesma requisição) e de itens na lista.
    INATIVACAO_MAX_CONTENT_LENGTH: int = 32 * 1024 * 1024
    MAX_INATIVACAO_ITENS: int = 500

    # Server (can be overridden by environment variables)
    DEBUG: bool = os.getenv("DEBUG", "false").lower() in ("1", "true", "yes")
    HOST: str = os.getenv("HOST", "0.0.0.0")
    PORT: int = int(os.getenv("PORT", "5000"))

    # Acesso (env: CORS_ORIGINS lista separada por virgula; vazio = mesma origem)
    CORS_ORIGINS: str = os.getenv("CORS_ORIGINS", "")
    # Token para DELETE /api/history fora de localhost (env: HISTORY_ADMIN_TOKEN)
    HISTORY_ADMIN_TOKEN: str = os.getenv("HISTORY_ADMIN_TOKEN", "")

    # Rotação da trilha de auditoria (JSONL)
    HISTORY_MAX_BYTES: int = int(os.getenv("HISTORY_MAX_BYTES", str(5 * 1024 * 1024)))
    HISTORY_MAX_ROWS: int = int(os.getenv("HISTORY_MAX_ROWS", "5000"))

    @property
    def cors_origins_list(self) -> list:
        return [o.strip() for o in self.CORS_ORIGINS.split(",") if o.strip()]

    def ensure_dirs(self):
        os.makedirs(self.UPLOAD_FOLDER, exist_ok=True)
        history_dir = os.path.dirname(self.HISTORY_LOG_FILE)
        if history_dir:
            os.makedirs(history_dir, exist_ok=True)
        return self


# Singleton-like settings instance for convenient import
settings = Settings().ensure_dirs()
