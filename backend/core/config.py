import os
from dataclasses import dataclass


@dataclass
class Settings:
    # Directories
    PROJECT_ROOT: str = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
    BACKEND_DIR: str = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    FRONTEND_DIR: str = os.path.abspath(os.path.join(BACKEND_DIR, '..', 'frontend'))
    FRONTEND_STATIC_DIR: str = os.path.abspath(os.path.join(FRONTEND_DIR, 'static'))

    # Uploads
    UPLOAD_FOLDER: str = os.path.join(BACKEND_DIR, 'tmp_uploads')
    MAX_CONTENT_LENGTH: int = 16 * 1024 * 1024

    # Server
    DEBUG: bool = True
    HOST: str = '0.0.0.0'
    PORT: int = 5000

    def ensure_dirs(self):
        os.makedirs(self.UPLOAD_FOLDER, exist_ok=True)
        # static dir is managed by frontend assets; no creation here.
        return self


# Singleton-like settings instance for convenient import
settings = Settings().ensure_dirs()
