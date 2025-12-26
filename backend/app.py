"""Aplicação principal Flask.

Responsável por criar e configurar a aplicação, registrar blueprints e
expor o objeto `app` para execução WSGI ou debug local.

Reorganizado para uma estrutura profissional:
backend/
  core/ (config, logging)
  api/  (blueprints por domínio)
  processor.py, utils.py, validators.py (camada de serviços)
  app.py (fábrica + bootstrap)
"""

import os
import sys
from flask import Flask
from flask_cors import CORS

# Garantir que o diretório pai esteja no sys.path para execução direta (python backend/app.py ou python app.py)
PARENT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if PARENT not in sys.path:
    sys.path.insert(0, PARENT)

from backend.core.config import settings
from backend.core.logging import get_logger
from backend.api import inativacao_bp, frontend_bp, cadastro_bp

logger = get_logger()


def create_app() -> Flask:
    """Factory para criar a aplicação Flask configurada."""
    app = Flask(__name__, static_folder=settings.FRONTEND_STATIC_DIR, static_url_path='/static')
    CORS(app)
    app.config['MAX_CONTENT_LENGTH'] = settings.MAX_CONTENT_LENGTH
    app.config['UPLOAD_FOLDER'] = settings.UPLOAD_FOLDER

    # Registrar blueprints
    app.register_blueprint(inativacao_bp)
    app.register_blueprint(cadastro_bp)
    app.register_blueprint(frontend_bp)

    logger.info('Aplicação Flask criada e blueprints registrados.')
    return app


# Instância global para uso por testes e servidores WSGI
app = create_app()


if __name__ == '__main__':
    logger.info('Iniciando servidor de desenvolvimento Flask...')
    app.run(debug=settings.DEBUG, host=settings.HOST, port=settings.PORT)