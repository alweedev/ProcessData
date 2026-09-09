"""Fixtures compartilhadas para a suíte de testes do backend.

Rode a partir da raiz do repositório: ``python -m pytest -q``.
"""
import pytest


@pytest.fixture(autouse=True)
def _isolate_state(tmp_path, monkeypatch):
    """Isola o estado em disco de cada teste.

    Repontar ``settings.UPLOAD_FOLDER`` e ``settings.HISTORY_LOG_FILE`` para
    dentro de ``tmp_path`` garante que nenhum teste escreva no repositório nem
    veja artefatos de outro teste. Todos os módulos importam a mesma instância
    ``settings``, e os handlers leem esses atributos em tempo de request, então
    alterar aqui basta.
    """
    from backend.core.config import settings

    upload_dir = tmp_path / "uploads"
    upload_dir.mkdir(parents=True, exist_ok=True)
    history_file = tmp_path / "history" / "history.log.jsonl"
    history_file.parent.mkdir(parents=True, exist_ok=True)

    monkeypatch.setattr(settings, "UPLOAD_FOLDER", str(upload_dir), raising=False)
    monkeypatch.setattr(settings, "HISTORY_LOG_FILE", str(history_file), raising=False)
    yield


@pytest.fixture
def app():
    import backend.app as app_module

    flask_app = app_module.app
    flask_app.config.update(TESTING=True)
    return flask_app


@pytest.fixture
def client(app):
    return app.test_client()
