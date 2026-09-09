"""CORS restrito + autorização do DELETE /api/history (P8e, P8f)."""


def test_cors_no_wildcard(client):
    resp = client.get("/api/health", headers={"Origin": "https://evil.example"})
    acao = resp.headers.get("Access-Control-Allow-Origin", "")
    assert acao != "*"
    assert "evil.example" not in acao


def test_history_delete_allowed_from_localhost(client):
    # test client usa REMOTE_ADDR 127.0.0.1 por padrão
    resp = client.delete("/api/history")
    assert resp.status_code == 200


def test_history_delete_denied_remote(client):
    resp = client.delete("/api/history", environ_base={"REMOTE_ADDR": "8.8.8.8"})
    assert resp.status_code == 403


def test_history_delete_allowed_with_token(client, monkeypatch):
    from backend.core.config import settings

    monkeypatch.setattr(settings, "HISTORY_ADMIN_TOKEN", "s3cr3t", raising=False)
    resp = client.delete(
        "/api/history",
        headers={"X-Admin-Token": "s3cr3t"},
        environ_base={"REMOTE_ADDR": "8.8.8.8"},
    )
    assert resp.status_code == 200


def test_history_get_still_open(client):
    resp = client.get("/api/history", environ_base={"REMOTE_ADDR": "8.8.8.8"})
    assert resp.status_code == 200
