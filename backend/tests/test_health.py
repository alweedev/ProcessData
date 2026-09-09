"""Health endpoint canônico é /api/health; /health duplicado foi removido (C2)."""


def test_api_health_ok(client):
    resp = client.get("/api/health")
    assert resp.status_code == 200
    assert resp.get_json()["status"] == "OK"


def test_bare_health_no_longer_serves_health_json(client):
    resp = client.get("/health")
    # agora cai no fallback de SPA (index.html), não mais no JSON de health
    payload = resp.get_json(silent=True)
    assert not (isinstance(payload, dict) and payload.get("status") == "OK")
