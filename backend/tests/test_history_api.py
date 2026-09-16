import backend.app as backend_app
from backend.services.audit_service import AuditService


def test_history_endpoints_work():
    app = backend_app.app
    with app.test_client() as client:
        res_clear = client.delete("/api/history")
        assert res_clear.status_code == 200

        res_get = client.get("/api/history?limit=10")
        assert res_get.status_code == 200
        payload = res_get.get_json()
        assert isinstance(payload, dict)
        assert "items" in payload
        assert "total" in payload


def test_get_history_reflects_real_content_and_order(client):
    AuditService.record("evt_um", "success", {"n": 1})
    AuditService.record("evt_dois", "success", {"n": 2})

    resp = client.get("/api/history?limit=10")
    assert resp.status_code == 200
    items = resp.get_json()["items"]
    assert len(items) >= 2
    # mais recente primeiro
    assert items[0]["event_type"] == "evt_dois"
    assert items[0]["details"] == {"n": 2}
    assert items[1]["event_type"] == "evt_um"


def test_get_history_bad_limit_falls_back_to_default(client):
    AuditService.record("evt", "success", {})
    resp = client.get("/api/history?limit=not-a-number")
    assert resp.status_code == 200
    assert resp.get_json()["items"]


def test_get_history_zero_limit_returns_everything(client):
    for i in range(3):
        AuditService.record(f"evt_{i}", "success", {})

    resp = client.get("/api/history?limit=0")
    assert resp.status_code == 200
    assert resp.get_json()["total"] >= 3
