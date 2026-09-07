import backend.app as backend_app


def test_history_endpoints_work():
    app = backend_app.app
    with app.test_client() as client:
        res_clear = client.delete('/api/history')
        assert res_clear.status_code == 200

        res_get = client.get('/api/history?limit=10')
        assert res_get.status_code == 200
        payload = res_get.get_json()
        assert isinstance(payload, dict)
        assert 'items' in payload
        assert 'total' in payload
