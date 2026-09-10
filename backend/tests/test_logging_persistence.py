"""Rotação da trilha de auditoria + respostas 5xx genéricas (P8g, P8h)."""

import os

import pandas as pd
from _helpers import valid_cpf, xlsx_upload


def test_history_rotation(monkeypatch):
    from backend.core.config import settings
    from backend.services.audit_service import AuditService

    monkeypatch.setattr(settings, "HISTORY_MAX_BYTES", 1500, raising=False)

    for i in range(60):
        AuditService.record("evt_teste", "success", {"i": i, "pad": "x" * 30})

    path = settings.HISTORY_LOG_FILE
    assert os.path.exists(path + ".1"), "rotação deveria ter criado history.log.jsonl.1"
    with open(path, encoding="utf-8") as fh:
        linhas_atuais = [ln for ln in fh if ln.strip()]
    assert len(linhas_atuais) < 60  # arquivo corrente foi truncado

    eventos = AuditService.list_events(limit=50)
    seq = [e["details"]["i"] for e in eventos]
    assert seq == sorted(seq, reverse=True)  # mais novo primeiro
    assert seq and seq[0] == 59


def test_append_failure_does_not_raise(monkeypatch):
    """Uma falha de I/O ao gravar a auditoria não pode propagar como exceção
    (isso substituiria uma resposta 500 controlada por um erro não tratado)."""
    from backend.infra.persistence.history_store import HistoryStore

    def boom(*a, **k):
        raise OSError("disco cheio")

    monkeypatch.setattr(os, "makedirs", boom)

    HistoryStore.append({"event_type": "evt", "status": "success", "details": {}})


def test_500_response_is_generic(client, monkeypatch):
    from backend.services.processing_service import ProcessingService

    def boom(*a, **k):
        raise RuntimeError("detalhe sensivel: C:\\segredo\\base_real.xlsx")

    monkeypatch.setattr(ProcessingService, "process_records_from_files", boom)

    df = pd.DataFrame(
        [
            {
                "CPF": valid_cpf(1),
                "NOME COMPLETO": "Ana",
                "EMAIL": "a@x.com",
                "EMPRESA": "E",
                "Centro de custo": "C",
                "SOLICITANTE? (S/N)": "S",
            }
        ]
    )
    data = {"files[]": xlsx_upload(df, "c.xlsx")}
    resp = client.post("/api/process_cadastro", data=data, content_type="multipart/form-data")
    assert resp.status_code == 500
    body = resp.get_data(as_text=True)
    assert "Erro interno" in resp.get_json()["error"]
    assert "segredo" not in body
    assert "Traceback" not in body
