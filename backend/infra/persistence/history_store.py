import json
import os
from collections import deque
from datetime import datetime, timezone

from backend.core.config import settings
from backend.core.logging import get_logger

logger = get_logger()


class HistoryStore:
    @staticmethod
    def _file_path():
        return settings.HISTORY_LOG_FILE

    @staticmethod
    def _rotate_if_needed(path: str) -> None:
        """Rotação simples de 1 geração quando o arquivo passa de HISTORY_MAX_BYTES."""
        max_bytes = getattr(settings, "HISTORY_MAX_BYTES", 5 * 1024 * 1024)
        try:
            if max_bytes and os.path.getsize(path) >= max_bytes:
                backup = path + ".1"
                if os.path.exists(backup):
                    os.remove(backup)
                os.replace(path, backup)
        except OSError:
            pass

    @staticmethod
    def append(event: dict):
        # Auditoria é best-effort: uma falha aqui (disco cheio, permissão, path
        # inválido) nunca pode virar uma exceção não tratada dentro do bloco
        # except de uma rota, substituindo a resposta 500 controlada.
        try:
            path = HistoryStore._file_path()
            os.makedirs(os.path.dirname(path), exist_ok=True)
            if os.path.exists(path):
                HistoryStore._rotate_if_needed(path)
            payload = {
                "timestamp": datetime.now(timezone.utc).isoformat(),
                **(event or {}),
            }
            with open(path, "a", encoding="utf-8") as fp:
                fp.write(json.dumps(payload, ensure_ascii=False) + "\n")
        except OSError:
            logger.warning("Falha ao gravar trilha de auditoria", exc_info=True)

    @staticmethod
    def list_all(limit: int = 200):
        path = HistoryStore._file_path()
        if not os.path.exists(path):
            return []
        max_rows = getattr(settings, "HISTORY_MAX_ROWS", 5000) or 5000
        with open(path, encoding="utf-8") as fp:
            tail = deque(fp, maxlen=max_rows)
        rows = []
        for line in tail:
            line = line.strip()
            if not line:
                continue
            try:
                rows.append(json.loads(line))
            except Exception:
                continue
        if limit and limit > 0:
            rows = rows[-limit:]
        return list(reversed(rows))

    @staticmethod
    def clear():
        path = HistoryStore._file_path()
        for p in (path, path + ".1"):
            if os.path.exists(p):
                os.remove(p)
