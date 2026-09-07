import json
import os
from datetime import datetime, timezone

from backend.core.config import settings


class HistoryStore:
    @staticmethod
    def _file_path():
        return settings.HISTORY_LOG_FILE

    @staticmethod
    def append(event: dict):
        path = HistoryStore._file_path()
        os.makedirs(os.path.dirname(path), exist_ok=True)
        payload = {
            "timestamp": datetime.now(timezone.utc).isoformat(),
            **(event or {}),
        }
        with open(path, "a", encoding="utf-8") as fp:
            fp.write(json.dumps(payload, ensure_ascii=False) + "\n")

    @staticmethod
    def list_all(limit: int = 200):
        path = HistoryStore._file_path()
        if not os.path.exists(path):
            return []
        rows = []
        with open(path, "r", encoding="utf-8") as fp:
            for line in fp:
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
        if os.path.exists(path):
            os.remove(path)
