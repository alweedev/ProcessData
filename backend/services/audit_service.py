from backend.infra.persistence.history_store import HistoryStore


class AuditService:
    @staticmethod
    def record(event_type: str, status: str, details: dict | None = None):
        HistoryStore.append(
            {
                "event_type": event_type,
                "status": status,
                "details": details or {},
            }
        )

    @staticmethod
    def list_events(limit: int = 200):
        return HistoryStore.list_all(limit=limit)

    @staticmethod
    def clear_events():
        HistoryStore.clear()
