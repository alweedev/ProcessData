import hashlib
import json
from typing import Any


def _canonico(value: Any) -> Any:
    if isinstance(value, dict):
        return {str(k): _canonico(v) for k, v in value.items()}
    if isinstance(value, (list, tuple, set, frozenset)):
        items = [_canonico(v) for v in value]
        return sorted(items, key=lambda v: json.dumps(v, sort_keys=True, ensure_ascii=False))
    return value


def impressao_digital(payload: dict[str, Any]) -> str:
    """SHA-256 do JSON canônico (chaves, listas e conjuntos ordenados): identifica o CONTEÚDO do diagnóstico."""
    canonico = json.dumps(_canonico(payload), sort_keys=True, separators=(",", ":"), ensure_ascii=False)
    return hashlib.sha256(canonico.encode("utf-8")).hexdigest()
