from __future__ import annotations

import json
from pathlib import Path
from typing import Any


ROOT_DIR = Path(__file__).resolve().parents[1]
DEFAULT_CONFIG_PATH = ROOT_DIR / "config" / "app_config.json"


def load_app_config(raw_bytes: bytes | None = None) -> dict[str, Any]:
    if raw_bytes is None:
        return _load_json_file(DEFAULT_CONFIG_PATH)

    try:
        payload = json.loads(raw_bytes.decode("utf-8"))
    except UnicodeDecodeError as exc:
        raise ValueError("JSON-конфиг должен быть в кодировке UTF-8.") from exc
    except json.JSONDecodeError as exc:
        raise ValueError(f"Не удалось разобрать JSON-конфиг: {exc.msg}.") from exc

    return _validate_top_level(payload)


def _load_json_file(path: Path) -> dict[str, Any]:
    if not path.exists():
        raise ValueError(f"Файл конфигурации не найден: {path}")

    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"Ошибка в {path.name}: {exc.msg}.") from exc

    return _validate_top_level(payload)


def _validate_top_level(payload: Any) -> dict[str, Any]:
    if not isinstance(payload, dict):
        raise ValueError("Корень JSON-конфига должен быть объектом.")

    pages = payload.get("pages")
    source_presets = payload.get("source_cols_preset")

    if not isinstance(pages, list):
        raise ValueError("Поле 'pages' должно быть массивом.")
    if not isinstance(source_presets, dict):
        raise ValueError("Поле 'source_cols_preset' должно быть объектом.")

    return payload
