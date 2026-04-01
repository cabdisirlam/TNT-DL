"""
KDL Resume State
Persists the last successfully started row for each workbook so the user
can restart a load from where it stopped across sessions.
"""

import json
import os
from typing import Optional

RESUME_FILE_NAME = "resume_state.json"


def _resume_path() -> str:
    appdata = os.getenv("APPDATA") or os.path.expanduser("~")
    return os.path.join(appdata, "KDL", RESUME_FILE_NAME)


def _load_all() -> dict:
    path = _resume_path()
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _save_all(data: dict) -> None:
    path = _resume_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    try:
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(data, fh, indent=2, ensure_ascii=True)
    except Exception:
        pass


def save_resume_row(workbook: str, last_row: int, load_mode: str = "") -> None:
    """
    Persist last_row (0-based) for workbook.
    Called after each row that starts loading so a crash/stop can be resumed.
    """
    key = _normalize_key(workbook)
    if not key:
        return
    data = _load_all()
    data[key] = {
        "last_row": last_row,
        "load_mode": load_mode,
    }
    _save_all(data)


def get_resume_row(workbook: str) -> Optional[int]:
    """
    Return the saved last_row (0-based) for workbook, or None if no saved state.
    """
    key = _normalize_key(workbook)
    if not key:
        return None
    data = _load_all()
    entry = data.get(key)
    if isinstance(entry, dict):
        val = entry.get("last_row")
        if isinstance(val, int):
            return val
    return None


def clear_resume_row(workbook: str) -> None:
    """Remove the saved resume position for a workbook (e.g. after successful completion)."""
    key = _normalize_key(workbook)
    if not key:
        return
    data = _load_all()
    data.pop(key, None)
    _save_all(data)


def _normalize_key(workbook: str) -> str:
    """Normalise a workbook path to a stable dict key."""
    if not workbook:
        return ""
    return os.path.normcase(os.path.normpath(workbook))
