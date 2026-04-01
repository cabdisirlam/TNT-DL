"""
KDL Load History
Persists a log of completed load runs to KDL AppData directory.
Each record captures workbook, rows, mode, result, and timing.
"""

import json
import os
from datetime import datetime
from typing import List, Dict, Any

HISTORY_FILE_NAME = "load_history.json"
MAX_HISTORY_ENTRIES = 500


def _history_path() -> str:
    appdata = os.getenv("APPDATA") or os.path.expanduser("~")
    return os.path.join(appdata, "KDL", HISTORY_FILE_NAME)


def load_history() -> List[Dict[str, Any]]:
    """Return all history entries, newest first."""
    path = _history_path()
    if not os.path.exists(path):
        return []
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        if isinstance(data, list):
            return data
    except Exception:
        pass
    return []


def _save_history(entries: List[Dict[str, Any]]) -> None:
    path = _history_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    try:
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(entries, fh, indent=2, ensure_ascii=True)
    except Exception:
        pass


def append_history_entry(
    workbook: str,
    load_mode: str,
    start_row: int,
    end_row: int,
    success_rows: int,
    failed_rows: int,
    duration_sec: float,
    target_title: str,
    result: str,
    dry_run: bool = False,
) -> None:
    """Append one record to the history log."""
    entry: Dict[str, Any] = {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "workbook": os.path.basename(workbook) if workbook else "(unsaved)",
        "load_mode": load_mode,
        "start_row": start_row + 1,   # 1-based for display
        "end_row": end_row + 1,
        "total_rows": max(0, end_row - start_row + 1),
        "success_rows": success_rows,
        "failed_rows": failed_rows,
        "duration_sec": round(duration_sec, 1),
        "target_title": target_title or "",
        "result": result,             # "success" | "stopped" | "error" | "dry_run"
        "dry_run": dry_run,
    }
    entries = load_history()
    entries.insert(0, entry)
    if len(entries) > MAX_HISTORY_ENTRIES:
        entries = entries[:MAX_HISTORY_ENTRIES]
    _save_history(entries)


def clear_history() -> None:
    _save_history([])
