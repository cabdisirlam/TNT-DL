"""Helpers for importing CSV/HTML tabular files into openpyxl workbooks."""

from __future__ import annotations

import csv
import os
from html.parser import HTMLParser


class _HTMLTableParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables: list[list[list[str]]] = []
        self._table_depth = 0
        self._current_table: list[list[str]] | None = None
        self._current_row: list[str] | None = None
        self._current_cell: list[str] | None = None

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag == "table":
            if self._table_depth == 0:
                self._current_table = []
            self._table_depth += 1
            return

        if self._table_depth != 1:
            return

        if tag == "tr":
            self._current_row = []
        elif tag in ("td", "th"):
            self._current_cell = []
        elif tag == "br" and self._current_cell is not None:
            self._current_cell.append("\n")

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag == "table":
            if self._table_depth == 1 and self._current_table:
                if any(row for row in self._current_table):
                    self.tables.append(self._current_table)
                self._current_table = None
            if self._table_depth > 0:
                self._table_depth -= 1
            return

        if self._table_depth != 1:
            return

        if tag in ("td", "th") and self._current_row is not None and self._current_cell is not None:
            text = "".join(self._current_cell)
            text = "\n".join(part.strip() for part in text.splitlines() if part.strip())
            self._current_row.append(text)
            self._current_cell = None
        elif tag == "tr" and self._current_table is not None and self._current_row is not None:
            if self._current_row:
                self._current_table.append(self._current_row)
            self._current_row = None

    def handle_data(self, data):
        if self._current_cell is not None:
            self._current_cell.append(data)


def _sheet_safe_name(name: str, fallback: str) -> str:
    cleaned = "".join("_" if ch in '[]:*?/\\\\' else ch for ch in str(name).strip())
    cleaned = cleaned.strip("'")
    return (cleaned or fallback)[:31]


def _base_name(filepath: str) -> str:
    return os.path.splitext(os.path.basename(filepath))[0]


def load_html_tables(filepath: str) -> list[list[list[str]]]:
    with open(filepath, "rb") as f:
        raw = f.read()

    for encoding in ("utf-8-sig", "utf-16", "cp1252", "latin-1"):
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = raw.decode("utf-8", errors="replace")

    parser = _HTMLTableParser()
    parser.feed(text)
    parser.close()
    return parser.tables


def list_source_sheet_names(filepath: str) -> list[str]:
    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".csv":
        return [_sheet_safe_name(_base_name(filepath), "Sheet1")]
    if ext in (".html", ".htm"):
        tables = load_html_tables(filepath)
        if not tables:
            raise RuntimeError("No HTML tables found in the selected file.")
        base = _base_name(filepath)
        names = []
        for idx, _rows in enumerate(tables, start=1):
            suffix = f"_{idx}" if len(tables) > 1 else ""
            names.append(_sheet_safe_name(f"{base}{suffix}", f"Table_{idx}"))
        return names
    raise ValueError(f"Unsupported tabular source: {filepath}")


def build_workbook_from_source(filepath: str):
    import openpyxl

    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".csv":
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = list_source_sheet_names(filepath)[0]
        with open(filepath, newline="", encoding="utf-8-sig") as f:
            for row_vals in csv.reader(f):
                ws.append(row_vals)
        return wb

    if ext in (".html", ".htm"):
        tables = load_html_tables(filepath)
        if not tables:
            raise RuntimeError("No HTML tables found in the selected file.")
        names = list_source_sheet_names(filepath)
        wb = openpyxl.Workbook()
        for idx, rows in enumerate(tables):
            ws = wb.active if idx == 0 else wb.create_sheet()
            ws.title = names[idx]
            for row_vals in rows:
                ws.append(row_vals)
        return wb

    raise ValueError(f"Unsupported tabular source: {filepath}")
