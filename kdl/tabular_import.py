"""Helpers for importing CSV/HTML tabular files into openpyxl workbooks."""

from __future__ import annotations

import csv
import os
import tempfile
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


def detect_source_format(filepath: str) -> str:
    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".csv":
        return "csv"
    if ext in (".html", ".htm"):
        return "html"

    try:
        with open(filepath, "rb") as f:
            sample = f.read(16384)
    except OSError:
        return "excel"

    if sample.startswith(b"PK\x03\x04"):
        return "excel"
    if sample.startswith(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"):
        return "excel"

    lowered = sample.decode("latin-1", errors="ignore").lower()
    html_markers = (
        "<html",
        "<table",
        "<tr",
        "<td",
        "<th",
        "<!doctype html",
        "urn:schemas-microsoft-com:office:excel",
        "content-type",
    )
    if any(marker in lowered for marker in html_markers):
        return "html"

    return "excel"


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
    source_kind = detect_source_format(filepath)
    if source_kind == "csv":
        return [_sheet_safe_name(_base_name(filepath), "Sheet1")]
    if source_kind == "html":
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

    source_kind = detect_source_format(filepath)
    if source_kind == "csv":
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = list_source_sheet_names(filepath)[0]
        with open(filepath, newline="", encoding="utf-8-sig") as f:
            for row_vals in csv.reader(f):
                ws.append(row_vals)
        return wb

    if source_kind == "html":
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


def convert_legacy_xls_to_temp_xlsx(filepath: str) -> str:
    excel_app = None
    excel_wb = None
    pythoncom = None
    temp_path = ""
    try:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        excel_app = win32com.client.DispatchEx("Excel.Application")
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        excel_wb = excel_app.Workbooks.Open(
            os.path.abspath(filepath),
            UpdateLinks=0,
            ReadOnly=True,
        )
        fd, temp_path = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        os.unlink(temp_path)
        excel_wb.SaveAs(temp_path, 51)
        return temp_path
    except Exception as exc:
        if temp_path and os.path.exists(temp_path):
            try:
                os.unlink(temp_path)
            except OSError:
                pass
        raise RuntimeError(
            "Could not convert this legacy .xls workbook. Open it in Excel and save as .xlsx, "
            "or make sure Excel is installed."
        ) from exc
    finally:
        if excel_wb is not None:
            try:
                excel_wb.Close(False)
            except Exception:
                pass
        if excel_app is not None:
            try:
                excel_app.Quit()
            except Exception:
                pass
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
