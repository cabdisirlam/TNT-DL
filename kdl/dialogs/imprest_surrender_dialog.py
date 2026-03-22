"""
Imprest Surrender AP Loader Dialog.

Workflow:
  1. Export blank Excel template  →  user fills it in
  2. User uploads the filled file →  system reads invoice rows
  3. Preview shows rows found
  4. Click "Load into Grid" → rows are loaded into the NT_DL spreadsheet
  5. User then presses F5 (Load), selects "Imprest Surrender" mode in
     Load Settings, and the keystrokes are sent to IFMIS from the grid.
"""

import os

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QGroupBox, QLabel,
    QPushButton, QLineEdit, QTextEdit,
    QFileDialog, QMessageBox,
)

from kdl.styles import accent_button_qss, dialog_qss, TEXT_MUTED


def _default_dir() -> str:
    downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    return downloads if os.path.isdir(downloads) else os.path.expanduser("~")


class ImprestSurrenderDialog(QDialog):
    """Converts a filled AP Imprest Surrender Excel template into grid rows."""

    load_into_grid = Signal(list)   # emits list[list] — one inner list per invoice row

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("NT_DL — Imprest Surrender AP Loader")
        self.setMinimumWidth(500)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.setWindowFlag(Qt.WindowCloseButtonHint, True)

        from kdl.config_store import get_dark_mode
        self.setStyleSheet(dialog_qss(dark=get_dark_mode()))

        self._rows: list = []       # list of dicts from read_invoice_rows
        self._filepath: str = ""

        self._build_ui()

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(6)
        layout.setContentsMargins(8, 8, 8, 8)

        # ── Step 1: Export template ──
        tpl_group = QGroupBox("Step 1 — Download Sample Template")
        tg = QVBoxLayout(tpl_group)
        tg.setSpacing(4)

        tpl_desc = QLabel(
            "Export a blank Excel template (Data_Entry sheet). "
            "Fill in the yellow cells — one row per invoice.")
        tpl_desc.setWordWrap(True)
        tg.addWidget(tpl_desc)

        self._export_btn = QPushButton("Export Sample Template…")
        self._export_btn.setFixedHeight(26)
        self._export_btn.clicked.connect(self._export_template)
        tg.addWidget(self._export_btn)
        layout.addWidget(tpl_group)

        # ── Step 2: Upload filled sheet ──
        upload_group = QGroupBox("Step 2 — Upload Filled Template")
        ug = QVBoxLayout(upload_group)
        ug.setSpacing(4)

        browse_row = QHBoxLayout()
        self._path_edit = QLineEdit()
        self._path_edit.setPlaceholderText("Select your filled .xlsx file…")
        self._path_edit.setReadOnly(True)
        browse_row.addWidget(self._path_edit, 1)

        self._browse_btn = QPushButton("Browse…")
        self._browse_btn.setFixedWidth(72)
        self._browse_btn.clicked.connect(self._browse_file)
        browse_row.addWidget(self._browse_btn)
        ug.addLayout(browse_row)

        self._upload_status = QLabel("")
        self._upload_status.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 12px;")
        ug.addWidget(self._upload_status)
        layout.addWidget(upload_group)

        # ── Step 3: Preview ──
        preview_group = QGroupBox("Step 3 — Invoice Preview")
        pg = QVBoxLayout(preview_group)
        pg.setSpacing(2)

        self._preview = QTextEdit()
        self._preview.setReadOnly(True)
        self._preview.setFixedHeight(96)
        self._preview.setPlaceholderText("Invoice rows will appear here after upload…")
        self._preview.setStyleSheet(
            "font-family: Consolas, 'Courier New', monospace; font-size: 11px;")
        pg.addWidget(self._preview)
        layout.addWidget(preview_group)

        # ── Info hint ──
        hint = QLabel(
            "After loading into the grid, press F5 → select Per Cell mode → "
            "After each row: None (the \\*dn keystroke advances DataLoad automatically).")
        hint.setWordWrap(True)
        hint.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 12px; margin-top: 2px;")
        layout.addWidget(hint)

        # ── Buttons ──
        btn_row = QHBoxLayout()
        btn_row.addStretch()

        close_btn = QPushButton("Close")
        close_btn.setFixedWidth(72)
        close_btn.clicked.connect(self.reject)
        btn_row.addWidget(close_btn)

        from kdl.config_store import get_dark_mode
        self._load_btn = QPushButton("  Load into Grid  ")
        self._load_btn.setDefault(True)
        self._load_btn.setFixedHeight(28)
        self._load_btn.setEnabled(False)
        self._load_btn.setStyleSheet(accent_button_qss(dark=get_dark_mode()))
        self._load_btn.clicked.connect(self._load_into_grid)
        btn_row.addWidget(self._load_btn)
        layout.addLayout(btn_row)

    # ── Export template ───────────────────────────────────────────────────────

    def _export_template(self):
        from kdl.engine.imprest_surrender_engine import export_template
        path, _ = QFileDialog.getSaveFileName(
            self, "Export Template",
            os.path.join(_default_dir(), "AP_Imprest_Surrender_Template.xlsx"),
            "Excel Files (*.xlsx)")
        if not path:
            return
        err = export_template(path)
        if err:
            QMessageBox.critical(self, "Export Error", err)
        else:
            self._upload_status.setText(f"Template saved: {os.path.basename(path)}")
            self._upload_status.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 12px;")

    # ── Browse / load rows ────────────────────────────────────────────────────

    def _browse_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open Filled Template",
            _default_dir(),
            "Excel Files (*.xlsx *.xls)")
        if not path:
            return
        self._path_edit.setText(path)
        self._filepath = path
        self._read_rows(path)

    def _read_rows(self, path: str):
        from kdl.engine.imprest_surrender_engine import read_invoice_rows, build_row_summary

        self._rows = []
        self._preview.clear()
        self._load_btn.setEnabled(False)
        self._upload_status.setText("Reading file…")
        self._upload_status.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 12px;")

        rows, err = read_invoice_rows(path)

        if err:
            self._upload_status.setText(f"Error: {err}")
            self._upload_status.setStyleSheet("color: #d9534f; font-size: 12px;")
            return

        if not rows:
            self._upload_status.setText(
                "No invoice rows found. Make sure data starts at row 5.")
            self._upload_status.setStyleSheet("color: #e8a900; font-size: 12px;")
            return

        self._rows = rows
        self._upload_status.setText(f"✓  {len(rows)} invoice(s) ready to load")
        self._upload_status.setStyleSheet("color: #5cb85c; font-size: 12px;")

        lines = [f"Row {i}: {build_row_summary(r)}" for i, r in enumerate(rows, 1)]
        self._preview.setPlainText("\n".join(lines))
        self._load_btn.setEnabled(True)

    # ── Load into grid ────────────────────────────────────────────────────────

    def _load_into_grid(self):
        if not self._rows:
            QMessageBox.warning(self, "No Data", "Upload a filled template first.")
            return

        from kdl.engine.imprest_surrender_engine import build_keystroke_row

        # One 68-cell keystroke row per invoice, starting at grid row 1.
        grid_rows = [build_keystroke_row(row) for row in self._rows]

        self.load_into_grid.emit(grid_rows)
        self.accept()
