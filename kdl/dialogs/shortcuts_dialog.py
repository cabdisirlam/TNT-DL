"""
NT_DL Shortcuts Dialog
Edit and manage shortcut commands and their keystroke equivalents.
"""

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QPushButton, QHeaderView, QMessageBox, QLabel
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QGuiApplication

from kdl.engine.keystroke_parser import DEFAULT_SHORTCUTS
from kdl.styles import dialog_qss, accent_button_qss, TEXT_MUTED


class ShortcutsDialog(QDialog):
    """Dialog to view and edit NT_DL shortcut commands."""

    def __init__(self, shortcuts: dict = None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("NT_DL - Edit Shortcuts / Commands")
        self.setMinimumSize(460, 360)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.setWindowFlag(Qt.WindowCloseButtonHint, True)

        self.shortcuts = dict(shortcuts or DEFAULT_SHORTCUTS)
        self.setStyleSheet(dialog_qss())
        self._build_ui()
        self._populate()
        self._fit_to_screen()

    def _build_ui(self):
        layout = QVBoxLayout(self)

        info_label = QLabel(
            "Edit shortcuts below. Use * prefix for shortcut names.\n"
            "Default standard uses \\{TAB} for Tab (e.g., *S, *DN, \\{TAB}). "
            "Plain tab is accepted and Convert Macros will normalize it."
        )
        info_label.setStyleSheet(f"color: {TEXT_MUTED}; margin-bottom: 8px;")
        layout.addWidget(info_label)

        # Table
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Shortcut", "Keystroke", "Description"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.setFont(QFont("Consolas", 13))
        layout.addWidget(self.table)

        # Buttons
        btn_layout = QHBoxLayout()

        add_btn = QPushButton("Add Shortcut")
        add_btn.clicked.connect(self._add_row)
        btn_layout.addWidget(add_btn)

        remove_btn = QPushButton("Remove Selected")
        remove_btn.clicked.connect(self._remove_row)
        btn_layout.addWidget(remove_btn)

        reset_btn = QPushButton("Reset to Defaults")
        reset_btn.clicked.connect(self._reset_defaults)
        btn_layout.addWidget(reset_btn)

        btn_layout.addStretch()

        cancel_btn = QPushButton("Close")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)

        save_btn = QPushButton("Save")
        save_btn.setStyleSheet(accent_button_qss())
        save_btn.clicked.connect(self._save)
        btn_layout.addWidget(save_btn)

        layout.addLayout(btn_layout)

    # Descriptions for built-in shortcuts
    _DESCRIPTIONS = {
        "*SP": "Save & Proceed",
        "*SV": "Save (Commit)",
        "*S": "Save (Commit)",
        "*NR": "New Record",
        "*NX": "Next row (Down + Home)",
        "*DN": "Down one step",
        "*PV": "Previous Record",
        "*NB": "Next Block/Field",
        "*CL": "Clear / Cancel",
        "*EX": "Exit Form",
        "*DL": "Delete Record",
        "*QR": "Enter Query",
        "*EQ": "Execute Query",
        "*CM": "Commit (Save)",
        "*DF": "Duplicate Field",
        "*DR": "Duplicate Record",
        "*LOV": "List of Values",
    }

    def _populate(self):
        self.table.setRowCount(len(self.shortcuts))
        for row, (shortcut, keystroke) in enumerate(sorted(self.shortcuts.items())):
            sc_item = QTableWidgetItem(shortcut)
            ks_item = QTableWidgetItem(keystroke)
            desc = self._DESCRIPTIONS.get(shortcut.upper(), "")
            desc_item = QTableWidgetItem(desc)
            self.table.setItem(row, 0, sc_item)
            self.table.setItem(row, 1, ks_item)
            self.table.setItem(row, 2, desc_item)

    def _add_row(self):
        row = self.table.rowCount()
        self.table.insertRow(row)
        self.table.setItem(row, 0, QTableWidgetItem("*NEW"))
        self.table.setItem(row, 1, QTableWidgetItem("\\"))
        self.table.setItem(row, 2, QTableWidgetItem(""))

    def _remove_row(self):
        row = self.table.currentRow()
        if row >= 0:
            self.table.removeRow(row)

    def _reset_defaults(self):
        self.shortcuts = dict(DEFAULT_SHORTCUTS)
        self._populate()

    def _save(self):
        """Save shortcuts and close."""
        self.shortcuts = {}
        for row in range(self.table.rowCount()):
            sc = self.table.item(row, 0)
            ks = self.table.item(row, 1)
            if sc and ks and sc.text().strip() and ks.text().strip():
                self.shortcuts[sc.text().strip().upper()] = ks.text().strip()
        self.accept()

    def get_shortcuts(self) -> dict:
        return self.shortcuts

    def _fit_to_screen(self):
        screen = self.screen() or QGuiApplication.primaryScreen()
        if not screen:
            return
        geo = screen.availableGeometry()
        max_w = max(380, geo.width() - 24)
        max_h = max(320, geo.height() - 24)
        self.setMaximumSize(max_w, max_h)
        hint = self.sizeHint()
        target_w = min(max(self.minimumWidth(), hint.width()), max_w)
        target_h = min(max(self.minimumHeight(), hint.height()), max_h)
        self.resize(target_w, target_h)
