"""
NT DL Load History Dialog
Shows a table of past load runs with timestamps, row counts, and outcomes.
"""

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QFrame,
    QMessageBox,
    QAbstractItemView,
)
from PySide6.QtGui import QColor, QFont

from kdl import __display_name__
from kdl.dialogs.dialog_sizing import fit_dialog_to_screen
from kdl.styles import dialog_qss, ACCENT


_COLUMNS = [
    ("Date / Time",    "timestamp",    140),
    ("Workbook",       "workbook",     160),
    ("Mode",           "load_mode",     90),
    ("Start",          "start_row",     52),
    ("End",            "end_row",       52),
    ("Total",          "total_rows",    52),
    ("OK",             "success_rows",  52),
    ("Fail",           "failed_rows",   52),
    ("Duration",       "duration_sec",  72),
    ("Target",         "target_title", 150),
    ("Result",         "result",        78),
]

_RESULT_COLORS = {
    "success":  ("#1a7a3d", "#d4f5e2"),   # dark-mode text, light-mode bg
    "dry_run":  ("#1155aa", "#daeaff"),
    "stopped":  ("#7a5a00", "#fff5cc"),
    "error":    ("#8a1a1a", "#fde8e8"),
}


class LoadHistoryDialog(QDialog):

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"{__display_name__} — Load History")
        self.setWindowFlag(Qt.WindowContextHelpButtonHint, False)
        self.setMinimumSize(860, 440)

        from kdl.config_store import get_dark_mode
        self._dark = get_dark_mode()
        self.setStyleSheet(dialog_qss(dark=self._dark))

        self._build_ui()
        self._load_data()
        fit_dialog_to_screen(
            self,
            min_width=860,
            min_height=440,
            preferred_width=1100,
            wide_width=1280,
            margin_width=60,
            margin_height=60,
        )

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setSpacing(8)
        root.setContentsMargins(12, 12, 12, 10)

        # Header row
        hdr = QHBoxLayout()
        title_lbl = QLabel("Load History")
        title_lbl.setStyleSheet("font-size: 15px; font-weight: 600;")
        hdr.addWidget(title_lbl)
        hdr.addStretch()

        self._count_lbl = QLabel("")
        self._count_lbl.setStyleSheet("color: #888; font-size: 12px;")
        hdr.addWidget(self._count_lbl)
        root.addLayout(hdr)

        # Separator
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setFrameShadow(QFrame.Sunken)
        root.addWidget(sep)

        # Table
        self.table = QTableWidget(0, len(_COLUMNS))
        self.table.setHorizontalHeaderLabels([c[0] for c in _COLUMNS])
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setHighlightSections(False)
        self.table.setShowGrid(True)
        self.table.setSortingEnabled(True)

        for col_idx, (_, _, width) in enumerate(_COLUMNS):
            self.table.setColumnWidth(col_idx, width)

        # Stretch the Workbook and Target columns
        hh = self.table.horizontalHeader()
        hh.setSectionResizeMode(1, QHeaderView.Stretch)   # Workbook
        hh.setSectionResizeMode(9, QHeaderView.Stretch)   # Target

        root.addWidget(self.table, 1)

        # Buttons
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        btn_row.addStretch()

        self.clear_btn = QPushButton("Clear History")
        self.clear_btn.setFixedHeight(34)
        self.clear_btn.clicked.connect(self._clear_history)
        btn_row.addWidget(self.clear_btn)

        close_btn = QPushButton("Close")
        close_btn.setFixedHeight(34)
        close_btn.setDefault(True)
        close_btn.clicked.connect(self.accept)
        btn_row.addWidget(close_btn)

        root.addLayout(btn_row)

    def _load_data(self):
        from kdl.engine.load_history import load_history
        entries = load_history()
        self._count_lbl.setText(f"{len(entries)} record{'s' if len(entries) != 1 else ''}")

        self.table.setSortingEnabled(False)
        self.table.setRowCount(len(entries))

        for row_idx, entry in enumerate(entries):
            for col_idx, (_, key, _) in enumerate(_COLUMNS):
                raw = entry.get(key, "")
                if key == "duration_sec":
                    text = f"{raw}s" if raw != "" else ""
                elif key == "load_mode":
                    text = {
                        "per_cell": "Per Cell",
                        "per_row": "Per Row",
                        "fast_send": "Fast Send",
                        "imprest_surrender": "Imprest",
                    }.get(str(raw), str(raw))
                elif key == "result":
                    is_dry = entry.get("dry_run", False)
                    if is_dry:
                        text = "Dry Run"
                    else:
                        text = str(raw).capitalize()
                else:
                    text = str(raw) if raw is not None else ""

                item = QTableWidgetItem(text)
                item.setTextAlignment(Qt.AlignCenter if col_idx not in (0, 1, 9) else Qt.AlignLeft | Qt.AlignVCenter)

                # Colour the Result column
                if key == "result":
                    result_key = "dry_run" if entry.get("dry_run") else str(entry.get("result", ""))
                    colors = _RESULT_COLORS.get(result_key)
                    if colors:
                        fg, bg = colors
                        item.setForeground(QColor(fg))
                        item.setBackground(QColor(bg))
                        bold_font = QFont()
                        bold_font.setBold(True)
                        item.setFont(bold_font)

                self.table.setItem(row_idx, col_idx, item)

        self.table.setSortingEnabled(True)
        self.table.sortByColumn(0, Qt.DescendingOrder)

    def _clear_history(self):
        reply = QMessageBox.question(
            self,
            "Clear History",
            "Clear all load history records?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            from kdl.engine.load_history import clear_history
            clear_history()
            self._load_data()
