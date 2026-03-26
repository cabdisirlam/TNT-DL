"""
Imprest Surrender AP Loader dialog.

Workflow:
1. Import the IFMIS export to create a prefilled template.
2. Open the completed template after the remaining amber fields are filled.
3. Review the invoice rows and either load them into the TNT DL grid or
   export a DataLoad fallback workbook.
"""

import os

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QDialog,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
)

from kdl.dialogs.dialog_sizing import create_hint_button, fit_dialog_to_screen
from kdl.styles import TEXT_MUTED, accent_button_qss, dialog_qss


def _default_dir() -> str:
    downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    return downloads if os.path.isdir(downloads) else os.path.expanduser("~")


class ImprestSurrenderDialog(QDialog):
    """Convert a completed AP Imprest Surrender workbook into grid rows."""

    load_into_grid = Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Imprest Surrender AP Loader")
        self.setMinimumWidth(640)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.setWindowFlag(Qt.WindowCloseButtonHint, True)

        from kdl.config_store import get_dark_mode

        self.setStyleSheet(dialog_qss(dark=get_dark_mode()))

        self._rows: list = []
        self._filepath: str = ""

        self._build_ui()
        self._fit_to_screen()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(16, 16, 16, 16)

        intro_row = QHBoxLayout()
        intro_row.setSpacing(8)
        intro = QLabel("Build and review the Imprest workbook, then send the rows to the grid.")
        intro.setObjectName("DialogIntro")
        intro.setWordWrap(True)
        intro_row.addWidget(intro, 1)
        intro_row.addWidget(
            create_hint_button(
                "Create a ready-to-load Imprest workbook in three steps: import the IFMIS "
                "export, open the completed template, then review the invoice rows.",
                label="i",
            )
        )
        layout.addLayout(intro_row)

        import_group = QGroupBox("Step 1 - Import IFMIS Export")
        import_layout = QVBoxLayout(import_group)
        import_layout.setSpacing(8)

        import_note_row = QHBoxLayout()
        import_note_row.setSpacing(8)
        import_note = QLabel("Import the IFMIS AP export to create the prefilled template.")
        import_note.setObjectName("DialogHint")
        import_note.setWordWrap(True)
        import_note_row.addWidget(import_note, 1)
        import_note_row.addWidget(
            create_hint_button(
                "TNT DL fills the known fields and marks the three values you still "
                "need to complete: Auth Ref, Admin Code, and Distribution Account.",
                label="i",
            )
        )
        import_layout.addLayout(import_note_row)

        self._import_ifmis_btn = QPushButton("Import IFMIS File...")
        self._import_ifmis_btn.setMinimumWidth(180)
        self._import_ifmis_btn.setMinimumHeight(38)
        self._import_ifmis_btn.clicked.connect(self._import_ifmis)
        import_layout.addWidget(self._import_ifmis_btn)

        self._import_status = QLabel("")
        self._import_status.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 12px;")
        self._import_status.setWordWrap(True)
        import_layout.addWidget(self._import_status)
        layout.addWidget(import_group)

        open_group = QGroupBox("Step 2 - Open Completed Template")
        open_layout = QVBoxLayout(open_group)
        open_layout.setSpacing(8)

        browse_row = QHBoxLayout()
        browse_row.setSpacing(10)
        self._path_edit = QLineEdit()
        self._path_edit.setPlaceholderText("Choose the completed .xlsx template...")
        self._path_edit.setReadOnly(True)
        browse_row.addWidget(self._path_edit, 1)

        self._browse_btn = QPushButton("Browse...")
        self._browse_btn.setMinimumWidth(112)
        self._browse_btn.setMinimumHeight(38)
        self._browse_btn.clicked.connect(self._browse_file)
        browse_row.addWidget(self._browse_btn)
        open_layout.addLayout(browse_row)

        self._upload_status = QLabel("")
        self._upload_status.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 12px;")
        self._upload_status.setWordWrap(True)
        open_layout.addWidget(self._upload_status)
        layout.addWidget(open_group)

        preview_group = QGroupBox("Step 3 - Review Invoice Rows")
        preview_layout = QVBoxLayout(preview_group)
        preview_layout.setSpacing(8)

        self._preview = QTextEdit()
        self._preview.setReadOnly(True)
        self._preview.setFixedHeight(128)
        self._preview.setPlaceholderText(
            "Loaded invoice rows will appear here after you open the template."
        )
        self._preview.setStyleSheet(
            "font-family: Consolas, 'Courier New', monospace; font-size: 12px;"
        )
        preview_layout.addWidget(self._preview)
        layout.addWidget(preview_group)

        button_row = QHBoxLayout()
        button_row.setSpacing(10)

        close_btn = QPushButton("Close")
        close_btn.setMinimumWidth(104)
        close_btn.clicked.connect(self.reject)
        button_row.addWidget(close_btn)

        self._ks_btn = QPushButton("Export DataLoad File...")
        self._ks_btn.setMinimumWidth(180)
        self._ks_btn.setMinimumHeight(38)
        self._ks_btn.setEnabled(False)
        self._ks_btn.setToolTip(
            "Save a workbook copy with a DL_Keystrokes sheet as a fallback.\n"
            "Load that sheet in DataLoad using Per Cell mode and Use Alternate Method."
        )
        self._ks_btn.clicked.connect(self._export_keystrokes)
        button_row.addWidget(self._ks_btn)

        from kdl.config_store import get_dark_mode

        self._load_btn = QPushButton("Load Rows into Grid")
        self._load_btn.setDefault(True)
        self._load_btn.setMinimumWidth(170)
        self._load_btn.setMinimumHeight(38)
        self._load_btn.setEnabled(False)
        self._load_btn.setStyleSheet(accent_button_qss(dark=get_dark_mode()))
        self._load_btn.setToolTip("Load the raw invoice values into columns A:K of the grid.")
        self._load_btn.clicked.connect(self._load_into_grid)
        button_row.addWidget(self._load_btn)
        button_row.addStretch()
        layout.addLayout(button_row)

    def _import_ifmis(self):
        from kdl.engine.imprest_surrender_engine import (
            IFMIS_BLANK_COLS,
            export_prefilled_template,
            import_ifmis_export,
        )

        src, _ = QFileDialog.getOpenFileName(
            self,
            "Open IFMIS Export",
            _default_dir(),
            "All Supported (*.xlsx *.xls *.csv);;Excel Files (*.xlsx *.xls);;CSV Files (*.csv);;All Files (*)",
        )
        if not src:
            return

        self._import_status.setText("Reading IFMIS export...")
        self._import_status.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 12px;")

        rows, skipped, err = import_ifmis_export(src)
        if err:
            self._import_status.setText(f"Error: {err}")
            self._import_status.setStyleSheet("color: #d9534f; font-size: 12px;")
            return

        if not rows:
            self._import_status.setText("No Prepayment rows found in the IFMIS export.")
            self._import_status.setStyleSheet("color: #e8a900; font-size: 12px;")
            return

        dest, _ = QFileDialog.getSaveFileName(
            self,
            "Save Prefilled Template",
            os.path.join(_default_dir(), "AP_Imprest_Surrender_Prefilled.xlsx"),
            "Excel Files (*.xlsx)",
        )
        if not dest:
            return

        err = export_prefilled_template(dest, rows)
        if err:
            self._import_status.setText(f"Export error: {err}")
            self._import_status.setStyleSheet("color: #d9534f; font-size: 12px;")
            return

        blank_names = ", ".join(sorted(IFMIS_BLANK_COLS))
        message = (
            f"{len(rows)} row(s) imported"
            f"{f', {skipped} non-Prepayment row(s) skipped' if skipped else ''}.\n"
            f"Template saved. Fill the amber columns ({blank_names}), then open it below."
        )
        self._import_status.setText(message)
        self._import_status.setStyleSheet("color: #5cb85c; font-size: 12px;")

    def _browse_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Open Completed Template",
            _default_dir(),
            "All Supported (*.xlsx *.xls *.csv);;Excel Files (*.xlsx *.xls);;CSV Files (*.csv);;All Files (*)",
        )
        if not path:
            return
        self._path_edit.setText(path)
        self._filepath = path
        self._read_rows(path)

    def _read_rows(self, path: str):
        from kdl.engine.imprest_surrender_engine import build_row_summary, read_invoice_rows

        self._rows = []
        self._preview.clear()
        self._load_btn.setEnabled(False)
        self._ks_btn.setEnabled(False)
        self._upload_status.setText("Reading workbook...")
        self._upload_status.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 12px;")

        rows, err = read_invoice_rows(path)

        if err:
            self._upload_status.setText(f"Error: {err}")
            self._upload_status.setStyleSheet("color: #d9534f; font-size: 12px;")
            return

        if not rows:
            self._upload_status.setText(
                "No invoice rows found. Confirm that the data starts on row 5."
            )
            self._upload_status.setStyleSheet("color: #e8a900; font-size: 12px;")
            return

        self._rows = rows
        self._upload_status.setText(f"Ready: {len(rows)} invoice row(s) loaded from the template.")
        self._upload_status.setStyleSheet("color: #5cb85c; font-size: 12px;")

        lines = [f"Row {index}: {build_row_summary(row)}" for index, row in enumerate(rows, 1)]
        self._preview.setPlainText("\n".join(lines))
        self._load_btn.setEnabled(True)
        self._ks_btn.setEnabled(True)

    def _export_keystrokes(self):
        if not self._rows or not self._filepath:
            QMessageBox.warning(self, "No Data", "Open a completed template first.")
            return

        from kdl.engine.imprest_surrender_engine import export_keystroke_sheet_to_workbook

        source_dir = os.path.dirname(self._filepath) or _default_dir()
        source_name, source_ext = os.path.splitext(os.path.basename(self._filepath))
        save_ext = source_ext if source_ext.lower() in (".xlsx", ".xlsm") else ".xlsx"
        default_path = os.path.join(
            source_dir,
            f"{source_name}_with_DL_Keystrokes{save_ext}",
        )

        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Workbook with DataLoad Sheet",
            default_path,
            "Excel Files (*.xlsx *.xlsm)",
        )
        if not path:
            return

        err = export_keystroke_sheet_to_workbook(self._filepath, path, self._rows)
        if err:
            QMessageBox.critical(self, "Export Error", err)
        else:
            self._upload_status.setText(
                f"Workbook saved with DL_Keystrokes sheet: {os.path.basename(path)}"
            )
            self._upload_status.setStyleSheet("color: #5cb85c; font-size: 12px;")

    def _load_into_grid(self):
        if not self._rows:
            QMessageBox.warning(self, "No Data", "Open a completed template first.")
            return

        from kdl.engine.imprest_surrender_engine import COLUMNS

        grid_rows = [[row.get(column, "") for column in COLUMNS] for row in self._rows]
        self.load_into_grid.emit(grid_rows)
        self.accept()

    def _fit_to_screen(self):
        fit_dialog_to_screen(
            self,
            min_width=760,
            min_height=540,
            preferred_width=940,
            wide_width=1080,
            margin_width=72,
            margin_height=80,
            extra_hint_width=40,
            extra_hint_height=28,
        )
