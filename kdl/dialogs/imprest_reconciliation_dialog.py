"""Dialog for the standalone Imprest reconciliation tool."""

from __future__ import annotations

import os

from PySide6.QtCore import Qt, QThread, QUrl, Signal
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from kdl.config_store import get_dark_mode
from kdl.dialogs.dialog_sizing import create_hint_button, fit_dialog_to_screen
from kdl.engine.imprest_reconciliation import (
    create_imprest_reconciliation,
    suggest_imprest_reconciliation_output,
)
from kdl.styles import accent_button_qss, dialog_qss, themed_button_qss
from kdl.tabular_import import list_excel_sheet_names


SUPPORTED_FILES = (
    "Supported Workbooks (*.xlsx *.xlsm *.xls *.xlsb *.xltx *.xltm *.xlt "
    "*.csv *.html *.htm);;Excel Workbooks (*.xlsx *.xlsm *.xls *.xlsb *.xltx "
    "*.xltm *.xlt);;CSV Files (*.csv);;HTML Files (*.html *.htm);;All Files (*)"
)


def _default_browse_dir() -> str:
    downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    return downloads if os.path.isdir(downloads) else os.path.expanduser("~")


class _SheetLoader(QThread):
    completed = Signal(list)
    failed = Signal(str)

    def __init__(self, source_path: str):
        super().__init__()
        self.source_path = source_path

    def run(self):
        try:
            names = list_excel_sheet_names(self.source_path)
            if not names:
                raise ValueError("The selected workbook contains no worksheets.")
        except Exception as exc:
            self.failed.emit(str(exc))
            return
        self.completed.emit(names)


class _ReconciliationWorker(QThread):
    completed = Signal(object)
    failed = Signal(str)

    def __init__(self, source_path: str, output_path: str, sheet_name: str):
        super().__init__()
        self.source_path = source_path
        self.output_path = output_path
        self.sheet_name = sheet_name

    def run(self):
        try:
            result = create_imprest_reconciliation(
                self.source_path,
                self.output_path,
                self.sheet_name,
            )
        except Exception as exc:
            self.failed.emit(str(exc))
            return
        self.completed.emit(result)


class ImprestReconciliationDialog(QDialog):
    """Independent Excel-to-Excel Imprest reconciliation workflow."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Imprest Reconciliation")
        self._sheet_loader = None
        self._worker = None
        self._output_path = ""
        self.setStyleSheet(dialog_qss(dark=get_dark_mode()))
        self._build_ui()
        fit_dialog_to_screen(
            self,
            min_width=540,
            min_height=440,
            preferred_width=680,
            wide_width=720,
            margin_width=48,
            margin_height=48,
            extra_hint_height=0,
        )

    def _build_ui(self):
        dark = get_dark_mode()
        primary_qss = accent_button_qss(dark=dark)
        secondary_qss = themed_button_qss(dark=dark)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(14, 14, 14, 14)
        outer.setSpacing(8)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(9)

        intro_row = QHBoxLayout()
        intro = QLabel(
            "Clean and reconcile Imprest debit and credit transactions in a separate Excel report."
        )
        intro.setObjectName("DialogIntro")
        intro.setWordWrap(True)
        intro_row.addWidget(intro, 1)
        intro_row.addWidget(
            create_hint_button(
                "This is a standalone spreadsheet tool. It does not use or change "
                "the Imprest Surrender and Imprest Old Date loading engines.",
                label="i",
            )
        )
        layout.addLayout(intro_row)

        source_group = QGroupBox("Source Transaction Workbook")
        source_layout = QVBoxLayout(source_group)
        source_row = QHBoxLayout()
        self._source_edit = QLineEdit()
        self._source_edit.setReadOnly(True)
        self._source_edit.setPlaceholderText("Choose the IFMIS transaction workbook or export...")
        browse_button = QPushButton("Browse...")
        browse_button.setMinimumWidth(112)
        browse_button.setStyleSheet(secondary_qss)
        browse_button.clicked.connect(self._browse_source)
        source_row.addWidget(self._source_edit, 1)
        source_row.addWidget(browse_button)
        source_layout.addLayout(source_row)

        sheet_row = QHBoxLayout()
        sheet_label = QLabel("Worksheet:")
        sheet_label.setMinimumWidth(78)
        self._sheet_combo = QComboBox()
        self._sheet_combo.setEnabled(False)
        self._sheet_combo.currentIndexChanged.connect(self._update_generate_state)
        sheet_row.addWidget(sheet_label)
        sheet_row.addWidget(self._sheet_combo, 1)
        source_layout.addLayout(sheet_row)
        layout.addWidget(source_group)

        output_group = QGroupBox("Output Workbook")
        output_row = QHBoxLayout(output_group)
        self._output_edit = QLineEdit()
        self._output_edit.setReadOnly(True)
        self._output_edit.setPlaceholderText(
            "Output location is set after selecting a source workbook..."
        )
        save_as_button = QPushButton("Save As...")
        save_as_button.setMinimumWidth(112)
        save_as_button.setStyleSheet(secondary_qss)
        save_as_button.clicked.connect(self._choose_output)
        output_row.addWidget(self._output_edit, 1)
        output_row.addWidget(save_as_button)
        layout.addWidget(output_group)

        sheets_group = QGroupBox("Generated Workbook")
        sheets_layout = QVBoxLayout(sheets_group)
        sheets_label = QLabel("Imprest_Reconciliation  •  Unmatched_Report")
        sheets_label.setWordWrap(True)
        sheets_layout.addWidget(sheets_label)
        note = QLabel(
            "Exact one-to-one matches are green. Unmatched rows are red, review "
            "items are amber, and the source workbook remains unchanged."
        )
        note.setObjectName("DialogHint")
        note.setWordWrap(True)
        sheets_layout.addWidget(note)
        layout.addWidget(sheets_group)

        action_row = QHBoxLayout()
        action_row.addStretch()
        self._generate_button = QPushButton("Generate Reconciliation")
        self._generate_button.setMinimumWidth(205)
        self._generate_button.setStyleSheet(primary_qss)
        self._generate_button.setEnabled(False)
        self._generate_button.clicked.connect(self._generate)
        action_row.addWidget(self._generate_button)
        layout.addLayout(action_row)

        result_group = QGroupBox("Reconciliation Result")
        result_layout = QVBoxLayout(result_group)
        self._result_text = QTextEdit()
        self._result_text.setReadOnly(True)
        self._result_text.setMinimumHeight(105)
        self._result_text.setPlaceholderText(
            "Matching totals and review warnings will appear here..."
        )
        result_layout.addWidget(self._result_text)
        layout.addWidget(result_group)

        scroll.setWidget(content)
        outer.addWidget(scroll, 1)

        footer = QHBoxLayout()
        self._open_button = QPushButton("Open Workbook")
        self._open_button.setStyleSheet(primary_qss)
        self._open_button.setEnabled(False)
        self._open_button.clicked.connect(self._open_workbook)
        close_button = QPushButton("Close")
        close_button.setStyleSheet(secondary_qss)
        close_button.clicked.connect(self.accept)
        footer.addWidget(self._open_button)
        footer.addStretch()
        footer.addWidget(close_button)
        outer.addLayout(footer)

    def _browse_source(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Imprest Transaction Workbook",
            _default_browse_dir(),
            SUPPORTED_FILES,
        )
        if not path:
            return

        self._source_edit.setText(path)
        self._output_path = suggest_imprest_reconciliation_output(path)
        self._output_edit.setText(self._output_path)
        self._sheet_combo.clear()
        self._sheet_combo.addItem("Loading worksheets...")
        self._sheet_combo.setEnabled(False)
        self._generate_button.setEnabled(False)
        self._open_button.setEnabled(False)
        self._result_text.setPlainText("Reading available worksheets...")

        loader = _SheetLoader(path)
        loader.completed.connect(self._on_sheets_loaded)
        loader.failed.connect(self._on_sheet_load_failed)
        loader.finished.connect(self._release_sheet_loader)
        self._sheet_loader = loader
        loader.start()

    def _on_sheets_loaded(self, names: list):
        self._sheet_combo.clear()
        self._sheet_combo.addItems([str(name) for name in names])
        self._sheet_combo.setEnabled(bool(names))
        self._result_text.setPlainText(
            "Source workbook is ready. Select the transaction worksheet and generate "
            "the separate two-sheet reconciliation report."
        )
        self._update_generate_state()

    def _on_sheet_load_failed(self, message: str):
        self._sheet_combo.clear()
        self._sheet_combo.setEnabled(False)
        self._result_text.setPlainText(f"ERROR:\n{message}")
        QMessageBox.warning(self, "Imprest Reconciliation", message)

    def _release_sheet_loader(self):
        loader = self._sheet_loader
        self._sheet_loader = None
        if loader is not None:
            loader.deleteLater()

    def _choose_output(self):
        source = self._source_edit.text().strip()
        initial = self._output_path or (
            suggest_imprest_reconciliation_output(source)
            if source
            else os.path.join(_default_browse_dir(), "Imprest_Reconciliation.xlsx")
        )
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Imprest Reconciliation",
            initial,
            "Excel Workbook (*.xlsx)",
        )
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        self._output_path = path
        self._output_edit.setText(path)
        self._open_button.setEnabled(False)
        self._update_generate_state()

    def _update_generate_state(self):
        self._generate_button.setEnabled(
            bool(
                self._source_edit.text().strip()
                and self._output_edit.text().strip()
                and self._sheet_combo.isEnabled()
                and self._sheet_combo.currentText().strip()
            )
            and self._worker is None
        )

    def _generate(self):
        source = self._source_edit.text().strip()
        output = self._output_edit.text().strip()
        sheet_name = self._sheet_combo.currentText().strip()
        if not source or not output or not sheet_name:
            return
        if os.path.normcase(os.path.abspath(source)) == os.path.normcase(os.path.abspath(output)):
            QMessageBox.warning(
                self,
                "Choose a Different Output",
                "The output workbook must be different from the source workbook.",
            )
            return
        if os.path.exists(output):
            answer = QMessageBox.question(
                self,
                "Replace Existing Workbook?",
                f"The output workbook already exists:\n{output}\n\nReplace it?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                return

        self._generate_button.setEnabled(False)
        self._generate_button.setText("Reconciling...")
        self._open_button.setEnabled(False)
        self._result_text.setPlainText(
            "Reading transactions, extracting final identifiers, matching debit and "
            "credit rows, and building the independent Excel report..."
        )
        worker = _ReconciliationWorker(source, output, sheet_name)
        worker.completed.connect(self._on_completed)
        worker.failed.connect(self._on_failed)
        worker.finished.connect(self._release_worker)
        self._worker = worker
        worker.start()

    def _on_completed(self, result):
        self._output_path = result.output_path
        self._output_edit.setText(result.output_path)
        self._result_text.setPlainText(result.message)
        self._open_button.setEnabled(os.path.isfile(result.output_path))
        self._reset_generate_button()
        if result.warnings:
            QMessageBox.warning(
                self,
                "Imprest Reconciliation - Review Required",
                "The workbook was created. Review the amber rows and the "
                "Unmatched_Report sheet before using the results.",
            )
        else:
            QMessageBox.information(
                self,
                "Imprest Reconciliation",
                "The separate two-sheet reconciliation workbook was created successfully.",
            )

    def _on_failed(self, message: str):
        self._result_text.setPlainText(f"ERROR:\n{message}")
        self._reset_generate_button()
        QMessageBox.critical(self, "Imprest Reconciliation Error", message)

    def _reset_generate_button(self):
        self._generate_button.setText("Generate Reconciliation")
        self._update_generate_state()

    def _release_worker(self):
        worker = self._worker
        self._worker = None
        if worker is not None:
            worker.deleteLater()
        self._update_generate_state()

    def _open_workbook(self):
        if self._output_path and os.path.isfile(self._output_path):
            QDesktopServices.openUrl(QUrl.fromLocalFile(self._output_path))

    def reject(self):
        running = (
            (self._sheet_loader is not None and self._sheet_loader.isRunning())
            or (self._worker is not None and self._worker.isRunning())
        )
        if running:
            QMessageBox.information(
                self,
                "Imprest Reconciliation",
                "Please wait for the current workbook operation to finish before closing.",
            )
            return
        super().reject()
