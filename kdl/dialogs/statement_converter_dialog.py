"""
Bank Statement Converter Dialog.
Lets the user pick an Excel file, choose sheets, and run the conversion.
"""

import os
import zipfile
import xml.etree.ElementTree as ET

from PySide6.QtCore import Qt, QThread, QTimer, Signal
from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from kdl.styles import accent_button_qss, dialog_qss


def _default_browse_dir() -> str:
    downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    return downloads if os.path.isdir(downloads) else os.path.expanduser("~")


def _fast_xlsx_sheet_names(filepath: str) -> list[str] | None:
    """Read sheet names from an xlsx/xlsm file by parsing xl/workbook.xml directly.
    Returns a list of names or None if the zip approach fails."""
    try:
        with zipfile.ZipFile(filepath, "r") as zf:
            with zf.open("xl/workbook.xml") as f:
                tree = ET.parse(f)
        ns = {"ns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        sheets = tree.findall(".//ns:sheet", ns)
        names = [s.get("name") for s in sheets if s.get("name")]
        return names if names else None
    except Exception:
        return None


class _SheetLoaderWorker(QThread):
    """Background worker that reads sheet names from an Excel file."""

    sheets_ready = Signal(list)   # list[str]
    load_error = Signal(str)

    def __init__(self, filepath: str):
        super().__init__()
        self.filepath = filepath

    def run(self):
        try:
            ext = os.path.splitext(self.filepath)[1].lower()

            # ── Fast path: xlsx / xlsm via zipfile (no workbook load) ──
            if ext in (".xlsx", ".xlsm"):
                names = _fast_xlsx_sheet_names(self.filepath)
                if names:
                    self.sheets_ready.emit(names)
                    return

            # ── .xls: try xlrd first (lightweight) ──
            if ext == ".xls":
                try:
                    import xlrd
                    wb = xlrd.open_workbook(self.filepath, on_demand=True)
                    try:
                        names = wb.sheet_names()
                    finally:
                        release = getattr(wb, "release_resources", None)
                        if callable(release):
                            release()
                    self.sheets_ready.emit(list(names))
                    return
                except Exception:
                    pass

                # Fallback for .xls: win32com
                excel = workbook = None
                try:
                    import win32com.client
                    excel = win32com.client.DispatchEx("Excel.Application")
                    excel.Visible = False
                    excel.DisplayAlerts = False
                    workbook = excel.Workbooks.Open(
                        os.path.abspath(self.filepath),
                        UpdateLinks=0,
                        ReadOnly=True,
                    )
                    names = [sheet.Name for sheet in workbook.Worksheets]
                    self.sheets_ready.emit(names)
                    return
                except Exception as exc:
                    raise RuntimeError(
                        "Could not read this .xls workbook. Open it in Excel and save as "
                        ".xlsx, or make sure Excel is installed for legacy .xls support."
                    ) from exc
                finally:
                    if workbook is not None:
                        try:
                            workbook.Close(False)
                        except Exception:
                            pass
                    if excel is not None:
                        try:
                            excel.Quit()
                        except Exception:
                            pass

            # ── Fallback: openpyxl read-only (slower but universal) ──
            import openpyxl
            wb = openpyxl.load_workbook(
                self.filepath, read_only=True, data_only=True, keep_links=False
            )
            try:
                self.sheets_ready.emit(list(wb.sheetnames))
            finally:
                wb.close()

        except Exception as exc:
            self.load_error.emit(str(exc))


class _ConverterWorker(QThread):
    def __init__(self, filepath: str, sheet_names: list[str], skip_contra: bool = True):
        super().__init__()
        self.filepath = filepath
        self.sheet_names = sheet_names if isinstance(sheet_names, list) else [sheet_names]
        self.skip_contra = skip_contra
        self.wb = None
        self.result = None
        self.error_message = ""

    def run(self):
        tmp_path = None
        excel_app = None
        excel_wb = None
        pythoncom = None
        try:
            import openpyxl
            import tempfile
            from kdl.engine.statement_converter import (
                AUDIT_DETAIL_FIRST_ROW,
                ConversionResult,
                _get_or_create_sheet,
                _write_audit_header,
                _write_rows_to_sheet,
                convert_statement,
            )

            source_ext = os.path.splitext(self.filepath)[1].lower()
            load_path = self.filepath
            if source_ext == ".xls":
                try:
                    import pythoncom
                    import win32com.client

                    pythoncom.CoInitialize()
                    excel_app = win32com.client.DispatchEx("Excel.Application")
                    excel_app.Visible = False
                    excel_app.DisplayAlerts = False
                    excel_wb = excel_app.Workbooks.Open(
                        os.path.abspath(self.filepath),
                        UpdateLinks=0,
                        ReadOnly=True,
                    )
                    fd, tmp_path = tempfile.mkstemp(suffix=".xlsx")
                    os.close(fd)
                    os.unlink(tmp_path)
                    excel_wb.SaveAs(tmp_path, 51)
                    load_path = tmp_path
                except Exception as exc:
                    self.error_message = (
                        "Could not convert this legacy .xls workbook for processing. "
                        "Open it in Excel and save it as .xlsx, or make sure Excel is installed."
                    )
                    raise RuntimeError(self.error_message) from exc
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

            keep_vba = source_ext == ".xlsm" and os.path.splitext(load_path)[1].lower() == ".xlsm"
            wb = openpyxl.load_workbook(
                load_path,
                data_only=True,
                keep_links=False,
                keep_vba=keep_vba,
            )
            multi = len(self.sheet_names) > 1
            all_output_data = []
            all_audit_rows = []
            all_messages = []
            any_success = False

            for sheet_name in self.sheet_names:
                result = convert_statement(wb, sheet_name, skip_contra=self.skip_contra)
                if result.success:
                    any_success = True
                    all_output_data.extend(result.output_data)
                    prefix = f"[{sheet_name}]\n" if multi else ""
                    all_messages.append(f"{prefix}{result.message}")
                    if multi:
                        all_audit_rows.extend(result.audit_rows)
                else:
                    prefix = f"[{sheet_name}] FAILED\n" if multi else "FAILED\n"
                    all_messages.append(f"{prefix}{result.message}")

            if any_success and multi:
                ws_out = _get_or_create_sheet(wb, "Output")
                _write_rows_to_sheet(ws_out, all_output_data)

                ws_audit = _get_or_create_sheet(wb, "Audit_Skipped")
                _write_audit_header(ws_audit)
                for i, row_vals in enumerate(all_audit_rows):
                    for c, val in enumerate(row_vals, 1):
                        ws_audit.cell(row=AUDIT_DETAIL_FIRST_ROW + i, column=c, value=val)

            sep = "\n\n" + ("-" * 40) + "\n\n" if multi else ""
            combined_msg = sep.join(all_messages)

            self.result = ConversionResult(
                success=any_success,
                message=combined_msg,
                output_data=all_output_data,
            )
            if any_success:
                self.wb = wb
        except Exception as exc:
            if not self.error_message:
                self.error_message = str(exc)
        finally:
            if tmp_path and os.path.exists(tmp_path):
                try:
                    os.unlink(tmp_path)
                except Exception:
                    pass


class StatementConverterDialog(QDialog):
    load_into_grid = Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Bank Statement Converter")
        self.setMinimumWidth(640)
        self.setWindowFlag(Qt.WindowCloseButtonHint, True)
        self._worker = None
        self._sheet_loader = None
        self._result = None
        self._wb = None
        self._sheet_checks = []

        from kdl.config_store import get_dark_mode

        self.setStyleSheet(dialog_qss(dark=get_dark_mode()))
        self._build_ui()
        self._fit_to_screen()

    def _release_worker(self):
        worker = self._worker
        self._worker = None
        if worker is None:
            return
        try:
            worker.deleteLater()
        except Exception:
            pass

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(14)
        layout.setContentsMargins(18, 18, 18, 18)

        intro = QLabel(
            "Choose the source workbook, select the sheets to convert, then review "
            "the summary before loading the output into the grid."
        )
        intro.setObjectName("DialogIntro")
        intro.setWordWrap(True)
        layout.addWidget(intro)

        file_group = QGroupBox("Source Workbook")
        file_layout = QHBoxLayout(file_group)
        file_layout.setSpacing(10)
        self._file_edit = QLineEdit()
        self._file_edit.setPlaceholderText("Choose the bank statement workbook (.xlsx, .xls, .xlsm)...")
        self._file_edit.setReadOnly(True)
        browse_btn = QPushButton("Browse...")
        browse_btn.setMinimumWidth(112)
        browse_btn.setMinimumHeight(38)
        browse_btn.clicked.connect(self._browse_file)
        file_layout.addWidget(self._file_edit)
        file_layout.addWidget(browse_btn)
        layout.addWidget(file_group)

        sheet_group = QGroupBox("Sheets")
        sheet_outer = QVBoxLayout(sheet_group)
        sheet_outer.setSpacing(8)

        sel_row = QHBoxLayout()
        sel_all_btn = QPushButton("Select All")
        sel_all_btn.setMinimumWidth(108)
        sel_all_btn.clicked.connect(self._select_all_sheets)
        sel_none_btn = QPushButton("Clear All")
        sel_none_btn.setMinimumWidth(108)
        sel_none_btn.clicked.connect(self._select_no_sheets)
        sel_row.addWidget(sel_all_btn)
        sel_row.addWidget(sel_none_btn)
        sel_row.addStretch()
        sheet_outer.addLayout(sel_row)

        self._sheet_check_container = QWidget()
        self._sheet_check_layout = QGridLayout(self._sheet_check_container)
        self._sheet_check_layout.setSpacing(8)
        self._sheet_check_layout.setContentsMargins(4, 2, 4, 2)
        sheet_outer.addWidget(self._sheet_check_container)
        layout.addWidget(sheet_group)

        options_group = QGroupBox("Conversion Options")
        options_layout = QVBoxLayout(options_group)
        self._skip_contra_check = QCheckBox("Skip contra/duplicate matched rows (CONTRA_MATCHED)")
        self._skip_contra_check.setChecked(True)
        self._skip_contra_check.setToolTip(
            "When checked, rows that match as debit/credit pairs on the same reference "
            "and amount are excluded from output. Uncheck to include all rows."
        )
        options_layout.addWidget(self._skip_contra_check)
        layout.addWidget(options_group)

        btn_row = QHBoxLayout()
        self._convert_btn = QPushButton("Convert Workbook")
        self._convert_btn.setEnabled(False)
        from kdl.config_store import get_dark_mode

        self._convert_btn.setMinimumWidth(170)
        self._convert_btn.setMinimumHeight(38)
        self._convert_btn.setStyleSheet(accent_button_qss(dark=get_dark_mode()))
        self._convert_btn.clicked.connect(self._run_conversion)
        btn_row.addStretch()
        btn_row.addWidget(self._convert_btn)
        layout.addLayout(btn_row)

        result_group = QGroupBox("Conversion Summary")
        result_layout = QVBoxLayout(result_group)
        self._result_text = QTextEdit()
        self._result_text.setReadOnly(True)
        self._result_text.setFixedHeight(136)
        self._result_text.setPlaceholderText("Conversion summary will appear here...")
        result_layout.addWidget(self._result_text)
        layout.addWidget(result_group)

        action_row = QHBoxLayout()
        self._load_grid_btn = QPushButton("Load Output to Grid")
        self._load_grid_btn.setMinimumWidth(158)
        self._load_grid_btn.setEnabled(False)
        self._load_grid_btn.setToolTip(
            "Load the converted Output sheet into the TNT DL grid and save the workbook."
        )
        self._load_grid_btn.clicked.connect(self._load_into_grid)

        self._close_btn = QPushButton("Close")
        self._close_btn.setMinimumWidth(104)
        self._close_btn.clicked.connect(self.accept)

        action_row.addWidget(self._load_grid_btn)
        action_row.addStretch()
        action_row.addWidget(self._close_btn)
        layout.addLayout(action_row)

    def _browse_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel File",
            _default_browse_dir(),
            "Excel Files (*.xlsx *.xls *.xlsm);;All Files (*)",
        )
        if not path:
            return
        self._file_edit.setText(path)
        self._start_sheet_loader(path)

    def _start_sheet_loader(self, filepath: str):
        # Cancel any in-progress loader
        if self._sheet_loader is not None and self._sheet_loader.isRunning():
            self._sheet_loader.quit()
            self._sheet_loader.wait()
        self._sheet_loader = None

        # Clear previous state
        for cb in self._sheet_checks:
            self._sheet_check_layout.removeWidget(cb)
            cb.deleteLater()
        self._sheet_checks.clear()
        self._convert_btn.setEnabled(False)
        self._result = None
        self._wb = None
        self._load_grid_btn.setEnabled(False)
        self._result_text.setPlainText("Reading sheets…")

        loader = _SheetLoaderWorker(filepath)
        loader.sheets_ready.connect(self._on_sheets_ready)
        loader.load_error.connect(self._on_sheets_error)
        loader.finished.connect(loader.deleteLater)
        self._sheet_loader = loader
        loader.start()

    def _on_sheets_ready(self, sheet_names: list):
        self._result_text.clear()
        cols = 3
        for i, name in enumerate(sheet_names):
            row, col = divmod(i, cols)
            cb = QCheckBox(name)
            is_generated = name == "Output" or name.startswith("Audit_")
            cb.setChecked(not is_generated)
            cb.stateChanged.connect(self._update_convert_btn)
            self._sheet_check_layout.addWidget(cb, row, col)
            self._sheet_checks.append(cb)
        self._update_convert_btn()
        if sum(1 for cb in self._sheet_checks if cb.isChecked()) == 1:
            QTimer.singleShot(0, self._run_conversion)

    def _on_sheets_error(self, message: str):
        self._result_text.clear()
        QMessageBox.warning(self, "File Error", f"Could not open file:\n{message}")

    def _select_all_sheets(self):
        for cb in self._sheet_checks:
            cb.setChecked(True)

    def _select_no_sheets(self):
        for cb in self._sheet_checks:
            cb.setChecked(False)

    def _update_convert_btn(self):
        has_file = bool(self._file_edit.text().strip())
        has_selection = any(cb.isChecked() for cb in self._sheet_checks)
        self._convert_btn.setEnabled(has_file and has_selection)

    def _run_conversion(self):
        filepath = self._file_edit.text().strip()
        selected = [cb.text() for cb in self._sheet_checks if cb.isChecked()]
        if not filepath or not selected:
            return

        self._convert_btn.setEnabled(False)
        self._convert_btn.setText("Converting...")
        self._result_text.setPlainText("Processing...")
        self._load_grid_btn.setEnabled(False)
        self._result = None
        self._wb = None

        self._worker = _ConverterWorker(
            filepath,
            selected,
            skip_contra=self._skip_contra_check.isChecked(),
        )
        self._worker.finished.connect(self._on_worker_finished)
        self._worker.start()

    def _on_worker_finished(self):
        worker = self._worker
        if worker is None:
            return

        self._convert_btn.setEnabled(True)
        self._convert_btn.setText("Convert Workbook")

        if worker.error_message:
            self._result = None
            self._wb = None
            self._result_text.setPlainText(f"ERROR:\n{worker.error_message}")
            QMessageBox.critical(self, "Conversion Error", worker.error_message)
            self._release_worker()
            return

        result = worker.result
        self._result = result
        if result is not None and result.success:
            self._wb = worker.wb

        self._release_worker()

        if result is None:
            self._result_text.setPlainText("ERROR:\nConversion did not return a result.")
            return

        self._result_text.setPlainText(result.message)
        if result.success:
            self._load_grid_btn.setEnabled(bool(result.output_data))
            self._result_text.append(
                '\nClick "Load Output to Grid" to load the result and save the workbook.'
            )

    def _load_into_grid(self):
        if not (self._result and self._result.output_data):
            return

        if self._wb is not None:
            filepath = self._file_edit.text().strip()
            source_ext = os.path.splitext(filepath)[1].lower()
            save_path = filepath
            if source_ext == ".xls":
                save_path = os.path.splitext(filepath)[0] + "_converted.xlsx"
            saved_name = None
            try:
                self._wb.save(save_path)
                saved_name = os.path.basename(save_path)
            except PermissionError:
                try:
                    import win32com.client

                    abs_path = os.path.abspath(save_path)
                    xl = win32com.client.GetActiveObject("Excel.Application")
                    for wb_com in list(xl.Workbooks):
                        if os.path.abspath(wb_com.FullName).lower() == abs_path.lower():
                            wb_com.Save()
                            wb_com.Close(SaveChanges=False)
                            break
                except Exception:
                    pass
                try:
                    self._wb.save(save_path)
                    saved_name = os.path.basename(save_path)
                except Exception as exc:
                    self._result_text.append(f"\nSave failed: {exc}")
            except Exception as exc:
                self._result_text.append(f"\nSave failed: {exc}")

            if saved_name:
                if save_path != filepath:
                    self._result_text.append(
                        f"Saved to: {saved_name} (legacy .xls source kept unchanged)"
                    )
                else:
                    self._result_text.append(f"Saved to: {saved_name}")

            close_wb = getattr(self._wb, "close", None)
            if callable(close_wb):
                try:
                    close_wb()
                except Exception:
                    pass

        self.load_into_grid.emit(self._result.output_data)
        self.accept()

    def closeEvent(self, event):
        if self._worker is not None and self._worker.isRunning():
            QMessageBox.warning(
                self,
                "Conversion In Progress",
                "Wait for the conversion to finish before closing this window.",
            )
            event.ignore()
            return
        self._release_worker()
        super().closeEvent(event)

    def _fit_to_screen(self):
        from PySide6.QtGui import QGuiApplication

        screen = self.screen() or QGuiApplication.primaryScreen()
        if screen:
            ag = screen.availableGeometry()
            self.resize(
                min(self.sizeHint().width() + 40, ag.width() - 80),
                min(self.sizeHint().height() + 40, ag.height() - 80),
            )
