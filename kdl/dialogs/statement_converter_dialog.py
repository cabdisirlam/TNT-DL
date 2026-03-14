"""
Bank Statement Converter Dialog.
Lets the user pick an Excel file, choose a sheet, and run the conversion.
"""

import os
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
    QListWidget, QAbstractItemView, QLineEdit, QFileDialog, QTextEdit, QMessageBox,
    QGroupBox, QSizePolicy, QCheckBox
)
from PySide6.QtCore import Qt, QThread, Signal
from kdl.styles import dialog_qss, accent_button_qss


# ── Worker thread so the UI doesn't freeze ────────────────

class _ConverterWorker(QThread):
    finished = Signal(object)   # ConversionResult
    error = Signal(str)

    def __init__(self, filepath: str, sheet_names: list, skip_contra: bool = True):
        super().__init__()
        self.filepath = filepath
        self.sheet_names = sheet_names if isinstance(sheet_names, list) else [sheet_names]
        self.skip_contra = skip_contra
        self.wb = None          # workbook kept alive so the dialog can save it

    def run(self):
        tmp_path = None
        try:
            import openpyxl, tempfile
            from kdl.engine.statement_converter import (
                convert_statement, ConversionResult, _get_or_create_sheet, OUTPUT_LAST_COL
            )

            # ── Pre-process via Excel COM ────────────────────────────────
            load_path = self.filepath
            try:
                import win32com.client
                _app = win32com.client.DispatchEx("Excel.Application")
                _app.Visible = False
                _app.DisplayAlerts = False
                _wb = _app.Workbooks.Open(
                    os.path.abspath(self.filepath),
                    UpdateLinks=0, ReadOnly=True
                )
                _app.CalculateFull()
                _fd, tmp_path = tempfile.mkstemp(suffix='.xlsx')
                os.close(_fd)
                os.unlink(tmp_path)
                _wb.SaveAs(tmp_path, 51)
                _wb.Close(False)
                _app.Quit()
                load_path = tmp_path
            except Exception:
                pass  # COM unavailable — fall back to direct openpyxl read

            # ── Convert each selected sheet ──────────────────────────────
            wb = openpyxl.load_workbook(load_path, data_only=True)
            multi = len(self.sheet_names) > 1
            all_output_data = []
            all_messages = []
            any_success = False

            for sheet_name in self.sheet_names:
                result = convert_statement(wb, sheet_name, skip_contra=self.skip_contra)
                if result.success:
                    any_success = True
                    all_output_data.extend(result.output_data)
                    prefix = f'[{sheet_name}]\n' if multi else ''
                    all_messages.append(f'{prefix}{result.message}')
                    # Rename Audit_Skipped so the next sheet doesn't overwrite it
                    if multi and 'Audit_Skipped' in wb.sheetnames:
                        safe = ('Audit_' + sheet_name)[:31]
                        # avoid duplicate sheet names
                        if safe in wb.sheetnames:
                            wb.remove(wb[safe])
                        wb['Audit_Skipped'].title = safe
                else:
                    prefix = f'[{sheet_name}] FAILED\n' if multi else 'FAILED\n'
                    all_messages.append(f'{prefix}{result.message}')

            # ── Rebuild combined Output sheet when multi-sheet ────────────
            if any_success and multi and all_output_data:
                ws_out = _get_or_create_sheet(wb, 'Output')
                for i, row_vals in enumerate(all_output_data):
                    for c, val in enumerate(row_vals, 1):
                        ws_out.cell(row=i + 2, column=c, value=val)

            sep = '\n\n' + '─' * 40 + '\n\n' if multi else ''
            combined_msg = sep.join(all_messages)

            combined = ConversionResult(
                success=any_success,
                message=combined_msg,
                output_data=all_output_data,
            )

            if any_success:
                self.wb = wb

            self.finished.emit(combined)
        except Exception as exc:
            self.error.emit(str(exc))
        finally:
            if tmp_path and os.path.exists(tmp_path):
                try:
                    os.unlink(tmp_path)
                except Exception:
                    pass


# ── Dialog ─────────────────────────────────────────────────

class StatementConverterDialog(QDialog):
    # Emitted when user wants to load Output into the TNT DL grid
    load_into_grid = Signal(list)   # list of rows (list of lists)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Bank Statement Converter')
        self.setMinimumWidth(560)
        self.setWindowFlag(Qt.WindowCloseButtonHint, True)
        self._worker = None
        self._result = None
        self._wb = None         # workbook ready to save on "Load into Grid"

        from kdl.config_store import get_dark_mode
        self.setStyleSheet(dialog_qss(dark=get_dark_mode()))

        self._build_ui()
        self._fit_to_screen()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(16, 16, 16, 16)

        # ── File picker ──
        file_group = QGroupBox('Excel File')
        file_layout = QHBoxLayout(file_group)
        self._file_edit = QLineEdit()
        self._file_edit.setPlaceholderText('Select an Excel file (.xlsx / .xls)...')
        self._file_edit.setReadOnly(True)
        browse_btn = QPushButton('Browse...')
        browse_btn.setFixedWidth(90)
        browse_btn.clicked.connect(self._browse_file)
        file_layout.addWidget(self._file_edit)
        file_layout.addWidget(browse_btn)
        layout.addWidget(file_group)

        # ── Sheet picker ──
        sheet_group = QGroupBox('Sheet / Tab  (Ctrl+click or Shift+click to select multiple)')
        sheet_outer = QVBoxLayout(sheet_group)
        sheet_row = QHBoxLayout()
        self._sheet_list = QListWidget()
        self._sheet_list.setEnabled(False)
        self._sheet_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self._sheet_list.setMaximumHeight(110)
        sheet_row.addWidget(self._sheet_list)
        sel_col = QVBoxLayout()
        sel_all_btn = QPushButton('All')
        sel_all_btn.setFixedWidth(48)
        sel_all_btn.clicked.connect(self._sheet_list.selectAll)
        sel_none_btn = QPushButton('None')
        sel_none_btn.setFixedWidth(48)
        sel_none_btn.clicked.connect(self._sheet_list.clearSelection)
        sel_col.addWidget(sel_all_btn)
        sel_col.addWidget(sel_none_btn)
        sel_col.addStretch()
        sheet_row.addLayout(sel_col)
        sheet_outer.addLayout(sheet_row)
        layout.addWidget(sheet_group)

        # ── Options ──
        options_group = QGroupBox('Options')
        options_layout = QVBoxLayout(options_group)
        self._skip_contra_check = QCheckBox('Skip contra/duplicate matched rows (CONTRA_MATCHED)')
        self._skip_contra_check.setChecked(True)
        self._skip_contra_check.setToolTip(
            'When checked, rows that match as debit/credit pairs on the same reference '
            'and amount are excluded from output. Uncheck to include all rows.'
        )
        options_layout.addWidget(self._skip_contra_check)
        layout.addWidget(options_group)

        # ── Convert button ──
        btn_row = QHBoxLayout()
        self._convert_btn = QPushButton('  Convert  ')
        self._convert_btn.setEnabled(False)
        from kdl.config_store import get_dark_mode
        self._convert_btn.setStyleSheet(accent_button_qss(dark=get_dark_mode()))
        self._convert_btn.clicked.connect(self._run_conversion)
        btn_row.addStretch()
        btn_row.addWidget(self._convert_btn)
        layout.addLayout(btn_row)

        # ── Results ──
        result_group = QGroupBox('Result Summary')
        result_layout = QVBoxLayout(result_group)
        self._result_text = QTextEdit()
        self._result_text.setReadOnly(True)
        self._result_text.setFixedHeight(160)
        self._result_text.setPlaceholderText('Conversion summary will appear here...')
        result_layout.addWidget(self._result_text)
        layout.addWidget(result_group)

        # ── Action buttons ──
        action_row = QHBoxLayout()
        self._load_grid_btn = QPushButton('Load Output into Grid')
        self._load_grid_btn.setEnabled(False)
        self._load_grid_btn.setToolTip(
            'Load the converted Output sheet into the TNT DL grid and save to Excel'
        )
        self._load_grid_btn.clicked.connect(self._load_into_grid)

        self._close_btn = QPushButton('Close')
        self._close_btn.clicked.connect(self.accept)

        action_row.addWidget(self._load_grid_btn)
        action_row.addStretch()
        action_row.addWidget(self._close_btn)
        layout.addLayout(action_row)

    # ── File browsing ──────────────────────────────────────

    def _browse_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, 'Select Excel File', '',
            'Excel Files (*.xlsx *.xls *.xlsm);;All Files (*)'
        )
        if not path:
            return
        self._file_edit.setText(path)
        self._populate_sheets(path)

    def _populate_sheets(self, filepath: str):
        self._sheet_list.clear()
        self._sheet_list.setEnabled(False)
        self._convert_btn.setEnabled(False)
        self._result = None
        self._wb = None
        self._load_grid_btn.setEnabled(False)
        try:
            import openpyxl
            wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
            for name in wb.sheetnames:
                self._sheet_list.addItem(name)
            wb.close()
            self._sheet_list.setEnabled(True)
            self._sheet_list.selectAll()
            self._convert_btn.setEnabled(True)
        except Exception as exc:
            QMessageBox.warning(self, 'File Error', f'Could not open file:\n{exc}')

    # ── Conversion ─────────────────────────────────────────

    def _run_conversion(self):
        filepath = self._file_edit.text().strip()
        selected = [item.text() for item in self._sheet_list.selectedItems()]
        if not filepath or not selected:
            return

        self._convert_btn.setEnabled(False)
        self._convert_btn.setText('Converting...')
        self._result_text.setPlainText('Processing...')
        self._load_grid_btn.setEnabled(False)
        self._result = None
        self._wb = None

        self._worker = _ConverterWorker(filepath, selected, skip_contra=self._skip_contra_check.isChecked())
        self._worker.finished.connect(self._on_finished)
        self._worker.error.connect(self._on_error)
        self._worker.start()

    def _on_finished(self, result):
        self._result = result
        self._convert_btn.setEnabled(True)
        self._convert_btn.setText('  Convert  ')

        # Keep the workbook so we can save it when the user clicks Load
        if result.success and self._worker is not None:
            self._wb = self._worker.wb

        self._result_text.setPlainText(result.message)
        if result.success:
            self._load_grid_btn.setEnabled(bool(result.output_data))
            self._result_text.append(
                '\nClick "Load Output into Grid" to load and save to Excel.'
            )

    def _on_error(self, msg: str):
        self._convert_btn.setEnabled(True)
        self._convert_btn.setText('  Convert  ')
        self._result_text.setPlainText(f'ERROR:\n{msg}')
        QMessageBox.critical(self, 'Conversion Error', msg)

    # ── Load into grid (+ save Excel) ──────────────────────

    def _load_into_grid(self):
        if not (self._result and self._result.output_data):
            return

        # Save Output + Audit_Skipped back to the Excel file
        if self._wb is not None:
            filepath = self._file_edit.text().strip()
            saved_name = None
            try:
                self._wb.save(filepath)
                saved_name = os.path.basename(filepath)
            except PermissionError:
                # File is open in Excel — save & close it via COM, then write output
                try:
                    import win32com.client
                    abs_path = os.path.abspath(filepath)
                    xl = win32com.client.GetActiveObject("Excel.Application")
                    for wb_com in list(xl.Workbooks):
                        if os.path.abspath(wb_com.FullName).lower() == abs_path.lower():
                            wb_com.Save()
                            wb_com.Close(SaveChanges=False)
                            break
                except Exception:
                    pass
                try:
                    self._wb.save(filepath)
                    saved_name = os.path.basename(filepath)
                except Exception as exc:
                    self._result_text.append(f'\nSave failed: {exc}')
            except Exception as exc:
                self._result_text.append(f'\nSave failed: {exc}')

            if saved_name:
                self._result_text.append(f'Saved to: {saved_name}')

        self.load_into_grid.emit(self._result.output_data)
        self.accept()

    # ── Fit to screen ──────────────────────────────────────

    def _fit_to_screen(self):
        from PySide6.QtGui import QGuiApplication
        screen = self.screen() or QGuiApplication.primaryScreen()
        if screen:
            ag = screen.availableGeometry()
            self.resize(min(self.sizeHint().width() + 40, ag.width() - 80),
                        min(self.sizeHint().height() + 40, ag.height() - 80))
