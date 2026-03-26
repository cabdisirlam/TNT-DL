"""
NT DL Multipurpose Tool load settings dialog.
"""

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QVBoxLayout,
)

from kdl import __display_name__
from kdl.dialogs.dialog_sizing import create_hint_button, fit_dialog_to_screen
from kdl.styles import TEXT_MUTED, accent_button_qss, dialog_qss
from kdl.window.window_manager import WindowManager


APP_TYPES = [
    "IFMIS (Oracle EBS R12+)",
    "Oracle EBS R12",
    "Oracle EBS 11i",
    "Oracle Cloud Apps",
    "Custom Application",
]

END_OF_ROW_ACTIONS = [
    ("None", "none"),
    ("Next Row", "new_record"),
    ("Next Row + Auto Save", "new_record_save_n"),
    ("Tab to Next Field", "tab"),
    ("Press Enter", "enter"),
    ("Save + Proceed", "save_proceed"),
]

LOAD_MODES = [
    ("Per Cell", "per_cell"),
    ("Per Row", "per_row"),
    ("Per Row (Fast Send)", "fast_send"),
    ("Imprest  (Alt+2 / Alt+D)", "imprest_surrender"),
]


class LoadSettingsDialog(QDialog):
    """Compact load settings dialog that fits on common laptop sizes."""

    load_requested = Signal(dict)

    def __init__(self, max_rows: int = 0, target_title: str = "",
                 target_hwnd=None, command_group: str = "", parent=None):
        super().__init__(parent)
        self.max_rows = max_rows
        self._target_title = target_title
        self._target_hwnd = target_hwnd
        self._command_group = command_group
        self._hourglass_before_load_control = True

        self.setWindowTitle(f"{__display_name__} - Load Settings")
        self.setMinimumSize(720, 470)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.setWindowFlag(Qt.WindowCloseButtonHint, True)

        from kdl.config_store import get_dark_mode

        self.setStyleSheet(dialog_qss(dark=get_dark_mode()))
        self._build_ui()
        self._refresh_windows()
        self._fit_to_screen()

        if target_title:
            idx = self.window_combo.findText(target_title)
            if idx >= 0:
                self.window_combo.setCurrentIndex(idx)

        if self._command_group:
            group = self._command_group.lower()
            if "cloud" in group:
                self.app_combo.setCurrentText("Oracle Cloud Apps")
            elif "11i" in group:
                self.app_combo.setCurrentText("Oracle EBS 11i")
            elif "r12" in group:
                self.app_combo.setCurrentText("Oracle EBS R12")

    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setSpacing(6)
        outer.setContentsMargins(8, 8, 8, 8)

        title_row = QHBoxLayout()
        title_row.setSpacing(8)

        title = QLabel("Load Settings")
        title.setStyleSheet("font-size: 16px; font-weight: 600;")
        title_row.addWidget(title)
        title_row.addWidget(
            create_hint_button(
                "1. Select the target window.\n"
                "2. Pick the load mode that matches your sheet.\n"
                "3. Choose rows to load.\n"
                "4. Adjust delays only when needed.\n"
                "5. Click Start.\n\n"
                "Load Control waits for app readiness after sends, tabs, saves, and row navigation.\n"
                "Pause on popup lets you dismiss a popup and continue manually.\n"
                "Stop on popup is best for unsupervised runs.",
                label="i",
            )
        )
        title_row.addStretch()
        outer.addLayout(title_row)

        body = QGridLayout()
        body.setHorizontalSpacing(8)
        body.setVerticalSpacing(6)
        body.setContentsMargins(0, 0, 0, 0)
        body.setColumnStretch(0, 1)
        body.setColumnStretch(1, 1)
        outer.addLayout(body)

        target_group = QGroupBox("Target Application")
        tg = QGridLayout(target_group)
        tg.setHorizontalSpacing(8)
        tg.setVerticalSpacing(6)

        tg.addWidget(QLabel("Window:"), 0, 0)
        self.window_combo = QComboBox()
        self.window_combo.setEditable(True)
        self.window_combo.setMinimumContentsLength(26)
        self.window_combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.window_combo.setMinimumHeight(28)
        tg.addWidget(self.window_combo, 0, 1)

        self.refresh_btn = QPushButton("Refresh")
        self.refresh_btn.setFixedSize(70, 28)
        self.refresh_btn.clicked.connect(self._refresh_windows)
        tg.addWidget(self.refresh_btn, 0, 2)

        tg.addWidget(QLabel("Version:"), 1, 0)
        self.app_combo = QComboBox()
        self.app_combo.addItems(APP_TYPES)
        self.app_combo.setMinimumContentsLength(22)
        self.app_combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.app_combo.setMinimumHeight(28)
        tg.addWidget(self.app_combo, 1, 1, 1, 2)
        body.addWidget(target_group, 0, 0)

        mode_group = QGroupBox("Loading Mode")
        mg = QGridLayout(mode_group)
        mg.setHorizontalSpacing(10)
        mg.setVerticalSpacing(6)

        self.radio_per_cell = QRadioButton(LOAD_MODES[0][0])
        self.radio_per_row = QRadioButton(LOAD_MODES[1][0])
        self.radio_fast_send = QRadioButton("Fast Send")
        self.radio_fast_send.setToolTip(LOAD_MODES[2][0])
        self.radio_imprest = QRadioButton("Imprest")
        self.radio_imprest.setToolTip(LOAD_MODES[3][0])
        self.radio_per_row.setChecked(True)

        mg.addWidget(self.radio_per_cell, 0, 0)
        mg.addWidget(self.radio_per_row, 0, 1)
        mg.addWidget(self.radio_fast_send, 1, 0)
        mg.addWidget(self.radio_imprest, 1, 1)

        mg.addWidget(QLabel("After each row:"), 2, 0)
        self.eor_combo = QComboBox()
        for text, _ in END_OF_ROW_ACTIONS:
            self.eor_combo.addItem(text)
        self.eor_combo.setMinimumContentsLength(18)
        self.eor_combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.eor_combo.setMinimumHeight(28)
        mg.addWidget(self.eor_combo, 2, 1)

        self._save_int_lbl = QLabel("Save every:")
        self.save_interval_input = QLineEdit("50")
        self.save_interval_input.setFixedWidth(56)
        self.save_interval_input.setFixedHeight(28)
        self.save_interval_input.setAlignment(Qt.AlignCenter)
        self._save_int_suffix = QLabel("rows")
        save_row = QHBoxLayout()
        save_row.setSpacing(6)
        save_row.addWidget(self._save_int_lbl)
        save_row.addWidget(self.save_interval_input)
        save_row.addWidget(self._save_int_suffix)
        save_row.addStretch()
        mg.addLayout(save_row, 3, 0, 1, 2)
        self._save_int_widgets = [self._save_int_lbl, self.save_interval_input, self._save_int_suffix]

        self.radio_per_cell.toggled.connect(self._update_mode_controls)
        self.radio_per_row.toggled.connect(self._update_mode_controls)
        self.radio_fast_send.toggled.connect(self._update_mode_controls)
        self.radio_imprest.toggled.connect(self._update_mode_controls)
        self.eor_combo.currentIndexChanged.connect(self._update_save_interval_visibility)
        body.addWidget(mode_group, 0, 1)

        range_group = QGroupBox("Rows to Load")
        rg = QGridLayout(range_group)
        rg.setHorizontalSpacing(8)
        rg.setVerticalSpacing(6)

        self.radio_all = QRadioButton("All Rows")
        self.radio_all.setChecked(True)
        rg.addWidget(self.radio_all, 0, 0, 1, 4)

        self.radio_selected = QRadioButton("Selected Rows / Columns")
        rg.addWidget(self.radio_selected, 1, 0, 1, 4)

        self.radio_range = QRadioButton("From Row")
        self.from_input = QLineEdit("1")
        self.from_input.setFixedWidth(64)
        self.from_input.setFixedHeight(28)
        self.from_input.setAlignment(Qt.AlignCenter)
        self.to_input = QLineEdit(str(max(1, self.max_rows)))
        self.to_input.setFixedWidth(64)
        self.to_input.setFixedHeight(28)
        self.to_input.setAlignment(Qt.AlignCenter)

        rg.addWidget(self.radio_range, 2, 0)
        rg.addWidget(self.from_input, 2, 1)
        rg.addWidget(QLabel("To Row"), 2, 2)
        rg.addWidget(self.to_input, 2, 3)
        body.addWidget(range_group, 1, 0)

        delay_group = QGroupBox("Delays && Options")
        dg = QGridLayout(delay_group)
        dg.setHorizontalSpacing(8)
        dg.setVerticalSpacing(6)

        dg.addWidget(QLabel("Cell delay:"), 0, 0)
        self.cell_delay_input = QLineEdit("0.2")
        self.cell_delay_input.setFixedWidth(60)
        self.cell_delay_input.setFixedHeight(28)
        self.cell_delay_input.setAlignment(Qt.AlignCenter)
        dg.addWidget(self.cell_delay_input, 0, 1)
        dg.addWidget(QLabel("seconds"), 0, 2)

        dg.addWidget(QLabel("Window delay:"), 1, 0)
        self.window_delay_input = QLineEdit("0.05")
        self.window_delay_input.setFixedWidth(60)
        self.window_delay_input.setFixedHeight(28)
        self.window_delay_input.setAlignment(Qt.AlignCenter)
        dg.addWidget(self.window_delay_input, 1, 1)
        dg.addWidget(QLabel("seconds"), 1, 2)

        self.hourglass_check = QCheckBox("Wait for hourglass cursor")
        self.hourglass_check.setChecked(True)
        dg.addWidget(self.hourglass_check, 2, 0, 1, 3)

        self.validate_check = QCheckBox("Validate data before run")
        self.validate_check.setChecked(True)
        dg.addWidget(self.validate_check, 3, 0, 1, 3)

        self.load_control_check = QCheckBox("Use Load Control (ready-state wait)")
        self.load_control_check.setChecked(False)
        self.load_control_check.setToolTip(
            "Monitors the target application and sends each cell only when it is ready.\n"
            "Also waits after tabs, saves, and row navigation.\n"
            "Automatically enables 'Wait if Cursor is Hour Glass'."
        )
        self.load_control_check.toggled.connect(self._sync_load_control_state)
        dg.addWidget(self.load_control_check, 4, 0, 1, 3)
        body.addWidget(delay_group, 1, 1)

        popup_group = QGroupBox("On Popup")
        pg = QVBoxLayout(popup_group)
        pg.setSpacing(6)
        pg.setContentsMargins(10, 10, 10, 10)

        self.radio_popup_pause = QRadioButton("Pause after dismissal")
        self.radio_popup_pause.setChecked(True)
        pg.addWidget(self.radio_popup_pause)

        self.radio_popup_stop = QRadioButton("Stop run on popup")
        pg.addWidget(self.radio_popup_stop)

        popup_note = QLabel("Use Stop for unsupervised or auto runs.")
        popup_note.setWordWrap(True)
        popup_note.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 12px;")
        pg.addWidget(popup_note)
        body.addWidget(popup_group, 2, 0, 1, 2)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        btn_row.addStretch()
        button_height = 40

        help_btn = QPushButton("Help")
        help_btn.setFixedSize(110, button_height)
        help_btn.clicked.connect(self._show_help)
        btn_row.addWidget(help_btn, 0, Qt.AlignVCenter)

        cancel_btn = QPushButton("Close")
        cancel_btn.setFixedSize(110, button_height)
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(cancel_btn, 0, Qt.AlignVCenter)

        start_btn = QPushButton("Start")
        start_btn.setDefault(True)
        start_btn.setFixedSize(146, button_height)
        from kdl.config_store import get_dark_mode
        start_btn.setStyleSheet(accent_button_qss(dark=get_dark_mode()))
        start_btn.clicked.connect(self._start_loading)
        btn_row.addWidget(start_btn, 0, Qt.AlignVCenter)
        outer.addLayout(btn_row)

        tip = QLabel("Tip: Press ESC twice quickly to stop loading")
        tip.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 12px;")
        outer.addWidget(tip)

        self._sync_load_control_state(self.load_control_check.isChecked())
        self._update_mode_controls()

    def _fit_to_screen(self):
        fit_dialog_to_screen(
            self,
            min_width=760,
            min_height=500,
            preferred_width=960,
            wide_width=1100,
            margin_width=64,
            margin_height=72,
            extra_hint_width=40,
            extra_hint_height=28,
        )

    def _selected_load_mode(self) -> str:
        if self.radio_imprest.isChecked():
            return "imprest_surrender"
        if self.radio_fast_send.isChecked():
            return "fast_send"
        if self.radio_per_row.isChecked():
            return "per_row"
        return "per_cell"

    def _update_mode_controls(self):
        is_form_mode = self.radio_per_row.isChecked() or self.radio_fast_send.isChecked()
        self.eor_combo.setEnabled(is_form_mode)

        if self.radio_imprest.isChecked():
            self.eor_combo.setCurrentIndex(0)
            self.cell_delay_input.setText("0.2")
        elif self.radio_fast_send.isChecked():
            self.eor_combo.setCurrentIndex(2)
            self.save_interval_input.setText("50")
            self.cell_delay_input.setText("0.05")
        elif self.radio_per_row.isChecked():
            if self.eor_combo.currentIndex() == 0:
                self.eor_combo.setCurrentIndex(2)
            self.cell_delay_input.setText("0.2")
        else:
            self.eor_combo.setCurrentIndex(0)
            self.cell_delay_input.setText("0.2")

        self._update_save_interval_visibility()

    def _update_save_interval_visibility(self):
        eor_idx = self.eor_combo.currentIndex()
        action = END_OF_ROW_ACTIONS[eor_idx][1] if 0 <= eor_idx < len(END_OF_ROW_ACTIONS) else "none"
        visible = action == "new_record_save_n" and self.eor_combo.isEnabled()
        for widget in self._save_int_widgets:
            widget.setVisible(visible)

    def _sync_load_control_state(self, checked: bool):
        checked = bool(checked)
        if checked:
            self._hourglass_before_load_control = self.hourglass_check.isChecked()
            self.hourglass_check.setChecked(True)
            self.hourglass_check.setEnabled(False)
            self.hourglass_check.setToolTip("Enabled automatically while Load Control is on.")
            return

        self.hourglass_check.setEnabled(True)
        self.hourglass_check.setChecked(bool(self._hourglass_before_load_control))
        self.hourglass_check.setToolTip("")

    @staticmethod
    def _is_oracle_like_window(title: str) -> bool:
        text = (title or "").lower()
        return any(key in text for key in ("oracle", "ifmis", "forms", "ebs", "responsibility"))

    def _refresh_windows(self):
        current_text = self.window_combo.currentText()
        self.window_combo.clear()
        try:
            windows = WindowManager.get_open_windows()
            for hwnd, title in windows:
                title_text = title or ""
                if (
                    ("NT_DL" in title_text or "NT DL" in title_text or "KDL" in title_text)
                    and ("Load Settings" in title_text or "Multipurpose Tool" in title_text)
                ):
                    continue
                self.window_combo.addItem(title_text, userData=hwnd)

            if current_text:
                idx = self.window_combo.findText(current_text)
                if idx >= 0:
                    self.window_combo.setCurrentIndex(idx)
                    return

            for idx in range(self.window_combo.count()):
                if self._is_oracle_like_window(self.window_combo.itemText(idx)):
                    self.window_combo.setCurrentIndex(idx)
                    return

            if self.window_combo.count() > 0:
                self.window_combo.setCurrentIndex(0)
        except Exception:
            pass

    def _show_help(self):
        QMessageBox.information(
            self,
            f"{__display_name__} Help",
            "LOAD SETTINGS QUICK GUIDE\n\n"
            "1. Select the target window.\n"
            "2. Pick the load mode that matches your sheet.\n"
            "3. Choose rows to load.\n"
            "4. Adjust delays only when needed.\n"
            "5. Click Start.\n\n"
            "NOTES\n"
            "  Load Control waits for application readiness after sends, tabs,\n"
            "  saves, and row navigation.\n"
            "  Pause on popup lets you dismiss a popup and continue manually.\n"
            "  Stop on popup is best for unsupervised runs.\n"
            "  Send errors still stop the run so you can verify partial data safely.\n\n"
            "STOP\n"
            "  Press ESC twice quickly or click Stop."
        )

    def _start_loading(self):
        target_hwnd = self.window_combo.currentData()
        target_title = self.window_combo.currentText()

        if not target_title:
            QMessageBox.warning(self, "No Target", "Please select a target window.")
            return

        try:
            cell_delay = float(self.cell_delay_input.text())
        except ValueError:
            cell_delay = 0.2

        try:
            window_delay = float(self.window_delay_input.text())
        except ValueError:
            window_delay = 0.05

        if self.radio_selected.isChecked():
            range_mode = "selected"
            from_row = 0
            to_row = 0
        elif self.radio_range.isChecked():
            range_mode = "range"
            try:
                from_row = max(0, int(self.from_input.text()) - 1)
                to_row = max(0, int(self.to_input.text()) - 1)
            except ValueError:
                QMessageBox.warning(self, "Invalid Range", "Please enter valid row numbers.")
                return
        else:
            range_mode = "all"
            from_row = 0
            to_row = self.max_rows - 1

        load_mode = self._selected_load_mode()
        eor_idx = self.eor_combo.currentIndex()
        end_of_row_action = END_OF_ROW_ACTIONS[eor_idx][1] if eor_idx >= 0 else "none"

        try:
            save_interval = max(1, int(self.save_interval_input.text()))
        except ValueError:
            save_interval = 50

        settings = {
            "range_mode": range_mode,
            "from_row": from_row,
            "to_row": to_row,
            "target_hwnd": target_hwnd,
            "target_title": target_title,
            "load_control": self.load_control_check.isChecked(),
            "wait_hourglass": self.hourglass_check.isChecked() or self.load_control_check.isChecked(),
            "speed_delay": cell_delay,
            "window_delay": window_delay,
            "load_mode": load_mode,
            "form_mode": load_mode in ("per_row", "fast_send", "imprest_surrender"),
            "end_of_row_action": end_of_row_action,
            "save_interval": save_interval,
            "validate_before_load": self.validate_check.isChecked(),
            "app_type": self.app_combo.currentText(),
            "popup_behavior": "stop" if self.radio_popup_stop.isChecked() else "pause",
        }

        self.load_requested.emit(settings)
        self.accept()
