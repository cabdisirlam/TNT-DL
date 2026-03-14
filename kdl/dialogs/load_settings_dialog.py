"""
NT_DL Load Settings Dialog
Professional loading dialog inspired by DataLoad Classic and FDL.
Includes Form Mode for IFMIS â€” load rows 1 to N with auto-Tab.
Default delays are tuned for IFMIS stability: 0.1s cell, 0.1s window.
"""

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QGroupBox, QRadioButton,
    QCheckBox, QLabel, QPushButton, QComboBox,
    QLineEdit, QMessageBox, QGridLayout
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont, QGuiApplication

from kdl.window.window_manager import WindowManager
from kdl.styles import dialog_qss, accent_button_qss, TEXT_MUTED


# Application types
APP_TYPES = [
    "IFMIS (Oracle EBS R12+)",
    "Oracle EBS R12",
    "Oracle EBS 11i",
    "Oracle Cloud Apps",
    "Custom Application",
]

# End-of-row actions for Form Mode
END_OF_ROW_ACTIONS = [
    ("None (Do Nothing)", "none"),
    ("Next Row (Down Arrow)", "new_record"),
    ("Next Row + Auto Save (Down Arrow + Ctrl+S every N rows)", "new_record_save_n"),
    ("Tab to Next Field", "tab"),
    ("Press Enter", "enter"),
    ("Save & Proceed (Ctrl+S then Down Arrow)", "save_proceed"),
]

LOAD_MODES = [
    ("Per Cell - one cell at a time (manual Tab/Enter in data)", "per_cell"),
    ("Per Row - auto-Tab between fields, action after each row", "per_row"),
    ("Per Row Paste - copy row and paste once (testing)", "per_row_paste"),
]


class LoadSettingsDialog(QDialog):
    """Loading settings dialog - professional DataLoad-style layout."""

    load_requested = Signal(dict)

    def __init__(self, max_rows: int = 0, target_title: str = "",
                 target_hwnd=None, command_group: str = "", parent=None):
        super().__init__(parent)
        self.max_rows = max_rows
        self._target_title = target_title
        self._target_hwnd = target_hwnd
        self._command_group = command_group

        self.setWindowTitle("NT_DL - Load Settings")
        self.setMinimumWidth(380)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.setWindowFlag(Qt.WindowCloseButtonHint, True)

        from kdl.config_store import get_dark_mode
        self.setStyleSheet(dialog_qss(dark=get_dark_mode()))
        self._build_ui()
        self._refresh_windows()
        self._fit_to_screen()

        # Pre-select target
        if target_title:
            idx = self.window_combo.findText(target_title)
            if idx >= 0:
                self.window_combo.setCurrentIndex(idx)

        # Pre-map app type from command group
        if self._command_group:
            group = self._command_group.lower()
            if "cloud" in group:
                self.app_combo.setCurrentText("Oracle Cloud Apps")
            elif "11i" in group:
                self.app_combo.setCurrentText("Oracle EBS 11i")
            elif "r12" in group:
                self.app_combo.setCurrentText("Oracle EBS R12")

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(4)
        layout.setContentsMargins(8, 8, 8, 8)

        # â”€â”€â”€ Target Window â”€â”€â”€
        target_group = QGroupBox("Target Application")
        tg = QVBoxLayout(target_group)
        tg.setSpacing(4)

        tw_row = QHBoxLayout()
        tw_label = QLabel("Window:")
        tw_row.addWidget(tw_label)

        self.window_combo = QComboBox()
        self.window_combo.setEditable(True)
        self.window_combo.setMinimumHeight(24)
        tw_row.addWidget(self.window_combo, 1)

        self.refresh_btn = QPushButton("Refresh")
        self.refresh_btn.setFixedSize(60, 24)
        self.refresh_btn.clicked.connect(self._refresh_windows)
        tw_row.addWidget(self.refresh_btn)
        tg.addLayout(tw_row)

        # Oracle version
        ver_row = QHBoxLayout()
        ver_label = QLabel("Version:")
        ver_row.addWidget(ver_label)
        self.app_combo = QComboBox()
        self.app_combo.addItems(APP_TYPES)
        self.app_combo.setMinimumHeight(24)
        ver_row.addWidget(self.app_combo, 1)
        tg.addLayout(ver_row)

        layout.addWidget(target_group)

        # â”€â”€â”€ Loading Mode â”€â”€â”€
        mode_group = QGroupBox("Loading Mode")
        mg = QVBoxLayout(mode_group)
        mg.setSpacing(3)

        self.radio_per_cell = QRadioButton(LOAD_MODES[0][0])
        self.radio_per_cell.setChecked(True)
        mg.addWidget(self.radio_per_cell)

        self.radio_per_row = QRadioButton(LOAD_MODES[1][0])
        mg.addWidget(self.radio_per_row)

        self.radio_per_row_paste = QRadioButton(LOAD_MODES[2][0])
        mg.addWidget(self.radio_per_row_paste)

        # End of row action (indent under Per Row)
        eor_row = QHBoxLayout()
        eor_row.addSpacing(22)
        eor_row.addWidget(QLabel("After each row:"))
        self.eor_combo = QComboBox()
        for text, _ in END_OF_ROW_ACTIONS:
            self.eor_combo.addItem(text)
        self.eor_combo.setMinimumHeight(24)
        eor_row.addWidget(self.eor_combo, 1)
        mg.addLayout(eor_row)

        # Save interval row (only visible when "Auto Save" action is selected)
        save_int_row = QHBoxLayout()
        save_int_row.addSpacing(22)
        self._save_int_lbl = QLabel("Save every:")
        save_int_row.addWidget(self._save_int_lbl)
        self.save_interval_input = QLineEdit("50")
        self.save_interval_input.setFixedWidth(52)
        self.save_interval_input.setFixedHeight(24)
        self.save_interval_input.setAlignment(Qt.AlignCenter)
        save_int_row.addWidget(self.save_interval_input)
        self._save_int_suffix = QLabel("rows  (also saves automatically at the last row)")
        save_int_row.addWidget(self._save_int_suffix)
        save_int_row.addStretch()
        mg.addLayout(save_int_row)
        self._save_int_widgets = [self._save_int_lbl, self.save_interval_input, self._save_int_suffix]

        self.radio_per_cell.toggled.connect(self._update_mode_controls)
        self.radio_per_row.toggled.connect(self._update_mode_controls)
        self.radio_per_row_paste.toggled.connect(self._update_mode_controls)
        self.eor_combo.currentIndexChanged.connect(self._update_save_interval_visibility)

        layout.addWidget(mode_group)

        # â”€â”€â”€ Row Range â”€â”€â”€
        range_group = QGroupBox("Rows to Load")
        rg = QVBoxLayout(range_group)
        rg.setSpacing(3)

        self.radio_all = QRadioButton("All Rows")
        self.radio_all.setChecked(True)
        rg.addWidget(self.radio_all)

        self.radio_selected = QRadioButton("Selected Rows and Columns")
        rg.addWidget(self.radio_selected)

        # Range row  â€” From [ _ ] To [ _ ]
        range_row = QHBoxLayout()
        self.radio_range = QRadioButton("From Row")

        self.from_input = QLineEdit("1")
        self.from_input.setFixedWidth(64)
        self.from_input.setFixedHeight(24)
        self.from_input.setAlignment(Qt.AlignCenter)

        to_lbl = QLabel("To Row")

        self.to_input = QLineEdit(str(max(1, self.max_rows)))
        self.to_input.setFixedWidth(64)
        self.to_input.setFixedHeight(24)
        self.to_input.setAlignment(Qt.AlignCenter)

        range_row.addWidget(self.radio_range)
        range_row.addWidget(self.from_input)
        range_row.addWidget(to_lbl)
        range_row.addWidget(self.to_input)
        range_row.addStretch()
        rg.addLayout(range_row)

        layout.addWidget(range_group)

        # â”€â”€â”€ Delays & Options â”€â”€â”€
        delay_group = QGroupBox("Delays && Options")
        dg = QGridLayout(delay_group)
        dg.setSpacing(4)

        # Delay after cell processed (recommended default: 0.1s)
        dg.addWidget(QLabel("Delay after cell processed:"), 0, 0)
        self.cell_delay_input = QLineEdit("0.1")
        self.cell_delay_input.setFixedWidth(52)
        self.cell_delay_input.setAlignment(Qt.AlignCenter)
        dg.addWidget(self.cell_delay_input, 0, 1)
        dg.addWidget(QLabel("seconds"), 0, 2)

        # Delay after window activated (recommended default: 0.05s)
        dg.addWidget(QLabel("Delay after window activated:"), 1, 0)
        self.window_delay_input = QLineEdit("0.05")
        self.window_delay_input.setFixedWidth(52)
        self.window_delay_input.setAlignment(Qt.AlignCenter)
        dg.addWidget(self.window_delay_input, 1, 1)
        dg.addWidget(QLabel("seconds"), 1, 2)

        # Options checkboxes
        self.hourglass_check = QCheckBox("Wait if Cursor is Hour Glass")
        self.hourglass_check.setChecked(True)
        dg.addWidget(self.hourglass_check, 2, 0, 1, 3)

        self.validate_check = QCheckBox("Validate data before start")
        self.validate_check.setChecked(True)
        dg.addWidget(self.validate_check, 3, 0, 1, 3)

        layout.addWidget(delay_group)

        # â”€â”€â”€ Buttons â”€â”€â”€
        btn_row = QHBoxLayout()
        btn_row.addStretch()

        help_btn = QPushButton("Help")
        help_btn.setFixedWidth(72)
        help_btn.clicked.connect(self._show_help)
        btn_row.addWidget(help_btn)

        cancel_btn = QPushButton("Close")
        cancel_btn.setFixedWidth(72)
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(cancel_btn)

        start_btn = QPushButton("  Start  ")
        start_btn.setDefault(True)
        start_btn.setFixedHeight(28)
        from kdl.config_store import get_dark_mode
        start_btn.setStyleSheet(accent_button_qss(dark=get_dark_mode()))
        start_btn.clicked.connect(self._start_loading)
        btn_row.addWidget(start_btn)

        layout.addLayout(btn_row)

        # ESC tip
        tip = QLabel("Tip: Press ESC once during loading to stop immediately")
        tip.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 12px; margin-top: 2px;")
        layout.addWidget(tip)
        self._update_mode_controls()

    def _fit_to_screen(self):
        screen = self.screen() or QGuiApplication.primaryScreen()
        if not screen:
            return
        geo = screen.availableGeometry()
        max_w = max(360, geo.width() - 24)
        max_h = max(320, geo.height() - 24)
        self.setMaximumSize(max_w, max_h)
        hint = self.sizeHint()
        target_w = min(max(self.minimumWidth(), hint.width()), max_w)
        target_h = min(max(320, hint.height()), max_h)
        self.resize(target_w, target_h)

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    # Actions
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def _selected_load_mode(self) -> str:
        if self.radio_per_row_paste.isChecked():
            return "per_row_paste"
        if self.radio_per_row.isChecked():
            return "per_row"
        return "per_cell"

    def _update_mode_controls(self):
        is_per_row = self.radio_per_row.isChecked() or self.radio_per_row_paste.isChecked()

        # End-of-row action only matters in Per Row mode
        self.eor_combo.setEnabled(is_per_row)

        if is_per_row:
            # Per Row default end-of-row = Auto Save every 50
            if self.eor_combo.currentIndex() == 0:  # "None" selected
                self.eor_combo.setCurrentIndex(2)    # auto-select "Auto Save every N"
        else:
            # Per Cell defaults: end-of-row disabled
            self.eor_combo.setCurrentIndex(0)  # reset to "None"
        self._update_save_interval_visibility()

    def _update_save_interval_visibility(self):
        eor_idx = self.eor_combo.currentIndex()
        action = END_OF_ROW_ACTIONS[eor_idx][1] if 0 <= eor_idx < len(END_OF_ROW_ACTIONS) else "none"
        visible = (action == "new_record_save_n") and self.eor_combo.isEnabled()
        for w in self._save_int_widgets:
            w.setVisible(visible)

    @staticmethod
    def _is_oracle_like_window(title: str) -> bool:
        text = (title or "").lower()
        return any(k in text for k in ("oracle", "ifmis", "forms", "ebs", "responsibility"))

    def _refresh_windows(self):
        current_text = self.window_combo.currentText()
        self.window_combo.clear()
        try:
            windows = WindowManager.get_open_windows()
            for hwnd, title in windows:
                # Filter out this app's own windows
                if ("NT_DL" in title or "KDL" in title) and ("Load Settings" in title or "Data Loader" in title):
                    continue
                self.window_combo.addItem(title, userData=hwnd)
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
            self, "NT_DL Help",
            "HOW TO LOAD DATA INTO IFMIS\n\n"
            "1. Import your data (Excel or CSV) into NT_DL\n"
            "2. Select the IFMIS Oracle window\n"
            "3. Choose Form Mode (recommended)\n"
            "4. Set row range (e.g. From 1 To 100)\n"
            "5. Click Start - NT_DL auto-Tabs between fields\n\n"
            "PROFESSIONAL OPTIONS\n"
            "  Wait Hour Glass: pauses while IFMIS is busy\n"
            "  Validate before start: catches data issues early\n\n"
            "LOAD MODES\n"
            "  Per Cell: sends one cell at a time (you control Tab/Enter)\n"
            "  Per Row: auto-Tabs between fields, runs end-of-row action\n"
            "  Per Row Paste: paste whole row once (testing mode)\n\n"
            "STOP: Press ESC once or click Stop button.\n\n"
            "Default delays in this build:\n"
            "  Cell processed: 0.1s\n"
            "  Window activated: 0.1s\n"
        )

    def _start_loading(self):
        target_hwnd = self.window_combo.currentData()
        target_title = self.window_combo.currentText()

        if not target_title:
            QMessageBox.warning(self, "No Target",
                                "Please select a target window.")
            return

        # Parse delays
        try:
            cell_delay = float(self.cell_delay_input.text())
        except ValueError:
            cell_delay = 0.1
        try:
            window_delay = float(self.window_delay_input.text())
        except ValueError:
            window_delay = 0.05

        # Range mode
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
                QMessageBox.warning(self, "Invalid Range",
                                    "Please enter valid row numbers.")
                return
        else:
            range_mode = "all"
            from_row = 0
            to_row = self.max_rows - 1

        load_mode = self._selected_load_mode()

        # End of row action
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
            "wait_hourglass": self.hourglass_check.isChecked(),
            "speed_delay": cell_delay,
            "window_delay": window_delay,
            "load_mode": load_mode,
            "form_mode": load_mode in {"per_row", "per_row_paste"},
            "end_of_row_action": end_of_row_action,
            "save_interval": save_interval,
            "validate_before_load": self.validate_check.isChecked(),
            "app_type": self.app_combo.currentText(),
        }

        self.load_requested.emit(settings)
        self.accept()
