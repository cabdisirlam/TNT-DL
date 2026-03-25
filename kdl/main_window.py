"""
NT DL Multipurpose Tool
Main window with spreadsheet, automation, and workflow utilities.
"""

import os
import re
import sys
import threading
from PySide6.QtWidgets import (
    QMainWindow, QMenuBar, QMenu, QToolBar, QStatusBar,
    QVBoxLayout, QHBoxLayout, QWidget, QFileDialog, QMessageBox,
    QLabel, QProgressBar, QComboBox, QApplication, QFrame,
    QDialog, QDialogButtonBox, QCheckBox, QLineEdit, QPushButton,
    QGridLayout, QListView, QStyle, QToolButton, QInputDialog,
    QSplitter, QTableWidgetItem, QRadioButton, QButtonGroup, QSpinBox
)
from PySide6.QtCore import Qt, QSize, QTimer, Signal
from PySide6.QtGui import QAction, QFont, QKeySequence, QColor, QIcon, QGuiApplication, QCloseEvent

from kdl.spreadsheet_widget import SpreadsheetWidget
from kdl.config_store import load_settings, save_settings
from kdl import __display_name__, __version__
from kdl.dialogs.load_settings_dialog import LoadSettingsDialog, END_OF_ROW_ACTIONS, LOAD_MODES
from kdl.dialogs.shortcuts_dialog import ShortcutsDialog
from kdl.dialogs.macro_recorder_dialog import MacroRecorderDialog
from kdl.dialogs.database_setup_dialog import DatabaseSetupDialog
from kdl.dialogs.load_result_dialog import LoadResultDialog
from kdl.dialogs.financial_report_dialog import FinancialReportDialog
from kdl.dialogs.budget_dialog import BudgetDialog
from kdl.engine.loader import LoaderThread
from kdl.engine.keystroke_parser import KeystrokeParser
from kdl.engine.validation import validate_ifmis_data
from kdl.window.window_manager import WindowManager
from kdl.styles import RED_BG, AMBER_BG


# Oracle / ERP Command Groups
COMMAND_GROUPS = [
    "Oracle Cloud Apps",
    "Oracle EBS R12 / 11i",
    "Oracle EBS 11.0 / 10.7",
    "Oracle EBS 10.7SC",
    "Siebel",
    "SAP",
    "Peoplesoft",
    "JDE",
    "Dynamics",
    "Other",
]

LOAD_DEFAULTS_VERSION = 6
VALID_LOAD_MODES = {"per_cell", "per_row", "fast_send", "imprest_surrender", "imprest_test"}
TABLE_FORMAT_HEADERS = [
    "Line",
    "Type",
    "Code",
    "Number",
    "Transaction Date",
    "Value Date",
    "Amount",
]
CELL_FORMAT_KEY_COLUMNS = {0, 1, 2, 3, 5, 7, 9, 11, 13, 14, 15}
DATE_TEXT_RE = re.compile(r"^\d{1,2}-[A-Za-z]{3}-\d{4}$")
AMOUNT_TEXT_RE = re.compile(r"^-?\d[\d,]*(?:\.\d+)?$")


class LoadProgressOverlay(QWidget):
    """Small always-on-top load counter, similar to DataLoad's floating monitor."""

    def __init__(self):
        super().__init__(None)
        self.setWindowFlags(
            Qt.Tool | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint
        )
        self.setAttribute(Qt.WA_ShowWithoutActivating, True)
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.setFixedSize(450, 126)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(8)

        self._mode_label = "Per Cell"
        self.title_label = QLabel("")
        self.title_label.setStyleSheet(
            "font-weight: 700; font-size: 13px; color: #4E6E8F; letter-spacing: 0.4px;"
        )
        layout.addWidget(self.title_label)
        self._refresh_title()

        cols_row = QHBoxLayout()
        cols_row.setSpacing(8)
        cols_lbl = QLabel("Cols")
        cols_lbl.setFixedWidth(30)
        self.cols_bar = QProgressBar()
        self.cols_bar.setTextVisible(False)
        self.cols_bar.setFixedHeight(16)
        self.cols_count = QLabel("0/0")
        self.cols_count.setFixedWidth(68)
        self.cols_count.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        cols_row.addWidget(cols_lbl)
        cols_row.addWidget(self.cols_bar, 1)
        cols_row.addWidget(self.cols_count)
        layout.addLayout(cols_row)

        rows_row = QHBoxLayout()
        rows_row.setSpacing(8)
        rows_lbl = QLabel("Rows")
        rows_lbl.setFixedWidth(30)
        self.rows_bar = QProgressBar()
        self.rows_bar.setTextVisible(False)
        self.rows_bar.setFixedHeight(16)
        self.rows_count = QLabel("0/0")
        self.rows_count.setFixedWidth(130)
        self.rows_count.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        rows_row.addWidget(rows_lbl)
        rows_row.addWidget(self.rows_bar, 1)
        rows_row.addWidget(self.rows_count)
        layout.addLayout(rows_row)

        cell_row = QHBoxLayout()
        cell_row.setSpacing(8)
        cell_lbl = QLabel("Cell")
        cell_lbl.setFixedWidth(30)
        self.cell_value = QLabel("")
        self.cell_value.setStyleSheet(
            "font-family: Consolas; font-size: 13px; color: #18324A; font-weight: 600;"
        )
        self.cell_value.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        cell_row.addWidget(cell_lbl)
        cell_row.addWidget(self.cell_value, 1)
        layout.addLayout(cell_row)

        self.setStyleSheet(
            """
            QWidget {
                background: qlineargradient(
                    x1:0, y1:0, x2:0, y2:1,
                    stop:0 #FFFFFF,
                    stop:1 #F3F8FC
                );
                border: 1px solid #BFD1E1;
                border-radius: 14px;
                color: #18324A;
                font-size: 13px;
            }
            QProgressBar {
                border: 1px solid #BDD0E0;
                border-radius: 6px;
                background-color: #EAF1F8;
            }
            QProgressBar::chunk {
                background-color: #1F7AB8;
                border-radius: 5px;
            }
            """
        )
        self._rows_current = 0
        self._rows_total = 0
        self._rows_marker = ""

    def _refresh_title(self):
        mode = (self._mode_label or "").strip() or "Per Cell"
        self.title_label.setText(f"NT DL  |  2x Esc to stop  |  Mode: {mode}")

    def _position_top_right(self):
        screen = QGuiApplication.primaryScreen()
        if not screen:
            return
        geo = screen.availableGeometry()
        self.move(geo.x() + geo.width() - self.width() - 16, geo.y() + 16)

    @staticmethod
    def _set_progress(bar: QProgressBar, count_lbl: QLabel, current: int, total: int):
        total = max(0, int(total))
        current = max(0, int(current))
        shown = min(current, total) if total > 0 else 0
        if total > 0:
            bar.setRange(0, total)
            bar.setValue(shown)
        else:
            bar.setRange(0, 1)
            bar.setValue(0)
        count_lbl.setText(f"{shown}/{total}")

    def _refresh_rows(self):
        total = max(0, int(self._rows_total))
        current = max(0, int(self._rows_current))
        shown = min(current, total) if total > 0 else 0
        if total > 0:
            self.rows_bar.setRange(0, total)
            self.rows_bar.setValue(shown)
        else:
            self.rows_bar.setRange(0, 1)
            self.rows_bar.setValue(0)
        text = f"{shown}/{total}"
        marker = (self._rows_marker or "").strip()
        if marker:
            text += f" ({marker})"
        self.rows_count.setText(text)

    def show_overlay(self, total_rows: int, total_cols: int):
        self._rows_current = 0
        self._rows_total = max(0, int(total_rows))
        self._rows_marker = ""
        self._refresh_title()
        self._refresh_rows()
        self._set_progress(self.cols_bar, self.cols_count, 0, total_cols)
        self.set_cell_text("")
        self._position_top_right()
        self.show()
        self.raise_()

    def set_mode_label(self, mode_label: str):
        self._mode_label = mode_label or "Per Cell"
        self._refresh_title()

    def update_rows(self, current: int, total: int):
        self._rows_current = max(0, int(current))
        self._rows_total = max(0, int(total))
        self._refresh_rows()

    def update_cols(self, current: int, total: int):
        self._set_progress(self.cols_bar, self.cols_count, current, total)

    def set_rows_marker(self, marker: str):
        self._rows_marker = marker or ""
        self._refresh_rows()

    def set_cell_text(self, text: str):
        clean = (text or "").replace("\n", " ").strip()
        if len(clean) > 70:
            clean = clean[:67] + "..."
        self.cell_value.setText(clean)


class MainWindow(QMainWindow):
    """NT_DL Main Application Window - FDL Style."""
    window_list_ready = Signal(list)

    def __init__(self):
        super().__init__()
        self._apply_window_title()
        self._responsive_show_applied = False
        self._compact_layout = False
        self._base_min_size = QSize(1100, 700)
        self._base_default_size = QSize(1280, 800)

        # Core objects
        self.parser = KeystrokeParser()
        self.loader_thread = None
        self.current_file = None
        self._is_paused = False
        self._has_header_row = False  # Track if row 1 is headers (templates)
        self._load_error_count = 0
        self._protect_load_enabled = False
        self._show_progress_bar = True
        self._compact_mode_enabled = False
        self._find_query = ""
        self._find_scope = "all"

        # Global load defaults (used to prefill Start Load popup)
        self._default_speed_delay = 0.1
        self._default_window_delay = 0.05
        self._default_wait_hourglass = True
        self._default_load_control = False
        self._default_form_mode = False
        self._default_load_mode = "per_cell"
        self._default_end_of_row_action = "none"
        self._default_validate_before_load = True
        self._default_popup_behavior = "pause"
        self._db_settings = {"mode": "ui_automation", "active_profile": "", "profiles": []}
        self._overlay_total_rows = 0
        self._overlay_total_cols = 0
        self._overlay_columns = []
        self._overlay_column_lookup = {}
        self._overlay_start_row = 0
        self._overlay_active_row_abs = 0
        self._load_visual_compact = False
        self._load_visual_cell_stride = 1
        self._load_visual_row_stride = 1
        self._load_visual_row_scroll_stride = 25
        self._load_visual_cells_seen = 0
        self._pending_load_result = None
        self._pending_load_thread = None
        self._deferred_load_result = None
        self._load_overlay = LoadProgressOverlay()
        self._window_refresh_inflight = False
        self._last_load_settings: dict = {}
        self._last_started_row: int = -1
        self._load_row_results: dict = {}
        self._freeze_header_enabled: bool = False
        self._apply_saved_settings()

        # Build UI
        self._build_menu_bar()
        self._build_toolbar()
        self._build_central()
        self._build_status_bar()
        self._apply_responsive_window_settings()

        # Apply styling
        self._apply_styles()

        # Cross-thread marshal for window enumeration results.
        self.window_list_ready.connect(self._apply_window_list)

        # Auto-refresh window list periodically
        self._refresh_timer = QTimer(self)
        self._refresh_timer.timeout.connect(self._refresh_windows)
        self._refresh_timer.start(10000)  # Every 10 seconds
        self._progress_hide_timer = QTimer(self)
        self._progress_hide_timer.setSingleShot(True)
        self._progress_hide_timer.timeout.connect(lambda: self.progress_bar.setVisible(False))

    def showEvent(self, event):
        super().showEvent(event)
        if not self._responsive_show_applied:
            self._apply_responsive_window_settings()
            self._responsive_show_applied = True
        self._align_top_rows_with_toolbar_anchor()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._align_top_rows_with_toolbar_anchor()

    def _apply_responsive_window_settings(self):
        """Fit the main window to the current screen and apply compact tweaks on small displays."""
        screen = None
        handle = self.windowHandle()
        if handle is not None:
            screen = handle.screen()
        if screen is None:
            screen = QGuiApplication.primaryScreen()
        if screen is None:
            self.setMinimumSize(self._base_min_size)
            self.resize(self._base_default_size)
            return

        available = screen.availableGeometry()
        compact_auto = available.width() <= 1440 or available.height() <= 820
        compact = bool(self._compact_mode_enabled or compact_auto)

        min_w = 960 if compact else self._base_min_size.width()
        min_h = 620 if compact else self._base_min_size.height()
        # Keep a practical lower bound, but never exceed the usable screen area.
        min_w = min(min_w, max(560, available.width() - 30))
        min_h = min(min_h, max(420, available.height() - 30))

        self.setMinimumSize(min_w, min_h)
        self._apply_compact_layout(compact)

        target_w = min(self._base_default_size.width(), available.width() - 24)
        target_h = min(self._base_default_size.height(), available.height() - 24)
        target_w = max(min_w, target_w)
        target_h = max(min_h, target_h)
        self.resize(target_w, target_h)

    def _apply_compact_layout(self, compact: bool):
        if compact == self._compact_layout:
            return

        self._compact_layout = compact

        if hasattr(self, "main_toolbar"):
            self.main_toolbar.setIconSize(QSize(20, 20) if compact else QSize(24, 24))
            self._normalize_toolbar_button_geometry()
        if hasattr(self, "window_combo"):
            self.window_combo.setMinimumWidth(280 if compact else 350)
        if hasattr(self, "command_group_combo"):
            self.command_group_combo.setMinimumWidth(160 if compact else 180)
        if hasattr(self, "refresh_btn"):
            self.refresh_btn.setFixedHeight(30 if compact else 34)
        if hasattr(self, "progress_bar"):
            self.progress_bar.setFixedWidth(180 if compact else 250)
        self._align_top_rows_with_toolbar_anchor()

    def _align_top_rows_with_toolbar_anchor(self):
        """
        Keep Window/Command/Notes block aligned with the right edge of the last toolbar icon.
        """
        if not hasattr(self, "top_rows_container") or not hasattr(self, "main_toolbar"):
            return
        if self.main_toolbar is None:
            return

        anchor_action = getattr(self, "convert_macros_btn", None)
        if anchor_action is None:
            actions = self.main_toolbar.actions()
            anchor_action = actions[-1] if actions else None
        if anchor_action is None:
            return

        btn = self.main_toolbar.widgetForAction(anchor_action)
        if btn is None or not isinstance(btn, QToolButton):
            return

        anchor_right = btn.mapTo(self, btn.rect().topRight()).x()
        block_left = self.top_rows_container.mapTo(self, self.top_rows_container.rect().topLeft()).x()
        target_width = max(920, anchor_right - block_left + 12)
        self.top_rows_container.setMaximumWidth(target_width)

    def _apply_saved_settings(self):
        """Load persisted app settings into runtime defaults."""
        settings = load_settings()
        load_defaults = settings.get("load_defaults", {})
        db = settings.get("database", {})
        ui = settings.get("ui", {})
        migrated_defaults = False

        try:
            defaults_version = int(load_defaults.get("defaults_version", 0))
        except Exception:
            defaults_version = 0

        if defaults_version < LOAD_DEFAULTS_VERSION:
            # Migrate old installs to the new IFMIS-friendly defaults once.
            self._default_speed_delay = 0.1
            self._default_window_delay = 0.05
            self._default_wait_hourglass = True
            migrated_defaults = True
        else:
            try:
                self._default_speed_delay = float(
                    load_defaults.get("speed_delay", self._default_speed_delay)
                )
            except Exception:
                pass
            try:
                self._default_window_delay = float(
                    load_defaults.get("window_delay", self._default_window_delay)
                )
            except Exception:
                pass

        self._default_wait_hourglass = bool(
            load_defaults.get("wait_hourglass", self._default_wait_hourglass)
        )
        self._default_load_control = bool(
            load_defaults.get("load_control", self._default_load_control)
        )
        self._default_form_mode = bool(load_defaults.get("form_mode", self._default_form_mode))
        saved_mode = str(load_defaults.get("load_mode", "")).strip().lower()
        if saved_mode in VALID_LOAD_MODES:
            self._default_load_mode = saved_mode
        else:
            self._default_load_mode = "per_cell"
        self._default_form_mode = self._default_load_mode in ("per_row", "fast_send", "imprest_surrender", "imprest_test")

        self._default_validate_before_load = bool(
            load_defaults.get("validate_before_load", self._default_validate_before_load)
        )
        self._default_end_of_row_action = str(
            load_defaults.get("end_of_row_action", self._default_end_of_row_action)
        )
        popup_behavior = str(
            load_defaults.get("popup_behavior", self._default_popup_behavior)
        ).strip().lower()
        self._default_popup_behavior = popup_behavior if popup_behavior in {"pause", "stop"} else "pause"
        if migrated_defaults:
            self._default_load_mode = "per_cell"
            self._default_form_mode = False
            self._default_end_of_row_action = "none"

        self._protect_load_enabled = bool(ui.get("protect_load_enabled", self._protect_load_enabled))
        self._show_progress_bar = bool(ui.get("show_progress_bar", self._show_progress_bar))
        self._compact_mode_enabled = bool(ui.get("compact_mode_enabled", self._compact_mode_enabled))

        if isinstance(db, dict):
            profiles = db.get("profiles", [])
            clean_profiles = []
            if isinstance(profiles, list):
                for profile in profiles:
                    if not isinstance(profile, dict):
                        continue
                    clean = dict(profile)
                    clean_profiles.append(clean)
            self._db_settings = {
                "mode": "ui_automation",
                "active_profile": db.get("active_profile", ""),
                "profiles": clean_profiles,
            }
        if migrated_defaults:
            self._persist_settings()

    def _persist_settings(self):
        """Persist runtime defaults and toggles."""
        payload = {
            "load_defaults": {
                "defaults_version": LOAD_DEFAULTS_VERSION,
                "speed_delay": self._default_speed_delay,
                "window_delay": self._default_window_delay,
                "wait_hourglass": self._default_wait_hourglass,
                "load_control": self._default_load_control,
                "form_mode": self._default_form_mode,
                "load_mode": self._default_load_mode,
                "validate_before_load": self._default_validate_before_load,
                "end_of_row_action": self._default_end_of_row_action,
                "popup_behavior": self._default_popup_behavior,
            },
            "ui": {
                "protect_load_enabled": self._protect_load_enabled,
                "show_progress_bar": self._show_progress_bar,
                "compact_mode_enabled": self._compact_mode_enabled,
            },
            "database": self._db_settings,
        }
        save_settings(payload)

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    # Menu Bar
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def _build_menu_bar(self):
        menubar = self.menuBar()

        # â”€â”€ File Menu â”€â”€
        file_menu = menubar.addMenu("&File")

        new_action = QAction("&New", self)
        new_action.setShortcut(QKeySequence.New)
        new_action.triggered.connect(self._new_file)
        file_menu.addAction(new_action)

        open_action = QAction("&Open...", self)
        open_action.setShortcut(QKeySequence.Open)
        open_action.triggered.connect(self._open_file)
        file_menu.addAction(open_action)

        save_action = QAction("&Save", self)
        save_action.setShortcut(QKeySequence.Save)
        save_action.triggered.connect(self._save_file)
        file_menu.addAction(save_action)

        save_as_action = QAction("Save &As...", self)
        save_as_action.setShortcut(QKeySequence("Ctrl+Shift+S"))
        save_as_action.triggered.connect(self._save_file_as)
        file_menu.addAction(save_as_action)

        file_menu.addSeparator()

        import_csv_action = QAction("Import &CSV...", self)
        import_csv_action.triggered.connect(self._import_csv)
        file_menu.addAction(import_csv_action)

        import_excel_action = QAction("Import &Excel...", self)
        import_excel_action.triggered.connect(self._import_excel)
        file_menu.addAction(import_excel_action)

        file_menu.addSeparator()

        exit_action = QAction("E&xit", self)
        exit_action.setShortcut(QKeySequence("Alt+F4"))
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # â”€â”€ Edit Menu â”€â”€
        edit_menu = menubar.addMenu("&Edit")

        undo_action = QAction("&Undo", self)
        undo_action.setShortcut(QKeySequence.Undo)
        undo_action.triggered.connect(self._undo_edit)
        edit_menu.addAction(undo_action)

        redo_action = QAction("&Redo", self)
        redo_action.setShortcut(QKeySequence.Redo)
        redo_action.triggered.connect(self._redo_edit)
        edit_menu.addAction(redo_action)

        edit_menu.addSeparator()

        back_action = QAction("&Back", self)
        back_action.setShortcut(QKeySequence("Alt+Left"))
        back_action.triggered.connect(self._go_back_position)
        edit_menu.addAction(back_action)

        forward_action = QAction("&Forward", self)
        forward_action.setShortcut(QKeySequence("Alt+Right"))
        forward_action.triggered.connect(self._go_forward_position)
        edit_menu.addAction(forward_action)

        edit_menu.addSeparator()

        cut_action = QAction("Cu&t", self)
        cut_action.setShortcut(QKeySequence.Cut)
        cut_action.triggered.connect(lambda: self.spreadsheet._cut_cells())
        edit_menu.addAction(cut_action)

        copy_action = QAction("&Copy", self)
        copy_action.setShortcut(QKeySequence.Copy)
        copy_action.triggered.connect(lambda: self.spreadsheet._copy_cells())
        edit_menu.addAction(copy_action)

        paste_action = QAction("&Paste", self)
        paste_action.setShortcut(QKeySequence.Paste)
        paste_action.triggered.connect(lambda: self.spreadsheet._paste_cells())
        edit_menu.addAction(paste_action)

        edit_menu.addSeparator()

        find_action = QAction("&Find...", self)
        find_action.setShortcut(QKeySequence.Find)
        find_action.triggered.connect(self._find_in_sheet)
        edit_menu.addAction(find_action)

        find_next_action = QAction("Find &Next", self)
        find_next_action.setShortcut(QKeySequence.FindNext)
        find_next_action.triggered.connect(self._find_next_in_sheet)
        edit_menu.addAction(find_next_action)

        replace_action = QAction("&Replace...", self)
        replace_action.setShortcut(QKeySequence("Ctrl+H"))
        replace_action.triggered.connect(self._find_replace_in_sheet)
        edit_menu.addAction(replace_action)

        fill_down_action = QAction("Fill &Down", self)
        fill_down_action.setShortcut(QKeySequence("Ctrl+D"))
        fill_down_action.triggered.connect(lambda: self.spreadsheet.fill_down())
        edit_menu.addAction(fill_down_action)

        edit_menu.addSeparator()

        clear_action = QAction("Clear &All", self)
        clear_action.triggered.connect(self._clear_all)
        edit_menu.addAction(clear_action)

        # â”€â”€ Tools Menu â”€â”€
        tools_menu = menubar.addMenu("&Tools")

        stmt_conv_action = QAction("Bank &Statement Converter...", self)
        stmt_conv_action.triggered.connect(self._open_statement_converter)
        tools_menu.addAction(stmt_conv_action)

        report_action = QAction("&Generate IFMIS Financial Statements...", self)
        report_action.triggered.connect(self._open_financial_report)
        tools_menu.addAction(report_action)

        budget_action = QAction("&Budget...", self)
        budget_action.triggered.connect(self._open_budget)
        tools_menu.addAction(budget_action)

        imprest_action = QAction("&Imprest Surrender AP Loader...", self)
        imprest_action.triggered.connect(self._open_imprest_surrender)
        tools_menu.addAction(imprest_action)

        tools_menu.addSeparator()

        start_load_action = QAction("&Start Load", self)
        start_load_action.setShortcut(QKeySequence("F5"))
        start_load_action.triggered.connect(self._start_loading)
        tools_menu.addAction(start_load_action)

        delays_action = QAction("&Delays && Timeouts", self)
        delays_action.triggered.connect(self._open_delays_timeouts)
        tools_menu.addAction(delays_action)

        options_action = QAction("&Options", self)
        options_action.triggered.connect(self._open_options)
        tools_menu.addAction(options_action)

        self.protect_load_action = QAction("&Protect Load", self)
        self.protect_load_action.setCheckable(True)
        self.protect_load_action.setChecked(self._protect_load_enabled)
        self.protect_load_action.toggled.connect(self._toggle_protect_load)
        tools_menu.addAction(self.protect_load_action)

        self.progress_bar_action = QAction("&Progress Bar", self)
        self.progress_bar_action.setCheckable(True)
        self.progress_bar_action.setChecked(self._show_progress_bar)
        self.progress_bar_action.toggled.connect(self._toggle_progress_bar_setting)
        tools_menu.addAction(self.progress_bar_action)

        excel_import_action = QAction("&Excel Import", self)
        excel_import_action.triggered.connect(self._import_excel)
        tools_menu.addAction(excel_import_action)

        data_menu = tools_menu.addMenu("&Data")
        data_menu.addAction(QAction("Import &CSV...", self, triggered=self._import_csv))
        data_menu.addAction(QAction("Import &Excel...", self, triggered=self._import_excel))
        data_menu.addSeparator()
        data_menu.addAction(QAction("&Save", self, triggered=self._save_file))
        data_menu.addAction(QAction("Save &As...", self, triggered=self._save_file_as))
        data_menu.addSeparator()
        data_menu.addAction(QAction("&Clear All", self, triggered=self._clear_all))

        validate_action = QAction("&Validate Data", self)
        validate_action.setShortcut(QKeySequence("Ctrl+Shift+V"))
        validate_action.triggered.connect(self._validate_current_data)
        tools_menu.addAction(validate_action)

        convert_table_action = QAction("Convert to &Table Format", self)
        convert_table_action.setShortcut(QKeySequence("Ctrl+Shift+T"))
        convert_table_action.triggered.connect(self._convert_to_table_format)
        tools_menu.addAction(convert_table_action)

        convert_cell_action = QAction("Convert to &Cell Format", self)
        convert_cell_action.setShortcut(QKeySequence("Ctrl+Shift+M"))
        convert_cell_action.triggered.connect(self._convert_to_cell_format)
        tools_menu.addAction(convert_cell_action)

        browser_menu = tools_menu.addMenu("&Browser Control")
        browser_menu.addAction(
            QAction("Activate &Selected Window", self, triggered=self._activate_selected_target)
        )
        browser_menu.addAction(
            QAction("Show &Oracle/IFMIS Windows", self, triggered=self._show_oracle_windows)
        )
        browser_menu.addAction(
            QAction("Show &Foreground Window", self, triggered=self._show_foreground_window)
        )

        macro_action = QAction("&Macro Recorder", self)
        macro_action.triggered.connect(self._open_macro_recorder)
        tools_menu.addAction(macro_action)

        setup_db_action = QAction("Setup &Databases", self)
        setup_db_action.triggered.connect(self._setup_databases)
        tools_menu.addAction(setup_db_action)

        tools_menu.addSeparator()

        shortcuts_action = QAction("Edit &Shortcuts...", self)
        shortcuts_action.triggered.connect(self._edit_shortcuts)
        tools_menu.addAction(shortcuts_action)

        refresh_win_action = QAction("&Refresh Windows", self)
        refresh_win_action.triggered.connect(self._refresh_windows)
        tools_menu.addAction(refresh_win_action)
        # â”€â”€ Key Columns Menu â”€â”€
        key_menu = menubar.addMenu("&Key Columns")
        mark_action = QAction("&Mark/Unmark Current Column as Key", self)
        mark_action.triggered.connect(self._toggle_key_column)
        key_menu.addAction(mark_action)

        clear_keys_action = QAction("&Clear All Key Columns", self)
        clear_keys_action.triggered.connect(self._clear_key_columns)
        key_menu.addAction(clear_keys_action)

        # -- View Menu --
        view_menu = menubar.addMenu('&View')
        self.dark_mode_action = QAction('&Dark Mode', self)
        self.dark_mode_action.setCheckable(True)
        from kdl.config_store import get_dark_mode
        self.dark_mode_action.setChecked(get_dark_mode())
        self.dark_mode_action.triggered.connect(self._toggle_dark_mode)
        view_menu.addAction(self.dark_mode_action)

        self.freeze_header_action = QAction('&Freeze Header Row (Row 1)', self)
        self.freeze_header_action.setCheckable(True)
        self.freeze_header_action.setChecked(False)
        self.freeze_header_action.triggered.connect(self._toggle_freeze_header)
        view_menu.addAction(self.freeze_header_action)

        # â”€â”€ Help Menu â”€â”€
        help_menu = menubar.addMenu('&Help')

        about_action = QAction(f"&About {__display_name__}", self)
        about_action.triggered.connect(self._show_about)
        help_menu.addAction(about_action)

        how_to_action = QAction("&How to Use", self)
        how_to_action.triggered.connect(self._show_how_to)
        help_menu.addAction(how_to_action)

        keystrokes_action = QAction("&Keystroke Reference", self)
        keystrokes_action.triggered.connect(self._show_keystrokes)
        help_menu.addAction(keystrokes_action)

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    # Toolbar
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def _icon(self, name: str) -> QIcon:
        """Load an SVG icon from kdl/assets, handling PyInstaller paths."""
        base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        path = os.path.join(base, "kdl", "assets", name)
        if os.path.exists(path):
            return QIcon(path)
        return QIcon()

    def _build_toolbar(self):
        toolbar = QToolBar("Main Toolbar")
        toolbar.setObjectName("MainToolbar")
        toolbar.setMovable(False)
        toolbar.setIconSize(QSize(24, 24))
        toolbar.setToolButtonStyle(Qt.ToolButtonIconOnly)
        self.addToolBar(toolbar)
        self.main_toolbar = toolbar

        # File actions
        new_act = QAction(self._icon("ic_new.svg"), "New", self, triggered=self._new_file)
        new_act.setToolTip("New  (Ctrl+N)")
        toolbar.addAction(new_act)

        open_act = QAction(self._icon("ic_open.svg"), "Open", self, triggered=self._open_file)
        open_act.setToolTip("Open  (Ctrl+O)")
        toolbar.addAction(open_act)

        save_act = QAction(self._icon("ic_save.svg"), "Save", self, triggered=self._save_file)
        save_act.setToolTip("Save  (Ctrl+S)")
        toolbar.addAction(save_act)

        toolbar.addSeparator()

        import_act = QAction(self._icon("ic_import.svg"), "Import Excel", self, triggered=self._import_excel)
        import_act.setToolTip("Import Excel")
        toolbar.addAction(import_act)

        shortcuts_act = QAction(self._icon("ic_shortcuts.svg"), "Edit Shortcuts", self, triggered=self._edit_shortcuts)
        shortcuts_act.setToolTip("Edit Shortcuts")
        toolbar.addAction(shortcuts_act)

        toolbar.addSeparator()

        # Loading controls
        self.start_btn = QAction(self._icon("ic_start.svg"), "Start", self, triggered=self._start_loading)
        self.start_btn.setToolTip("Start Load  (F5)")
        toolbar.addAction(self.start_btn)

        self.stop_btn = QAction(self._icon("ic_stop.svg"), "Stop", self, triggered=self._stop_loading)
        self.stop_btn.setToolTip("Stop Load")
        self.stop_btn.setEnabled(False)

        self.pause_btn = QAction(self._icon("ic_pause.svg"), "Pause", self, triggered=self._pause_resume)
        self.pause_btn.setShortcut(QKeySequence("F6"))
        self.pause_btn.setToolTip("Pause / Resume  (F6)")
        self.pause_btn.setEnabled(False)

        self.step_btn = QAction(self._icon("ic_step.svg"), "Step", self, triggered=self._next_step)
        self.step_btn.setToolTip("Step Forward")
        self.step_btn.setEnabled(False)

        self.statement_btn = QAction(
            self._icon("ic_statement.svg"),
            "Bank Statement",
            self,
            triggered=self._open_statement_converter,
        )
        self.statement_btn.setToolTip("Bank Statement Converter")
        toolbar.addAction(self.statement_btn)

        self.report_btn = QAction(
            self._icon("ic_report.svg"),
            "IFMIS Report",
            self,
            triggered=self._open_financial_report,
        )
        self.report_btn.setToolTip("Generate IFMIS Financial Statements from Notes")
        toolbar.addAction(self.report_btn)

        self.budget_btn = QAction(
            self._icon("ic_budget.svg"),
            "Budget",
            self,
            triggered=self._open_budget,
        )
        self.budget_btn.setToolTip("GOK IFMIS Budget Processor")
        toolbar.addAction(self.budget_btn)

        self.imprest_btn = QAction(
            self._icon("ic_imprest.svg"),
            "Imprest Surrender",
            self,
            triggered=self._open_imprest_surrender,
        )
        self.imprest_btn.setToolTip("Imprest Surrender AP Invoice Loader")
        toolbar.addAction(self.imprest_btn)

        toolbar.addSeparator()

        # Edit controls
        self.tb_undo_btn = QAction(self._icon("ic_undo.svg"), "Undo", self, triggered=self._undo_edit)
        self.tb_undo_btn.setToolTip("Undo  (Ctrl+Z)")
        toolbar.addAction(self.tb_undo_btn)

        self.tb_redo_btn = QAction(self._icon("ic_redo.svg"), "Redo", self, triggered=self._redo_edit)
        self.tb_redo_btn.setToolTip("Redo  (Ctrl+Y)")
        toolbar.addAction(self.tb_redo_btn)

        self.tb_back_btn = QAction(self._icon("ic_back.svg"), "Back", self, triggered=self._go_back_position)
        self.tb_back_btn.setToolTip("Back  (Alt+Left)")
        toolbar.addAction(self.tb_back_btn)

        self.tb_forward_btn = QAction(self._icon("ic_forward.svg"), "Forward", self, triggered=self._go_forward_position)
        self.tb_forward_btn.setToolTip("Forward  (Alt+Right)")
        toolbar.addAction(self.tb_forward_btn)

        self.tb_clear_btn = QAction(self._icon("ic_clear.svg"), "Clear All", self, triggered=self._clear_all)
        self.tb_clear_btn.setToolTip("Clear All  (clear entire grid)")
        toolbar.addAction(self.tb_clear_btn)

        toolbar.addSeparator()

        # REC button
        self.rec_btn = QAction(self._icon("ic_rec.svg"), "REC", self, triggered=self._open_macro_recorder)
        self.rec_btn.setToolTip("Macro Recorder")
        toolbar.addAction(self.rec_btn)

        self.convert_table_btn = QAction(
            self._icon("ic_table_format.svg"),
            "To Table",
            self,
            triggered=self._convert_to_table_format,
        )
        self.convert_table_btn.setToolTip("Convert to Table Format  (Ctrl+Shift+T)")
        toolbar.addAction(self.convert_table_btn)

        self.convert_cell_btn = QAction(
            self._icon("ic_cell_format.svg"),
            "To Cell",
            self,
            triggered=self._convert_to_cell_format,
        )
        self.convert_cell_btn.setToolTip("Convert to Cell Format  (Ctrl+Shift+M)")
        toolbar.addAction(self.convert_cell_btn)
        # Backward-compatible anchor name used by layout alignment.
        self.convert_macros_btn = self.convert_cell_btn
        self._normalize_toolbar_button_geometry()
        self._apply_toolbar_action_colors()

    def _normalize_toolbar_button_geometry(self):
        """Force uniform button and icon size for every toolbar action."""
        if not self.main_toolbar:
            return
        icon_size = self.main_toolbar.iconSize()
        compact = icon_size.width() <= 20
        for action in self.main_toolbar.actions():
            btn = self.main_toolbar.widgetForAction(action)
            if btn is None or not isinstance(btn, QToolButton):
                continue
            btn.setAutoRaise(False)
            btn.setCursor(Qt.PointingHandCursor)
            if compact:
                btn.setFixedSize(44, 40)
            else:
                btn.setFixedSize(50, 46)
            btn.setIconSize(icon_size)

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    # Central Widget
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def _build_central(self):
        central = QWidget()
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(12, 8, 12, 0)
        main_layout.setSpacing(6)

        self.top_rows_container = QFrame()
        self.top_rows_container.setObjectName("TopRowsCard")
        top_rows_layout = QVBoxLayout(self.top_rows_container)
        top_rows_layout.setContentsMargins(14, 12, 14, 10)
        top_rows_layout.setSpacing(0)

        shell_row = QHBoxLayout()
        shell_row.setSpacing(12)

        win_label = QLabel("Window")
        win_label.setObjectName("ShellFieldLabel")
        win_label.setMinimumWidth(56)
        shell_row.addWidget(win_label)

        self.window_combo = QComboBox()
        self.window_combo.setObjectName("ShellCombo")
        self.window_combo.setEditable(True)
        self.window_combo.setInsertPolicy(QComboBox.NoInsert)
        self.window_combo.setMaxVisibleItems(24)
        self.window_combo.setIconSize(QSize(16, 16))
        self.window_combo.setMinimumContentsLength(42)
        self.window_combo.setPlaceholderText("Select IFMIS/Oracle target window")
        self.window_combo.setView(QListView())
        self.window_combo.currentTextChanged.connect(self._update_window_combo_tooltip)
        self.window_combo.setMinimumWidth(280)
        shell_row.addWidget(self.window_combo, 5)

        self.refresh_btn = QPushButton("Refresh")
        self.refresh_btn.setObjectName("ShellRefreshButton")
        self.refresh_btn.setFixedHeight(38)
        self.refresh_btn.setCursor(Qt.PointingHandCursor)
        self.refresh_btn.clicked.connect(self._refresh_windows)
        shell_row.addWidget(self.refresh_btn)

        cg_label = QLabel("Command Group")
        cg_label.setObjectName("ShellFieldLabel")
        cg_label.setMinimumWidth(112)
        shell_row.addWidget(cg_label)

        self.command_group_combo = QComboBox()
        self.command_group_combo.setObjectName("ShellCombo")
        self.command_group_combo.setMaxVisibleItems(20)
        self.command_group_combo.setIconSize(QSize(16, 16))
        self.command_group_combo.setView(QListView())
        self.command_group_combo.addItems(COMMAND_GROUPS)
        for idx, group_name in enumerate(COMMAND_GROUPS):
            self.command_group_combo.setItemIcon(idx, self._icon_for_command_group(group_name))
            self.command_group_combo.setItemData(idx, group_name, Qt.ToolTipRole)
        self.command_group_combo.setCurrentIndex(1)  # Default: Oracle EBS R12 / 11i
        self.command_group_combo.setMinimumWidth(210)
        shell_row.addWidget(self.command_group_combo, 2)

        top_rows_layout.addLayout(shell_row)

        # â”€â”€ Separator line â”€â”€
        sep = QFrame()
        sep.setObjectName("TopRowsDivider")
        sep.setFrameShape(QFrame.HLine)
        top_rows_layout.addWidget(sep)
        main_layout.addWidget(self.top_rows_container, 0)

        # â”€â”€ Spreadsheet Grid â”€â”€
        # Formula Bar
        formula_card = QFrame()
        formula_card.setObjectName("FormulaCard")
        _fb_row = QHBoxLayout(formula_card)
        _fb_row.setContentsMargins(10, 8, 10, 8)
        _fb_row.setSpacing(8)
        self._cell_ref_label = QLabel("R1 C1")
        self._cell_ref_label.setObjectName("CellRefPill")
        self._cell_ref_label.setFixedWidth(74)
        self._cell_ref_label.setAlignment(Qt.AlignCenter)
        self._formula_bar = QLineEdit()
        self._formula_bar.setObjectName("FormulaBar")
        self._formula_bar.setPlaceholderText("Cell value")
        _fb_row.addWidget(self._cell_ref_label)
        _fb_row.addWidget(self._formula_bar, 1)
        main_layout.addWidget(formula_card, 0)

        self.spreadsheet = SpreadsheetWidget()
        main_layout.addWidget(self.spreadsheet, 1)

        self.setCentralWidget(central)

        # Connect cell selection to the formula bar
        self.spreadsheet.currentCellChanged.connect(self._update_formula_bar)
        self._formula_bar.returnPressed.connect(self._apply_formula_bar)
        self.spreadsheet.data_changed.connect(self._update_row_count)
        self.spreadsheet.paste_completed.connect(self._on_paste_completed)

        QTimer.singleShot(0, self._align_top_rows_with_toolbar_anchor)

        # Initial window refresh
        QTimer.singleShot(500, self._refresh_windows)

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    # Status Bar
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def _build_status_bar(self):
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        self.rows_label = QLabel("Rows: 0")
        self.rows_label.setStyleSheet(
            "font-weight: 600; font-size: 15px; color: #FFFFFF; padding: 0 8px;"
        )
        self.status_bar.addWidget(self.rows_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedWidth(250)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setVisible(False)
        self.status_bar.addPermanentWidget(self.progress_bar)

        self.eta_label = QLabel("")
        self.eta_label.setStyleSheet("font-weight: 600; font-size: 16px; color: #FFFFFF; padding: 0 8px;")
        self.eta_label.setVisible(False)
        self.status_bar.addPermanentWidget(self.eta_label)

        self.status_label = QLabel("Ready")
        self.status_label.setStyleSheet("font-weight: 600; font-size: 18px; color: #FFFFFF;")
        self.status_bar.addPermanentWidget(self.status_label)

    def _update_formula_bar(self, row, col, *_):
        self._cell_ref_label.setText(f"R{row + 1} C{col + 1}")
        item = self.spreadsheet.item(row, col)
        self._formula_bar.blockSignals(True)
        self._formula_bar.setText(item.text() if item else "")
        self._formula_bar.blockSignals(False)

    def _apply_formula_bar(self):
        row = self.spreadsheet.currentRow()
        col = self.spreadsheet.currentColumn()
        if row < 0 or col < 0:
            return
        item = self.spreadsheet._ensure_item(row, col)
        item.setText(self._formula_bar.text())
        self.spreadsheet.setFocus()

    def _update_row_count(self):
        count = self.spreadsheet.get_row_count_with_data()
        self.rows_label.setText(f"Rows: {count}")

    def _on_paste_completed(self, rows: int, cols: int):
        self.status_label.setText(f"Pasted {rows} row(s) x {cols} col(s)")

    def _undo_edit(self):
        if self.spreadsheet.undo():
            self.status_label.setText("Undo applied")
        else:
            self.status_label.setText("Nothing to undo")

    def _redo_edit(self):
        if self.spreadsheet.redo():
            self.status_label.setText("Redo applied")
        else:
            self.status_label.setText("Nothing to redo")

    def _go_back_position(self):
        if self.spreadsheet.go_back():
            row = self.spreadsheet.currentRow() + 1
            col = self.spreadsheet.currentColumn() + 1
            self.status_label.setText(f"Back to R{row} C{col}")
        else:
            self.status_label.setText("No previous position")

    def _go_forward_position(self):
        if self.spreadsheet.go_forward():
            row = self.spreadsheet.currentRow() + 1
            col = self.spreadsheet.currentColumn() + 1
            self.status_label.setText(f"Forward to R{row} C{col}")
        else:
            self.status_label.setText("No forward position")

    def _find_in_sheet(self):
        query, ok = QInputDialog.getText(
            self,
            "Find",
            "Search text:",
            QLineEdit.Normal,
            self._find_query,
        )
        if not ok:
            return

        query = query.strip()
        if not query:
            self.status_label.setText("Find: enter text to search.")
            return

        labels = ["All cells", "Current row", "Current column"]
        scope_to_label = {
            "all": "All cells",
            "row": "Current row",
            "column": "Current column",
        }
        default_label = scope_to_label.get(self._find_scope, "All cells")
        try:
            default_idx = labels.index(default_label)
        except ValueError:
            default_idx = 0

        chosen, ok = QInputDialog.getItem(
            self,
            "Find Scope",
            "Search in:",
            labels,
            default_idx,
            False,
        )
        if not ok:
            return

        label_to_scope = {
            "All cells": "all",
            "Current row": "row",
            "Current column": "column",
        }
        self._find_query = query
        self._find_scope = label_to_scope.get(chosen, "all")
        self._find_next_in_sheet()

    def _find_next_in_sheet(self):
        query = (self._find_query or "").strip()
        if not query:
            self._find_in_sheet()
            return

        scope = self._find_scope if self._find_scope in {"all", "row", "column"} else "all"
        match = self.spreadsheet.find_next_match(query, scope)
        scope_label = {
            "all": "all cells",
            "row": "current row",
            "column": "current column",
        }.get(scope, "all cells")

        if match is None:
            self.status_label.setText(f"Find: '{query}' not found in {scope_label}.")
            return

        row, col = match
        self.status_label.setText(
            f"Find: '{query}' at R{row + 1} C{col + 1} ({scope_label})."
        )

    def _find_replace_in_sheet(self):
        """Open a Find & Replace dialog."""
        dlg = QDialog(self)
        dlg.setWindowTitle("Find & Replace")
        dlg.setWindowFlag(Qt.WindowContextHelpButtonHint, False)
        dlg.setMinimumWidth(380)

        from kdl.styles import dialog_qss
        from kdl.config_store import get_dark_mode
        dlg.setStyleSheet(dialog_qss(dark=get_dark_mode()))

        layout = QVBoxLayout(dlg)
        layout.setSpacing(8)

        grid = QGridLayout()
        grid.addWidget(QLabel("Find:"), 0, 0)
        find_edit = QLineEdit(self._find_query or "")
        grid.addWidget(find_edit, 0, 1)
        grid.addWidget(QLabel("Replace with:"), 1, 0)
        replace_edit = QLineEdit()
        grid.addWidget(replace_edit, 1, 1)
        grid.addWidget(QLabel("Search in:"), 2, 0)
        scope_combo = QComboBox()
        scope_combo.addItems(["All cells", "Current row", "Current column"])
        scope_map = {"all": 0, "row": 1, "column": 2}
        scope_combo.setCurrentIndex(scope_map.get(self._find_scope or "all", 0))
        grid.addWidget(scope_combo, 2, 1)
        layout.addLayout(grid)

        btn_row = QHBoxLayout()
        find_next_btn = QPushButton("Find Next")
        replace_btn = QPushButton("Replace")
        replace_all_btn = QPushButton("Replace All")
        close_btn = QPushButton("Close")
        btn_row.addWidget(find_next_btn)
        btn_row.addWidget(replace_btn)
        btn_row.addWidget(replace_all_btn)
        btn_row.addStretch()
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

        result_label = QLabel("")
        result_label.setWordWrap(True)
        layout.addWidget(result_label)

        idx_to_scope = {0: "all", 1: "row", 2: "column"}

        def _scope():
            return idx_to_scope.get(scope_combo.currentIndex(), "all")

        def _find_next():
            q = find_edit.text().strip()
            if not q:
                result_label.setText("Enter text to find.")
                return
            self._find_query = q
            self._find_scope = _scope()
            match = self.spreadsheet.find_next_match(q, self._find_scope)
            if match is None:
                result_label.setText(f"'{q}' not found.")
            else:
                result_label.setText(f"Found at R{match[0]+1} C{match[1]+1}.")

        def _replace_one():
            q = find_edit.text().strip()
            r = replace_edit.text()
            if not q:
                result_label.setText("Enter text to find.")
                return
            self._find_query = q
            self._find_scope = _scope()
            # Replace only in current cell if it matches, then advance
            row = self.spreadsheet.currentRow()
            col = self.spreadsheet.currentColumn()
            item = self.spreadsheet.item(row, col)
            if item and q.casefold() in item.text().casefold():
                import re as _re
                new_text = _re.sub(_re.escape(q), r, item.text(), flags=_re.IGNORECASE)
                item.setText(new_text)
                result_label.setText("Replaced 1 occurrence.")
            _find_next()

        def _replace_all():
            q = find_edit.text().strip()
            r = replace_edit.text()
            if not q:
                result_label.setText("Enter text to find.")
                return
            self._find_query = q
            self._find_scope = _scope()
            count = self.spreadsheet.replace_in_cells(q, r, self._find_scope)
            result_label.setText(f"Replaced {count} occurrence(s).")

        find_next_btn.clicked.connect(_find_next)
        replace_btn.clicked.connect(_replace_one)
        replace_all_btn.clicked.connect(_replace_all)
        close_btn.clicked.connect(dlg.accept)
        dlg.exec()

    def _icon_for_command_group(self, group_name: str) -> QIcon:
        name = (group_name or "").lower()
        style = self.style()
        if "cloud" in name:
            return style.standardIcon(QStyle.SP_DriveNetIcon)
        if "oracle" in name:
            return style.standardIcon(QStyle.SP_DriveHDIcon)
        if "sap" in name or "siebel" in name or "peoplesoft" in name or "jde" in name:
            return style.standardIcon(QStyle.SP_ComputerIcon)
        if "dynamics" in name:
            return style.standardIcon(QStyle.SP_DesktopIcon)
        return style.standardIcon(QStyle.SP_FileIcon)

    def _icon_for_window_title(self, title: str, process_name: str = "") -> QIcon:
        text = f"{title} {process_name}".lower()
        style = self.style()
        if any(k in text for k in ("oracle", "ifmis", "forms", "ebs", "responsibility")):
            return style.standardIcon(QStyle.SP_DriveHDIcon)
        if any(k in text for k in ("chrome", "msedge", "firefox", "iexplore", "browser")):
            return style.standardIcon(QStyle.SP_DirOpenIcon)
        if any(k in text for k in ("excel", "sheet", ".xlsx", ".csv")):
            return style.standardIcon(QStyle.SP_FileIcon)
        return style.standardIcon(QStyle.SP_ComputerIcon)

    def _update_window_combo_tooltip(self, text: str):
        tip = text.strip() if text else "Select IFMIS/Oracle target window"
        self.window_combo.setToolTip(tip)

    def _is_oracle_like_window(self, title: str, process_name: str = "") -> bool:
        text = f"{title} {process_name}".lower()
        return any(k in text for k in ("oracle", "ifmis", "forms", "ebs", "responsibility"))

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    # Window Management
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def _refresh_windows(self):
        """Refresh the list of available target windows (enumeration runs off main thread)."""
        if self.loader_thread and self.loader_thread.isRunning():
            return
        if self._window_refresh_inflight:
            return
        self._window_refresh_inflight = True

        def _enumerate():
            try:
                windows = WindowManager.get_open_windows()
                # Collect process names while still in background
                result = []
                for hwnd, title in windows:
                    title_text = title or ""
                    if (
                        __display_name__ in title_text
                        or (("NT_DL" in title_text or "KDL" in title_text) and "Data Loader" in title_text)
                    ):
                        continue
                    pname = WindowManager.get_window_process_name(hwnd)
                    result.append((hwnd, title, pname))
                self.window_list_ready.emit(result)
            except Exception:
                self.window_list_ready.emit([])

        threading.Thread(target=_enumerate, daemon=True).start()

    def _apply_window_list(self, windows: list):
        """Apply enumerated window list to combo (runs on main thread)."""
        self._window_refresh_inflight = False
        current_text = self.window_combo.currentText()
        current_hwnd = self.window_combo.currentData()
        self.window_combo.clear()
        count = 0
        for hwnd, title, process_name in windows:
            self.window_combo.addItem(title, userData=hwnd)
            idx = self.window_combo.count() - 1
            self.window_combo.setItemIcon(idx, self._icon_for_window_title(title, process_name))
            self.window_combo.setItemData(idx, title, Qt.ToolTipRole)
            self.window_combo.setItemData(idx, process_name, Qt.UserRole + 1)
            count += 1

        restored = False
        if current_hwnd is not None:
            for idx in range(self.window_combo.count()):
                if self.window_combo.itemData(idx) == current_hwnd:
                    self.window_combo.setCurrentIndex(idx)
                    self._update_window_combo_tooltip(self.window_combo.currentText())
                    restored = True
                    break

        if not restored and current_text:
            idx = self.window_combo.findText(current_text)
            if idx >= 0:
                self.window_combo.setCurrentIndex(idx)
                self._update_window_combo_tooltip(current_text)
                restored = True

        if not restored:
            for idx in range(self.window_combo.count()):
                title = self.window_combo.itemText(idx)
                process_name = self.window_combo.itemData(idx, Qt.UserRole + 1) or ""
                if self._is_oracle_like_window(title, process_name):
                    self.window_combo.setCurrentIndex(idx)
                    self._update_window_combo_tooltip(title)
                    restored = True
                    break

        if not restored and self.window_combo.count() > 0:
            self.window_combo.setCurrentIndex(0)
            self._update_window_combo_tooltip(self.window_combo.currentText())

        if count == 0:
            self.window_combo.setEditText("")
            self._update_window_combo_tooltip("")
            self.status_label.setText("No target windows detected. Open IFMIS then click refresh.")
        else:
            self.status_label.setText(f"Detected {count} window(s).")

    def _open_delays_timeouts(self):
        """Set default delays used by the Start Load popup."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Delays & Timeouts")
        dialog.setFixedWidth(360)
        dialog.setWindowFlag(Qt.WindowCloseButtonHint, True)

        layout = QVBoxLayout(dialog)
        grid = QGridLayout()

        grid.addWidget(QLabel("Delay after cell processed (seconds):"), 0, 0)
        cell_edit = QLineEdit(f"{self._default_speed_delay:g}")
        cell_edit.setAlignment(Qt.AlignCenter)
        grid.addWidget(cell_edit, 0, 1)

        grid.addWidget(QLabel("Delay after window activated (seconds):"), 1, 0)
        window_edit = QLineEdit(f"{self._default_window_delay:g}")
        window_edit.setAlignment(Qt.AlignCenter)
        grid.addWidget(window_edit, 1, 1)

        layout.addLayout(grid)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        cancel_button = buttons.button(QDialogButtonBox.Cancel)
        if cancel_button is not None:
            cancel_button.setText("Close")
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        if dialog.exec() != QDialog.Accepted:
            return

        try:
            cell_delay = float(cell_edit.text().strip())
            window_delay = float(window_edit.text().strip())
        except ValueError:
            QMessageBox.warning(self, "Invalid Value", "Please enter valid numeric delay values.")
            return

        self._default_speed_delay = max(0.0, min(2.0, cell_delay))
        self._default_window_delay = max(0.0, min(5.0, window_delay))
        self._persist_settings()
        self.status_label.setText(
            f"Defaults updated: cell={self._default_speed_delay:g}s, window={self._default_window_delay:g}s"
        )

    def _open_options(self):
        """Set default load options used by the Start Load popup."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Options")
        dialog.setFixedWidth(500)
        dialog.setWindowFlag(Qt.WindowCloseButtonHint, True)

        layout = QVBoxLayout(dialog)

        wait_check = QCheckBox("Wait if Cursor is Hour Glass")
        wait_check.setChecked(self._default_wait_hourglass)
        layout.addWidget(wait_check)

        mode_row = QHBoxLayout()
        mode_row.addWidget(QLabel("Default load mode:"))
        mode_combo = QComboBox()
        for text, key in LOAD_MODES:
            mode_combo.addItem(text, userData=key)
        mode_index = max(0, mode_combo.findData(self._default_load_mode))
        mode_combo.setCurrentIndex(mode_index)
        mode_row.addWidget(mode_combo, 1)
        layout.addLayout(mode_row)

        validate_check = QCheckBox("Validate data before load")
        validate_check.setChecked(self._default_validate_before_load)
        layout.addWidget(validate_check)

        compact_check = QCheckBox("Enable Compact Mode (best for 13-inch screens)")
        compact_check.setChecked(self._compact_mode_enabled)
        layout.addWidget(compact_check)

        eor_row = QHBoxLayout()
        eor_row.addWidget(QLabel("Default after each row:"))
        eor_combo = QComboBox()
        selected_idx = 0
        for idx, (text, key) in enumerate(END_OF_ROW_ACTIONS):
            eor_combo.addItem(text, userData=key)
            if key == self._default_end_of_row_action:
                selected_idx = idx
        eor_combo.setCurrentIndex(selected_idx)
        eor_row.addWidget(eor_combo, 1)
        layout.addLayout(eor_row)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        cancel_button = buttons.button(QDialogButtonBox.Cancel)
        if cancel_button is not None:
            cancel_button.setText("Close")
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        if dialog.exec() != QDialog.Accepted:
            return

        chosen_mode = mode_combo.currentData() or "per_cell"
        self._default_wait_hourglass = wait_check.isChecked()
        self._default_load_mode = chosen_mode if chosen_mode in VALID_LOAD_MODES else "per_cell"
        self._default_form_mode = self._default_load_mode in ("per_row", "fast_send", "imprest_surrender", "imprest_test")
        self._default_validate_before_load = validate_check.isChecked()
        self._compact_mode_enabled = compact_check.isChecked()
        self._default_end_of_row_action = eor_combo.currentData()
        self._persist_settings()
        self._apply_responsive_window_settings()
        self.status_label.setText("Default load options updated")

    def _toggle_protect_load(self, checked: bool):
        self._protect_load_enabled = bool(checked)
        self._persist_settings()
        self.status_label.setText(
            "Protect Load enabled" if self._protect_load_enabled else "Protect Load disabled"
        )

    def _toggle_progress_bar_setting(self, checked: bool):
        self._show_progress_bar = bool(checked)
        self._persist_settings()
        if not self._show_progress_bar:
            self.progress_bar.setVisible(False)
        self.status_label.setText(
            "Progress bar enabled" if self._show_progress_bar else "Progress bar hidden"
        )

    def _activate_selected_target(self):
        hwnd = self.window_combo.currentData()
        title = self.window_combo.currentText().strip()

        if hwnd is None and title:
            hwnd = WindowManager.find_window_containing_title(title)

        if hwnd is None:
            self._show_styled_message(
                "No Target",
                "Please select a valid target window first.",
                status="warning",
            )
            return

        if WindowManager.activate_window(hwnd):
            self.status_label.setText(f"Activated target window: {title or hwnd}")
        else:
            self._show_styled_message(
                "Activation Failed",
                "Could not activate the selected window.",
                status="warning",
            )

    def _show_oracle_windows(self):
        windows = WindowManager.find_oracle_windows()
        if not windows:
            QMessageBox.information(
                self,
                "Oracle/IFMIS Windows",
                "No Oracle/IFMIS windows detected right now."
            )
            return

        lines = [f"{idx + 1}. {title}" for idx, (_, title) in enumerate(windows[:20])]
        QMessageBox.information(
            self,
            "Oracle/IFMIS Windows",
            "Detected windows:\n\n" + "\n".join(lines)
        )

    def _show_foreground_window(self):
        title = WindowManager.get_foreground_window_title().strip()
        hwnd = WindowManager.get_foreground_window_handle()
        if not title:
            QMessageBox.information(self, "Foreground Window", "No active foreground window detected.")
            return

        match_idx = -1
        for idx in range(self.window_combo.count()):
            item_hwnd = self.window_combo.itemData(idx)
            item_title = self.window_combo.itemText(idx).strip()
            if hwnd and item_hwnd == hwnd:
                match_idx = idx
                break
            if item_title == title:
                match_idx = idx
                break

        if match_idx >= 0:
            self.window_combo.setCurrentIndex(match_idx)
        else:
            self.window_combo.addItem(title, userData=hwnd if hwnd else None)
            idx = self.window_combo.count() - 1
            process_name = WindowManager.get_window_process_name(hwnd) if hwnd else ""
            self.window_combo.setItemIcon(idx, self._icon_for_window_title(title, process_name))
            self.window_combo.setItemData(idx, title, Qt.ToolTipRole)
            self.window_combo.setCurrentIndex(idx)

        self._update_window_combo_tooltip(title)
        self.status_label.setText(f"Foreground window selected: {title}")
        QMessageBox.information(self, "Foreground Window", f"Current foreground window:\n{title}")

    def _open_macro_recorder(self):
        dialog = MacroRecorderDialog(self)
        if dialog.exec() != QDialog.Accepted:
            return

        raw_macro_text = dialog.get_macro_text().strip()
        if not raw_macro_text:
            return
        macro_text = self._normalize_macro_cell_value(raw_macro_text)

        shortcut_saved = ""
        if dialog.should_save_shortcut():
            shortcut_key = dialog.get_shortcut_key()
            if shortcut_key:
                # Keep shortcut value in recorder keystroke form, but insert
                # normalized App Format into the sheet.
                self.parser.update_shortcut(shortcut_key, raw_macro_text)
                shortcut_saved = shortcut_key

        if dialog.apply_to_selection():
            sel = self.spreadsheet.get_selected_range()
            if not sel:
                QMessageBox.warning(
                    self,
                    "No Selection",
                    "Select a range first, or choose Current Cell in Macro Recorder."
                )
                return
            top, bottom, left, right = sel
            changed = 0
            for row in range(top, bottom + 1):
                for col in range(left, right + 1):
                    item = self.spreadsheet.item(row, col)
                    if item is None:
                        item = QTableWidgetItem("")
                        self.spreadsheet.setItem(row, col, item)
                    item.setText(macro_text)
                    changed += 1
            self.spreadsheet._refresh_highlighting()
            if shortcut_saved:
                self.status_label.setText(
                    f"Macro inserted into {changed} selected cell(s), saved as {shortcut_saved}"
                )
            else:
                self.status_label.setText(f"Macro inserted into {changed} selected cell(s)")
            return

        row = self.spreadsheet.currentRow()
        col = self.spreadsheet.currentColumn()
        if row < 0:
            row = 0
        if col < 0:
            col = 0

        item = self.spreadsheet.item(row, col)
        if item is None:
            item = QTableWidgetItem("")
            self.spreadsheet.setItem(row, col, item)
        item.setText(macro_text)
        self.spreadsheet._refresh_highlighting()
        if shortcut_saved:
            self.status_label.setText(
                f"Macro inserted at R{row + 1} C{col + 1}, saved as {shortcut_saved}"
            )
        else:
            self.status_label.setText(f"Macro inserted at R{row + 1} C{col + 1}")

    def _setup_databases(self):
        dialog = DatabaseSetupDialog({"database": self._db_settings}, self)
        if dialog.exec() != QDialog.Accepted:
            return

        updated = dialog.get_settings().get("database", {})
        if isinstance(updated, dict):
            self._db_settings = updated
            self._persist_settings()

            profile = self._db_settings.get("active_profile", "")
            mode_text = "UI Automation"
            if profile:
                self.status_label.setText(f"Database setup saved: {mode_text}, profile '{profile}'")
            else:
                self.status_label.setText(f"Database setup saved: {mode_text}")

    def _open_statement_converter(self):
        from kdl.dialogs.statement_converter_dialog import StatementConverterDialog
        dlg = StatementConverterDialog(self)
        dlg.load_into_grid.connect(self._load_statement_output_into_grid)
        dlg.exec()

    def _open_financial_report(self):
        dlg = FinancialReportDialog(self)
        dlg.exec()

    def _open_budget(self):
        dlg = BudgetDialog(self)
        dlg.exec()

    def _open_imprest_surrender(self):
        from kdl.dialogs.imprest_surrender_dialog import ImprestSurrenderDialog
        dlg = ImprestSurrenderDialog(self)
        dlg.load_into_grid.connect(self._load_imprest_output_into_grid)
        dlg.exec()

    def _load_imprest_output_into_grid(self, rows: list):
        if not rows:
            return
        try:
            self.spreadsheet.load_from_rows(rows)
            self.status_label.setText(
                f"Imprest Surrender loaded: {len(rows)} invoice(s). "
                "Press F5, select 'Imprest Surrender' mode, then Load.")
        except Exception as exc:
            self.status_label.setText("Imprest Surrender load failed.")
            QMessageBox.critical(
                self,
                "Imprest Surrender Error",
                f"Failed to load invoices into the grid:\n{exc}",
            )

    @staticmethod
    def _default_file_dialog_dir() -> str:
        downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        return downloads if os.path.isdir(downloads) else os.path.expanduser("~")

    def _load_statement_output_into_grid(self, rows: list):
        if not rows:
            return
        try:
            self.spreadsheet.load_from_rows(rows)
            self.status_label.setText(f"Statement Output loaded: {len(rows)} row(s)")
        except Exception as exc:
            self.status_label.setText("Statement Output load failed.")
            QMessageBox.critical(
                self,
                "Bank Statement Error",
                f"Failed to load the converted Output into the grid:\n{exc}",
            )

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    # File Operations
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def _apply_window_title(self, detail: str | None = None):
        if detail:
            self.setWindowTitle(f"{__display_name__}  |  {detail}")
            return
        self.setWindowTitle(__display_name__)

    def _new_file(self):
        if self._confirm_discard():
            self.spreadsheet.clear_all()
            self.current_file = None
            self._has_header_row = False
            self._apply_window_title("New File")
            self.status_label.setText("New file created")

    def _open_file(self):
        if not self._confirm_discard():
            return
        filepath, _ = QFileDialog.getOpenFileName(
            self, "Open File", self._default_file_dialog_dir(),
            "NT_DL Files (*.csv *.xlsx);;CSV Files (*.csv);;Excel Files (*.xlsx);;All Files (*)"
        )
        if filepath:
            self.spreadsheet.clear_all()
            self._has_header_row = False
            ext = os.path.splitext(filepath)[1].lower()
            if ext == ".xlsx":
                self.spreadsheet.import_excel(filepath)
            else:
                self.spreadsheet.import_csv(filepath)
            self.current_file = filepath
            self._apply_window_title(os.path.basename(filepath))
            self.status_label.setText(f"Opened: {os.path.basename(filepath)}")

    def _save_file(self) -> bool:
        if self.current_file:
            ext = os.path.splitext(self.current_file)[1].lower()
            if ext == ".xlsx":
                self.spreadsheet.export_excel(self.current_file)
            else:
                self.spreadsheet.export_csv(self.current_file)
            self.status_label.setText(f"Saved: {os.path.basename(self.current_file)}")
            return True
        else:
            return self._save_file_as()

    def _save_file_as(self) -> bool:
        filepath, selected_filter = QFileDialog.getSaveFileName(
            self, "Save File As", self._default_file_dialog_dir(),
            "CSV Files (*.csv);;Excel Files (*.xlsx);;All Files (*)"
        )
        if filepath:
            ext = os.path.splitext(filepath)[1].lower()
            if ext == ".xlsx":
                self.spreadsheet.export_excel(filepath)
            else:
                self.spreadsheet.export_csv(filepath)
            self.current_file = filepath
            self._apply_window_title(os.path.basename(filepath))
            self.status_label.setText(f"Saved: {os.path.basename(filepath)}")
            return True
        return False

    def _import_csv(self):
        filepath, _ = QFileDialog.getOpenFileName(
            self, "Import CSV", self._default_file_dialog_dir(), "CSV Files (*.csv);;All Files (*)"
        )
        if filepath:
            self.spreadsheet.import_csv(filepath)
            self._has_header_row = False
            self.status_label.setText(f"Imported CSV: {os.path.basename(filepath)}")

    def _import_excel(self):
        filepath, _ = QFileDialog.getOpenFileName(
            self, "Import Excel", self._default_file_dialog_dir(), "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        if filepath:
            self.spreadsheet.import_excel(filepath)
            self._has_header_row = False
            self.status_label.setText(f"Imported Excel: {os.path.basename(filepath)}")

    def _clear_all(self):
        reply = QMessageBox.question(
            self, "Clear All",
            "Are you sure you want to clear all data?",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.spreadsheet.clear_all()
            self._has_header_row = False
            self.status_label.setText("Cleared all data")

    def _confirm_discard(self) -> bool:
        return self._ask_save_before_close(
            title="Unsaved Changes",
            text="Do you want to save the current spreadsheet?",
        )

    def _ask_save_before_close(self, title: str, text: str) -> bool:
        has_data = self.spreadsheet.get_row_count_with_data() > 0
        if not has_data:
            return True

        reply = QMessageBox.question(
            self,
            title,
            text,
            QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
            QMessageBox.Yes,
        )
        if reply == QMessageBox.Cancel:
            return False
        if reply == QMessageBox.Yes:
            return self._save_file()
        return True

    def closeEvent(self, event: QCloseEvent):
        self._pending_load_result = None
        self._pending_load_thread = None
        self._deferred_load_result = None
        self._progress_hide_timer.stop()
        if self.loader_thread and self.loader_thread.isRunning():
            self.loader_thread.stop()
            self.loader_thread.wait(3000)
            if self.loader_thread.isRunning():
                self._show_styled_message(
                    "Please Wait",
                    "Load thread is still stopping. Try closing again in a moment.",
                    status="warning",
                )
                event.ignore()
                return

        if not self._ask_save_before_close(
            title="Exit NT_DL",
            text="Do you want to save the current spreadsheet?",
        ):
            event.ignore()
            return

        event.accept()

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    # Loading Operations
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def _start_loading(self):
        """Show load settings dialog and start loading."""
        data_rows = self.spreadsheet.get_row_count_with_data()
        if data_rows == 0:
            QMessageBox.warning(self, "No Data",
                                "No data to load. Please enter data or import a file first.")
            return

        # Pre-set target from the Window combo
        dialog = LoadSettingsDialog(
            max_rows=data_rows,
            target_title=self.window_combo.currentText(),
            target_hwnd=self.window_combo.currentData(),
            command_group=self.command_group_combo.currentText(),
            parent=self
        )

        # Keep current Start popup flow, but prefill from Tools defaults.
        dialog.cell_delay_input.setText(f"{self._default_speed_delay:g}")
        dialog.window_delay_input.setText(f"{self._default_window_delay:g}")
        dialog.hourglass_check.setChecked(self._default_wait_hourglass)
        dialog.load_control_check.setChecked(self._default_load_control)
        if self._default_load_mode == "imprest_surrender":
            dialog.radio_imprest.setChecked(True)
        elif self._default_load_mode == "imprest_test":
            dialog.radio_imprest_test.setChecked(True)
        elif self._default_load_mode == "per_row":
            dialog.radio_per_row.setChecked(True)
        elif self._default_load_mode == "fast_send":
            dialog.radio_fast_send.setChecked(True)
        else:
            dialog.radio_per_cell.setChecked(True)
        dialog._update_mode_controls()
        dialog._sync_load_control_state(self._default_load_control)
        dialog.validate_check.setChecked(self._default_validate_before_load)
        if self._default_popup_behavior == "stop":
            dialog.radio_popup_stop.setChecked(True)
        else:
            dialog.radio_popup_pause.setChecked(True)
        for idx, (_, key) in enumerate(END_OF_ROW_ACTIONS):
            if key == self._default_end_of_row_action:
                dialog.eor_combo.setCurrentIndex(idx)
                break

        dialog.load_requested.connect(self._execute_load)
        dialog.exec()

    def _detect_delay_columns(self, grid_data: list) -> set:
        """Detect delay columns by header labels such as 'Delay'."""
        delay_columns = set()
        if not grid_data:
            return delay_columns

        first_row = grid_data[0]
        for col_idx, cell_value in enumerate(first_row):
            header = str(cell_value).strip().lower()
            if not header:
                continue

            normalized = header.replace("_", " ").replace("-", " ")
            if "delay" in normalized or normalized in {"wait", "pause", "sleep"}:
                delay_columns.add(col_idx)

        return delay_columns

    @staticmethod
    def _col_label(col_idx: int) -> str:
        col_idx = max(0, int(col_idx))
        label = ""
        n = col_idx
        while True:
            label = chr((n % 26) + 65) + label
            n = n // 26 - 1
            if n < 0:
                break
        return label

    def _apply_validation_issue_markers(self, issues: list):
        """Mark validation problem cells for easy correction."""
        self.spreadsheet._refresh_highlighting()
        for issue in issues[:2000]:
            row = int(issue.row)
            col = int(issue.col)
            if row < 0 or col < 0 or row >= self.spreadsheet.rowCount() or col >= self.spreadsheet.columnCount():
                continue
            item = self.spreadsheet.item(row, col)
            if item is None:
                item = QTableWidgetItem("")
                self.spreadsheet.setItem(row, col, item)
            if issue.severity == "error":
                item.setBackground(QColor(RED_BG))
            else:
                item.setBackground(QColor(AMBER_BG))

    def _detect_header_row_for_validation(self, grid_data: list) -> bool:
        if self._has_header_row:
            return True
        if not grid_data:
            return False
        first = [str(v or "").strip().lower() for v in grid_data[0]]
        hits = 0
        for token in ("type", "code", "transaction date", "value date", "amount"):
            if any(token == cell or token in cell for cell in first):
                hits += 1
        return hits >= 2

    def _validate_rows_for_load(
        self, grid_data: list, from_row: int, to_row: int, selected_cols
    ) -> bool:
        issues = validate_ifmis_data(
            grid_data,
            has_header_row=self._detect_header_row_for_validation(grid_data),
            start_row=from_row,
            end_row=to_row,
            selected_columns=selected_cols,
        )
        if not issues:
            self.status_label.setText("Validation passed: no issues found.")
            return True

        self._apply_validation_issue_markers(issues)
        errors = [i for i in issues if i.severity == "error"]
        warnings = [i for i in issues if i.severity != "error"]

        preview = []
        for issue in issues[:12]:
            preview.append(
                f"R{issue.row + 1} {self._col_label(issue.col)}: {issue.message}"
            )
        details = "\n".join(preview)
        if len(issues) > 12:
            details += f"\n... and {len(issues) - 12} more."

        if errors:
            self._show_styled_message(
                "Validation Failed",
                f"Found {len(errors)} error(s) and {len(warnings)} warning(s).\n"
                "Fix errors before loading.\n\n"
                f"{details}",
                status="error",
            )
            self.status_label.setText(
                f"Validation failed: {len(errors)} errors, {len(warnings)} warnings."
            )
            return False

        proceed = self._ask_styled_message(
            "Validation Warnings",
            f"Found {len(warnings)} warning(s), no blocking errors.\n"
            "Continue with loading?\n\n"
            f"{details}",
            status="warning",
            accept_text="Continue",
            reject_text="Cancel",
        )
        if not proceed:
            self.status_label.setText("Load canceled after validation warnings.")
            return False
        return True

    def _validate_current_data(self):
        """Validate current sheet using built-in IFMIS rules."""
        grid_data = self.spreadsheet.get_grid_data()
        if not grid_data:
            self._show_styled_message("Validation", "No data found to validate.", status="info")
            return
        last_row = max(0, len(grid_data) - 1)
        if self._validate_rows_for_load(grid_data, 0, last_row, None):
            self._show_styled_message("Validation", "Validation passed.", status="success")

    def _execute_load(self, settings: dict):
        """Execute the data load with given settings."""
        self._last_load_settings = dict(settings)
        self._load_row_results = {}
        self._last_started_row = -1
        grid_data = self.spreadsheet.get_grid_data()

        # Keep latest settings as defaults for next Start popup.
        self._default_speed_delay = settings.get("speed_delay", self._default_speed_delay)
        self._default_window_delay = settings.get("window_delay", self._default_window_delay)
        self._default_wait_hourglass = settings.get("wait_hourglass", self._default_wait_hourglass)
        self._default_load_control = settings.get("load_control", self._default_load_control)
        chosen_mode = str(settings.get("load_mode", "")).strip().lower()
        if chosen_mode not in VALID_LOAD_MODES:
            chosen_mode = "per_cell"
        self._default_load_mode = chosen_mode
        self._default_form_mode = chosen_mode in ("per_row", "fast_send", "imprest_surrender", "imprest_test")
        self._default_validate_before_load = settings.get(
            "validate_before_load", self._default_validate_before_load
        )
        self._default_end_of_row_action = settings.get(
            "end_of_row_action", self._default_end_of_row_action
        )
        popup_behavior = str(settings.get("popup_behavior", self._default_popup_behavior)).strip().lower()
        self._default_popup_behavior = popup_behavior if popup_behavior in {"pause", "stop"} else "pause"
        self._persist_settings()

        from_row = settings.get("from_row", 0)
        to_row = settings.get("to_row", len(grid_data) - 1)

        # If we have a header row and mode is 'all', skip row 0 (headers)
        if settings.get("range_mode") == "all" and self._has_header_row:
            from_row = max(from_row, 1)

        if settings.get("range_mode") == "selected":
            sel = self.spreadsheet.get_selected_range()
            if sel:
                from_row, to_row = sel[0], sel[1]
            else:
                self._show_styled_message(
                    "No Selection",
                    "Please select rows to load.",
                    status="warning",
                )
                return

        load_mode = str(settings.get("load_mode", "")).strip().lower()
        if load_mode not in VALID_LOAD_MODES:
            load_mode = "per_cell"



        # If target window not set in dialog, use main window combo
        target_hwnd = settings.get("target_hwnd") or self.window_combo.currentData()
        target_title = settings.get("target_title") or self.window_combo.currentText()
        target_title = (target_title or "").strip()

        if not target_title:
            self._show_styled_message(
                "No Target",
                "Please select a target window first.",
                status="warning",
            )
            return

        if target_hwnd is None:
            target_hwnd = WindowManager.find_window_containing_title(target_title)

        if target_hwnd is None:
            self._show_styled_message(
                "Target Not Found",
                "Could not resolve the selected target window.\n"
                "Please click Refresh Windows and select the target from the list.",
                status="warning",
            )
            self.status_label.setText("Load canceled: target window not found.")
            return

        if self._protect_load_enabled:
            rows_to_load = max(0, to_row - from_row + 1)
            mode_label = {
                "per_cell": "Per Cell",
                "per_row": "Per Row",
            }.get(load_mode, "Per Cell")
            reply = QMessageBox.question(
                self,
                "Confirm Start Load",
                f"Mode: {mode_label}\n"
                f"Start loading {rows_to_load} row(s) to:\n{target_title or 'Selected target'} ?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                self.status_label.setText("Load canceled (Protect Load)")
                return

        # Get selected columns if in 'selected' mode
        selected_cols = None
        if settings.get("range_mode") == "selected":
            sel = self.spreadsheet.get_selected_range()
            if sel and len(sel) >= 4:
                selected_cols = set(range(sel[2], sel[3] + 1))

        if settings.get("validate_before_load", True):
            if not self._validate_rows_for_load(grid_data, from_row, to_row, selected_cols):
                return
        else:
            self.spreadsheet._refresh_highlighting()

        delay_cols = self._detect_delay_columns(grid_data)
        if selected_cols is not None:
            delay_cols = {c for c in delay_cols if c in selected_cols}

        # Clean up previous loader thread only after it has fully exited.
        if self.loader_thread is not None:
            if self.loader_thread.isRunning():
                self._show_styled_message(
                    "Please Wait",
                    "The previous load is still shutting down. Try again in a moment.",
                    status="warning",
                )
                self.status_label.setText("Previous load is still shutting down.")
                return
            self.loader_thread.setParent(None)
            self.loader_thread.deleteLater()
            self.loader_thread = None
        self.loader_thread = LoaderThread(self)
        self._load_error_count = 0
        self._pending_load_result = None
        self._pending_load_thread = None
        self._deferred_load_result = None

        self.loader_thread.configure(
            grid_data=grid_data,
            start_row=from_row,
            end_row=to_row,
            target_hwnd=target_hwnd,
            target_title=target_title,
            speed_delay=settings.get("speed_delay", 0.1),
            window_delay=settings.get("window_delay", 0.1),
            wait_hourglass=settings.get("wait_hourglass", False),
            key_columns=list(self.spreadsheet.key_columns),
            selected_columns=list(selected_cols) if selected_cols else None,
            delay_columns=list(delay_cols),
            form_mode=load_mode in ("per_row", "fast_send", "imprest_surrender", "imprest_test"),
            load_mode=load_mode,
            end_of_row_action=settings.get("end_of_row_action", "none"),
            save_interval=settings.get("save_interval", 50),
            db_settings=self._db_settings,
            use_fast_send=load_mode in ("fast_send", "imprest_surrender", "imprest_test"),
            popup_stop_on_error=settings.get("popup_behavior", "pause") == "stop",
            load_control=settings.get("load_control", False),
        )
        self.loader_thread.parser.shortcuts = self.parser.shortcuts

        # Connect signals
        self.loader_thread.progress_updated.connect(self._on_progress)
        self.loader_thread.cell_processed.connect(self._on_cell_processed)
        self.loader_thread.loading_complete.connect(self._on_loading_complete)
        self.loader_thread.finished.connect(self._on_loader_thread_finished)
        self.loader_thread.row_started.connect(self._on_row_started)
        self.loader_thread.step_waiting.connect(self._on_step_waiting)
        self.loader_thread.popup_paused.connect(self._on_popup_paused)

        # UI state
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.pause_btn.setEnabled(True)
        self.step_btn.setEnabled(False)
        planned_rows = max(0, to_row - from_row + 1)
        if selected_cols is not None:
            overlay_cols = sorted(selected_cols)
        elif grid_data and from_row < len(grid_data):
            overlay_cols = [
                idx for idx, val in enumerate(grid_data[from_row])
                if str(val).strip() != ""
            ]
        else:
            overlay_cols = []
        if not overlay_cols:
            overlay_cols = list(range(self.spreadsheet.columnCount()))
        self._overlay_columns = overlay_cols
        self._overlay_column_lookup = {col: idx + 1 for idx, col in enumerate(overlay_cols)}
        self._overlay_total_cols = len(overlay_cols)
        self._overlay_total_rows = max(0, to_row + 1)
        self._overlay_start_row = from_row
        self._overlay_active_row_abs = max(0, from_row)
        planned_visual_cells = max(1, planned_rows) * max(1, self._overlay_total_cols)
        self._load_visual_compact = planned_visual_cells >= 8000
        self._load_visual_cell_stride = 25 if self._load_visual_compact else 1
        self._load_visual_row_stride = 10 if self._load_visual_compact else 1
        self._load_visual_cells_seen = 0
        self._load_overlay.show_overlay(self._overlay_total_rows, self._overlay_total_cols)
        self._refresh_timer.stop()
        self._progress_hide_timer.stop()

        self.progress_bar.setVisible(True)
        if planned_rows > 0:
            self.progress_bar.setRange(0, planned_rows)
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat(f"0/{planned_rows} (%p%)")
        else:
            self.progress_bar.setRange(0, 0)
            self.progress_bar.setFormat("0/0")
        mode_label = {
            "per_cell":          "Per Cell",
            "per_row":           "Per Row",
            "fast_send":         "Fast Send",
            "imprest_surrender": "Imprest",
            "imprest_test":      "Imprest Test",
        }.get(load_mode, "Per Cell")
        self._load_overlay.set_mode_label(mode_label)
        self.status_label.setText(f"Loading ({mode_label})... Switch to target window!")

        self.loader_thread.start()

    def _stop_loading(self):
        if self.loader_thread and self.loader_thread.isRunning():
            self.loader_thread.stop()
            self.stop_btn.setEnabled(False)
            self.status_label.setText("Stopping...")

    def _pause_resume(self):
        if self.loader_thread and self.loader_thread.isRunning():
            if not self._is_paused:
                self.loader_thread.pause()
                self.pause_btn.setText("Resume")
                self.status_label.setText("Paused")
                self._is_paused = True
            else:
                self.loader_thread.resume()
                self.pause_btn.setText("Pause")
                self.status_label.setText("Resumed...")
                self._is_paused = False

    def _on_popup_paused(self, popup_title: str):
        """Handle automatic pause triggered by a detected popup/LOV.

        Shows a dialog letting the user dismiss the popup in IFMIS,
        then choose to resume from the current row, a different row, or stop.
        """
        self._is_paused = True
        self.pause_btn.setText("Resume")
        self.status_label.setText(f"Paused — popup: {popup_title}")

        # Bring KDL to front so the user sees the dialog
        self.showNormal()
        self.raise_()
        self.activateWindow()

        grid_data = self.spreadsheet.get_grid_data()
        total_rows = len(grid_data)
        paused_row = max(0, self._last_started_row)

        dlg = QDialog(self)
        dlg.setWindowTitle("Popup Detected — Load Paused")
        dlg.setWindowFlag(Qt.WindowContextHelpButtonHint, False)
        dlg.setMinimumWidth(440)

        from kdl.styles import dialog_qss
        from kdl.config_store import get_dark_mode
        dlg.setStyleSheet(dialog_qss(dark=get_dark_mode()))

        layout = QVBoxLayout(dlg)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        lbl_status = QLabel(
            f"<b>Popup detected:</b> {popup_title}<br>"
            f"Load paused at <b>Row {paused_row + 1}</b>."
        )
        lbl_status.setWordWrap(True)
        layout.addWidget(lbl_status)

        # Step-by-step guidance
        from PySide6.QtWidgets import QGroupBox as _GB
        guide_box = _GB("Steps:")
        guide_layout = QVBoxLayout(guide_box)
        guide_layout.setSpacing(4)
        steps = [
            "1.  Minimise this dialog  (click the  ─  button above)",
            "2.  In IFMIS, dismiss the popup (Cancel / Escape / OK)",
            "3.  Clear any partial data if needed",
            "4.  Click the <b>first field</b> of the row you want to continue from",
            "5.  Return here and choose how to continue",
        ]
        for step in steps:
            lbl = QLabel(step)
            lbl.setWordWrap(True)
            guide_layout.addWidget(lbl)
        layout.addWidget(guide_box)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setFrameShadow(QFrame.Sunken)
        layout.addWidget(sep)

        resume_lbl = QLabel("Continue from:")
        layout.addWidget(resume_lbl)

        btn_group = QButtonGroup(dlg)

        rb_current = QRadioButton(f"Current row  (Row {paused_row + 1} — retry)")
        rb_current.setChecked(True)
        btn_group.addButton(rb_current, 0)
        layout.addWidget(rb_current)

        next_row = min(paused_row + 1, max(0, total_rows - 1))
        rb_next = QRadioButton(f"Next row  (Row {next_row + 1} — skip current)")
        btn_group.addButton(rb_next, 1)
        layout.addWidget(rb_next)

        rb_custom = QRadioButton("Custom row:")
        btn_group.addButton(rb_custom, 2)
        custom_row = QHBoxLayout()
        custom_row.addWidget(rb_custom)
        spin = QSpinBox()
        spin.setMinimum(1)
        spin.setMaximum(max(1, total_rows))
        spin.setValue(paused_row + 1)
        spin.setEnabled(False)
        custom_row.addWidget(spin)
        custom_row.addStretch()
        layout.addLayout(custom_row)

        rb_custom.toggled.connect(spin.setEnabled)

        btns = QDialogButtonBox()
        btns.addButton("Resume", QDialogButtonBox.AcceptRole)
        btns.addButton("Stop Load", QDialogButtonBox.RejectRole)
        btns.accepted.connect(dlg.accept)
        btns.rejected.connect(dlg.reject)
        layout.addWidget(btns)

        result = dlg.exec()

        if result == QDialog.Accepted and self.loader_thread:
            chosen_id = btn_group.checkedId()
            if chosen_id == 0:
                resume_row = paused_row
            elif chosen_id == 1:
                resume_row = next_row
            else:
                resume_row = spin.value() - 1

            if resume_row != paused_row:
                # User chose a different row — stop current load and relaunch
                self.loader_thread.stop()
                self.loader_thread.wait(2000)
                settings = dict(self._last_load_settings)
                settings["from_row"] = resume_row
                settings["range_mode"] = "custom"
                QTimer.singleShot(200, lambda: self._execute_load(settings))
            else:
                # Resume from exactly where we paused
                self.loader_thread.resume()
                self.pause_btn.setText("Pause")
                self.status_label.setText("Resumed...")
                self._is_paused = False
        else:
            # User chose Stop
            if self.loader_thread:
                self.loader_thread.stop()

    def _next_step(self):
        if self.loader_thread:
            self.loader_thread.resume()

    def _on_progress(self, current, total, message):
        current = max(0, int(current))
        total = max(0, int(total))
        if total > 0:
            if self.progress_bar.maximum() != total:
                self.progress_bar.setRange(0, total)
            shown = min(current, total)
            self.progress_bar.setValue(shown)
            self.progress_bar.setFormat(f"{shown}/{total} (%p%)")
        else:
            if self.progress_bar.maximum() != 0:
                self.progress_bar.setRange(0, 0)
            self.progress_bar.setFormat("Loading...")
        # Extract ETA from message and show in dedicated label
        eta_text = ""
        if "ETA:" in message:
            try:
                eta_text = "ETA: " + message.split("ETA:")[1].strip()
                message = message.split("| ETA:")[0].strip()
            except Exception:
                eta_text = ""
        self.eta_label.setText(eta_text)
        self.eta_label.setVisible(bool(eta_text))
        self.status_label.setText(message)
        absolute_by_progress = max(0, self._overlay_start_row + current)
        overlay_current = max(absolute_by_progress, self._overlay_active_row_abs)
        self._load_overlay.update_rows(overlay_current, self._overlay_total_rows)

    def _on_cell_processed(self, row, col, success):
        # Track per-row result for status coloring
        if not success:
            self._load_row_results[row] = False
        elif row not in self._load_row_results:
            self._load_row_results[row] = True
        self._load_visual_cells_seen += 1
        update_visuals = (
            not self._load_visual_compact
            or not success
            or (self._load_visual_cells_seen % self._load_visual_cell_stride) == 0
        )

        item = self.spreadsheet.item(row, col)
        if update_visuals:
            if not self._load_visual_compact or not success:
                self.spreadsheet.set_loading_position(row, col, keep_visible=False)
            current_col = self._overlay_column_lookup.get(col, col + 1)
            self._load_overlay.update_cols(current_col, self._overlay_total_cols)
            if item is not None:
                self._load_overlay.set_cell_text(item.text())
            else:
                self._load_overlay.set_cell_text(f"R{row + 1} C{col + 1}")

        if success:
            return

        self._load_error_count += 1
        if item is not None:
            item.setBackground(QColor(RED_BG))

        self.status_label.setText(
            f"Error at R{row + 1} C{col + 1} (errors: {self._load_error_count})"
        )

    def _on_row_started(self, row):
        self._last_started_row = row
        should_highlight_row = (
            not self._load_visual_compact
            or row == self._overlay_start_row
            or ((row - self._overlay_start_row) % self._load_visual_row_stride) == 0
        )
        if should_highlight_row:
            keep_visible = (
                not self._load_visual_compact
                or ((row - self._overlay_start_row) % self._load_visual_row_scroll_stride) == 0
            )
            self.spreadsheet.highlight_loading_row(row, keep_visible=keep_visible)
        self._overlay_active_row_abs = max(0, row + 1)
        self._load_overlay.set_rows_marker(f"R{row + 1}")

    def _on_step_waiting(self, row, col):
        self.status_label.setText(f"Step mode: Waiting at R{row+1} C{col+1}. Click 'Step' to continue.")
        self._load_overlay.set_cell_text(f"Waiting R{row+1} C{col+1}")

    def _on_loading_complete(self, success, message):
        self.stop_btn.setEnabled(False)
        self.pause_btn.setEnabled(False)
        self.pause_btn.setText("Pause")
        self._is_paused = False
        self.step_btn.setEnabled(False)
        self._pending_load_result = (bool(success), str(message or ""))
        self._pending_load_thread = self.sender()
        self.spreadsheet.clear_loading_position()
        self._overlay_active_row_abs = 0
        self._load_visual_compact = False
        self._load_visual_row_stride = 1
        self._load_visual_cells_seen = 0
        self._load_overlay.set_rows_marker("")
        self._load_overlay.hide()
        self.eta_label.setText("")
        self.eta_label.setVisible(False)
        self.status_label.setText(message)

    def _on_loader_thread_finished(self):
        finished_thread = self.sender()

        self.start_btn.setEnabled(True)
        # Only snap to 100% if load completed successfully; keep partial value on stopped loads
        was_success = self._pending_load_result[0] if self._pending_load_result else True
        maximum = self.progress_bar.maximum()
        if maximum > 0:
            if was_success:
                self.progress_bar.setValue(maximum)
                self.progress_bar.setFormat(f"{maximum}/{maximum} (100%)")
            # else: leave the bar showing actual rows reached
            self._progress_hide_timer.start(3000)
        else:
            self.progress_bar.setVisible(False)

        self.spreadsheet.clear_loading_position()
        self._overlay_active_row_abs = 0
        self._load_visual_compact = False
        self._load_visual_row_stride = 1
        self._load_visual_cells_seen = 0
        self._load_overlay.set_rows_marker("")
        self._load_overlay.hide()

        if not self._refresh_timer.isActive():
            self._refresh_timer.start(10000)

        if finished_thread is self.loader_thread:
            self.loader_thread = None

        if finished_thread is not None:
            finished_thread.setParent(None)
            finished_thread.deleteLater()

        if self._pending_load_thread is finished_thread and self._pending_load_result is not None:
            self._deferred_load_result = self._pending_load_result
        self._pending_load_result = None
        self._pending_load_thread = None

        if self._deferred_load_result is not None:
            QTimer.singleShot(150, self._show_deferred_load_result)

    def _show_deferred_load_result(self):
        if self._deferred_load_result is None:
            return

        success, message = self._deferred_load_result
        self._deferred_load_result = None
        self.status_label.setText(message)
        self._apply_load_row_colors()
        self.showNormal()
        self.raise_()
        self.activateWindow()
        self._show_load_result(success, message)

    def _apply_load_row_colors(self):
        """Apply green/red row highlights based on per-row load results."""
        for row, ok in self._load_row_results.items():
            self.spreadsheet.highlight_row_result(row, ok)

    def _toggle_freeze_header(self, checked: bool):
        self._freeze_header_enabled = checked
        self.spreadsheet.set_freeze_header(checked)

    def _show_styled_message(self, title: str, message: str, status: str = "info"):
        dialog = LoadResultDialog(title=title, message=message, status=status, parent=self)
        dialog.exec()

    def _ask_styled_message(
        self,
        title: str,
        message: str,
        status: str = "warning",
        accept_text: str = "Yes",
        reject_text: str = "No",
    ) -> bool:
        dialog = LoadResultDialog(
            title=title,
            message=message,
            status=status,
            parent=self,
            confirm=True,
            accept_text=accept_text,
            reject_text=reject_text,
        )
        result: int = dialog.exec()
        return bool(result == QDialog.Accepted)

    def _show_load_result(self, success: bool, message: str):
        text = (message or "").strip()
        lowered = text.lower()

        if success:
            title = "Loading Complete"
            status = "success"
        elif "load stopped" in lowered or "loading stopped" in lowered:
            title = "Load Stopped"
            status = "stopped"
        else:
            title = "Loading Result"
            status = "error"

        self._show_styled_message(title, text, status)

        # Offer resume after a stopped (non-success) load
        if not success and self._last_load_settings:
            QTimer.singleShot(100, self._show_resume_dialog)

    def _show_resume_dialog(self):
        """Ask user if they want to resume from where the load stopped."""
        if not self._last_load_settings:
            return

        grid_data = self.spreadsheet.get_grid_data()
        total_rows = len(grid_data)
        last_row = self._last_started_row  # 0-based last row that started

        dlg = QDialog(self)
        dlg.setWindowTitle("Load Stopped — Action Required")
        dlg.setWindowFlag(Qt.WindowContextHelpButtonHint, False)
        dlg.setMinimumWidth(420)

        from kdl.styles import dialog_qss
        from kdl.config_store import get_dark_mode
        dlg.setStyleSheet(dialog_qss(dark=get_dark_mode()))

        layout = QVBoxLayout(dlg)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        # ── Status line ──
        lbl_status = QLabel("<b>Load stopped.</b>")
        lbl_status.setWordWrap(True)
        layout.addWidget(lbl_status)

        # ── Step-by-step guidance ──
        from PySide6.QtWidgets import QGroupBox as _GB, QFrame as _Fr
        guide_box = _GB("Before resuming or closing — do the following steps:")
        guide_layout = QVBoxLayout(guide_box)
        guide_layout.setSpacing(4)
        steps = [
            "1.  Minimise this dialog  (click the  ─  button above)",
            "2.  In IFMIS/Oracle, clear any partial data from the row you want to continue from",
            "3.  If you want to close and save first, do that now",
            "4.  Click the <b>first field</b> of the row you want to continue from",
            "5.  Return here and choose how you want to resume",
        ]
        for step in steps:
            lbl = QLabel(step)
            lbl.setWordWrap(True)
            guide_layout.addWidget(lbl)
        layout.addWidget(guide_box)

        # ── Separator ──
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setFrameShadow(QFrame.Sunken)
        layout.addWidget(sep)

        # ── Resume options ──
        resume_lbl = QLabel("Resume options:")
        layout.addWidget(resume_lbl)

        btn_group = QButtonGroup(dlg)

        rb_stopped = QRadioButton("Retry the stopped row  (re-enter the same grid row)")
        rb_stopped.setChecked(True)
        btn_group.addButton(rb_stopped, 0)
        layout.addWidget(rb_stopped)

        rb_begin = QRadioButton("Row 1  (start from the beginning)")
        btn_group.addButton(rb_begin, 1)
        layout.addWidget(rb_begin)

        rb_custom = QRadioButton("Custom row:")
        btn_group.addButton(rb_custom, 2)
        custom_row = QHBoxLayout()
        custom_row.addWidget(rb_custom)
        spin = QSpinBox()
        spin.setMinimum(1)
        spin.setMaximum(max(1, total_rows))
        spin.setValue(1)
        spin.setEnabled(False)
        custom_row.addWidget(spin)
        custom_row.addStretch()
        layout.addLayout(custom_row)

        rb_custom.toggled.connect(spin.setEnabled)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.button(QDialogButtonBox.Ok).setText("Resume")
        btns.accepted.connect(dlg.accept)
        btns.rejected.connect(dlg.reject)
        layout.addWidget(btns)

        if dlg.exec() != QDialog.Accepted:
            return

        chosen_id = btn_group.checkedId()
        if chosen_id == 0:
            from_row = max(0, last_row)
        elif chosen_id == 1:
            from_row = 0
        else:
            from_row = spin.value() - 1  # convert to 0-based

        settings = dict(self._last_load_settings)
        settings["from_row"] = from_row
        settings["range_mode"] = "custom"
        self._execute_load(settings)

    def _apply_toolbar_action_colors(self):
        """Apply high-contrast colours while keeping one consistent toolbar button system."""
        if not self.main_toolbar:
            return

        # Baseline for every toolbar button so icons look from one family.
        for action in self.main_toolbar.actions():
            btn = self.main_toolbar.widgetForAction(action)
            if btn is None or not isinstance(btn, QToolButton):
                continue
            btn.setStyleSheet(
                """
                QToolButton {
                    background-color: rgba(255, 255, 255, 0.12);
                    color: #FFFFFF;
                    border: 1px solid rgba(255, 255, 255, 0.22);
                    border-radius: 12px;
                    padding: 5px;
                    min-width: 44px;
                    min-height: 40px;
                }
                QToolButton:hover {
                    background-color: rgba(255, 255, 255, 0.18);
                    border: 1px solid rgba(255, 255, 255, 0.30);
                }
                QToolButton:pressed {
                    background-color: rgba(0, 0, 0, 0.18);
                }
                QToolButton:disabled {
                    background-color: rgba(209, 213, 219, 0.72);
                    color: #6B7280;
                    border: 1px solid #9CA3AF;
                }
                """
            )

        def style_action(action: QAction, bg: str, fg: str = "#FFFFFF", border: str = "#1F2937"):
            btn = self.main_toolbar.widgetForAction(action)
            if btn is None or not isinstance(btn, QToolButton):
                return
            base = QColor(bg)
            top = base.lighter(118).name()
            bottom = base.darker(108).name()
            hover_top = base.lighter(128).name()
            hover_bottom = base.darker(100).name()
            pressed_bg = base.darker(120).name()
            hover_border = QColor(border).darker(110).name()
            btn.setStyleSheet(
                f"""
                QToolButton {{
                    background-color: qlineargradient(
                        x1:0, y1:0, x2:0, y2:1,
                        stop:0 {top},
                        stop:1 {bottom}
                    );
                    color: {fg};
                    border: 1px solid {border};
                    border-radius: 12px;
                    padding: 5px;
                    min-width: 44px;
                    min-height: 40px;
                }}
                QToolButton:hover {{
                    background-color: qlineargradient(
                        x1:0, y1:0, x2:0, y2:1,
                        stop:0 {hover_top},
                        stop:1 {hover_bottom}
                    );
                    border: 1px solid {hover_border};
                }}
                QToolButton:pressed {{
                    background-color: {pressed_bg};
                    border: 1px solid {border};
                    color: {fg};
                }}
                QToolButton:disabled {{
                    background-color: #D1D5DB;
                    color: #6B7280;
                    border: 1px solid #9CA3AF;
                }}
                """
            )

        style_action(self.start_btn, "#16A34A", "#FFFFFF", "#166534")
        style_action(self.stop_btn, "#DC2626", "#FFFFFF", "#7F1D1D")
        style_action(self.pause_btn, "#D97706", "#111827", "#92400E")
        style_action(self.step_btn, "#2563EB", "#FFFFFF", "#1E3A8A")
        style_action(self.statement_btn, "#2563EB", "#FFFFFF", "#1D4ED8")
        style_action(self.report_btn, "#0F766E", "#FFFFFF", "#115E59")
        style_action(self.budget_btn, "#C77A11", "#FFFFFF", "#9A5A09")
        style_action(self.imprest_btn, "#0891B2", "#FFFFFF", "#0E7490")
        style_action(self.rec_btn, "#BE123C", "#FFFFFF", "#881337")
        style_action(self.convert_table_btn, "#0F766E", "#FFFFFF", "#115E59")
        style_action(self.convert_cell_btn, "#0EA5E9", "#0F172A", "#0369A1")

    @staticmethod
    def _normalize_backslash_macro(text: str) -> str:
        """
        Normalize single-token brace macros toward NT_DL standard keywords.
        Example: \\{TAB} -> \\{TAB}, \\{DOWN} -> dn
        """
        match = re.fullmatch(r"\\\{([^{}]+)\}", text.strip(), re.IGNORECASE)
        if not match:
            return text.strip()

        inner = " ".join(match.group(1).strip().split())
        if not inner:
            return text.strip()

        parts = inner.split(" ", 1)
        key = parts[0].upper()
        alias = {
            "DN": "DOWN",
            "ESCAPE": "ESC",
            "DEL": "DELETE",
            "INS": "INSERT",
            "BKSP": "BACKSPACE",
            "PAGEDOWN": "PGDN",
            "PAGEUP": "PGUP",
        }
        key = alias.get(key, key)

        # Keep counted forms explicit, e.g. \{TAB 5}.
        if len(parts) > 1:
            count = parts[1].strip()
            if count.isdigit():
                return f"\\{{{key} {count}}}"
            return text.strip()

        keyword_map = {
            "TAB": "\\{TAB}",
            "ENTER": "enter",
            "DOWN": "dn",
            "UP": "up",
            "LEFT": "left",
            "RIGHT": "right",
            "ESC": "esc",
        }
        if key in keyword_map:
            return keyword_map[key]

        if len(parts) == 1:
            return f"\\{{{key}}}"
        return text.strip()

    @classmethod
    def _normalize_macro_cell_value(cls, value: str) -> str:
        """
        Convert values into NT_DL standard format for clarity and IFMIS stability.
        Examples:
          tab    -> \\{TAB}
          \\*s    -> *S
          r    -> \\r
        """
        text = str(value or "").strip()
        if not text:
            return ""

        shortcut_aliases = {
            "*SV": "*S",
        }

        # Keep forced literal data unchanged.
        if text.startswith("'"):
            return text

        # Legacy escaped shortcuts (e.g. \*s) -> standard app shortcut (*S).
        escaped_shortcut = re.fullmatch(r"\\\*([A-Za-z0-9_]+)", text)
        if escaped_shortcut:
            resolved = f"*{escaped_shortcut.group(1).upper()}"
            return shortcut_aliases.get(resolved, resolved)

        # Direct ctrl+s keystroke becomes app shortcut save.
        if text.lower() == r"\^s":
            return "*S"

        # Existing keystroke macros: normalize simple brace forms to standard keywords.
        if text.startswith("\\"):
            return cls._normalize_backslash_macro(text)

        # Shortcut commands use *NAME in app format.
        if text.startswith("*"):
            body = text[1:]
            if body and all(ch.isalnum() or ch == "_" for ch in body):
                resolved = f"*{body.upper()}"
                return shortcut_aliases.get(resolved, resolved)
            return text

        keyword = " ".join(text.lower().split())
        keyword_map = {
            "tab": "\\{TAB}",
            "enter": "enter",
            "dn": "dn",
            "down": "dn",
            "up": "up",
            "left": "left",
            "right": "right",
            "esc": "esc",
            "escape": "esc",
        }
        if keyword in keyword_map:
            return keyword_map[keyword]

        if re.fullmatch(r"f(?:[1-9]|1[0-6])", keyword):
            return f"\\{{{keyword.upper()}}}"

        # Single letter shorthand should type the key directly (e.g. r -> \r).
        if len(text) == 1 and text.isalpha():
            return f"\\{text.lower()}"

        return text

    @staticmethod
    def _normalize_header_text(value) -> str:
        text = str(value or "").strip().lower()
        text = re.sub(r"[^a-z0-9]+", " ", text)
        return re.sub(r"\s+", " ", text).strip()

    @staticmethod
    def _clean_cell_text(value) -> str:
        return str(value or "").strip()

    @staticmethod
    def _normalize_date_text(token: str) -> str:
        """Return a date-like token, trimming common trailing time fragments when present."""
        text = str(token or "").strip()
        if not text:
            return ""
        if DATE_TEXT_RE.fullmatch(text):
            return text
        first = text.split()[0]
        if DATE_TEXT_RE.fullmatch(first):
            return first
        return text

    @classmethod
    def _extract_dates_from_payload(cls, payload: list[str], amount: str) -> tuple[str, str]:
        """Extract transaction/value dates with a permissive fallback for non-standard formats."""
        rest = [cls._clean_cell_text(v) for v in payload[1:]]
        rest = [v for v in rest if v and v != amount]
        if not rest:
            return "", ""

        normalized = [cls._normalize_date_text(v) for v in rest]
        strict = [v for v in normalized if DATE_TEXT_RE.fullmatch(v)]

        if strict:
            t_date = strict[0]
            v_date = strict[1] if len(strict) >= 2 else t_date
            return t_date, v_date

        # Fallback: keep first two non-control payload values as dates instead of blanking.
        t_date = normalized[0]
        v_date = normalized[1] if len(normalized) >= 2 else t_date
        return t_date, v_date

    @classmethod
    def _detect_table_col_map(cls, grid_data: list) -> tuple[dict, bool]:
        aliases = {
            "line": {"line", "line no", "line number", "no", "no.", "serial", "serial no"},
            "type": {"type", "transaction type", "trx type"},
            "code": {"code", "transaction code", "trx code"},
            "number": {"number", "reference", "reference number", "transaction number"},
            "transaction_date": {"transaction date", "trx date", "date"},
            "value_date": {"value date"},
            "amount": {"amount", "transaction amount"},
        }
        if grid_data:
            normalized = [cls._normalize_header_text(v) for v in grid_data[0]]
            found = {}
            for idx, name in enumerate(normalized):
                if not name:
                    continue
                for key, names in aliases.items():
                    if key in found:
                        continue
                    if name in names:
                        found[key] = idx
            required = {"type", "code", "number", "transaction_date", "value_date", "amount"}
            if required.issubset(set(found)):
                return found, True

        # Default fallback by positional layout.
        return {
            "line": 0,
            "type": 1,
            "code": 2,
            "number": 3,
            "transaction_date": 4,
            "value_date": 5,
            "amount": 6,
        }, False

    @staticmethod
    def _normalize_table_type(raw_type: str, raw_code: str) -> str:
        lowered = (raw_type or "").strip().lower()
        if lowered in {"receipt", "r"}:
            return "Receipt"
        if lowered in {"payment", "p"}:
            return "Payment"
        if (raw_code or "").strip().upper() == "TRFC":
            return "Receipt"
        return "Payment"

    @staticmethod
    def _is_cell_format_row(row: list) -> bool:
        if len(row) < 9:
            return False
        vals = [str(v or "").strip().lower() for v in row]
        tab_hits = 0
        for token in vals:
            if token in {"tab", r"\{tab}"}:
                tab_hits += 1
        has_code = any(token in {"trfd", "trfc"} for token in vals)
        has_tail_macro = any(token in {"*s", "*sv", "*dn", "*nx"} for token in vals)
        return tab_hits >= 3 and has_code and has_tail_macro

    @staticmethod
    def _is_control_token(token: str) -> bool:
        low = token.strip().lower()
        if not low:
            return True
        if low in {
            "tab", r"\{tab}", "dn", "down", "up", "left", "right", "enter", "esc", "escape",
            r"\r", r"\n",
            "*s", "*sv", "*sp", "*dn", "*nx", "*nr", "*pv", "*up", "*cl", "*nb",
        }:
            return True
        if low.startswith("*"):
            return True
        if low.startswith("\\{") and low.endswith("}"):
            return True
        return False

    @classmethod
    def _extract_table_fields_from_cell_row(cls, row: list) -> dict | None:
        if not cls._is_cell_format_row(row):
            return None
        vals = [cls._clean_cell_text(v) for v in row]
        lower_vals = [v.lower() for v in vals]

        code_idx = -1
        for idx, token in enumerate(lower_vals):
            if token in {"trfd", "trfc"}:
                code_idx = idx
                break
        if code_idx < 0:
            return None

        code = vals[code_idx].upper()
        type_val = "Receipt" if code == "TRFC" else "Payment"

        # Receipt hints may appear before code in macro rows.
        prefix = lower_vals[:code_idx]
        if any(tok in {"*dn", "dn", "down", r"\r", "r", "receipt"} for tok in prefix):
            type_val = "Receipt"

        payload = []
        for token in vals[code_idx + 1:]:
            if cls._is_control_token(token):
                continue
            payload.append(token)

        if not payload:
            return None

        number = payload[0]
        amount = ""
        for p in reversed(payload):
            if p == number:
                continue
            amount = p
            break
        if not amount and len(payload) > 1:
            amount = payload[-1]

        t_date, v_date = cls._extract_dates_from_payload(payload, amount)

        if not any([code, number, t_date, v_date, amount]):
            return None
        return {
            "type": type_val,
            "code": code,
            "number": number,
            "transaction_date": t_date,
            "value_date": v_date,
            "amount": amount,
        }

    @classmethod
    def _extract_table_fields_from_table_row(cls, row: list, col_map: dict) -> dict | None:
        def cell(key: str) -> str:
            idx = col_map.get(key)
            if idx is None or idx < 0 or idx >= len(row):
                return ""
            return cls._clean_cell_text(row[idx])

        row_vals = [cls._clean_cell_text(v) for v in row]

        def normalize_data_token(token: str) -> str:
            if cls._is_control_token(token):
                return ""
            return token

        def normalize_date_token(token: str) -> str:
            token = normalize_data_token(token)
            return cls._normalize_date_text(token) if token else ""

        def normalize_amount_token(token: str) -> str:
            token = normalize_data_token(token)
            return token if token and AMOUNT_TEXT_RE.fullmatch(token) else ""

        raw_code = normalize_data_token(cell("code"))
        code_idx = col_map.get("code", -1)
        if not raw_code:
            for idx, token in enumerate(row_vals):
                low = token.lower()
                if low in {"trfd", "trfc"}:
                    raw_code = token
                    code_idx = idx
                    break

        raw_number = normalize_data_token(cell("number"))
        raw_tdate = normalize_date_token(cell("transaction_date"))
        raw_vdate = normalize_date_token(cell("value_date"))
        raw_amount = normalize_amount_token(cell("amount"))

        # Fallback from row payload after code when mapped columns are noisy/misaligned.
        payload = []
        if code_idx is not None and isinstance(code_idx, int) and code_idx >= 0:
            for token in row_vals[code_idx + 1:]:
                if cls._is_control_token(token):
                    continue
                payload.append(token)

        if not raw_number and payload:
            raw_number = payload[0]

        date_candidates = [cls._normalize_date_text(p) for p in payload]
        date_candidates = [p for p in date_candidates if p and DATE_TEXT_RE.fullmatch(p)]
        if not raw_tdate and date_candidates:
            raw_tdate = date_candidates[0]
        if not raw_vdate:
            if len(date_candidates) > 1:
                raw_vdate = date_candidates[1]
            elif raw_tdate:
                raw_vdate = raw_tdate
        if not raw_tdate or not raw_vdate:
            fallback_t, fallback_v = cls._extract_dates_from_payload(["", *payload], raw_amount)
            if not raw_tdate:
                raw_tdate = fallback_t
            if not raw_vdate:
                raw_vdate = fallback_v

        if not raw_amount:
            for p in reversed(payload):
                if p in {raw_number, raw_tdate, raw_vdate}:
                    continue
                if AMOUNT_TEXT_RE.fullmatch(p):
                    raw_amount = p
                    break
            if not raw_amount and payload:
                candidate = payload[-1]
                if candidate not in {raw_number, raw_tdate, raw_vdate} and AMOUNT_TEXT_RE.fullmatch(candidate):
                    raw_amount = candidate

        values = {
            "type": normalize_data_token(cell("type")),
            "code": raw_code,
            "number": raw_number,
            "transaction_date": raw_tdate,
            "value_date": raw_vdate,
            "amount": raw_amount,
        }
        if not any(values.values()):
            return None
        return values

    @classmethod
    def _normalize_table_fields(cls, fields: dict) -> dict:
        tx_type = cls._normalize_table_type(fields.get("type", ""), fields.get("code", ""))
        code = cls._clean_cell_text(fields.get("code", "")).upper()
        if not code:
            code = "TRFC" if tx_type == "Receipt" else "TRFD"
        return {
            "type": tx_type,
            "code": code,
            "number": cls._clean_cell_text(fields.get("number", "")),
            "transaction_date": cls._clean_cell_text(fields.get("transaction_date", "")),
            "value_date": cls._clean_cell_text(fields.get("value_date", "")),
            "amount": cls._clean_cell_text(fields.get("amount", "")),
        }

    def _replace_sheet_data(self, rows: list[list[str]], *, key_columns: set | None, has_header_row: bool):
        max_cols = max((len(r) for r in rows), default=0)
        target_rows = max(1, len(rows) + 10)
        target_cols = max(1, max_cols + 5)

        self.spreadsheet._begin_history_action()
        self.spreadsheet.setUpdatesEnabled(False)
        self.spreadsheet.blockSignals(True)
        try:
            self.spreadsheet.clearContents()
            self.spreadsheet.setRowCount(target_rows)
            self.spreadsheet.setColumnCount(target_cols)
            self.spreadsheet._update_headers()

            for r, row in enumerate(rows):
                for c, value in enumerate(row):
                    text = str(value) if value is not None else ""
                    if text == "":
                        continue
                    self.spreadsheet.setItem(r, c, QTableWidgetItem(text))
        finally:
            self.spreadsheet.blockSignals(False)
            self.spreadsheet.setUpdatesEnabled(True)
            self.spreadsheet._end_history_action()

        self.spreadsheet._rebuild_cell_cache()
        if key_columns is not None:
            self.spreadsheet.set_key_columns(set(key_columns))
        self.spreadsheet._refresh_highlighting()
        self._has_header_row = bool(has_header_row)
        self.spreadsheet.data_changed.emit()

    def _normalize_sheet_cell_macros(self) -> tuple[bool, int]:
        """Normalize macro tokens across non-empty cells. Returns (has_content, changed_count)."""
        grid_data = self.spreadsheet.get_grid_data()
        if not grid_data:
            return False, 0

        targets: list[QTableWidgetItem] = []
        max_row = len(grid_data) - 1
        max_col = max((len(r) for r in grid_data), default=0) - 1
        if max_row < 0 or max_col < 0:
            return False, 0

        for row in range(max_row + 1):
            for col in range(max_col + 1):
                item = self.spreadsheet.item(row, col)
                if item and item.text().strip():
                    targets.append(item)

        if not targets:
            return False, 0

        changed = 0
        self.spreadsheet._begin_history_action()
        self.spreadsheet.setUpdatesEnabled(False)
        self.spreadsheet.blockSignals(True)
        try:
            for item in targets:
                before = item.text()
                after = self._normalize_macro_cell_value(before)
                if after != before:
                    item.setText(after)
                    changed += 1
        finally:
            self.spreadsheet.blockSignals(False)
            self.spreadsheet.setUpdatesEnabled(True)
            self.spreadsheet._end_history_action()

        if changed > 0:
            self.spreadsheet._rebuild_cell_cache()
            self.spreadsheet._refresh_highlighting()
        return True, changed

    @staticmethod
    def _is_dn_navigation_token(text: str) -> bool:
        token = str(text or "").strip().lower()
        return token in {"*dn", "dn", "down", r"\{down}"}

    @staticmethod
    def _is_receipt_token(text: str) -> bool:
        token = str(text or "").strip().lower()
        return token in {r"\r", "receipt", "r"}

    def _remove_dn_before_receipt_tokens(self) -> int:
        """
        Permanently remove *DN/dn directly before receipt tokens and shift row cells left.
        This keeps data contiguous so no empty gap is left after deletion.
        """
        grid_data = self.spreadsheet.get_grid_data()
        if not grid_data:
            return 0

        row_count = len(grid_data)
        max_cols = max((len(r) for r in grid_data), default=0)
        if row_count <= 0 or max_cols <= 0:
            return 0

        removed = 0
        self.spreadsheet._begin_history_action()
        self.spreadsheet.setUpdatesEnabled(False)
        self.spreadsheet.blockSignals(True)
        try:
            for row in range(row_count):
                row_values: list[str] = []
                for col in range(max_cols):
                    item = self.spreadsheet.item(row, col)
                    row_values.append(item.text() if item else "")

                changed_row = False
                col = 0
                while col < max_cols:
                    current = row_values[col]
                    if not self._is_dn_navigation_token(current):
                        col += 1
                        continue

                    next_col = col + 1
                    while next_col < max_cols and not str(row_values[next_col] or "").strip():
                        next_col += 1

                    if next_col < max_cols and self._is_receipt_token(row_values[next_col]):
                        for shift_col in range(col, max_cols - 1):
                            row_values[shift_col] = row_values[shift_col + 1]
                        row_values[max_cols - 1] = ""
                        removed += 1
                        changed_row = True
                        # Re-check same index in case multiple *DN tokens appear in sequence.
                        continue

                    col += 1

                if not changed_row:
                    continue

                for col in range(max_cols):
                    text = row_values[col]
                    item = self.spreadsheet.item(row, col)
                    if text:
                        if item is None:
                            self.spreadsheet.setItem(row, col, QTableWidgetItem(text))
                        elif item.text() != text:
                            item.setText(text)
                    elif item is not None and item.text():
                        item.setText("")
        finally:
            self.spreadsheet.blockSignals(False)
            self.spreadsheet.setUpdatesEnabled(True)
            self.spreadsheet._end_history_action()

        if removed > 0:
            self.spreadsheet._rebuild_cell_cache()
            self.spreadsheet._refresh_highlighting()
        return removed

    def _convert_to_cell_format(self):
        """Normalize keystroke tokens in-place for cell workflows (no column remapping)."""
        has_content, changed = self._normalize_sheet_cell_macros()
        if not has_content:
            self.status_label.setText("To Cell: nothing to normalize.")
            return

        removed_dn = self._remove_dn_before_receipt_tokens()

        if changed > 0 or removed_dn > 0:
            self.status_label.setText(
                f"To Cell: normalized {changed} cell(s), removed {removed_dn} *DN token(s) before \\r."
            )
        else:
            self.status_label.setText("To Cell: no changes needed.")

    def _convert_to_table_format(self):
        """Convert legacy cell-format rows into table-style rows for per-row workflows."""
        try:
            has_content, normalized = self._normalize_sheet_cell_macros()
            if not has_content:
                self.status_label.setText("Convert to Table: nothing to convert.")
                return

            grid_data = self.spreadsheet.get_grid_data()
            col_map, has_header = self._detect_table_col_map(grid_data)
            start_row = 1 if has_header else 0
            table_rows: list[list[str]] = []
            line_no = 1

            for row in grid_data[start_row:]:
                if not any(self._clean_cell_text(v) for v in row):
                    continue

                fields = self._extract_table_fields_from_cell_row(row)
                if fields is None:
                    fields = self._extract_table_fields_from_table_row(row, col_map)
                if not fields:
                    continue

                norm = self._normalize_table_fields(fields)
                if not any(
                    [
                        norm["number"],
                        norm["transaction_date"],
                        norm["value_date"],
                        norm["amount"],
                        norm["code"],
                    ]
                ):
                    continue

                table_rows.append(
                    [
                        "tab",
                        norm["type"],
                        norm["code"],
                        norm["number"],
                        norm["transaction_date"],
                        norm["value_date"],
                        norm["amount"],
                    ]
                )
                line_no += 1

            if line_no == 1:
                self.status_label.setText("Convert to Table: no compatible rows found.")
                return

            self._replace_sheet_data(table_rows, key_columns=set(), has_header_row=False)
            if normalized > 0:
                self.status_label.setText(
                    f"Converted {line_no - 1} row(s) to Table Format (normalized {normalized} cell(s) first)."
                )
            else:
                self.status_label.setText(f"Converted {line_no - 1} row(s) to Table Format.")
        except Exception as exc:
            self.status_label.setText("Convert to Table failed.")
            QMessageBox.critical(
                self,
                "Convert to Table Error",
                f"Failed to convert the current grid to Table Format:\n{exc}",
            )

    def _convert_macros_to_app_format(self):
        """Backward-compatible alias: Convert to Cell Format."""
        self._convert_to_cell_format()

    # Shortcuts
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def _edit_shortcuts(self):
        dialog = ShortcutsDialog(shortcuts=self.parser.shortcuts, parent=self)
        if dialog.exec():
            self.parser.shortcuts = dialog.get_shortcuts()
            self.status_label.setText("Shortcuts updated")

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    # Templates
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def _load_template(self, template: dict):
        reply = QMessageBox.question(
            self, "Load Template",
            f"Load template: {template['name']}?\nThis will replace current data.",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        self.spreadsheet.clear_all()
        headers = template.get("headers", [])
        sample_data = template.get("sample_data", [])
        key_cols = set(template.get("key_columns", []))

        self.spreadsheet.blockSignals(True)
        for c, h in enumerate(headers):
            item = QTableWidgetItem(h)
            item.setFont(QFont("Segoe UI", 10, QFont.Bold))
            self.spreadsheet.setItem(0, c, item)

        for r, row_data in enumerate(sample_data):
            for c, val in enumerate(row_data):
                item = QTableWidgetItem(str(val))
                self.spreadsheet.setItem(r + 1, c, item)

        self.spreadsheet.blockSignals(False)
        self.spreadsheet._rebuild_cell_cache()
        self.spreadsheet.set_key_columns(key_cols)
        self._has_header_row = True  # Mark that row 1 is headers
        self.status_label.setText(
            f"Template loaded: {template['name']}  |  Data starts at row 2 (row 1 = headers)"
        )
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    # Key Columns
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def _toggle_key_column(self):
        col = self.spreadsheet.currentColumn()
        if col < 0:
            self.status_label.setText("Select a column first")
            return
        self.spreadsheet._toggle_key_column(col)
        col_letter = self._col_label(col)
        if col in self.spreadsheet.key_columns:
            self.status_label.setText(f"Column {col_letter} marked as Key Column")
        else:
            self.status_label.setText(f"Column {col_letter} unmarked as Key Column")

    def _clear_key_columns(self):
        self.spreadsheet.key_columns.clear()
        self.spreadsheet._refresh_highlighting()
        self.status_label.setText("All key column markings cleared")

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    # Help Dialogs
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def _show_about(self):
        QMessageBox.about(
            self, f"About {__display_name__}",
            f"<h2>{__display_name__}</h2>"
            f"<p>Version {__version__}</p>"
            "<p>Multipurpose desktop tool for automation, reporting, conversions, "
            "and Oracle/ERP workflow support.</p>"
            "<p>Includes data loading, financial utilities, statement conversion, "
            "and task-specific workflow tools in one application.</p>"
            "<p><b>Purchase this app:</b> USD 1,000.<br>"
            "<b>Discounted offer:</b> USD 900.</p>"
            "<hr>"
        )



    def _show_how_to(self):
        QMessageBox.information(
            self, f"How to Use {__display_name__}",
            "<h3>Quick Start Guide</h3>"
            "<ol>"
            "<li><b>Select Target Window:</b> Choose the target form or application "
            "from the Window dropdown at the top</li>"
            "<li><b>Select Command Group:</b> Choose the command group that matches "
            "your Oracle or ERP screen</li>"
            "<li><b>Prepare Data:</b> Enter transaction data in the grid, "
            "or import from Excel/CSV</li>"
            "<li><b>Add Navigation:</b> Use keystrokes in key columns:<br>"
            "&nbsp;&nbsp;- <code>\\{TAB}</code> = Tab to next field (default)<br>"
            "&nbsp;&nbsp;- <code>*DN</code> = Down one step (dropdown move)<br>"
            "&nbsp;&nbsp;- <code>*S</code> = Save (Ctrl+S)<br>"
            "&nbsp;&nbsp;- <code>*NX</code> = Next row (Down + Home)<br>"
            "&nbsp;&nbsp;- <code>\\r</code> = Type r for Receipt (after <code>*DN</code>)</li>"
            "<li><b>Prepare Format:</b> Use <b>To Table</b> (Ctrl+Shift+T) for Per Row sheets "
            "or <b>To Cell</b> (Ctrl+Shift+M) for legacy macro sheets.</li>"
            "<li><b>Default Standard:</b> In Cell Format sheets use <code>\\{TAB}</code> for Tab. "
            "Plain <code>tab</code> is accepted and To Cell normalizes it to <code>\\{TAB}</code>.</li>"
            "<li><b>Start Loading:</b> Click Start, set row range "
            "(e.g. rows 1 to 500), and click Start</li>"
            "</ol>"
            "<p><b>Loading Modes:</b></p>"
            "<ul>"
            "<li><b>Per Cell:</b> Loads one cell at a time (step-by-step)</li>"
            "<li><b>Row Range:</b> Loads rows 1 to 500 at once (batch)</li>"
            "</ul>"
        )

    def _show_keystrokes(self):
        QMessageBox.information(
            self, "Keystroke Reference",
            "<h3>NT_DL Keystroke Reference</h3>"
            "<table border='1' cellpadding='4'>"
            "<tr><th>Default Standard</th><th>Action</th></tr>"
            "<tr><td><code>\\{TAB}</code></td><td>Tab key (default)</td></tr>"
            "<tr><td><code>enter</code></td><td>Enter key</td></tr>"
            "<tr><td><code>dn</code></td><td>Down arrow</td></tr>"
            "<tr><td><code>up</code></td><td>Up arrow</td></tr>"
            "<tr><td><code>left</code></td><td>Left arrow</td></tr>"
            "<tr><td><code>right</code></td><td>Right arrow</td></tr>"
            "<tr><td><code>\\r</code></td><td>Type r directly (dropdown type-ahead)</td></tr>"
            "<tr><td><code>\\{TAB 5}</code></td><td>Tab 5 times (advanced)</td></tr>"
            "<tr><td><code>\\%key</code></td><td>Alt + key (advanced)</td></tr>"
            "<tr><td><code>\\^key</code></td><td>Ctrl + key (advanced)</td></tr>"
            "<tr><td><code>\\+key</code></td><td>Shift + key (advanced)</td></tr>"
            "<tr><td><code>*MC(x,y)</code></td><td>Mouse click</td></tr>"
            "</table>"
            "<p><b>Default for this app:</b> use standard format in sheets. "
            "<code>*S</code> is the default Save command for IFMIS. "
            "<code>*SV</code> is an alias for Save. "
            "If you need right then tab, use two steps: <code>right</code> then <code>\\{TAB}</code>. "
            "Plain <code>tab</code> is compatibility input and To Cell normalizes it to <code>\\{TAB}</code>.</p>"
            "<h4>Shortcuts</h4>"
            "<table border='1' cellpadding='4'>"
            "<tr><th>Shortcut</th><th>Action</th></tr>"
            "<tr><td><code>*SP</code></td><td>Save & Proceed</td></tr>"
            "<tr><td><code>*S</code> (default), <code>*SV</code> (alias)</td><td>Save</td></tr>"
            "<tr><td><code>*DN</code></td><td>Down one step</td></tr>"
            "<tr><td><code>*NR</code></td><td>New Record</td></tr>"
            "<tr><td><code>*NX</code></td><td>Next row (Down + Home)</td></tr>"
            "<tr><td><code>*QR</code></td><td>Enter Query</td></tr>"
            "<tr><td><code>*EQ</code></td><td>Execute Query</td></tr>"
            "</table>"
        )

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    # Styling
    # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€








    def _apply_styles(self):
        from kdl.styles import main_window_qss
        from kdl.config_store import get_dark_mode
        dark = get_dark_mode()
        self.setStyleSheet(main_window_qss(dark=dark))
        self._apply_toolbar_action_colors()

    def _toggle_dark_mode(self):
        from kdl.config_store import get_dark_mode, set_dark_mode
        from kdl.spreadsheet_widget import apply_spreadsheet_theme
        dark = not get_dark_mode()
        set_dark_mode(dark)
        if hasattr(self, 'dark_mode_action'):
            self.dark_mode_action.setChecked(dark)
        self._apply_styles()
        apply_spreadsheet_theme(dark)
        self.spreadsheet.viewport().update()
        self.spreadsheet.update()
