"""
KDL Global Style System
Blue enterprise theme with light-blue chrome and larger text.
"""

import os
import sys

# ── Color Palette ──────────────────────────────────────────
# Backgrounds
BG_BASE = "#F4F9FF"            # App background
BG_SURFACE = "#FFFFFF"         # Cards, panels, dialogs — clean white
BG_ELEVATED = "#F8FCFF"        # Raised inputs, group boxes
BG_HOVER = "#EAF4FE"           # Hover state
BG_ACTIVE = "#DCEEFE"          # Pressed / active

# Navy header/toolbar
NAVY = "#69B4E8"               # Menu bar, toolbar background
NAVY_LIGHT = "#82C4F0"         # Toolbar hover
NAVY_DARK = "#4F9FD4"          # Toolbar pressed

# Borders
BORDER_SUBTLE = "#D7E7F6"      # Light borders (grid lines)
BORDER_DEFAULT = "#B6D3EE"     # Default borders
BORDER_STRONG = "#8FBDE4"      # Focus/hover borders

# Text
TEXT_PRIMARY = "#1B3550"        # Primary text
TEXT_SECONDARY = "#4E6E8F"      # Muted text
TEXT_MUTED = "#7C95AF"          # Disabled / hint text
TEXT_ON_NAVY = "#FFFFFF"        # Text on navy backgrounds
TEXT_ON_ACCENT = "#FFFFFF"

# Accent (Light blue)
ACCENT = "#69B4E8"
ACCENT_HOVER = "#82C4F0"
ACCENT_PRESSED = "#5FADE3"
ACCENT_MUTED = "rgba(105, 180, 232, 0.18)"
ACCENT_LIGHT = "#E6F3FD"       # Light blue tint

# Semantic
GREEN = "#16A34A"
GREEN_BG = "#EAF8EF"
RED = "#DC2626"
RED_BG = "#FDECEA"
AMBER = "#D97706"
AMBER_BG = "#FFF4E5"


def _asset_path(filename: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    return os.path.join(base, "kdl", "assets", filename).replace("\\", "/")


def _arrow_rule() -> str:
    arrow = _asset_path("arrow_down.svg")
    if os.path.exists(arrow):
        return f'QComboBox::down-arrow {{ image: url("{arrow}"); width: 12px; height: 12px; }}'
    return "QComboBox::down-arrow { width: 12px; height: 12px; }"


# ── Main Window QSS ───────────────────────────────────────
def main_window_qss() -> str:
    return f"""
        /* ─── Base ─── */
        QMainWindow {{
            background-color: {BG_BASE};
            font-family: "Segoe UI", "Helvetica Neue", sans-serif;
        }}
        QWidget {{
            color: {TEXT_PRIMARY};
            font-family: "Segoe UI", "Helvetica Neue", sans-serif;
            font-size: 15px;
        }}

        /* ─── Menu Bar (Light Blue Header) ─── */
        QMenuBar {{
            background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                              stop:0 #72BBE9, stop:1 #5BABE3);
            border-bottom: 1px solid #4A9DD0;
            padding: 2px 8px;
            font-size: 18px;
            font-weight: 500;
            min-height: 38px;
        }}
        QMenuBar::item {{
            padding: 8px 14px;
            border-radius: 4px;
            background: transparent;
            color: #FFFFFF;
            font-size: 18px;
        }}
        QMenuBar::item:selected {{
            background-color: rgba(255, 255, 255, 0.20);
            color: #FFFFFF;
        }}
        QMenuBar::item:pressed {{
            background-color: rgba(0, 0, 0, 0.10);
            color: #FFFFFF;
        }}

        /* ─── Dropdown Menus ─── */
        QMenu {{
            background-color: {BG_SURFACE};
            border: 1px solid {BORDER_DEFAULT};
            border-radius: 8px;
            padding: 6px 0;
        }}
        QMenu::item {{
            padding: 8px 32px 8px 20px;
            font-size: 16px;
            color: {TEXT_PRIMARY};
        }}
        QMenu::item:selected {{
            background-color: {ACCENT_LIGHT};
            color: {ACCENT};
        }}
        QMenu::separator {{
            height: 1px;
            background-color: {BORDER_SUBTLE};
            margin: 4px 10px;
        }}

        /* ─── Toolbar (Light Blue Header) ─── */
        QToolBar {{
            background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                              stop:0 #6DB8E8, stop:1 #58A8DE);
            border-bottom: 1px solid #4A9DD0;
            padding: 6px 12px;
            spacing: 4px;
        }}
        QToolBar::separator {{
            width: 1px;
            background-color: rgba(255, 255, 255, 0.30);
            margin: 6px 8px;
        }}
        QToolBar QToolButton {{
            padding: 6px;
            border-radius: 7px;
            border: 1px solid transparent;
            background-color: transparent;
            color: #FFFFFF;
            font-weight: 600;
            min-width: 40px;
            min-height: 36px;
        }}
        QToolBar QToolButton:hover {{
            background-color: rgba(255, 255, 255, 0.22);
            border: 1px solid rgba(255, 255, 255, 0.30);
        }}
        QToolBar QToolButton:pressed {{
            background-color: rgba(0, 0, 0, 0.10);
        }}
        QToolBar QToolButton:disabled {{
            opacity: 0.35;
        }}

        /* ─── Inputs ─── */
        QComboBox, QLineEdit {{
            padding: 6px 10px;
            border: 1px solid {BORDER_DEFAULT};
            border-radius: 5px;
            background-color: {BG_SURFACE};
            color: {TEXT_PRIMARY};
            font-size: 16px;
            min-height: 32px;
            selection-background-color: {ACCENT_LIGHT};
            selection-color: {TEXT_PRIMARY};
        }}
        QComboBox:hover, QLineEdit:hover {{
            border-color: {BORDER_STRONG};
        }}
        QComboBox:focus, QLineEdit:focus {{
            border: 2px solid {ACCENT};
            padding: 5px 9px;
        }}
        QComboBox::drop-down {{
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 28px;
            border-left: 1px solid {BORDER_DEFAULT};
            border-top-right-radius: 4px;
            border-bottom-right-radius: 4px;
            background-color: {BG_ELEVATED};
        }}
        QComboBox::drop-down:hover {{
            background-color: {BG_HOVER};
        }}
        {_arrow_rule()}
        QComboBox QAbstractItemView {{
            border: 1px solid {BORDER_DEFAULT};
            border-radius: 6px;
            background: {BG_SURFACE};
            color: {TEXT_PRIMARY};
            selection-background-color: {ACCENT_LIGHT};
            selection-color: {ACCENT};
            font-size: 16px;
            outline: 0;
        }}
        QComboBox QAbstractItemView::item {{
            min-height: 30px;
            padding: 6px 10px;
        }}

        /* ─── Buttons ─── */
        QPushButton {{
            background-color: {BG_SURFACE};
            border: 1px solid {BORDER_DEFAULT};
            border-radius: 5px;
            padding: 6px 16px;
            font-size: 15px;
            font-weight: 500;
            color: {TEXT_PRIMARY};
        }}
        QPushButton:hover {{
            background-color: {BG_HOVER};
            border-color: {BORDER_STRONG};
        }}
        QPushButton:pressed {{
            background-color: {BG_ACTIVE};
        }}
        QPushButton:disabled {{
            background-color: {BG_ELEVATED};
            color: {TEXT_MUTED};
            border-color: {BORDER_SUBTLE};
        }}

        /* ─── Status Bar ─── */
        QStatusBar {{
            background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                              stop:0 #5BABE3, stop:1 #4FA0D8);
            border-top: 1px solid #4A9DD0;
            font-size: 18px;
            color: #FFFFFF;
            min-height: 34px;
            padding: 2px 8px;
        }}
        QProgressBar {{
            border: 1px solid {BORDER_DEFAULT};
            border-radius: 4px;
            background-color: {BG_ELEVATED};
            text-align: center;
            height: 18px;
            font-size: 15px;
            color: {TEXT_PRIMARY};
            font-weight: 600;
        }}
        QProgressBar::chunk {{
            background-color: {ACCENT};
            border-radius: 3px;
        }}

        /* ─── Table / Spreadsheet ─── */
        QTableWidget {{
            gridline-color: {BORDER_SUBTLE};
            selection-background-color: #FFE082;
            selection-color: {TEXT_PRIMARY};
            border: 1px solid {BORDER_DEFAULT};
            border-radius: 5px;
            background-color: {BG_SURFACE};
            font-family: "Consolas", "Cascadia Code", monospace;
            font-size: 15px;
            color: {TEXT_PRIMARY};
        }}
        QTableWidget::item:selected {{
            background-color: #FFE082;
            color: {TEXT_PRIMARY};
        }}
        QHeaderView::section {{
            background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                              stop:0 #F1F7FD, stop:1 #E3EEF9);
            border: none;
            border-right: 1px solid {BORDER_SUBTLE};
            border-bottom: 1px solid {BORDER_DEFAULT};
            padding: 6px 4px;
            font-weight: 600;
            font-size: 15px;
            color: {TEXT_SECONDARY};
        }}
        QTableCornerButton::section {{
            background-color: {BG_ELEVATED};
            border: none;
            border-right: 1px solid {BORDER_SUBTLE};
            border-bottom: 1px solid {BORDER_DEFAULT};
        }}

        /* ─── Scrollbars ─── */
        QScrollBar:vertical {{
            background: {BG_BASE};
            width: 10px;
            margin: 0;
        }}
        QScrollBar::handle:vertical {{
            background: {BORDER_DEFAULT};
            border-radius: 5px;
            min-height: 30px;
        }}
        QScrollBar::handle:vertical:hover {{
            background: {BORDER_STRONG};
        }}
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
            height: 0;
        }}
        QScrollBar:horizontal {{
            background: {BG_BASE};
            height: 10px;
            margin: 0;
        }}
        QScrollBar::handle:horizontal {{
            background: {BORDER_DEFAULT};
            border-radius: 5px;
            min-width: 30px;
        }}
        QScrollBar::handle:horizontal:hover {{
            background: {BORDER_STRONG};
        }}
        QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
            width: 0;
        }}

        /* ─── Group Boxes ─── */
        QGroupBox {{
            border: 1px solid {BORDER_DEFAULT};
            border-radius: 8px;
            margin-top: 16px;
            padding-top: 18px;
            font-weight: 600;
            color: {TEXT_PRIMARY};
            background-color: {BG_SURFACE};
        }}
        QGroupBox::title {{
            subcontrol-origin: margin;
            left: 14px;
            padding: 0 8px;
            color: {ACCENT};
            font-size: 14px;
            font-weight: 700;
        }}

        /* ─── Checkboxes & Radio Buttons ─── */
        QCheckBox, QRadioButton {{
            font-size: 14px;
            spacing: 8px;
            color: {TEXT_PRIMARY};
        }}
        QCheckBox::indicator, QRadioButton::indicator {{
            width: 16px;
            height: 16px;
            border: 2px solid {BORDER_STRONG};
            background-color: {BG_SURFACE};
        }}
        QCheckBox::indicator {{
            border-radius: 4px;
        }}
        QRadioButton::indicator {{
            border-radius: 9px;
        }}
        QCheckBox::indicator:checked, QRadioButton::indicator:checked {{
            background-color: {ACCENT};
            border-color: {ACCENT};
        }}
        QCheckBox::indicator:hover, QRadioButton::indicator:hover {{
            border-color: {ACCENT};
        }}

        /* ─── Tooltips ─── */
        QToolTip {{
            background-color: {NAVY};
            color: {TEXT_ON_NAVY};
            border: 1px solid {NAVY_DARK};
            border-radius: 5px;
            padding: 6px 10px;
            font-size: 14px;
        }}

        /* ─── Tabs ─── */
        QTabWidget::pane {{
            border: 1px solid {BORDER_DEFAULT};
            border-radius: 6px;
            background-color: {BG_SURFACE};
        }}
        QTabBar::tab {{
            background-color: {BG_ELEVATED};
            border: 1px solid {BORDER_SUBTLE};
            padding: 8px 16px;
            margin-right: 2px;
            border-top-left-radius: 5px;
            border-top-right-radius: 5px;
            color: {TEXT_SECONDARY};
        }}
        QTabBar::tab:selected {{
            background-color: {BG_SURFACE};
            color: {ACCENT};
            border-bottom-color: {BG_SURFACE};
            font-weight: 600;
        }}
        QTabBar::tab:hover {{
            background-color: {BG_HOVER};
            color: {TEXT_PRIMARY};
        }}

        /* ─── Frames ─── */
        QFrame[frameShape="4"] {{
            color: {BORDER_SUBTLE};
        }}
        QFrame[frameShape="5"] {{
            color: {BORDER_SUBTLE};
        }}

        /* ─── Labels ─── */
        QLabel {{
            color: {TEXT_PRIMARY};
        }}

        /* ─── PlainTextEdit ─── */
        QPlainTextEdit {{
            background-color: {BG_SURFACE};
            color: {TEXT_PRIMARY};
            border: 1px solid {BORDER_DEFAULT};
            border-radius: 5px;
            font-family: "Consolas", "Cascadia Code", monospace;
            font-size: 15px;
            padding: 6px;
            selection-background-color: {ACCENT_LIGHT};
        }}

        /* ─── Dialog Buttons ─── */
        QDialogButtonBox QPushButton {{
            min-width: 80px;
            min-height: 30px;
        }}
    """


# ── Accent Button Style ──────────────────────────────────
def accent_button_qss() -> str:
    """Primary action button (Start, Save, OK)."""
    return f"""
        QPushButton {{
            background-color: {ACCENT};
            color: {TEXT_ON_ACCENT};
            border: none;
            border-radius: 5px;
            font-weight: 600;
            font-size: 15px;
            padding: 8px 24px;
        }}
        QPushButton:hover {{
            background-color: {ACCENT_HOVER};
        }}
        QPushButton:pressed {{
            background-color: {ACCENT_PRESSED};
        }}
        QPushButton:disabled {{
            background-color: {BG_HOVER};
            color: {TEXT_MUTED};
        }}
    """


# ── Dialog Base QSS ──────────────────────────────────────
def dialog_qss() -> str:
    """Base style for all dialogs."""
    return f"""
        QDialog {{
            background-color: {BG_BASE};
            color: {TEXT_PRIMARY};
            font-size: 14px;
        }}
        QLabel {{
            color: {TEXT_PRIMARY};
        }}
        QGroupBox {{
            border: 1px solid {BORDER_DEFAULT};
            border-radius: 8px;
            margin-top: 16px;
            padding-top: 18px;
            font-weight: 600;
            color: {TEXT_PRIMARY};
            background-color: {BG_SURFACE};
        }}
        QGroupBox::title {{
            subcontrol-origin: margin;
            left: 14px;
            padding: 0 8px;
            color: {ACCENT};
            font-size: 14px;
            font-weight: 700;
        }}
        QComboBox, QLineEdit {{
            padding: 6px 10px;
            border: 1px solid {BORDER_DEFAULT};
            border-radius: 5px;
            background-color: {BG_SURFACE};
            color: {TEXT_PRIMARY};
            font-size: 14px;
            min-height: 24px;
            selection-background-color: {ACCENT_LIGHT};
        }}
        QComboBox:hover, QLineEdit:hover {{
            border-color: {BORDER_STRONG};
        }}
        QComboBox:focus, QLineEdit:focus {{
            border: 2px solid {ACCENT};
            padding: 5px 9px;
        }}
        QComboBox::drop-down {{
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 26px;
            border-left: 1px solid {BORDER_DEFAULT};
            border-top-right-radius: 4px;
            border-bottom-right-radius: 4px;
            background-color: {BG_ELEVATED};
        }}
        QComboBox QAbstractItemView {{
            background: {BG_SURFACE};
            color: {TEXT_PRIMARY};
            border: 1px solid {BORDER_DEFAULT};
            selection-background-color: {ACCENT_LIGHT};
            selection-color: {ACCENT};
        }}
        QPushButton {{
            background-color: {BG_SURFACE};
            border: 1px solid {BORDER_DEFAULT};
            border-radius: 5px;
            padding: 6px 16px;
            font-size: 14px;
            font-weight: 500;
            color: {TEXT_PRIMARY};
        }}
        QPushButton:hover {{
            background-color: {BG_HOVER};
            border-color: {BORDER_STRONG};
        }}
        QPushButton:pressed {{
            background-color: {BG_ACTIVE};
        }}
        QCheckBox, QRadioButton {{
            font-size: 14px;
            spacing: 8px;
            color: {TEXT_PRIMARY};
        }}
        QCheckBox::indicator, QRadioButton::indicator {{
            width: 16px;
            height: 16px;
            border: 2px solid {BORDER_STRONG};
            background-color: {BG_SURFACE};
        }}
        QCheckBox::indicator {{ border-radius: 4px; }}
        QRadioButton::indicator {{ border-radius: 9px; }}
        QCheckBox::indicator:checked, QRadioButton::indicator:checked {{
            background-color: {ACCENT};
            border-color: {ACCENT};
        }}
        QTableWidget {{
            gridline-color: {BORDER_SUBTLE};
            background-color: {BG_SURFACE};
            color: {TEXT_PRIMARY};
            border: 1px solid {BORDER_DEFAULT};
            border-radius: 5px;
            selection-background-color: {ACCENT_LIGHT};
            selection-color: {TEXT_PRIMARY};
            font-family: "Consolas", monospace;
            font-size: 14px;
        }}
        QHeaderView::section {{
            background-color: {BG_ELEVATED};
            color: {TEXT_SECONDARY};
            border: none;
            border-right: 1px solid {BORDER_SUBTLE};
            border-bottom: 1px solid {BORDER_DEFAULT};
            padding: 6px 4px;
            font-weight: 600;
            font-size: 14px;
        }}
        QFrame[frameShape="4"] {{
            color: {BORDER_SUBTLE};
        }}
        QToolTip {{
            background-color: {NAVY};
            color: {TEXT_ON_NAVY};
            border: 1px solid {NAVY_DARK};
            border-radius: 5px;
            padding: 6px 10px;
            font-size: 14px;
        }}
        QPlainTextEdit {{
            background-color: {BG_SURFACE};
            color: {TEXT_PRIMARY};
            border: 1px solid {BORDER_DEFAULT};
            border-radius: 5px;
            padding: 6px;
            font-family: "Consolas", monospace;
            font-size: 14px;
        }}
        QScrollBar:vertical {{
            background: {BG_BASE}; width: 8px;
        }}
        QScrollBar::handle:vertical {{
            background: {BORDER_DEFAULT}; border-radius: 4px; min-height: 30px;
        }}
        QScrollBar::handle:vertical:hover {{
            background: {BORDER_STRONG};
        }}
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
        QScrollBar:horizontal {{
            background: {BG_BASE}; height: 8px;
        }}
        QScrollBar::handle:horizontal {{
            background: {BORDER_DEFAULT}; border-radius: 4px; min-width: 30px;
        }}
        QScrollBar::handle:horizontal:hover {{
            background: {BORDER_STRONG};
        }}
        QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0; }}
    """


# ── Load Result Dialog QSS ───────────────────────────────
def load_result_qss(accent: str, panel_bg: str) -> str:
    return f"""
        QDialog {{
            background-color: {BG_BASE};
        }}
        QFrame#ResultPanel {{
            background-color: {panel_bg};
            border: 1px solid {BORDER_DEFAULT};
            border-left: 5px solid {accent};
            border-radius: 8px;
        }}
        QLabel#ResultHeading {{
            font-size: 18px;
            font-weight: 700;
            color: {accent};
        }}
        QLabel#ResultSubheading {{
            font-size: 15px;
            color: {TEXT_SECONDARY};
        }}
        QLabel#ResultSummary {{
            font-size: 15px;
            color: {TEXT_PRIMARY};
            font-weight: 600;
        }}
        QFrame#DetailsBox {{
            background-color: {BG_SURFACE};
            border: 1px solid {BORDER_SUBTLE};
            border-radius: 6px;
        }}
        QLabel#DetailKey {{
            color: {TEXT_SECONDARY};
            font-size: 15px;
            font-weight: 600;
            min-width: 105px;
        }}
        QLabel#DetailVal {{
            color: {TEXT_PRIMARY};
            font-size: 15px;
            font-weight: 600;
        }}
        QPushButton {{
            min-width: 88px;
            min-height: 32px;
            font-weight: 600;
            border-radius: 5px;
            border: 1px solid {BORDER_DEFAULT};
            background-color: {BG_SURFACE};
            color: {TEXT_PRIMARY};
            padding: 4px 14px;
        }}
        QPushButton:hover {{
            background-color: {BG_HOVER};
            border-color: {BORDER_STRONG};
        }}
        QPushButton:pressed {{
            background-color: {BG_ACTIVE};
        }}
    """
