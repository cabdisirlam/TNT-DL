"""
KDL Global Style System
Refined desktop theme with cleaner chrome, softer surfaces, and stronger hierarchy.
"""

import os
import sys

# Light palette
BG_BASE = "#F1F6FA"
BG_SURFACE = "#FFFFFF"
BG_ELEVATED = "#F7FAFD"
BG_HOVER = "#EBF3F9"
BG_ACTIVE = "#DCEAF5"

NAVY = "#1D6B9F"
NAVY_LIGHT = "#2D81B9"
NAVY_DARK = "#14527E"

BORDER_SUBTLE = "#D7E1EB"
BORDER_DEFAULT = "#BDD0E0"
BORDER_STRONG = "#7FA7C6"

TEXT_PRIMARY = "#18324A"
TEXT_SECONDARY = "#58718A"
TEXT_MUTED = "#7B92A8"
TEXT_ON_NAVY = "#FFFFFF"
TEXT_ON_ACCENT = "#FFFFFF"

ACCENT = "#1F7AB8"
ACCENT_HOVER = "#2B8BCF"
ACCENT_PRESSED = "#17689D"
ACCENT_MUTED = "rgba(31, 122, 184, 0.16)"
ACCENT_LIGHT = "#E6F1F8"

GREEN = "#12915A"
GREEN_BG = "#E7F7EF"
RED = "#C83A3A"
RED_BG = "#FBECEC"
AMBER = "#C77A11"
AMBER_BG = "#FFF4E3"

# Dark palette
DARK_BG_BASE = "#1C2330"
DARK_BG_SURFACE = "#232C3B"
DARK_BG_ELEVATED = "#293343"
DARK_BG_HOVER = "#313C50"
DARK_BG_ACTIVE = "#39465F"
DARK_NAVY = "#18304D"
DARK_NAVY_LIGHT = "#21436A"
DARK_NAVY_DARK = "#12243A"
DARK_BORDER_SUBTLE = "#2B3750"
DARK_BORDER_DEFAULT = "#35435F"
DARK_BORDER_STRONG = "#50627F"
DARK_TEXT_PRIMARY = "#D9E5F1"
DARK_TEXT_SECONDARY = "#93A9BD"
DARK_TEXT_MUTED = "#60758D"
DARK_ACCENT = "#4A8CC5"
DARK_ACCENT_HOVER = "#5BA0D8"
DARK_ACCENT_PRESSED = "#3C76AA"
DARK_ACCENT_MUTED = "rgba(74, 140, 197, 0.16)"
DARK_ACCENT_LIGHT = "#1A2C42"


def _palette(dark: bool):
    """Return a dict of color values for the given theme."""
    if dark:
        return dict(
            BG_BASE=DARK_BG_BASE,
            BG_SURFACE=DARK_BG_SURFACE,
            BG_ELEVATED=DARK_BG_ELEVATED,
            BG_HOVER=DARK_BG_HOVER,
            BG_ACTIVE=DARK_BG_ACTIVE,
            NAVY=DARK_NAVY,
            NAVY_LIGHT=DARK_NAVY_LIGHT,
            NAVY_DARK=DARK_NAVY_DARK,
            BORDER_SUBTLE=DARK_BORDER_SUBTLE,
            BORDER_DEFAULT=DARK_BORDER_DEFAULT,
            BORDER_STRONG=DARK_BORDER_STRONG,
            TEXT_PRIMARY=DARK_TEXT_PRIMARY,
            TEXT_SECONDARY=DARK_TEXT_SECONDARY,
            TEXT_MUTED=DARK_TEXT_MUTED,
            TEXT_ON_NAVY=TEXT_ON_NAVY,
            TEXT_ON_ACCENT=TEXT_ON_ACCENT,
            ACCENT=DARK_ACCENT,
            ACCENT_HOVER=DARK_ACCENT_HOVER,
            ACCENT_PRESSED=DARK_ACCENT_PRESSED,
            ACCENT_MUTED=DARK_ACCENT_MUTED,
            ACCENT_LIGHT=DARK_ACCENT_LIGHT,
            GREEN=GREEN,
            GREEN_BG="#0F2B1E",
            RED=RED,
            RED_BG="#301417",
            AMBER=AMBER,
            AMBER_BG="#31230D",
        )
    return dict(
        BG_BASE=BG_BASE,
        BG_SURFACE=BG_SURFACE,
        BG_ELEVATED=BG_ELEVATED,
        BG_HOVER=BG_HOVER,
        BG_ACTIVE=BG_ACTIVE,
        NAVY=NAVY,
        NAVY_LIGHT=NAVY_LIGHT,
        NAVY_DARK=NAVY_DARK,
        BORDER_SUBTLE=BORDER_SUBTLE,
        BORDER_DEFAULT=BORDER_DEFAULT,
        BORDER_STRONG=BORDER_STRONG,
        TEXT_PRIMARY=TEXT_PRIMARY,
        TEXT_SECONDARY=TEXT_SECONDARY,
        TEXT_MUTED=TEXT_MUTED,
        TEXT_ON_NAVY=TEXT_ON_NAVY,
        TEXT_ON_ACCENT=TEXT_ON_ACCENT,
        ACCENT=ACCENT,
        ACCENT_HOVER=ACCENT_HOVER,
        ACCENT_PRESSED=ACCENT_PRESSED,
        ACCENT_MUTED=ACCENT_MUTED,
        ACCENT_LIGHT=ACCENT_LIGHT,
        GREEN=GREEN,
        GREEN_BG=GREEN_BG,
        RED=RED,
        RED_BG=RED_BG,
        AMBER=AMBER,
        AMBER_BG=AMBER_BG,
    )


def _asset_path(filename: str) -> str:
    base = getattr(
        sys,
        "_MEIPASS",
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    )
    return os.path.join(base, "kdl", "assets", filename).replace("\\", "/")


def _arrow_rule() -> str:
    arrow = _asset_path("arrow_down.svg")
    if os.path.exists(arrow):
        return (
            'QComboBox::down-arrow { image: url("'
            + arrow
            + '"); width: 12px; height: 12px; }'
        )
    return "QComboBox::down-arrow { width: 12px; height: 12px; }"


def main_window_qss(dark: bool = False) -> str:
    p = _palette(dark)
    bg_base = p["BG_BASE"]
    bg_surface = p["BG_SURFACE"]
    bg_elevated = p["BG_ELEVATED"]
    bg_hover = p["BG_HOVER"]
    bg_active = p["BG_ACTIVE"]
    navy = p["NAVY"]
    navy_light = p["NAVY_LIGHT"]
    navy_dark = p["NAVY_DARK"]
    border_subtle = p["BORDER_SUBTLE"]
    border_default = p["BORDER_DEFAULT"]
    border_strong = p["BORDER_STRONG"]
    text_primary = p["TEXT_PRIMARY"]
    text_secondary = p["TEXT_SECONDARY"]
    text_muted = p["TEXT_MUTED"]
    accent = p["ACCENT"]
    accent_hover = p["ACCENT_HOVER"]
    accent_pressed = p["ACCENT_PRESSED"]
    accent_light = p["ACCENT_LIGHT"]
    return f"""
        QMainWindow {{
            background-color: {bg_base};
            font-family: "Segoe UI", "Helvetica Neue", sans-serif;
        }}
        QWidget {{
            color: {text_primary};
            font-family: "Segoe UI", "Helvetica Neue", sans-serif;
            font-size: 14px;
        }}

        QMenuBar {{
            background-color: qlineargradient(
                x1:0, y1:0, x2:0, y2:1,
                stop:0 {navy_light},
                stop:1 {navy_dark}
            );
            border-bottom: 1px solid {navy_dark};
            padding: 4px 12px 6px 12px;
            font-size: 16px;
            font-weight: 600;
            min-height: 40px;
        }}
        QMenuBar::item {{
            padding: 9px 14px;
            border-radius: 10px;
            background: transparent;
            color: #FFFFFF;
            font-size: 16px;
        }}
        QMenuBar::item:selected {{
            background-color: rgba(255, 255, 255, 0.18);
        }}
        QMenuBar::item:pressed {{
            background-color: rgba(0, 0, 0, 0.16);
        }}

        QMenu {{
            background-color: {bg_surface};
            border: 1px solid {border_default};
            border-radius: 12px;
            padding: 8px 0;
        }}
        QMenu::item {{
            padding: 9px 34px 9px 18px;
            font-size: 14px;
            color: {text_primary};
        }}
        QMenu::item:selected {{
            background-color: {accent_light};
            color: {text_primary};
        }}
        QMenu::separator {{
            height: 1px;
            background-color: {border_subtle};
            margin: 6px 12px;
        }}

        QToolBar {{
            background-color: qlineargradient(
                x1:0, y1:0, x2:0, y2:1,
                stop:0 {navy_light},
                stop:1 {navy_dark}
            );
            border-bottom: 1px solid {navy_dark};
            padding: 10px 14px 12px 14px;
            spacing: 6px;
        }}
        QToolBar::separator {{
            width: 1px;
            background-color: rgba(255, 255, 255, 0.24);
            margin: 8px 10px;
        }}
        QToolBar QToolButton {{
            padding: 6px;
            border-radius: 12px;
            border: 1px solid transparent;
            background-color: transparent;
            color: #FFFFFF;
            font-weight: 600;
            min-width: 44px;
            min-height: 40px;
        }}
        QToolBar QToolButton:hover {{
            background-color: rgba(255, 255, 255, 0.18);
            border: 1px solid rgba(255, 255, 255, 0.26);
        }}
        QToolBar QToolButton:pressed {{
            background-color: rgba(0, 0, 0, 0.14);
        }}
        QToolBar QToolButton:disabled {{
            opacity: 0.35;
        }}

        QComboBox, QLineEdit {{
            padding: 8px 12px;
            border: 1px solid {border_default};
            border-radius: 10px;
            background-color: {bg_surface};
            color: {text_primary};
            font-size: 14px;
            min-height: 34px;
            selection-background-color: {accent_light};
            selection-color: {text_primary};
        }}
        QComboBox:hover, QLineEdit:hover {{
            border-color: {border_strong};
        }}
        QComboBox:focus, QLineEdit:focus {{
            border: 2px solid {accent};
            padding: 7px 11px;
        }}
        QComboBox::drop-down {{
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 30px;
            border-left: 1px solid {border_default};
            border-top-right-radius: 9px;
            border-bottom-right-radius: 9px;
            background-color: {bg_elevated};
        }}
        QComboBox::drop-down:hover {{
            background-color: {bg_hover};
        }}
        {_arrow_rule()}
        QComboBox QAbstractItemView {{
            border: 1px solid {border_default};
            border-radius: 10px;
            background: {bg_surface};
            color: {text_primary};
            selection-background-color: {accent_light};
            selection-color: {text_primary};
            font-size: 14px;
            outline: 0;
        }}
        QComboBox QAbstractItemView::item {{
            min-height: 32px;
            padding: 7px 10px;
        }}

        QPushButton {{
            background-color: {bg_surface};
            border: 1px solid {border_default};
            border-radius: 10px;
            padding: 8px 18px;
            font-size: 14px;
            font-weight: 600;
            color: {text_primary};
        }}
        QPushButton:hover {{
            background-color: {bg_hover};
            border-color: {border_strong};
        }}
        QPushButton:pressed {{
            background-color: {bg_active};
        }}
        QPushButton:disabled {{
            background-color: {bg_elevated};
            color: {text_muted};
            border-color: {border_subtle};
        }}

        QStatusBar {{
            background-color: qlineargradient(
                x1:0, y1:0, x2:0, y2:1,
                stop:0 {navy},
                stop:1 {navy_dark}
            );
            border-top: 1px solid {navy_dark};
            font-size: 15px;
            color: #FFFFFF;
            min-height: 38px;
            padding: 4px 10px;
        }}
        QProgressBar {{
            border: 1px solid {border_default};
            border-radius: 7px;
            background-color: {bg_elevated};
            text-align: center;
            height: 18px;
            font-size: 13px;
            color: {text_primary};
            font-weight: 600;
        }}
        QProgressBar::chunk {{
            background-color: {accent};
            border-radius: 6px;
        }}

        QTableWidget {{
            gridline-color: {border_subtle};
            selection-background-color: #FFE082;
            selection-color: {text_primary};
            border: 1px solid {border_default};
            border-radius: 12px;
            background-color: {bg_surface};
            font-family: "Consolas", "Cascadia Code", monospace;
            font-size: 14px;
            color: {text_primary};
        }}
        QTableWidget::item:selected {{
            background-color: #FFE082;
            color: {text_primary};
        }}
        QHeaderView::section {{
            background-color: qlineargradient(
                x1:0, y1:0, x2:0, y2:1,
                stop:0 {bg_elevated},
                stop:1 {bg_hover}
            );
            border: none;
            border-right: 1px solid {border_subtle};
            border-bottom: 1px solid {border_default};
            padding: 8px 6px;
            font-weight: 600;
            font-size: 13px;
            color: {text_secondary};
        }}
        QTableCornerButton::section {{
            background-color: {bg_elevated};
            border: none;
            border-right: 1px solid {border_subtle};
            border-bottom: 1px solid {border_default};
        }}

        QScrollBar:vertical {{
            background: {bg_base};
            width: 10px;
            margin: 0;
        }}
        QScrollBar::handle:vertical {{
            background: {border_default};
            border-radius: 5px;
            min-height: 30px;
        }}
        QScrollBar::handle:vertical:hover {{
            background: {border_strong};
        }}
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
            height: 0;
        }}
        QScrollBar:horizontal {{
            background: {bg_base};
            height: 10px;
            margin: 0;
        }}
        QScrollBar::handle:horizontal {{
            background: {border_default};
            border-radius: 5px;
            min-width: 30px;
        }}
        QScrollBar::handle:horizontal:hover {{
            background: {border_strong};
        }}
        QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
            width: 0;
        }}

        QGroupBox {{
            border: 1px solid {border_default};
            border-radius: 14px;
            margin-top: 18px;
            padding-top: 20px;
            font-weight: 600;
            color: {text_primary};
            background-color: {bg_surface};
        }}
        QGroupBox::title {{
            subcontrol-origin: margin;
            left: 16px;
            padding: 0 10px;
            color: {accent_pressed};
            font-size: 13px;
            font-weight: 700;
        }}

        QCheckBox, QRadioButton {{
            font-size: 13px;
            spacing: 10px;
            color: {text_primary};
        }}
        QCheckBox::indicator, QRadioButton::indicator {{
            width: 18px;
            height: 18px;
            border: 2px solid {border_strong};
            background-color: {bg_surface};
        }}
        QCheckBox::indicator {{
            border-radius: 4px;
        }}
        QRadioButton::indicator {{
            border-radius: 9px;
        }}
        QCheckBox::indicator:checked, QRadioButton::indicator:checked {{
            background-color: {accent};
            border-color: {accent};
        }}
        QCheckBox::indicator:hover, QRadioButton::indicator:hover {{
            border-color: {accent};
        }}

        QToolTip {{
            background-color: {navy};
            color: {TEXT_ON_NAVY};
            border: 1px solid {navy_dark};
            border-radius: 8px;
            padding: 6px 10px;
            font-size: 13px;
        }}

        QTabWidget::pane {{
            border: 1px solid {border_default};
            border-radius: 10px;
            background-color: {bg_surface};
        }}
        QTabBar::tab {{
            background-color: {bg_elevated};
            border: 1px solid {border_subtle};
            padding: 9px 16px;
            margin-right: 4px;
            border-top-left-radius: 8px;
            border-top-right-radius: 8px;
            color: {text_secondary};
        }}
        QTabBar::tab:selected {{
            background-color: {bg_surface};
            color: {accent_pressed};
            border-bottom-color: {bg_surface};
            font-weight: 700;
        }}
        QTabBar::tab:hover {{
            background-color: {bg_hover};
            color: {text_primary};
        }}

        QFrame#TopRowsCard, QFrame#FormulaCard {{
            background-color: {bg_surface};
            border: 1px solid {border_default};
            border-radius: 14px;
        }}
        QFrame#TopRowsDivider {{
            color: {border_subtle};
            max-height: 1px;
        }}

        QLabel#ShellFieldLabel, QLabel#ShellFieldLabelStacked {{
            color: {text_secondary};
            font-size: 12px;
            font-weight: 700;
        }}
        QLabel#CellRefPill {{
            background-color: {accent};
            color: {TEXT_ON_ACCENT};
            border-radius: 10px;
            padding: 6px 10px;
            font-family: "Consolas", "Cascadia Code", monospace;
            font-size: 12px;
            font-weight: 700;
        }}
        QComboBox#ShellCombo, QLineEdit#FormulaBar {{
            min-height: 36px;
            font-size: 14px;
            font-weight: 600;
        }}
        QPushButton#ShellRefreshButton {{
            background-color: {accent};
            color: {TEXT_ON_ACCENT};
            border: 1px solid {accent_pressed};
            border-radius: 11px;
            padding: 0 16px;
            min-height: 36px;
            font-size: 14px;
            font-weight: 700;
        }}
        QPushButton#ShellRefreshButton:hover {{
            background-color: {accent_hover};
            border-color: {accent};
        }}
        QPushButton#ShellRefreshButton:pressed {{
            background-color: {accent_pressed};
        }}

        QLabel {{
            color: {text_primary};
        }}
        QPlainTextEdit {{
            background-color: {bg_surface};
            color: {text_primary};
            border: 1px solid {border_default};
            border-radius: 10px;
            font-family: "Consolas", "Cascadia Code", monospace;
            font-size: 14px;
            padding: 8px;
            selection-background-color: {accent_light};
        }}
        QDialogButtonBox QPushButton {{
            min-width: 84px;
            min-height: 34px;
        }}
    """


def accent_button_qss(dark: bool = False) -> str:
    """Primary action button (Start, Save, OK)."""
    p = _palette(dark)
    return f"""
        QPushButton {{
            background-color: {p["ACCENT"]};
            color: {p["TEXT_ON_ACCENT"]};
            border: 1px solid {p["ACCENT_PRESSED"]};
            border-radius: 10px;
            font-weight: 700;
            font-size: 14px;
            padding: 9px 24px;
        }}
        QPushButton:hover {{
            background-color: {p["ACCENT_HOVER"]};
            border-color: {p["ACCENT"]};
        }}
        QPushButton:pressed {{
            background-color: {p["ACCENT_PRESSED"]};
        }}
        QPushButton:disabled {{
            background-color: {p["BG_HOVER"]};
            color: {p["TEXT_MUTED"]};
            border-color: {p["BORDER_SUBTLE"]};
        }}
    """


def themed_button_qss(dark: bool = False) -> str:
    """Secondary button that still uses the active app palette."""
    p = _palette(dark)
    return f"""
        QPushButton {{
            background-color: {p["BG_SURFACE"]};
            color: {p["TEXT_PRIMARY"]};
            border: 1px solid {p["ACCENT"]};
            border-radius: 10px;
            font-weight: 600;
            font-size: 13px;
            padding: 8px 18px;
        }}
        QPushButton:hover {{
            background-color: {p["ACCENT_LIGHT"]};
            border-color: {p["ACCENT_HOVER"]};
        }}
        QPushButton:pressed {{
            background-color: {p["BG_ACTIVE"]};
            border-color: {p["ACCENT_PRESSED"]};
        }}
        QPushButton:disabled {{
            background-color: {p["BG_ELEVATED"]};
            color: {p["TEXT_MUTED"]};
            border-color: {p["BORDER_SUBTLE"]};
        }}
    """


def dialog_qss(dark: bool = False) -> str:
    """Base style for all dialogs."""
    p = _palette(dark)
    bg_base = p["BG_BASE"]
    bg_surface = p["BG_SURFACE"]
    bg_elevated = p["BG_ELEVATED"]
    bg_hover = p["BG_HOVER"]
    bg_active = p["BG_ACTIVE"]
    border_subtle = p["BORDER_SUBTLE"]
    border_default = p["BORDER_DEFAULT"]
    border_strong = p["BORDER_STRONG"]
    text_primary = p["TEXT_PRIMARY"]
    text_secondary = p["TEXT_SECONDARY"]
    text_on_navy = p["TEXT_ON_NAVY"]
    navy = p["NAVY"]
    navy_dark = p["NAVY_DARK"]
    accent = p["ACCENT"]
    accent_pressed = p["ACCENT_PRESSED"]
    accent_light = p["ACCENT_LIGHT"]
    return f"""
        QDialog {{
            background-color: {bg_base};
            color: {text_primary};
            font-size: 14px;
            font-family: "Segoe UI", "Helvetica Neue", sans-serif;
        }}
        QLabel {{
            color: {text_primary};
        }}
        QLabel#DialogIntro {{
            color: {text_secondary};
            font-size: 13px;
            line-height: 1.35;
            padding: 0 2px 4px 2px;
        }}
        QLabel#DialogHint {{
            color: {p["TEXT_MUTED"]};
            font-size: 12px;
            padding: 2px 2px 0 2px;
        }}
        QGroupBox {{
            border: 1px solid {border_default};
            border-radius: 16px;
            margin-top: 16px;
            padding-top: 18px;
            font-weight: 600;
            color: {text_primary};
            background-color: {bg_surface};
        }}
        QGroupBox::title {{
            subcontrol-origin: margin;
            left: 14px;
            padding: 0 9px;
            color: {accent_pressed};
            font-size: 13px;
            font-weight: 700;
        }}
        QComboBox, QLineEdit {{
            padding: 8px 12px;
            border: 1px solid {border_default};
            border-radius: 10px;
            background-color: {bg_surface};
            color: {text_primary};
            font-size: 13px;
            min-height: 34px;
            selection-background-color: {accent_light};
        }}
        QComboBox:hover, QLineEdit:hover {{
            border-color: {border_strong};
        }}
        QComboBox:focus, QLineEdit:focus {{
            border: 2px solid {accent};
            padding: 7px 11px;
        }}
        QComboBox::drop-down {{
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 28px;
            border-left: 1px solid {border_default};
            border-top-right-radius: 9px;
            border-bottom-right-radius: 9px;
            background-color: {bg_elevated};
        }}
        {_arrow_rule()}
        QComboBox QAbstractItemView {{
            background: {bg_surface};
            color: {text_primary};
            border: 1px solid {border_default};
            selection-background-color: {accent_light};
            selection-color: {text_primary};
        }}
        QPushButton {{
            background-color: {bg_surface};
            border: 1px solid {border_default};
            border-radius: 10px;
            padding: 8px 18px;
            font-size: 13px;
            font-weight: 600;
            color: {text_primary};
            min-height: 36px;
            min-width: 96px;
        }}
        QPushButton:hover {{
            background-color: {bg_hover};
            border-color: {border_strong};
        }}
        QPushButton:pressed {{
            background-color: {bg_active};
        }}
        QPushButton:disabled {{
            background-color: {bg_elevated};
            color: {p["TEXT_MUTED"]};
            border-color: {border_subtle};
        }}
        QCheckBox, QRadioButton {{
            font-size: 13px;
            spacing: 10px;
            color: {text_primary};
        }}
        QCheckBox::indicator, QRadioButton::indicator {{
            width: 18px;
            height: 18px;
            border: 2px solid {border_strong};
            background-color: {bg_surface};
        }}
        QCheckBox::indicator {{
            border-radius: 4px;
        }}
        QRadioButton::indicator {{
            border-radius: 9px;
        }}
        QCheckBox::indicator:checked, QRadioButton::indicator:checked {{
            background-color: {accent};
            border-color: {accent};
        }}
        QTableWidget {{
            gridline-color: {border_subtle};
            background-color: {bg_surface};
            color: {text_primary};
            border: 1px solid {border_default};
            border-radius: 10px;
            selection-background-color: {accent_light};
            selection-color: {text_primary};
            font-family: "Consolas", monospace;
            font-size: 13px;
        }}
        QHeaderView::section {{
            background-color: {bg_elevated};
            color: {text_secondary};
            border: none;
            border-right: 1px solid {border_subtle};
            border-bottom: 1px solid {border_default};
            padding: 8px 6px;
            font-weight: 600;
            font-size: 13px;
        }}
        QFrame[frameShape="4"] {{
            color: {border_subtle};
        }}
        QToolTip {{
            background-color: {navy};
            color: {text_on_navy};
            border: 1px solid {navy_dark};
            border-radius: 8px;
            padding: 6px 10px;
            font-size: 13px;
        }}
        QTextEdit, QPlainTextEdit {{
            background-color: {bg_surface};
            color: {text_primary};
            border: 1px solid {border_default};
            border-radius: 12px;
            padding: 10px;
            font-family: "Consolas", monospace;
            font-size: 13px;
        }}
        QTextEdit:focus, QPlainTextEdit:focus {{
            border: 2px solid {accent};
        }}
        QScrollBar:vertical {{
            background: {bg_base};
            width: 10px;
        }}
        QScrollBar::handle:vertical {{
            background: {border_default};
            border-radius: 5px;
            min-height: 30px;
        }}
        QScrollBar::handle:vertical:hover {{
            background: {border_strong};
        }}
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
            height: 0;
        }}
        QScrollBar:horizontal {{
            background: {bg_base};
            height: 10px;
        }}
        QScrollBar::handle:horizontal {{
            background: {border_default};
            border-radius: 5px;
            min-width: 30px;
        }}
        QScrollBar::handle:horizontal:hover {{
            background: {border_strong};
        }}
        QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
            width: 0;
        }}
    """


def load_result_qss(accent: str, panel_bg: str, dark: bool = False) -> str:
    p = _palette(dark)
    bg_base = p["BG_BASE"]
    bg_surface = p["BG_SURFACE"]
    bg_hover = p["BG_HOVER"]
    bg_active = p["BG_ACTIVE"]
    border_subtle = p["BORDER_SUBTLE"]
    border_default = p["BORDER_DEFAULT"]
    border_strong = p["BORDER_STRONG"]
    text_primary = p["TEXT_PRIMARY"]
    text_secondary = p["TEXT_SECONDARY"]
    return f"""
        QDialog {{
            background-color: {bg_base};
        }}
        QFrame#ResultPanel {{
            background-color: {panel_bg};
            border: 1px solid {border_default};
            border-left: 5px solid {accent};
            border-radius: 12px;
        }}
        QLabel#ResultHeading {{
            font-size: 18px;
            font-weight: 700;
            color: {accent};
        }}
        QLabel#ResultSubheading {{
            font-size: 15px;
            color: {text_secondary};
        }}
        QLabel#ResultSummary {{
            font-size: 15px;
            color: {text_primary};
            font-weight: 600;
        }}
        QFrame#DetailsBox {{
            background-color: {bg_surface};
            border: 1px solid {border_subtle};
            border-radius: 10px;
        }}
        QLabel#DetailKey {{
            color: {text_secondary};
            font-size: 14px;
            font-weight: 600;
            min-width: 105px;
        }}
        QLabel#DetailVal {{
            color: {text_primary};
            font-size: 14px;
            font-weight: 600;
        }}
        QPushButton {{
            min-width: 90px;
            min-height: 34px;
            font-weight: 700;
            border-radius: 10px;
            border: 1px solid {border_default};
            background-color: {bg_surface};
            color: {text_primary};
            padding: 5px 16px;
        }}
        QPushButton:hover {{
            background-color: {bg_hover};
            border-color: {border_strong};
        }}
        QPushButton:pressed {{
            background-color: {bg_active};
        }}
    """
