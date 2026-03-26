"""Helpers for sizing dialogs to desktop screens without clipping action rows."""

from PySide6.QtCore import Qt
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import QToolButton, QToolTip


def fit_dialog_to_screen(
    dialog,
    *,
    min_width: int,
    min_height: int,
    preferred_width: int,
    wide_width: int | None = None,
    margin_width: int = 72,
    margin_height: int = 72,
    extra_hint_width: int = 24,
    extra_hint_height: int = 24,
) -> None:
    """Resize a dialog to a comfortable desktop width while clamping to screen bounds."""
    screen = dialog.screen() or QGuiApplication.primaryScreen()
    if not screen:
        return

    geo = screen.availableGeometry()
    max_w = max(320, geo.width() - margin_width)
    max_h = max(240, geo.height() - margin_height)
    clamped_min_w = min(min_width, max_w)
    clamped_min_h = min(min_height, max_h)

    hint = dialog.sizeHint()
    desktop_width = wide_width if wide_width is not None and geo.width() >= 1600 else preferred_width

    target_w = max(clamped_min_w, hint.width() + extra_hint_width, desktop_width)
    target_h = max(clamped_min_h, hint.height() + extra_hint_height)

    dialog.setMinimumSize(clamped_min_w, clamped_min_h)
    dialog.setMaximumSize(max_w, max_h)
    dialog.resize(min(target_w, max_w), min(target_h, max_h))


def create_hint_button(text: str, label: str = "?") -> QToolButton:
    """Create a compact hover-help button for longer dialog instructions."""
    help_text = (text or "").strip()
    button = QToolButton()
    button.setText(label)
    button.setAutoRaise(True)
    button.setCursor(Qt.PointingHandCursor)
    button.setToolTip(help_text)
    button.setToolTipDuration(30000)
    button.setWhatsThis(help_text)
    button.setFocusPolicy(Qt.NoFocus)
    button.setFixedSize(24, 24)
    button.setStyleSheet(
        """
        QToolButton {
            border: 1px solid #C7D2DE;
            border-radius: 12px;
            font-weight: 700;
            font-size: 12px;
            color: #27445C;
            background: #F4F8FB;
        }
        QToolButton:hover {
            background: #E6EFF6;
            border-color: #98B0C3;
        }
        """
    )
    button.clicked.connect(
        lambda: QToolTip.showText(
            button.mapToGlobal(button.rect().bottomLeft()),
            help_text,
            button,
        )
    )
    return button
