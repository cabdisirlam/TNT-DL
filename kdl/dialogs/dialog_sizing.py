"""Helpers for sizing dialogs to desktop screens without clipping action rows."""

from PySide6.QtCore import Qt
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import QFrame, QScrollArea, QToolButton, QToolTip, QVBoxLayout, QWidget


def create_scrollable_dialog_layout(
    dialog,
    *,
    content_margins: tuple[int, int, int, int] = (16, 16, 16, 16),
    spacing: int = 12,
) -> QVBoxLayout:
    """Return a scrollable top-level layout for taller dialogs."""
    root = QVBoxLayout(dialog)
    root.setContentsMargins(0, 0, 0, 0)
    root.setSpacing(0)

    scroll = QScrollArea(dialog)
    scroll.setWidgetResizable(True)
    scroll.setFrameShape(QFrame.NoFrame)
    scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
    scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
    scroll.setStyleSheet("QScrollArea { background: transparent; }")
    scroll.viewport().setAutoFillBackground(False)

    content = QWidget(scroll)
    content.setObjectName("DialogScrollContent")
    content.setAttribute(Qt.WA_StyledBackground, True)
    content.setStyleSheet("QWidget#DialogScrollContent { background: transparent; }")
    layout = QVBoxLayout(content)
    layout.setContentsMargins(*content_margins)
    layout.setSpacing(spacing)

    scroll.setWidget(content)
    root.addWidget(scroll)

    return layout


def fit_dialog_to_screen(
    dialog,
    *,
    min_width: int,
    min_height: int,
    preferred_width: int,
    wide_width: int | None = None,
    preferred_height: int | None = None,
    tall_height: int | None = None,
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
    desktop_height = (
        tall_height if tall_height is not None and geo.height() >= 1000 else preferred_height
    )

    target_w = max(clamped_min_w, hint.width() + extra_hint_width, desktop_width)
    hinted_h = hint.height() + extra_hint_height
    if desktop_height is None:
        target_h = max(clamped_min_h, hinted_h)
    else:
        target_h = max(clamped_min_h, min(hinted_h, desktop_height))

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
