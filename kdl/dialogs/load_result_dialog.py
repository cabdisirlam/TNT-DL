"""
Professional load result dialog for success, stop, and error outcomes.
"""

from PySide6.QtCore import Qt
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QFrame,
    QDialogButtonBox,
    QStyle,
    QSizePolicy,
)
from kdl.styles import load_result_qss, ACCENT


class LoadResultDialog(QDialog):
    """Styled result dialog shown after load completion."""

    def __init__(
        self,
        title: str,
        message: str,
        status: str = "info",
        parent=None,
        confirm: bool = False,
        accept_text: str = "OK",
        reject_text: str = "Cancel",
    ):
        super().__init__(parent)
        self._title = title.strip() or "Load Result"
        self._message = (message or "").strip()
        self._status = (status or "info").lower()
        self._confirm = bool(confirm)
        self._accept_text = (accept_text or "OK").strip()
        self._reject_text = (reject_text or "Cancel").strip()

        self.setWindowTitle(self._title)
        self.setModal(True)
        self.setMinimumWidth(380)
        self.setWindowFlag(Qt.WindowContextHelpButtonHint, False)
        self.setWindowFlag(Qt.WindowCloseButtonHint, True)

        self._build_ui()
        self._apply_style()
        self._fit_to_screen()

    def _status_palette(self):
        def tint(hex_color: str) -> str:
            raw = (hex_color or "").strip().lstrip("#")
            if len(raw) != 6:
                return "rgba(0, 0, 0, 0.08)"
            try:
                r = int(raw[0:2], 16)
                g = int(raw[2:4], 16)
                b = int(raw[4:6], 16)
            except ValueError:
                return "rgba(0, 0, 0, 0.08)"
            return f"rgba({r}, {g}, {b}, 0.12)"

        if self._status == "success":
            accent = ACCENT
            return {
                "icon": self.style().standardIcon(QStyle.SP_DialogApplyButton),
                "accent": accent,
                "panel": tint(accent),
                "title": "Load completed",
            }
        if self._status == "stopped":
            accent = ACCENT
            return {
                "icon": self.style().standardIcon(QStyle.SP_MessageBoxInformation),
                "accent": accent,
                "panel": tint(accent),
                "title": "Load stopped",
            }
        if self._status == "warning":
            accent = ACCENT
            return {
                "icon": self.style().standardIcon(QStyle.SP_MessageBoxWarning),
                "accent": accent,
                "panel": tint(accent),
                "title": "Attention required",
            }
        if self._status == "error":
            accent = ACCENT
            return {
                "icon": self.style().standardIcon(QStyle.SP_MessageBoxCritical),
                "accent": accent,
                "panel": tint(accent),
                "title": "Action needed",
            }
        if self._status == "info":
            accent = ACCENT
            return {
                "icon": self.style().standardIcon(QStyle.SP_MessageBoxInformation),
                "accent": accent,
                "panel": tint(accent),
                "title": "Information",
            }
        accent = ACCENT
        return {
            "icon": self.style().standardIcon(QStyle.SP_MessageBoxWarning),
            "accent": accent,
            "panel": tint(accent),
            "title": "Load result",
        }

    def _parse_message(self):
        lines = [line.strip() for line in self._message.splitlines() if line.strip()]
        if not lines:
            return "", []
        summary = lines[0]
        kv_pairs = []
        for line in lines[1:]:
            if ":" not in line:
                continue
            key, value = line.split(":", 1)
            key = key.strip()
            value = value.strip()
            if key:
                kv_pairs.append((key, value))
        return summary, kv_pairs

    def _build_ui(self):
        palette = self._status_palette()
        summary, kv_pairs = self._parse_message()

        root = QVBoxLayout(self)
        root.setContentsMargins(14, 12, 14, 10)
        root.setSpacing(10)

        panel = QFrame()
        panel.setObjectName("ResultPanel")
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(12, 10, 12, 10)
        panel_layout.setSpacing(10)

        top = QHBoxLayout()
        top.setSpacing(10)

        icon_label = QLabel()
        icon_label.setPixmap(palette["icon"].pixmap(28, 28))
        icon_label.setAlignment(Qt.AlignTop)
        top.addWidget(icon_label, 0, Qt.AlignTop)

        text_col = QVBoxLayout()
        text_col.setSpacing(2)

        heading = QLabel(palette["title"])
        heading.setObjectName("ResultHeading")
        text_col.addWidget(heading)

        subheading = QLabel(self._title)
        subheading.setObjectName("ResultSubheading")
        text_col.addWidget(subheading)

        top.addLayout(text_col, 1)
        panel_layout.addLayout(top)

        if summary:
            summary_label = QLabel(summary)
            summary_label.setObjectName("ResultSummary")
            summary_label.setWordWrap(True)
            panel_layout.addWidget(summary_label)

        if kv_pairs:
            details_box = QFrame()
            details_box.setObjectName("DetailsBox")
            details_layout = QVBoxLayout(details_box)
            details_layout.setContentsMargins(10, 8, 10, 8)
            details_layout.setSpacing(5)
            for key, value in kv_pairs:
                row = QHBoxLayout()
                row.setSpacing(8)

                key_label = QLabel(f"{key}:")
                key_label.setObjectName("DetailKey")
                key_label.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Preferred)
                row.addWidget(key_label)

                val_label = QLabel(value)
                val_label.setObjectName("DetailVal")
                val_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
                row.addWidget(val_label, 1)

                details_layout.addLayout(row)

            panel_layout.addWidget(details_box)

        root.addWidget(panel)

        if self._confirm:
            buttons = QDialogButtonBox(QDialogButtonBox.Yes | QDialogButtonBox.No)
            yes_btn = buttons.button(QDialogButtonBox.Yes)
            no_btn = buttons.button(QDialogButtonBox.No)
            if yes_btn is not None:
                yes_btn.setText(self._accept_text)
                yes_btn.setDefault(True)
            if no_btn is not None:
                no_btn.setText(self._reject_text)
            buttons.accepted.connect(self.accept)
            buttons.rejected.connect(self.reject)
        else:
            buttons = QDialogButtonBox(QDialogButtonBox.Ok)
            ok_btn = buttons.button(QDialogButtonBox.Ok)
            if ok_btn is not None:
                ok_btn.setText(self._accept_text)
                ok_btn.setDefault(True)
            buttons.accepted.connect(self.accept)
        root.addWidget(buttons)

        panel.setProperty("accent", palette["accent"])
        panel.setProperty("panel", palette["panel"])
        panel.style().unpolish(panel)
        panel.style().polish(panel)

    def _apply_style(self):
        palette = self._status_palette()
        self.setStyleSheet(load_result_qss(palette["accent"], palette["panel"]))

    def _fit_to_screen(self):
        screen = self.screen() or QGuiApplication.primaryScreen()
        if not screen:
            return
        geo = screen.availableGeometry()
        max_w = max(320, geo.width() - 24)
        max_h = max(260, geo.height() - 24)
        self.setMaximumSize(max_w, max_h)
        hint = self.sizeHint()
        target_w = min(max(self.minimumWidth(), hint.width()), max_w)
        target_h = min(max(240, hint.height()), max_h)
        self.resize(target_w, target_h)
