"""
Macro recorder dialog for capturing keyboard input into KDL keystroke syntax.
"""

import re
import threading
from typing import List, Optional, Set

import pyautogui
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QTextCursor, QGuiApplication
from PySide6.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPlainTextEdit,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from kdl.dialogs.load_result_dialog import LoadResultDialog
from kdl.styles import dialog_qss

try:
    import keyboard
except Exception:  # pragma: no cover - environment dependent
    keyboard = None

try:
    from pynput import keyboard as pk
except Exception:  # pragma: no cover - environment dependent
    pk = None


_SPECIAL_KEY_MAP = {
    "tab": "TAB",
    "enter": "ENTER",
    "esc": "ESC",
    "escape": "ESC",
    "up": "UP",
    "down": "DOWN",
    "left": "LEFT",
    "right": "RIGHT",
    "home": "HOME",
    "end": "END",
    "page up": "PGUP",
    "page down": "PGDN",
    "backspace": "BACKSPACE",
    "delete": "DELETE",
    "del": "DELETE",
    "insert": "INSERT",
    "space": "SPACE",
    "caps lock": "CAPSLOCK",
    "print screen": "PRTSC",
}

for _i in range(1, 17):
    _SPECIAL_KEY_MAP[f"f{_i}"] = f"F{_i}"


class MacroRecorderDialog(QDialog):
    """Capture global keyboard presses and convert to KDL keystroke string."""

    token_captured = Signal(str)
    stop_requested = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Macro Recorder")
        self.setMinimumWidth(440)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.setWindowFlag(Qt.WindowCloseButtonHint, True)

        self._hook = None
        self._hook_backend = ""
        self._recording = False
        self._tokens: List[str] = []
        self._modifiers_down: Set[str] = set()
        self._modifiers_lock = threading.Lock()
        self._mouse_macro_text: str = ""

        from kdl.config_store import get_dark_mode
        self.setStyleSheet(dialog_qss(dark=get_dark_mode()))
        self._build_ui()
        self.token_captured.connect(self._append_token_ui)
        self.stop_requested.connect(self._stop_recording)
        self._fit_to_screen()

    def _build_ui(self):
        dialog_layout = QVBoxLayout(self)
        dialog_layout.setContentsMargins(0, 0, 0, 0)
        dialog_layout.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.NoFrame)
        dialog_layout.addWidget(scroll)

        scroll_widget = QWidget()
        scroll.setWidget(scroll_widget)
        layout = QVBoxLayout(scroll_widget)
        layout.setSpacing(8)

        title = QLabel(
            "Record keyboard actions, then insert as App Format macro (e.g. tab, *S, *DN, \\r)."
        )
        title.setWordWrap(True)
        layout.addWidget(title)

        btn_row = QHBoxLayout()
        self.start_btn = QPushButton("Start Recording")
        self.start_btn.clicked.connect(self._start_recording)
        btn_row.addWidget(self.start_btn)

        self.stop_btn = QPushButton("Stop")
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self._stop_recording)
        btn_row.addWidget(self.stop_btn)

        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(self._clear_recording)
        btn_row.addWidget(clear_btn)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        helper_row = QHBoxLayout()
        helper_row.addWidget(QLabel("Mouse Macro:"))
        left_click_btn = QPushButton("Use Left Click Position")
        left_click_btn.clicked.connect(lambda: self._capture_mouse_click("left"))
        helper_row.addWidget(left_click_btn)
        right_click_btn = QPushButton("Use Right Click Position")
        right_click_btn.clicked.connect(lambda: self._capture_mouse_click("right"))
        helper_row.addWidget(right_click_btn)
        helper_row.addStretch()
        layout.addLayout(helper_row)

        self.state_label = QLabel("Status: idle (press F8 to stop while recording)")
        layout.addWidget(self.state_label)

        self.preview = QPlainTextEdit()
        self.preview.setPlaceholderText("Recorded macro will appear here...")
        self.preview.setFixedHeight(90)
        layout.addWidget(self.preview)

        mode_row = QHBoxLayout()
        mode_row.addWidget(QLabel("Insert into:"))
        self.mode_group = QButtonGroup(self)
        self.current_cell_radio = QRadioButton("Current Cell")
        self.current_cell_radio.setChecked(True)
        self.mode_group.addButton(self.current_cell_radio)
        mode_row.addWidget(self.current_cell_radio)
        self.selection_radio = QRadioButton("Selected Range")
        self.mode_group.addButton(self.selection_radio)
        mode_row.addWidget(self.selection_radio)
        mode_row.addStretch()
        layout.addLayout(mode_row)

        shortcut_row = QHBoxLayout()
        self.save_shortcut_check = QCheckBox("Save as shortcut:")
        self.save_shortcut_check.toggled.connect(self._toggle_shortcut_input)
        shortcut_row.addWidget(self.save_shortcut_check)
        self.shortcut_input = QLineEdit("*M1")
        self.shortcut_input.setEnabled(False)
        self.shortcut_input.setFixedWidth(84)
        self.shortcut_input.setToolTip("Example: *M1, *BANK1")
        shortcut_row.addWidget(self.shortcut_input)
        shortcut_row.addWidget(QLabel("(for keystroke macros starting with \\ only)"))
        shortcut_row.addStretch()
        layout.addLayout(shortcut_row)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        cancel_button = buttons.button(QDialogButtonBox.Cancel)
        if cancel_button is not None:
            cancel_button.setText("Close")
        buttons.accepted.connect(self._accept_if_valid)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

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
        target_h = min(max(360, hint.height()), max_h)
        self.resize(target_w, target_h)

    def _start_recording(self):
        if self._recording:
            return

        self._tokens = []
        self._modifiers_down.clear()
        self._mouse_macro_text = ""
        self.preview.clear()

        keyboard_error = ""
        try:
            if keyboard is not None:
                self._hook = keyboard.hook(self._on_key_event, suppress=False)
                self._hook_backend = "keyboard"
        except Exception as exc:
            keyboard_error = str(exc)
            self._hook = None
            self._hook_backend = ""

        if self._hook is None:
            try:
                if pk is not None:
                    self._hook = pk.Listener(
                        on_press=self._on_press_pynput,
                        on_release=self._on_release_pynput,
                    )
                    self._hook.start()
                    self._hook_backend = "pynput"
            except Exception as exc:
                detail = f"keyboard: {keyboard_error or 'not available'}\n"
                detail += f"pynput: {exc}"
                self._show_styled_message(
                    "Recorder Error",
                    "Failed to start keyboard recording.\n"
                    f"{detail}\n\nTry running NT_DL as Administrator.",
                    status="warning",
                )
                return

        if self._hook is None:
            self._show_styled_message(
                "Recorder Error",
                "Failed to start keyboard recording.\n"
                "No supported keyboard hook backend is available.",
                status="warning",
            )
            return

        self._recording = True
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.state_label.setText("Status: recording... (press F8 to stop)")

    def _stop_recording(self):
        self._unhook()
        self._recording = False
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        macro = self._render_macro_string()
        if macro:
            self.preview.setPlainText(macro)
        self.state_label.setText("Status: stopped")

    def _clear_recording(self):
        self._tokens = []
        with self._modifiers_lock:
            self._modifiers_down.clear()
        self._mouse_macro_text = ""
        self.preview.clear()
        self.state_label.setText("Status: idle")

    def _accept_if_valid(self):
        if self._recording:
            self._stop_recording()

        text = self.preview.toPlainText().strip()
        if not text:
            self._show_styled_message(
                "No Macro",
                "Record or type a macro first.",
                status="warning",
            )
            return

        if self.save_shortcut_check.isChecked():
            key = self.get_shortcut_key()
            if not key:
                self._show_styled_message(
                    "Invalid Shortcut",
                    "Shortcut must start with '*' and have at least one more character.",
                    status="warning",
                )
                return
            if not text.startswith("\\"):
                self._show_styled_message(
                    "Unsupported Shortcut Value",
                    "Only keystroke macros (starting with \\) can be saved as shortcuts.",
                    status="warning",
                )
                return
        self.accept()

    def _show_styled_message(self, title: str, message: str, status: str = "warning"):
        dialog = LoadResultDialog(title=title, message=message, status=status, parent=self)
        dialog.exec()

    def _on_key_event(self, event):
        if not self._recording:
            return

        name = (event.name or "").lower().strip()
        if not name:
            return

        modifier = self._get_modifier_name(name)
        if event.event_type == "down":
            if name == "f8":
                self.stop_requested.emit()
                return
            with self._modifiers_lock:
                if modifier:
                    self._modifiers_down.add(modifier)
                    return
                mods = set(self._modifiers_down)

            token = self._token_from_key(name, mods)
            if token:
                self.token_captured.emit(token)
        elif event.event_type == "up" and modifier:
            with self._modifiers_lock:
                self._modifiers_down.discard(modifier)

    @staticmethod
    def _normalize_pynput_key(key) -> str:
        try:
            if hasattr(key, "char") and key.char:
                return str(key.char).lower().strip()
            if hasattr(key, "name") and key.name:
                return str(key.name).lower().strip()
            text = str(key).lower().strip()
            if text.startswith("key."):
                text = text[4:]
            return text
        except Exception:
            return ""

    def _on_press_pynput(self, key):
        if not self._recording:
            return

        name = self._normalize_pynput_key(key)
        if not name:
            return

        modifier = self._get_modifier_name(name)
        if name == "f8":
            self.stop_requested.emit()
            return
        with self._modifiers_lock:
            if modifier:
                self._modifiers_down.add(modifier)
                return
            mods = set(self._modifiers_down)

        token = self._token_from_key(name, mods)
        if token:
            self.token_captured.emit(token)

    def _on_release_pynput(self, key):
        if not self._recording:
            return

        name = self._normalize_pynput_key(key)
        if not name:
            return

        modifier = self._get_modifier_name(name)
        if modifier:
            with self._modifiers_lock:
                self._modifiers_down.discard(modifier)

    @staticmethod
    def _get_modifier_name(name: str) -> Optional[str]:
        if "ctrl" in name:
            return "ctrl"
        if name in {"alt", "alt gr", "left alt", "right alt"}:
            return "alt"
        if "shift" in name:
            return "shift"
        return None

    def _token_from_key(self, name: str, modifiers: Set[str]) -> Optional[str]:
        normalized = name.replace("_", " ").replace("-", " ")
        normalized = re.sub(r"\s+", " ", normalized).strip()

        # Build combined modifier prefix (supports Ctrl+Shift+Alt combos)
        prefix = ""
        if "ctrl" in modifiers:
            prefix += "^"
        if "alt" in modifiers:
            prefix += "%"
        if "shift" in modifiers:
            prefix += "+"

        if len(normalized) == 1 and normalized.isprintable():
            if prefix:
                return f"{prefix}{normalized.lower()}"
            return normalized

        key_name = _SPECIAL_KEY_MAP.get(normalized)
        if not key_name:
            if re.fullmatch(r"f([1-9]|1[0-6])", normalized):
                key_name = normalized.upper()
            else:
                return None

        if prefix:
            return f"{prefix}{{{key_name}}}"
        return f"{{{key_name}}}"

    def _render_macro_string(self) -> str:
        if self._mouse_macro_text:
            return self._mouse_macro_text
        if not self._tokens:
            return ""
        return "\\" + "".join(self._tokens)

    def _append_token_ui(self, token: str):
        self._tokens.append(token)
        macro = self._render_macro_string()
        self.preview.setPlainText(macro)
        self.preview.moveCursor(QTextCursor.End)

    def _capture_mouse_click(self, button: str):
        x, y = pyautogui.position()
        if button == "right":
            self._mouse_macro_text = f"*MR({x},{y})"
        else:
            self._mouse_macro_text = f"*MC({x},{y})"
        self._tokens = []
        self.preview.setPlainText(self._mouse_macro_text)
        self.state_label.setText(f"Status: mouse macro captured at ({x}, {y})")

    def _toggle_shortcut_input(self, checked: bool):
        self.shortcut_input.setEnabled(checked)

    def get_macro_text(self) -> str:
        return self.preview.toPlainText().strip()

    def apply_to_selection(self) -> bool:
        return self.selection_radio.isChecked()

    def should_save_shortcut(self) -> bool:
        return self.save_shortcut_check.isChecked()

    def get_shortcut_key(self) -> str:
        key = self.shortcut_input.text().strip().upper()
        if not key:
            return ""
        if not key.startswith("*"):
            key = "*" + key
        if len(key) < 2:
            return ""
        return key

    def _unhook(self):
        if self._hook is None:
            return
        if self._hook_backend == "keyboard":
            try:
                if keyboard is not None:
                    keyboard.unhook(self._hook)
            except Exception:
                pass
        elif self._hook_backend == "pynput":
            try:
                self._hook.stop()
            except Exception:
                pass
        self._hook = None
        self._hook_backend = ""

    def closeEvent(self, event):
        self._unhook()
        super().closeEvent(event)
