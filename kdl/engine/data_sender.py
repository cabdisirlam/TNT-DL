"""
KDL Data Sender
Sends data and keystrokes to the target application window.
Uses pyautogui for keystroke simulation and pyperclip for clipboard-based data pasting.
"""

import time
import pyautogui
import pyperclip
from typing import Optional, Callable

from kdl.engine.keystroke_parser import ParsedCell, CellType
from kdl.window.window_manager import WindowManager


# Safety: keep pyautogui fail-safe enabled (mouse to corner aborts)
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0  # Disable built-in pause; loader uses its own speed_delay


class DataSender:
    """Sends parsed cell data to a target application window."""

    def __init__(self):
        self.target_hwnd: Optional[int] = None
        self.target_title: str = ""
        self.speed_delay: float = 0.1  # Delay between actions in seconds
        self.window_delay: float = 0.1  # Delay after activating target window
        self.wait_for_hourglass: bool = False
        self.hourglass_timeout: float = 30.0  # Max wait time for hourglass
        self.stop_requested_cb: Optional[Callable[[], bool]] = None
        self.last_error: str = ""
        self._clipboard_session_active = False
        self._clipboard_session_value: Optional[str] = None
        self._clipboard_session_had_value = False

    def set_stop_checker(self, checker: Optional[Callable[[], bool]]):
        """Set callback used to abort long waits when stop is requested."""
        self.stop_requested_cb = checker

    def _is_stop_requested(self) -> bool:
        if not self.stop_requested_cb:
            return False
        try:
            return bool(self.stop_requested_cb())
        except Exception:
            return False

    def _sleep_interruptible(self, seconds: float) -> bool:
        """Sleep in small chunks so stop requests can interrupt long waits."""
        seconds = max(0.0, float(seconds))
        end_time = time.time() + seconds
        while time.time() < end_time:
            if self._is_stop_requested():
                return False
            time.sleep(max(0, min(0.01, end_time - time.time())))
        return True

    def set_target(self, hwnd: int, title: str):
        """Set the target window."""
        self.target_hwnd = hwnd
        self.target_title = title

    def set_speed(self, delay_seconds: float):
        """Set the delay between actions (0.01 = fast, 1.0 = slow)."""
        self.speed_delay = max(0.0, min(2.0, delay_seconds))

    def set_window_delay(self, delay_seconds: float):
        """Set delay after window activation."""
        self.window_delay = max(0.0, min(5.0, delay_seconds))

    def begin_clipboard_session(self):
        """
        Capture clipboard once for the whole load session.
        Restoring per-cell can race with Ctrl+V in some apps and paste stale content.
        """
        if self._clipboard_session_active:
            return
        self._clipboard_session_active = True
        self._clipboard_session_value = None
        self._clipboard_session_had_value = False
        try:
            self._clipboard_session_value = pyperclip.paste()
            self._clipboard_session_had_value = True
        except Exception:
            self._clipboard_session_value = None
            self._clipboard_session_had_value = False

    def end_clipboard_session(self):
        """Restore clipboard captured at session start, or clear it."""
        if not self._clipboard_session_active:
            return
        try:
            if self._clipboard_session_had_value:
                pyperclip.copy(self._clipboard_session_value if self._clipboard_session_value is not None else "")
            else:
                # Initial capture failed — clear clipboard so loaded data doesn't linger
                pyperclip.copy("")
        except Exception:
            pass
        finally:
            self._clipboard_session_active = False
            self._clipboard_session_value = None
            self._clipboard_session_had_value = False

    def activate_target(self) -> bool:
        """Activate the target window. Returns True if successful."""
        # Fallback if user manually typed a title with no HWND
        if not self.target_hwnd and self.target_title:
            self.target_hwnd = WindowManager.find_window_containing_title(self.target_title)

        if not self.target_hwnd:
            return False

        if self.target_title and self.verify_target_active():
            return True

        activated_any = False
        for _ in range(2):
            result = WindowManager.activate_window(self.target_hwnd)
            if not result:
                continue
            activated_any = True
            if self.window_delay > 0 and not self._sleep_interruptible(self.window_delay):
                return False
            if not self.target_title or self.verify_target_active():
                return True

        if not self.target_title:
            return activated_any
        return False

    def verify_target_active(self) -> bool:
        """Check if the target window is currently the foreground window."""
        fg_hwnd = WindowManager.get_foreground_window_handle()
        if not fg_hwnd:
            return False

        try:
            if self.target_hwnd and int(fg_hwnd) == int(self.target_hwnd):
                return True
        except Exception:
            pass

        if self.target_title:
            target = self.target_title.strip().lower()
            fg_title = WindowManager.get_foreground_window_title().strip().lower()
            if target and fg_title and (target in fg_title or fg_title in target):
                return True

        return False

    def _is_excel_target(self) -> bool:
        """Best-effort check for Microsoft Excel target window."""
        try:
            if not self.target_hwnd:
                return False
            pname = WindowManager.get_window_process_name(self.target_hwnd).strip().lower()
            return pname == "excel.exe"
        except Exception:
            return False

    def _wait_if_hourglass(self) -> bool:
        """Wait until the cursor is no longer an hourglass; return False if interrupted/timed out."""
        if not self.wait_for_hourglass:
            return True

        start = time.time()
        while WindowManager.is_cursor_hourglass():
            if self._is_stop_requested():
                self.last_error = "hourglass wait interrupted"
                return False
            if time.time() - start > self.hourglass_timeout:
                self.last_error = f"hourglass wait timeout ({self.hourglass_timeout:.1f}s)"
                return False
            if not self._sleep_interruptible(0.01):
                self.last_error = "hourglass wait interrupted"
                return False
        return True

    def send_cell(self, parsed: ParsedCell) -> bool:
        """
        Send a single parsed cell to the target application.
        Returns True if successful.
        """
        self.last_error = ""
        try:
            if parsed.cell_type == CellType.EMPTY:
                return True

            if parsed.cell_type == CellType.DATA:
                return self._send_data(parsed.data_text)

            if parsed.cell_type == CellType.KEYSTROKE:
                return self._send_keystrokes(parsed)

            if parsed.cell_type == CellType.MOUSE_LEFT:
                return self._send_mouse_click(parsed.mouse_x, parsed.mouse_y, button='left')

            if parsed.cell_type == CellType.MOUSE_RIGHT:
                return self._send_mouse_click(parsed.mouse_x, parsed.mouse_y, button='right')

            if parsed.cell_type == CellType.DELAY:
                if not self._sleep_interruptible(parsed.delay_ms / 1000.0):
                    self.last_error = "delay interrupted"
                    return False
                return True

            return True

        except Exception as e:
            self.last_error = f"send_cell: {e}"
            print(f"Error sending cell: {e}")
            return False

    def _send_data(self, text: str) -> bool:
        """Send plain text data using clipboard copy-paste."""
        try:
            if not self._clipboard_session_active:
                self.begin_clipboard_session()

            # Keep one-cell semantics; newline pastes can spill into multiple rows in Excel.
            safe_text = (text or "").replace("\r\n", " ").replace("\n", " ").replace("\r", " ")

            # Excel is sensitive to clipboard state while rapidly automating.
            # For ASCII-only text, typewrite avoids "cannot paste" popups.
            # For Unicode text, fall through to clipboard paste.
            if self._is_excel_target() and safe_text.isascii():
                pyautogui.typewrite(safe_text, interval=0.01)
                if not self._sleep_interruptible(self.speed_delay):
                    self.last_error = "send_data interrupted"
                    return False
                return self._wait_if_hourglass()

            # Copy data to clipboard
            pyperclip.copy(safe_text)
            if not self._sleep_interruptible(0.01):
                self.last_error = "send_data interrupted"
                return False

            # Paste into target using Ctrl+V
            pyautogui.hotkey('ctrl', 'v')
            if not self._sleep_interruptible(self.speed_delay):
                self.last_error = "send_data interrupted"
                return False

            return self._wait_if_hourglass()
        except Exception as e:
            self.last_error = f"send_data: {e}"
            print(f"Error sending data: {e}")
            return False

    def _send_keystrokes(self, parsed: ParsedCell) -> bool:
        """Send keystroke actions to the target."""
        try:
            for action in parsed.key_actions:
                action_type = action.get("type", "")

                if action_type == "key":
                    # Single key press
                    key = action.get("key", "")
                    pyautogui.press(key)

                elif action_type == "hotkey":
                    # Key with modifiers
                    modifiers = action.get("modifiers", [])
                    key = action.get("key", "")
                    keys = modifiers + [key]
                    pyautogui.hotkey(*keys)

                elif action_type == "type":
                    # Type text characters
                    text = action.get("text", "")
                    pyautogui.typewrite(text, interval=0.02)

                if not self._sleep_interruptible(self.speed_delay):
                    self.last_error = "send_keystrokes interrupted"
                    return False
                if not self._wait_if_hourglass():
                    return False

            return True
        except Exception as e:
            self.last_error = f"send_keystrokes: {e}"
            print(f"Error sending keystrokes: {e}")
            return False

    def _send_mouse_click(self, x: int, y: int, button: str = 'left') -> bool:
        """Simulate a mouse click at screen coordinates."""
        try:
            pyautogui.click(x, y, button=button)
            if not self._sleep_interruptible(self.speed_delay):
                self.last_error = "send_mouse_click interrupted"
                return False
            return self._wait_if_hourglass()
        except Exception as e:
            self.last_error = f"send_mouse_click: {e}"
            print(f"Error sending mouse click: {e}")
            return False
