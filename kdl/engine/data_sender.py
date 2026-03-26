"""
KDL Data Sender
Sends data and keystrokes to the target application window.
Uses pyautogui for keystroke simulation and pyperclip for clipboard-based data pasting.
"""

import ctypes
import ctypes.wintypes
import time
import pyautogui
import pyperclip
from typing import Optional, Callable

from kdl.engine.keystroke_parser import ParsedCell, CellType
from kdl.window.window_manager import WindowManager


# Safety: keep pyautogui fail-safe enabled (mouse to corner aborts)
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0  # Disable built-in pause; loader uses its own speed_delay

# ─── SendInput (Fast Send) Win32 structures ───────────────────────────────────
_INPUT_KEYBOARD   = 1
_KEYEVENTF_KEYUP  = 0x0002
_KEYEVENTF_UNICODE = 0x0004
_KEYEVENTF_EXTENDEDKEY = 0x0001


class _KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk",         ctypes.wintypes.WORD),
        ("wScan",       ctypes.wintypes.WORD),
        ("dwFlags",     ctypes.wintypes.DWORD),
        ("time",        ctypes.wintypes.DWORD),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]


class _MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx",          ctypes.c_long),
        ("dy",          ctypes.c_long),
        ("mouseData",   ctypes.wintypes.DWORD),
        ("dwFlags",     ctypes.wintypes.DWORD),
        ("time",        ctypes.wintypes.DWORD),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]


class _HARDWAREINPUT(ctypes.Structure):
    _fields_ = [
        ("uMsg",    ctypes.wintypes.DWORD),
        ("wParamL", ctypes.wintypes.WORD),
        ("wParamH", ctypes.wintypes.WORD),
    ]


class _INPUT_UNION(ctypes.Union):
    _fields_ = [
        ("mi", _MOUSEINPUT),
        ("ki", _KEYBDINPUT),
        ("hi", _HARDWAREINPUT),
    ]


class _INPUT(ctypes.Structure):
    _fields_ = [
        ("type",   ctypes.wintypes.DWORD),
        ("_input", _INPUT_UNION),
    ]


# pyautogui key name → Win32 Virtual Key code
_SI_VK_MAP: dict = {
    "tab": 0x09, "enter": 0x0D, "return": 0x0D, "escape": 0x1B, "esc": 0x1B,
    "backspace": 0x08, "delete": 0x2E, "insert": 0x2D,
    "home": 0x24, "end": 0x23, "pageup": 0x21, "pagedown": 0x22,
    "left": 0x25, "up": 0x26, "right": 0x27, "down": 0x28,
    "f1": 0x70, "f2": 0x71, "f3": 0x72,  "f4": 0x73,
    "f5": 0x74, "f6": 0x75, "f7": 0x76,  "f8": 0x77,
    "f9": 0x78, "f10": 0x79, "f11": 0x7A, "f12": 0x7B,
    "shift": 0x10, "ctrl": 0x11, "alt": 0x12,
    "space": 0x20, "capslock": 0x14,
}

# VK codes that need KEYEVENTF_EXTENDEDKEY (arrow keys, nav cluster)
_SI_EXTENDED_KEYS: set = {
    0x21, 0x22, 0x23, 0x24,   # pageup, pagedown, end, home
    0x25, 0x26, 0x27, 0x28,   # left, up, right, down
    0x2D, 0x2E,               # insert, delete
}


def _si_make_ki(vk: int = 0, scan: int = 0, flags: int = 0) -> _INPUT:
    """Build a keyboard INPUT struct for SendInput."""
    inp = _INPUT()
    inp.type = _INPUT_KEYBOARD
    inp._input.ki.wVk = vk
    inp._input.ki.wScan = scan
    inp._input.ki.dwFlags = flags
    inp._input.ki.time = 0
    inp._input.ki.dwExtraInfo = ctypes.pointer(ctypes.c_ulong(0))
    return inp

# ─────────────────────────────────────────────────────────────────────────────


class DataSender:
    """Sends parsed cell data to a target application window."""

    def __init__(self):
        self.target_hwnd: Optional[int] = None
        self.target_title: str = ""
        self.speed_delay: float = 0.01  # Default non-fast-send delay between actions
        self.window_delay: float = 0.1  # Delay after activating target window
        self.wait_for_hourglass: bool = False
        self.hourglass_timeout: float = 30.0  # Max wait time for hourglass
        self.stop_requested_cb: Optional[Callable[[], bool]] = None
        self.pause_requested_cb: Optional[Callable[[], bool]] = None
        self.last_error: str = ""
        self.use_fast_send: bool = False
        self.fast_send_row_mode: bool = False  # True = minimal per-cell delay, full delay at row end only
        self.load_control: bool = False   # Adaptive: skip fixed delay, wait for app readiness
        self._clipboard_session_active = False
        self._clipboard_session_value: Optional[str] = None
        self._clipboard_session_had_value = False

    def set_stop_checker(self, checker: Optional[Callable[[], bool]]):
        """Set callback used to abort long waits when stop is requested."""
        self.stop_requested_cb = checker

    def set_pause_checker(self, checker: Optional[Callable[[], bool]]):
        """Set callback used to pause long waits while the loader is paused."""
        self.pause_requested_cb = checker

    def _is_stop_requested(self) -> bool:
        if not self.stop_requested_cb:
            return False
        try:
            return bool(self.stop_requested_cb())
        except Exception:
            return False

    def _is_pause_requested(self) -> bool:
        if not self.pause_requested_cb:
            return False
        try:
            return bool(self.pause_requested_cb())
        except Exception:
            return False

    def _wait_while_paused(self) -> bool:
        """Yield while paused so manual Pause also affects sender waits."""
        while self._is_pause_requested():
            if self._is_stop_requested():
                return False
            time.sleep(0.01)
        return not self._is_stop_requested()

    def _sleep_interruptible(self, seconds: float) -> bool:
        """Sleep in small chunks so stop requests can interrupt long waits."""
        seconds = max(0.0, float(seconds))
        end_time = time.time() + seconds
        while time.time() < end_time:
            if not self._wait_while_paused():
                return False
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
        if not self._wait_while_paused():
            self.last_error = "hourglass wait interrupted"
            return False
        while WindowManager.is_cursor_hourglass():
            if not self._wait_while_paused():
                self.last_error = "hourglass wait interrupted"
                return False
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

    def _wait_for_ready(self) -> bool:
        """Load Control: short settle then wait until the app is no longer busy."""
        if not self.load_control:
            return True
        if not self._sleep_interruptible(0.008):   # 8 ms settle before polling
            return False
        return self._wait_if_hourglass()

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
                if self.use_fast_send:
                    return self._send_data_fast(parsed.data_text)
                return self._send_data(parsed.data_text)

            if parsed.cell_type == CellType.KEYSTROKE:
                if self.use_fast_send:
                    return self._send_keystroke_fast(parsed)
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
            _delay = 0.008 if self.load_control else self.speed_delay
            if self._is_excel_target() and safe_text.isascii():
                pyautogui.typewrite(safe_text, interval=0.01)
                if not self._sleep_interruptible(_delay):
                    self.last_error = "send_data interrupted"
                    return False
                return self._wait_for_ready() if self.load_control else self._wait_if_hourglass()

            # Copy data to clipboard
            pyperclip.copy(safe_text)
            if not self._sleep_interruptible(0.01):
                self.last_error = "send_data interrupted"
                return False

            # Paste into target using Ctrl+V
            pyautogui.hotkey('ctrl', 'v')
            if not self._sleep_interruptible(_delay):
                self.last_error = "send_data interrupted"
                return False

            return self._wait_for_ready() if self.load_control else self._wait_if_hourglass()
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

                _delay = 0.008 if self.load_control else self.speed_delay
                if not self._sleep_interruptible(_delay):
                    self.last_error = "send_keystrokes interrupted"
                    return False
                ready = self._wait_for_ready() if self.load_control else self._wait_if_hourglass()
                if not ready:
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

    # ─── Fast Send (SendInput) methods ────────────────────────────────────────

    def _si_send_unicode(self, text: str) -> bool:
        """Inject unicode text via SendInput KEYEVENTF_UNICODE (no clipboard)."""
        if not text:
            return True
        try:
            events = []
            for ch in text:
                scan = ord(ch)
                events.append(_si_make_ki(0, scan, _KEYEVENTF_UNICODE))
                events.append(_si_make_ki(0, scan, _KEYEVENTF_UNICODE | _KEYEVENTF_KEYUP))
            arr = (_INPUT * len(events))(*events)
            sent = ctypes.windll.user32.SendInput(len(events), arr, ctypes.sizeof(_INPUT))
            return sent == len(events)
        except Exception as e:
            self.last_error = f"_si_send_unicode: {e}"
            return False

    def _si_send_vk(self, vk: int) -> bool:
        """Send a virtual key down+up pair via SendInput."""
        try:
            ext = _KEYEVENTF_EXTENDEDKEY if vk in _SI_EXTENDED_KEYS else 0
            events = [
                _si_make_ki(vk, 0, ext),
                _si_make_ki(vk, 0, ext | _KEYEVENTF_KEYUP),
            ]
            arr = (_INPUT * 2)(*events)
            sent = ctypes.windll.user32.SendInput(2, arr, ctypes.sizeof(_INPUT))
            return sent == 2
        except Exception as e:
            self.last_error = f"_si_send_vk: {e}"
            return False

    def _si_send_hotkey(self, modifiers: list, key: str) -> bool:
        """Send a modifier+key combo (e.g. Ctrl+S) via SendInput."""
        try:
            events = []
            mod_vks = [_SI_VK_MAP[m.lower()] for m in modifiers if m.lower() in _SI_VK_MAP]
            key_vk = _SI_VK_MAP.get(key, 0)
            if not key_vk and len(key) == 1:
                ctypes.windll.user32.VkKeyScanW.argtypes = [ctypes.c_wchar]
                vks_result = ctypes.windll.user32.VkKeyScanW(key)
                if vks_result != -1:
                    key_vk = vks_result & 0xFF
                    if (vks_result >> 8) & 0x01 and 0x10 not in mod_vks:
                        mod_vks.append(0x10)  # shift required for this char
            for vk in mod_vks:
                events.append(_si_make_ki(vk, 0, 0))
            if key_vk:
                ext = _KEYEVENTF_EXTENDEDKEY if key_vk in _SI_EXTENDED_KEYS else 0
                events.append(_si_make_ki(key_vk, 0, ext))
                events.append(_si_make_ki(key_vk, 0, ext | _KEYEVENTF_KEYUP))
            for vk in reversed(mod_vks):
                events.append(_si_make_ki(vk, 0, _KEYEVENTF_KEYUP))
            if not events:
                return True
            arr = (_INPUT * len(events))(*events)
            sent = ctypes.windll.user32.SendInput(len(events), arr, ctypes.sizeof(_INPUT))
            return sent == len(events)
        except Exception as e:
            self.last_error = f"_si_send_hotkey: {e}"
            return False

    def _send_data_fast(self, text: str) -> bool:
        """Send plain text via SendInput unicode injection instead of clipboard."""
        try:
            safe = (text or "").replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
            if not safe:
                return True
            if not self._si_send_unicode(safe):
                # Fallback to clipboard if SendInput fails
                return self._send_data(text)
            # In fast_send_row_mode, use minimal 2ms settle per cell;
            # the full delay is applied once at end-of-row instead.
            if self.fast_send_row_mode:
                _delay = 0.002
            elif self.load_control:
                _delay = 0.008
            else:
                _delay = self.speed_delay
            if not self._sleep_interruptible(_delay):
                self.last_error = "send_data_fast interrupted"
                return False
            return self._wait_for_ready() if self.load_control else self._wait_if_hourglass()
        except Exception as e:
            self.last_error = f"send_data_fast: {e}"
            return self._send_data(text)  # fallback

    def _send_keystroke_fast(self, parsed: ParsedCell) -> bool:
        """Send keystroke actions via SendInput instead of pyautogui."""
        try:
            for action in parsed.key_actions:
                action_type = action.get("type", "")
                if action_type == "key":
                    key = str(action.get("key", "")).lower()
                    vk = _SI_VK_MAP.get(key, 0)
                    ok = self._si_send_vk(vk) if vk else self._si_send_unicode(key)
                    if not ok:
                        return False
                elif action_type == "hotkey":
                    ok = self._si_send_hotkey(
                        action.get("modifiers", []),
                        str(action.get("key", "")).lower(),
                    )
                    if not ok:
                        return False
                elif action_type == "type":
                    if not self._si_send_unicode(action.get("text", "")):
                        return False
                # In fast_send_row_mode, use minimal 2ms settle per cell;
                # the full delay is applied once at end-of-row instead.
                if self.fast_send_row_mode:
                    _delay = 0.002
                elif self.load_control:
                    _delay = 0.008
                else:
                    _delay = self.speed_delay
                if not self._sleep_interruptible(_delay):
                    self.last_error = "send_keystroke_fast interrupted"
                    return False
                ready = self._wait_for_ready() if self.load_control else self._wait_if_hourglass()
                if not ready:
                    return False
            return True
        except Exception as e:
            self.last_error = f"send_keystroke_fast: {e}"
            return self._send_keystrokes(parsed)  # fallback
