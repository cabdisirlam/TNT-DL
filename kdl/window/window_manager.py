"""
KDL Window Manager
Enumerates open application windows on Windows 11, activates target windows,
and monitors cursor state. Behaves like FDL/DataLoad Classic window detection.
"""

import ctypes
import ctypes.wintypes
import time
from typing import List, Tuple, Optional


# Windows API constants
SW_RESTORE = 9
SW_SHOW = 5
SW_SHOWNOACTIVATE = 4
WS_VISIBLE = 0x10000000
WS_EX_TOOLWINDOW = 0x00000080
WS_EX_APPWINDOW = 0x00040000
WS_EX_NOACTIVATE = 0x08000000
GWL_EXSTYLE = -20
GWL_STYLE = -16
GW_OWNER = 4
GA_ROOT = 2
GA_ROOTOWNER = 3
IDC_WAIT = 32514
IDC_APPSTARTING = 32650

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32


# System windows to exclude (these show up but aren't real app windows)
EXCLUDED_TITLES = {
    "Program Manager",
    "Windows Input Experience",
    "Windows Shell Experience Host",
    "Microsoft Text Input Application",
    "NVIDIA GeForce Overlay",
    "Search",
    "Start",
    "Task View",
    "Task Switching",
    "Desktop Window Manager",
}

# Partial title matches to exclude
EXCLUDED_PARTIAL = [
    "Antigravity -",  # Our own IDE
]

# Popup title keywords that usually indicate IFMIS/Oracle validation dialogs.
POPUP_KEYWORDS = [
    "error", "warning", "invalid", "failed", "failure", "cannot", "can't",
    "required", "mandatory", "exception", "message", "alert", "not allowed",
]


class WindowManager:
    """
    Manages target window enumeration, activation, and monitoring.
    Detects open windows the same way FDL / DataLoad Classic do —
    listing all visible application windows with their full titles.
    """

    @staticmethod
    def get_open_windows() -> List[Tuple[int, str]]:
        """
        Get a list of all open, visible APPLICATION windows with titles.
        Filters out system/tool windows to match FDL/DataLoad behavior.
        Returns list of (hwnd, title) tuples sorted alphabetically.
        """
        windows = []

        def _is_app_window(hwnd) -> bool:
            """Check if a window is a real application window (not a tool/system window)."""
            # Must be visible
            if not user32.IsWindowVisible(hwnd):
                return False

            # Get extended style
            ex_style = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)

            # Skip tool windows (floating toolbars, etc.) unless they are app windows
            if ex_style & WS_EX_TOOLWINDOW and not (ex_style & WS_EX_APPWINDOW):
                return False

            # Skip windows with no-activate style (system overlay windows)
            if ex_style & WS_EX_NOACTIVATE:
                return False

            # Must have no owner (top-level window) OR be an app window
            owner = user32.GetWindow(hwnd, GW_OWNER)
            if owner and not (ex_style & WS_EX_APPWINDOW):
                return False

            # Must have a title
            length = user32.GetWindowTextLengthW(hwnd)
            if length <= 0:
                return False

            return True

        def enum_callback(hwnd, _):
            if _is_app_window(hwnd):
                length = user32.GetWindowTextLengthW(hwnd)
                buf = ctypes.create_unicode_buffer(length + 1)
                user32.GetWindowTextW(hwnd, buf, length + 1)
                title = buf.value.strip()

                if title:
                    # Exclude known system windows
                    if title in EXCLUDED_TITLES:
                        return True

                    # Exclude partial matches
                    for partial in EXCLUDED_PARTIAL:
                        if partial in title:
                            return True

                    windows.append((hwnd, title))
            return True

        WNDENUMPROC = ctypes.WINFUNCTYPE(
            ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM
        )
        # Hold reference to prevent GC during enumeration
        cb = WNDENUMPROC(enum_callback)
        user32.EnumWindows(cb, 0)

        # Sort alphabetically by title (like FDL/DataLoad Classic)
        windows.sort(key=lambda w: w[1].lower())

        return windows

    @staticmethod
    def find_window_by_title(title: str) -> Optional[int]:
        """Find a window handle by its exact title."""
        windows = WindowManager.get_open_windows()
        for hwnd, wtitle in windows:
            if wtitle == title:
                return hwnd
        return None

    @staticmethod
    def find_window_containing_title(partial_title: str) -> Optional[int]:
        """Find a window handle by partial title match (case-insensitive)."""
        windows = WindowManager.get_open_windows()
        partial_lower = partial_title.lower()
        for hwnd, wtitle in windows:
            if partial_lower in wtitle.lower():
                return hwnd
        return None

    @staticmethod
    def find_oracle_windows() -> List[Tuple[int, str]]:
        """
        Find Oracle Applications / IFMIS windows specifically.
        Looks for common Oracle Forms window title patterns.
        """
        oracle_keywords = [
            "oracle", "ifmis", "gok", "forms", "navigator",
            "responsibility", "applications", "e-business",
        ]
        windows = WindowManager.get_open_windows()
        results = []
        for hwnd, title in windows:
            title_lower = title.lower()
            for keyword in oracle_keywords:
                if keyword in title_lower:
                    results.append((hwnd, title))
                    break
        return results

    @staticmethod
    def activate_window(hwnd: int) -> bool:
        """
        Bring a window to the foreground and activate it.
        Uses AttachThreadInput trick for reliable foreground activation,
        which is needed for Oracle Forms running in Java/JInitiator.
        Returns True if successful.
        """
        try:
            # Restore if minimized
            if user32.IsIconic(hwnd):
                user32.ShowWindow(hwnd, SW_RESTORE)

            # Get thread IDs
            foreground_hwnd = user32.GetForegroundWindow()
            current_thread = kernel32.GetCurrentThreadId()
            target_thread = user32.GetWindowThreadProcessId(hwnd, None)
            foreground_thread = user32.GetWindowThreadProcessId(foreground_hwnd, None)

            # Attach input threads for reliable foreground switch
            # This is the same technique FDL/DataLoad use
            attached_current = False
            attached_foreground = False
            try:
                if current_thread != target_thread:
                    user32.AttachThreadInput(current_thread, target_thread, True)
                    attached_current = True
                if foreground_thread != target_thread:
                    user32.AttachThreadInput(foreground_thread, target_thread, True)
                    attached_foreground = True

                # Bring to front
                user32.BringWindowToTop(hwnd)
                user32.ShowWindow(hwnd, SW_SHOW)
                user32.SetForegroundWindow(hwnd)

                # Set focus
                user32.SetFocus(hwnd)
            finally:
                # Always detach threads to avoid corrupting Windows input routing
                if attached_current:
                    user32.AttachThreadInput(current_thread, target_thread, False)
                if attached_foreground:
                    user32.AttachThreadInput(foreground_thread, target_thread, False)

            # Small delay for window to become active
            time.sleep(0.15)

            return True
        except Exception:
            # Fallback: simple activation
            try:
                user32.SetForegroundWindow(hwnd)
                time.sleep(0.1)
                return True
            except Exception:
                return False

    @staticmethod
    def is_cursor_hourglass() -> bool:
        """
        Check if the current cursor is an hourglass/busy cursor.
        Detects both the wait cursor and the app-starting cursor.
        Used for the 'Wait if Cursor is Hour Glass' feature.
        """
        try:
            class CURSORINFO(ctypes.Structure):
                _fields_ = [
                    ("cbSize", ctypes.c_uint),
                    ("flags", ctypes.c_uint),
                    ("hCursor", ctypes.c_void_p),
                    ("ptScreenPos", ctypes.wintypes.POINT),
                ]

            ci = CURSORINFO()
            ci.cbSize = ctypes.sizeof(CURSORINFO)

            if user32.GetCursorInfo(ctypes.byref(ci)):
                # Check for wait cursor (hourglass)
                wait_cursor = user32.LoadCursorW(None, IDC_WAIT)
                # Check for app-starting cursor (arrow + hourglass)
                appstart_cursor = user32.LoadCursorW(None, IDC_APPSTARTING)

                return ci.hCursor == wait_cursor or ci.hCursor == appstart_cursor

            return False
        except Exception:
            return False

    @staticmethod
    def get_foreground_window_title() -> str:
        """Get the title of the currently active foreground window."""
        hwnd = user32.GetForegroundWindow()
        if hwnd:
            length = user32.GetWindowTextLengthW(hwnd)
            if length > 0:
                buf = ctypes.create_unicode_buffer(length + 1)
                user32.GetWindowTextW(hwnd, buf, length + 1)
                return buf.value
        return ""

    @staticmethod
    def get_foreground_window_handle() -> int:
        """Get handle of current foreground window."""
        try:
            return int(user32.GetForegroundWindow() or 0)
        except Exception:
            return 0

    @staticmethod
    def get_window_title(hwnd: int) -> str:
        """Get title text for a specific window handle."""
        try:
            if not hwnd:
                return ""
            length = user32.GetWindowTextLengthW(hwnd)
            if length <= 0:
                return ""
            buf = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buf, length + 1)
            return buf.value
        except Exception:
            return ""

    @staticmethod
    def get_window_process_id(hwnd: int) -> int:
        """Get process id owning a window handle."""
        try:
            if not hwnd:
                return 0
            pid = ctypes.wintypes.DWORD()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            return int(pid.value)
        except Exception:
            return 0

    @staticmethod
    def get_window_class_name(hwnd: int) -> str:
        """Get Win32 class name for a window."""
        try:
            if not hwnd:
                return ""
            buf = ctypes.create_unicode_buffer(256)
            if user32.GetClassNameW(hwnd, buf, 256):
                return buf.value
        except Exception:
            pass
        return ""

    @staticmethod
    def detect_blocking_popup(target_hwnd: int, target_title: str = "") -> str:
        """
        Detect a likely blocking popup dialog belonging to the same process as target.
        Returns popup title if detected, else empty string.
        """
        try:
            if not target_hwnd:
                return ""

            target_pid = WindowManager.get_window_process_id(target_hwnd)
            if not target_pid:
                return ""
            target_root = user32.GetAncestor(target_hwnd, GA_ROOTOWNER)

            fg_hwnd = WindowManager.get_foreground_window_handle()
            if fg_hwnd and fg_hwnd != target_hwnd and user32.IsWindowVisible(fg_hwnd):
                fg_pid = WindowManager.get_window_process_id(fg_hwnd)
                if fg_pid == target_pid:
                    title = WindowManager.get_window_title(fg_hwnd).strip()
                    lowered = title.lower()
                    owner = user32.GetWindow(fg_hwnd, GW_OWNER)
                    root_owner = user32.GetAncestor(fg_hwnd, GA_ROOTOWNER)
                    class_name = WindowManager.get_window_class_name(fg_hwnd)
                    class_lower = class_name.lower()
                    is_dialog_class = class_name == "#32770" or "dialog" in class_lower
                    is_owned_popup = bool(owner) or (
                        root_owner and target_root and root_owner == target_root and fg_hwnd != target_root
                    )
                    keyword_hit = any(k in lowered for k in POPUP_KEYWORDS)

                    # Treat as blocking when it's a dialog/keyword match, or an owned popup with visible title.
                    if is_dialog_class or keyword_hit or (is_owned_popup and bool(title)):
                        # Avoid reporting the exact same target title as popup.
                        if target_title and title and target_title.strip().lower() == lowered:
                            return ""
                        return title or f"(popup: {class_name or 'unknown class'})"

            # Fallback: scan all visible top-level windows in same process for modal/popups.
            found_popup = ""

            def enum_callback(hwnd, _):
                nonlocal found_popup
                if found_popup:
                    return False
                if hwnd == target_hwnd:
                    return True
                if not user32.IsWindowVisible(hwnd):
                    return True
                if WindowManager.get_window_process_id(hwnd) != target_pid:
                    return True

                title = WindowManager.get_window_title(hwnd).strip()
                lowered = title.lower()
                class_name = WindowManager.get_window_class_name(hwnd)
                class_lower = class_name.lower()
                owner = user32.GetWindow(hwnd, GW_OWNER)
                root_owner = user32.GetAncestor(hwnd, GA_ROOTOWNER)

                is_dialog_class = class_name == "#32770" or "dialog" in class_lower
                is_owned_popup = bool(owner) or (
                    root_owner and target_root and root_owner == target_root and hwnd != target_root
                )
                keyword_hit = any(k in lowered for k in POPUP_KEYWORDS)
                same_title = bool(target_title and title and target_title.strip().lower() == lowered)

                if (is_dialog_class or keyword_hit or (is_owned_popup and bool(title))) and not same_title:
                    found_popup = title or f"(popup: {class_name or 'unknown class'})"
                    return False
                return True

            WNDENUMPROC = ctypes.WINFUNCTYPE(
                ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM
            )
            cb = WNDENUMPROC(enum_callback)
            user32.EnumWindows(cb, 0)
            if found_popup:
                return found_popup

            # If target becomes disabled, Oracle often has a modal child dialog open.
            if not user32.IsWindowEnabled(target_hwnd):
                child_popup = WindowManager._detect_child_dialog(target_hwnd, target_title)
                if child_popup:
                    return child_popup
                return "(modal dialog)"
        except Exception:
            return ""

        return ""

    @staticmethod
    def _detect_child_dialog(parent_hwnd: int, target_title: str = "") -> str:
        """Best-effort detection of visible child dialogs inside the target window."""
        try:
            found = ""
            target_lower = (target_title or "").strip().lower()

            def enum_child_callback(hwnd, _):
                nonlocal found
                if found:
                    return False
                if not user32.IsWindowVisible(hwnd):
                    return True

                class_name = WindowManager.get_window_class_name(hwnd)
                class_lower = class_name.lower()
                title = WindowManager.get_window_title(hwnd).strip()
                lowered = title.lower()

                is_dialog_class = class_name == "#32770" or "dialog" in class_lower
                keyword_hit = any(k in lowered for k in POPUP_KEYWORDS)
                same_title = bool(target_lower and lowered and target_lower == lowered)

                if (is_dialog_class or keyword_hit) and not same_title:
                    found = title or f"(child dialog: {class_name or 'unknown class'})"
                    return False
                return True

            ENUMCHILDPROC = ctypes.WINFUNCTYPE(
                ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM
            )
            cb = ENUMCHILDPROC(enum_child_callback)
            user32.EnumChildWindows(parent_hwnd, cb, 0)
            return found
        except Exception:
            return ""

    @staticmethod
    def get_window_process_name(hwnd: int) -> str:
        """Get the process name of a window (e.g., 'chrome.exe', 'java.exe')."""
        try:
            pid = ctypes.wintypes.DWORD()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))

            PROCESS_QUERY_INFORMATION = 0x0400
            PROCESS_VM_READ = 0x0010

            h_process = kernel32.OpenProcess(
                PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, False, pid.value
            )
            if h_process:
                try:
                    buf = ctypes.create_unicode_buffer(260)
                    psapi = ctypes.windll.psapi
                    psapi.GetModuleBaseNameW(h_process, None, buf, 260)
                    return buf.value
                finally:
                    kernel32.CloseHandle(h_process)
        except Exception:
            pass
        return ""
