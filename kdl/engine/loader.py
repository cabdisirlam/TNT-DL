"""
KDL Loader Engine
Main loading orchestrator that iterates through spreadsheet rows,
parses each cell, and sends data to the target application.
Runs in a QThread to keep the UI responsive.

Supports two modes:
  - Cell Mode: processes each cell as-is (data, keystrokes, shortcuts)
  - Form Mode: treats each row as a complete form entry, auto-TABs
    between columns and sends end-of-row action (New Record/Save/Enter)
"""

import time
import re
import ctypes
import pyautogui
from pyautogui import FailSafeException
from PySide6.QtCore import QThread, Signal, QMutex, QWaitCondition

from kdl.engine.keystroke_parser import KeystrokeParser, CellType, ParsedCell
from kdl.engine.data_sender import DataSender
from kdl.window.window_manager import WindowManager

try:
    import keyboard
except Exception:  # pragma: no cover - environment dependent
    keyboard = None

try:
    from pynput import keyboard as pk
except Exception:  # pragma: no cover - environment dependent
    pk = None

class LoaderThread(QThread):
    """Background thread that loads data into the target application."""

    _FORM_DATA_TOKEN_RE = re.compile(
        r"(?i)(?:(?<=^)|(?<=[,\s;|]))(?:\*[A-Z0-9_]+|\\\{[^{}]+\}|tab|enter|dn|down|up|left|right|esc|escape)(?=$|(?=[,\s;|]))"
    )
    _FORM_HEADER_ALIASES = {
        "type": {"type", "line type", "transaction type", "trx type"},
        "code": {"code", "transaction code", "trx code"},
        "no": {"no", "no.", "line", "line no", "line number", "serial", "serial no"},
    }

    # Signals
    progress_updated = Signal(int, int, str)  # current_row, total_rows, status_msg
    cell_processed = Signal(int, int, bool)   # row, col, success
    loading_complete = Signal(bool, str)       # success, message
    row_started = Signal(int)                  # row_number
    step_waiting = Signal(int, int)            # row, col - waiting for user in step mode
    popup_paused = Signal(str)                  # popup_title - emitted when loader pauses on popup

    def __init__(self, parent=None):
        super().__init__(parent)
        self.parser = KeystrokeParser()
        self.sender = DataSender()

        # Loading parameters
        self.grid_data = []        # 2D list of cell values
        self.start_row = 0
        self.end_row = 0
        self.key_columns = set()   # Columns that are key/keystroke columns

        # Form Mode parameters
        self.form_mode = False     # If True, auto-Tab between cols, action at end of row
        self.load_mode = "per_cell"
        self.end_of_row_action = "none"  # What to do at end of each row

        # Control flags
        self._stop_requested = False
        self._pause_requested = False
        self._step_mode = False
        self._step_advance_requested = False
        self._stop_reason = ""
        self._mutex = QMutex()
        self._wait_condition = QWaitCondition()
        self.db_settings = {}
        self._popup_auto_pause_enabled = True
        self._popup_stop_on_error = False  # True = stop on popup, False = pause
        self._last_popup_check_at = 0.0
        self._last_popup_title = ""
        self._esc_was_down = False
        self._esc_guard_until = 0.0
        self._esc_first_time = 0.0       # timestamp of first ESC press (double-press gate)
        self._esc_hook = None
        self._esc_hook_backend = ""
        self._cell_send_retries = 0  # Blind retries are unsafe: a failed send may have
                                     # already typed/pasted into IFMIS, so retrying
                                     # duplicates data. Keep at 0; raise only for
                                     # pre-send activation failures, not mid-send ones.
        self._form_type_col = None
        self._form_code_col = None
        self._form_no_col = None
        self._form_first_data_row = 0

    def configure(self, grid_data, start_row, end_row, target_hwnd, target_title,
                  speed_delay=0.2, wait_hourglass=False,
                  key_columns=None, selected_columns=None, delay_columns=None,
                  form_mode=False, load_mode="per_cell", end_of_row_action="none",
                  window_delay=0.1, save_interval=50, db_settings=None,
                  popup_stop_on_error=False, use_fast_send=False,
                  load_control=False):
        """Configure the loader before starting."""
        self.grid_data = grid_data
        self.start_row = start_row
        self.end_row = end_row
        # Step mode is disabled in this build.
        self._step_mode = False

        self.sender.set_target(target_hwnd, target_title)
        self.sender.set_speed(speed_delay)
        self.sender.set_window_delay(window_delay)
        self.sender.load_control = bool(load_control)
        self.sender.wait_for_hourglass = wait_hourglass or bool(load_control)
        self.sender.set_stop_checker(self._is_stop_requested)
        self.sender.set_pause_checker(self._is_pause_requested)

        self.key_columns = set(key_columns) if key_columns else set()
        self.selected_columns = set(selected_columns) if selected_columns else None
        self.delay_columns = set(delay_columns) if delay_columns else set()

        # Form mode
        self.form_mode = form_mode
        self.load_mode = str(load_mode or "per_cell").strip().lower()
        self.end_of_row_action = end_of_row_action
        self.save_interval = max(1, int(save_interval or 50))
        self.db_settings = db_settings or {}
        self._popup_stop_on_error = bool(popup_stop_on_error)
        self.sender.use_fast_send = bool(use_fast_send)
        self.sender.fast_send_row_mode = (self.load_mode == "fast_send")

        self._stop_requested = False
        self._pause_requested = False
        self._step_advance_requested = False
        self._stop_reason = ""
        self._last_popup_check_at = 0.0
        self._last_popup_title = ""
        self._esc_was_down = False
        self._esc_guard_until = 0.0
        self._esc_first_time = 0.0
        self._esc_hook = None
        self._esc_hook_backend = ""
        self._detect_form_business_columns()

    def stop(self):
        """Request the loader to stop."""
        self._stop_requested = True
        if not self._stop_reason:
            self._stop_reason = "user request"
        self._pause_requested = False
        self._mutex.lock()
        self._step_advance_requested = True
        self._wait_condition.wakeAll()
        self._mutex.unlock()

    def pause(self):
        """Pause the loader."""
        self._pause_requested = True

    def resume(self):
        """Resume the loader from pause or step-by-step wait."""
        self._pause_requested = False
        self._mutex.lock()
        self._step_advance_requested = True
        self._wait_condition.wakeAll()
        self._mutex.unlock()

    _ESC_DOUBLE_WINDOW = 0.5   # seconds: two ESC presses within this window = stop

    def _check_esc_stop(self) -> bool:
        """Check if user pressed ESC twice quickly to stop loading."""
        try:
            if time.time() < self._esc_guard_until:
                return False
            # VK_ESCAPE = 0x1B
            is_down = bool(ctypes.windll.user32.GetAsyncKeyState(0x1B) & 0x8000)
            if is_down and not self._esc_was_down:
                now = time.time()
                if self._esc_first_time and (now - self._esc_first_time) <= self._ESC_DOUBLE_WINDOW:
                    # Second press within window â†’ stop
                    self._esc_first_time = 0.0
                    self._stop_requested = True
                    if not self._stop_reason:
                        self._stop_reason = "ESC key"
                    self._esc_was_down = True
                    return True
                else:
                    # First press â€” record time, wait for second
                    self._esc_first_time = now
            self._esc_was_down = is_down
        except Exception:
            pass
        return False

    def _request_esc_stop(self):
        """Stop the load from a global ESC hook callback (requires 2 quick presses)."""
        if time.time() < self._esc_guard_until or self._stop_requested:
            return
        now = time.time()
        if self._esc_first_time and (now - self._esc_first_time) <= self._ESC_DOUBLE_WINDOW:
            # Second press â†’ stop
            self._esc_first_time = 0.0
            self._request_stop("ESC key")
            self._pause_requested = False
            self._mutex.lock()
            self._step_advance_requested = True
            self._wait_condition.wakeAll()
            self._mutex.unlock()
        else:
            # First press â€” wait for second
            self._esc_first_time = now

    def _on_keyboard_event(self, event):
        name = str(getattr(event, "name", "") or "").lower().strip()
        if getattr(event, "event_type", "") != "down":
            return
        if name in {"esc", "escape"}:
            self._request_esc_stop()

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
        name = self._normalize_pynput_key(key)
        if name in {"esc", "escape"}:
            self._request_esc_stop()

    def _start_esc_listener(self):
        """Start a best-effort global ESC hook; polling remains as fallback."""
        try:
            if keyboard is not None:
                self._esc_hook = keyboard.hook(self._on_keyboard_event, suppress=False)
                self._esc_hook_backend = "keyboard"
                return
        except Exception as exc:
            self._esc_hook = None
            self._esc_hook_backend = ""

        try:
            if pk is not None:
                self._esc_hook = pk.Listener(on_press=self._on_press_pynput)
                self._esc_hook.start()
                self._esc_hook_backend = "pynput"
                return
        except Exception:
            self._esc_hook = None
            self._esc_hook_backend = ""

        # Keep silent in UI if no hook backend is available; polling remains active.

    def _stop_esc_listener(self):
        """Stop the global ESC hook if one was started."""
        if self._esc_hook is None:
            return
        if self._esc_hook_backend == "keyboard":
            try:
                if keyboard is not None:
                    keyboard.unhook(self._esc_hook)
            except Exception:
                pass
        elif self._esc_hook_backend == "pynput":
            try:
                self._esc_hook.stop()
            except Exception:
                pass
        self._esc_hook = None
        self._esc_hook_backend = ""

    def _is_stop_requested(self) -> bool:
        """Unified stop check for button stop and ESC stop."""
        if self._stop_requested:
            return True
        return self._check_esc_stop()

    def _is_pause_requested(self) -> bool:
        return bool(self._pause_requested)

    def _format_elapsed(self, started_at: float) -> str:
        elapsed_sec = max(0, int(time.time() - started_at))
        mins = elapsed_sec // 60
        secs = elapsed_sec % 60
        return f"{mins} min(s), {secs} sec(s)"

    def _format_eta(self, started_at: float, rows_processed: int, total_rows: int) -> str:
        if rows_processed <= 0:
            return "Calculating..."
        elapsed = time.time() - started_at
        avg = elapsed / rows_processed
        remaining = avg * (total_rows - rows_processed)
        if remaining <= 0:
            return "Done"
        mins, secs = divmod(int(remaining), 60)
        return f"{mins}m {secs}s" if mins else f"{secs}s"

    def _build_stopped_message(self, rows_processed: int, total_rows: int, started_at: float) -> str:
        reason = self._stop_reason or "user request"
        elapsed = self._format_elapsed(started_at)
        return (
            "Load stopped.\n"
            f"Reason: {reason}\n"
            f"Rows loaded: {rows_processed}/{total_rows}\n"
            f"Time: {elapsed}"
        )

    def _request_stop(self, reason: str):
        """Mark stop requested once with a readable reason."""
        self._stop_requested = True
        if not self._stop_reason:
            self._stop_reason = (reason or "user request").strip() or "user request"

    def _interruptible_delay(self, seconds: float, *, wait_hourglass: bool = False) -> bool:
        """
        Delay helper that exits quickly when stop is requested.
        Optionally waits for busy cursor and stops on timeout.
        """
        if not self.sender._sleep_interruptible(seconds):
            return False
        if wait_hourglass and not self.sender._wait_if_hourglass():
            if not self._is_stop_requested():
                reason = (self.sender.last_error or "target remained busy").strip()
                self._request_stop(reason)
            return False
        return not self._is_stop_requested()

    def _wait_after_ui_action(self, fallback_delay: float) -> bool:
        """
        Wait after navigation keys like Tab, Down, Enter, or Save.
        With Load Control on, rely on app readiness instead of fixed delays.
        """
        if self.sender.load_control:
            if not self.sender._wait_for_ready():
                if not self._is_stop_requested():
                    reason = (self.sender.last_error or "target remained busy").strip()
                    self._request_stop(reason or "target remained busy")
                return False
            return True

        if not self._interruptible_delay(fallback_delay, wait_hourglass=True):
            if not self._is_stop_requested():
                reason = (self.sender.last_error or "target remained busy").strip()
                self._request_stop(reason or "target remained busy")
            return False
        return True

    def _fast_send_popup_check_interval(self) -> float:
        if self.sender.fast_send_row_mode:
            return 0.5
        return 0.15

    def run(self):
        """Main orchestrator for loading."""
        try:
            self._esc_guard_until = time.time() + 0.2
            try:
                self._esc_was_down = bool(ctypes.windll.user32.GetAsyncKeyState(0x1B) & 0x8000)
            except Exception:
                self._esc_was_down = False
            self._start_esc_listener()
            if self.sender.uses_clipboard_session():
                try:
                    self.sender.begin_clipboard_session()
                except Exception as e:
                    self.loading_complete.emit(False, f"Loading failed: clipboard unavailable ({e})")
                    return
            mode = self.db_settings.get("mode", "ui_automation")
            if mode != "ui_automation":
                self.progress_updated.emit(
                    0, 1,
                    "Oracle Direct mode is disabled in this build; using UI Automation."
                )
            self._run_ui_mode()
        except FailSafeException:
            # Mouse moved to screen corner (fail-safe)
            self.loading_complete.emit(
                False,
                "Loading stopped: you moved the mouse to a screen corner (fail-safe)."
            )
        except Exception as e:
            self.loading_complete.emit(False, f"Loading failed: {str(e)}")
        finally:
            self._stop_esc_listener()
            self.sender.end_clipboard_session()

    def _run_db_mode(self):
        """Deprecated path: Oracle direct mode is intentionally disabled."""
        self.loading_complete.emit(False, "Oracle Direct mode is disabled in this build.")

    def _send_cell_with_retry(self, parsed: ParsedCell) -> bool:
        # Retries are intentionally disabled.
        # A send failure may have already typed/pasted into IFMIS; retrying
        # blindly would duplicate or corrupt the field.  The popup-pause
        # mechanism handles the real recovery path.
        if self._is_stop_requested():
            return False
        return self.sender.send_cell(parsed)

    def _parse_cell_for_column(self, col_idx: int, cell_value, is_delay: bool = False) -> ParsedCell:
        """
        Parse by column context:
        - Delay columns always parse as delay-aware.
        - All other columns use automatic parser detection.
        Key columns are optional visual aids; parsing is auto-detected.
        """
        if is_delay:
            return self.parser.parse_cell(cell_value, is_delay_column=True)
        return self.parser.parse_cell(cell_value, is_delay_column=False)

    @staticmethod
    def _normalize_header_text(value) -> str:
        text = str(value or "").strip().lower()
        text = re.sub(r"[^a-z0-9]+", " ", text)
        return re.sub(r"\s+", " ", text).strip()

    def _detect_form_business_columns(self):
        """
        Best-effort detection for table-style IFMIS sheets, e.g:
        No | Type | Number | Date | Date | Amount
        """
        self._form_type_col = None
        self._form_code_col = None
        self._form_no_col = None
        self._form_first_data_row = self.start_row

        if not self.form_mode or not self.grid_data:
            return

        # Header-based detection (preferred).
        header = self.grid_data[0] if self.grid_data else []
        normalized = [self._normalize_header_text(v) for v in header]
        for idx, name in enumerate(normalized):
            if not name:
                continue
            if self._form_type_col is None and name in self._FORM_HEADER_ALIASES["type"]:
                self._form_type_col = idx
            if self._form_code_col is None and name in self._FORM_HEADER_ALIASES["code"]:
                self._form_code_col = idx
            if self._form_no_col is None and name in self._FORM_HEADER_ALIASES["no"]:
                self._form_no_col = idx

        if self._form_type_col is not None or self._form_code_col is not None or self._form_no_col is not None:
            self._form_first_data_row = max(self.start_row, 1)
            return

        # Fallback: common no-header layout where col0=No and col1=Type.
        check_end = min(self.end_row, len(self.grid_data) - 1)
        type_hits = 0
        for row_idx in range(self.start_row, check_end + 1):
            row = self.grid_data[row_idx] if row_idx < len(self.grid_data) else []
            if len(row) < 2:
                continue
            type_raw = str(row[1] if row[1] is not None else "").strip().lower()
            if type_raw in {"payment", "receipt", "p", "r"}:
                type_hits += 1
                if type_hits >= 1:
                    self._form_no_col = 0
                    self._form_type_col = 1
                    self._form_code_col = 2
                    break

    def _form_field_extra_settle(self, col_idx: int) -> float:
        if not self.form_mode or self.sender.fast_send_row_mode:
            return 0.0
        if self._form_code_col is not None and col_idx == self._form_code_col:
            return max(float(self.sender.speed_delay), 0.03)
        return 0.0

    def _apply_form_business_rules(self, row_idx: int, col_idx: int, parsed: ParsedCell) -> ParsedCell:
        """
        Business rules for Per Row table mode:
        - Type column:
          - Receipt/r => type the full word "Receipt" directly into the LOV field.
          - Payment/p => type the full word "Payment" directly into the LOV field.
          Both become DATA cells so auto-tab handles field navigation normally.
        - No/Line column: numeric serials are treated as Tab (skip auto-numbered field).
        """
        if not self.form_mode or parsed.cell_type != CellType.DATA:
            return parsed

        text = (parsed.data_text or "").strip()
        if not text:
            return parsed
        lowered = text.lower()

        if self._form_type_col is not None and col_idx == self._form_type_col:
            if lowered in {"receipt", "r"}:
                return ParsedCell(
                    cell_type=CellType.DATA,
                    raw_value=parsed.raw_value,
                    data_text="Receipt",
                )
            if lowered in {"payment", "p"}:
                return ParsedCell(
                    cell_type=CellType.DATA,
                    raw_value=parsed.raw_value,
                    data_text="Payment",
                )
            return parsed

        # Fallback: only reinterpret p/r/payment/receipt as IFMIS type tokens when
        # header-based detection actually ran and confirmed a type column exists
        # (self._form_type_col is not None means detection succeeded but col_idx
        # just didn't match â€” so skip the fallback).
        # When detection found nothing (_form_type_col is None AND _form_no_col is
        # None), it means the sheet layout is unknown; applying the fallback silently
        # to any sheet with "r"/"p" in an early column is dangerous for non-IFMIS
        # layouts and is therefore disabled.
        # Users who need this behaviour should add a "Type" header to their sheet.

        if (
            self._form_no_col is not None
            and col_idx == self._form_no_col
            and re.fullmatch(r"\d+", lowered)
        ):
            return ParsedCell(
                cell_type=CellType.KEYSTROKE,
                raw_value=parsed.raw_value,
                key_actions=[{"type": "key", "key": "tab"}],
            )

        return parsed

    def _sanitize_form_data_cell(self, parsed: ParsedCell) -> ParsedCell:
        """
        Per-row mode guard:
        remove pasted tab characters and standalone macro-like navigation tokens
        from DATA cells so they are not typed into target fields.
        """
        if parsed.cell_type != CellType.DATA:
            return parsed

        original = parsed.data_text or ""
        if not original:
            return parsed

        # Remove literal tab characters copied from TSV/Excel text blobs.
        cleaned = original.replace("\t", " ")
        # Remove standalone tokens such as *DN, tab, \{TAB}, etc.
        cleaned = self._FORM_DATA_TOKEN_RE.sub(" ", cleaned)
        cleaned = re.sub(r"\s+", " ", cleaned).strip(" ,;|")

        if not cleaned:
            return ParsedCell(cell_type=CellType.EMPTY, raw_value=parsed.raw_value)
        if cleaned == original:
            return parsed

        return ParsedCell(
            cell_type=CellType.DATA,
            raw_value=parsed.raw_value,
            data_text=cleaned,
        )

    @staticmethod
    def _typed_text_from_keystroke(parsed: ParsedCell) -> str:
        """
        Return concatenated text for pure type actions.
        Any non-type key action is treated as non-pasteable and returns empty.
        """
        if parsed.cell_type != CellType.KEYSTROKE:
            return ""
        out = []
        for action in parsed.key_actions:
            if action.get("type") != "type":
                return ""
            out.append(str(action.get("text", "")))
        return "".join(out)

    def _build_row_paste_payload(self, row_idx: int, row_data: list) -> tuple[str, bool]:
        """
        Build tab-delimited text for whole-row paste mode.
        Returns (payload, ok). ok=False means interrupted.
        """
        row_tokens: list[str] = []
        for col_idx, cell_value in enumerate(row_data):
            if self.selected_columns is not None and col_idx not in self.selected_columns:
                continue

            is_delay = col_idx in self.delay_columns
            parsed = self._parse_cell_for_column(col_idx, cell_value, is_delay=is_delay)
            parsed = self._apply_form_business_rules(row_idx, col_idx, parsed)
            parsed = self._sanitize_form_data_cell(parsed)

            if parsed.cell_type == CellType.DELAY:
                if not self._interruptible_delay(parsed.delay_ms / 1000.0):
                    return "", False
                continue

            if parsed.cell_type == CellType.EMPTY:
                row_tokens.append("")
                continue

            if parsed.cell_type == CellType.DATA:
                row_tokens.append(parsed.data_text or "")
                continue

            if parsed.cell_type == CellType.KEYSTROKE:
                row_tokens.append(self._typed_text_from_keystroke(parsed))
                continue

            row_tokens.append("")

        while row_tokens and not str(row_tokens[-1]).strip():
            row_tokens.pop()

        return "\t".join(row_tokens), True

    def _perform_end_of_row_action(self, rows_processed: int, is_last_row: bool = False) -> bool:
        eor_delay = max(0.0, float(self.sender.speed_delay))
        save_settle = max(0.1, float(self.sender.speed_delay))
        if self.end_of_row_action == "new_record":
            pyautogui.press('down')
            return self._wait_after_ui_action(eor_delay)
        elif self.end_of_row_action in ("new_record_save_n", "new_record_save_50"):
            completed_rows = rows_processed + 1
            interval = self.save_interval if self.end_of_row_action == "new_record_save_n" else 50
            if completed_rows % interval == 0 or is_last_row:
                if is_last_row and not self._wait_after_ui_action(0.5):
                    return False
                pyautogui.hotkey('ctrl', 's')
                if not self._wait_after_ui_action(save_settle):
                    return False
            pyautogui.press('down')
            return self._wait_after_ui_action(eor_delay)
        elif self.end_of_row_action == "save_proceed":
            if is_last_row and not self._wait_after_ui_action(0.5):
                return False
            pyautogui.hotkey('ctrl', 's')
            if not self._wait_after_ui_action(save_settle):
                return False
            pyautogui.press('down')
            return self._wait_after_ui_action(eor_delay)
        elif self.end_of_row_action == "enter":
            pyautogui.press('enter')
            return self._wait_after_ui_action(eor_delay)
        elif self.end_of_row_action == "tab":
            pyautogui.press('tab')
            return self._wait_after_ui_action(self.sender.speed_delay)
        elif self.end_of_row_action == "none":
            return self._wait_after_ui_action(0.0)
        return True

    def _run_ui_mode(self):
        """Legacy sequential UI automation load loop."""
        total_rows = max(0, self.end_row - self.start_row + 1)
        rows_processed = 0
        errors = 0
        partial_row_loaded = False
        started_at = time.time()
        if total_rows <= 0:
            self.loading_complete.emit(False, "No rows selected to load.")
            return

        self.progress_updated.emit(0, total_rows, f"Starting load... (0/{total_rows})")

        # Initial delay to let user switch to target
        if not self._interruptible_delay(1.5):
            self.loading_complete.emit(False, self._build_stopped_message(0, total_rows, started_at))
            return

        # Activate target window
        if not self.sender.activate_target():
            self.loading_complete.emit(False, "Failed to activate target window")
            return

        # Fast-send mode emits signals at a throttled interval to avoid flooding
        # the Qt main-thread event queue with thousands of queued signals, which
        # would make the overlay appear frozen and stay visible after loading.
        _is_fast = self.sender.fast_send_row_mode
        _FAST_OVERLAY_INTERVAL = 0.15   # seconds between overlay refreshes
        _last_overlay_update = 0.0

        for row_idx in range(self.start_row, self.end_row + 1):
            if self._is_stop_requested():
                break

            self._check_pause()

            if row_idx >= len(self.grid_data):
                break

            row_data = self.grid_data[row_idx]
            row_had_activity = False
            elapsed = self._format_elapsed(started_at)
            eta = self._format_eta(started_at, rows_processed, total_rows)
            _now = time.time()
            _emit_overlay = not _is_fast or (_now - _last_overlay_update) >= _FAST_OVERLAY_INTERVAL
            if _emit_overlay:
                _last_overlay_update = _now
                self.row_started.emit(row_idx)
                self.progress_updated.emit(
                    rows_processed, total_rows,
                    f"Loading row {row_idx + 1}... ({rows_processed}/{total_rows}) | Elapsed: {elapsed} | ETA: {eta}"
                )

            if self._check_blocking_popup(rows_processed, total_rows):
                if self._is_stop_requested():
                    break
                self._check_pause()

            if self.form_mode:
                if not self.sender.activate_target():
                    self.loading_complete.emit(False, "Lost focus on target window - stopped.")
                    return

                if self.load_mode == "imprest_surrender":
                    from kdl.engine.imprest_surrender_engine import (
                        COLUMNS, execute_row_for_loader,
                    )
                    row_dict = {
                        col: (str(row_data[i]).strip()
                              if i < len(row_data) and row_data[i] is not None else "")
                        for i, col in enumerate(COLUMNS)
                    }
                    sup = row_dict.get("Supplier_Num", "")
                    # Auto-detect per-cell (DL macro) format.
                    # When col 0 is a DL macro string the user loaded a per-cell
                    # keystroke row â€” extract the real invoice data from the known
                    # fixed column positions inside that row.
                    if sup.startswith("\\") or (sup.startswith("{") and "}" in sup):
                        _d = [
                            str(row_data[i]).strip()
                            if i < len(row_data) and row_data[i] is not None else ""
                            for i in range(len(row_data))
                        ]
                        # Locate the save command (\^s or \*s) then back-calculate
                        # GL_Date (save-4) and Distribution_Account (save-2).
                        save_idx = next(
                            (i for i, v in enumerate(_d)
                             if v in ("\\^s", "\\*s", "*s")),
                            len(_d)
                        )
                        row_dict = {
                            "Supplier_Num":         _d[10] if len(_d) > 10 else "",
                            "Invoice_Date":         _d[15] if len(_d) > 15 else "",
                            "Invoice_Num":          _d[17] if len(_d) > 17 else "",
                            "Invoice_Amount":       _d[20] if len(_d) > 20 else "",
                            "Description":          _d[28] if len(_d) > 28 else "",
                            "Payment_Method":       _d[34] if len(_d) > 34 else "",
                            "Terms_Date":           "",
                            "Auth_Ref_No":          _d[52] if len(_d) > 52 else "",
                            "Administrative_Code":  _d[54] if len(_d) > 54 else "",
                            "GL_Date":              _d[save_idx - 4] if save_idx >= 4 else "",
                            "Distribution_Account": _d[save_idx - 2] if save_idx >= 2 else "",
                            "Old_Imprest_No":       _d[80] if len(_d) > 80 else "",
                        }
                        sup = row_dict.get("Supplier_Num", "")
                        self.progress_updated.emit(
                            rows_processed, total_rows,
                            f"[Imprest] Row {row_idx + 1} | per-cell fmt"
                            f" | Supplier: {sup or '(empty)'}"
                        )
                    else:
                        self.progress_updated.emit(
                            rows_processed, total_rows,
                            f"[Imprest] Row {row_idx + 1} | Target: {self.sender.target_title!r}"
                            f" | Supplier: {sup or '(empty)'}"
                        )
                    def _imprest_popup_fn(popup_title):
                        """Auto-handle LOV (e.g. Supplier Site): accept the
                        highlighted row and close it.

                        Oracle LOV accept strategy:
                          1. Enter  — works when focus is already on a result row.
                          2. Down + Enter  — when focus is on the Find field (text
                             was typed in it), Down moves focus to the first result
                             row and Enter accepts it.
                          3. Tab + Enter  — last resort before manual pause.
                        Falls back to a manual pause only when all attempts fail.
                        """
                        from kdl.window.window_manager import WindowManager

                        if not self._interruptible_delay(0.15):
                            return False

                        # Attempt 1: plain Enter (works when focus is on result row)
                        self.sender._si_send_vk(0x0D)   # VK_RETURN
                        if not self._wait_after_ui_action(0.40):
                            return False
                        still_open = WindowManager.detect_blocking_popup(
                            self.sender.target_hwnd, self.sender.target_title)
                        if not still_open:
                            self.progress_updated.emit(
                                rows_processed, total_rows,
                                f"[Imprest] Auto-accepted LOV ‘{popup_title}’ → Enter")
                            return not self._is_stop_requested()

                        # Attempt 2: Down Arrow then Enter.
                        # When the LOV Find field has focus (text was typed into it),
                        # Enter only re-queries; Down Arrow moves focus to the first
                        # result row so the following Enter accepts it.
                        self.sender._si_send_vk(0x28)   # VK_DOWN
                        if not self._wait_after_ui_action(0.20):
                            return False
                        self.sender._si_send_vk(0x0D)   # VK_RETURN
                        if not self._wait_after_ui_action(0.40):
                            return False
                        still_open = WindowManager.detect_blocking_popup(
                            self.sender.target_hwnd, self.sender.target_title)
                        if not still_open:
                            self.progress_updated.emit(
                                rows_processed, total_rows,
                                f"[Imprest] Auto-accepted LOV ‘{popup_title}’ → Down+Enter")
                            return not self._is_stop_requested()

                        # Attempt 3: Tab then Enter (alternative focus move)
                        self.sender._si_send_vk(0x09)   # VK_TAB
                        if not self._wait_after_ui_action(0.20):
                            return False
                        self.sender._si_send_vk(0x0D)   # VK_RETURN
                        if not self._wait_after_ui_action(0.40):
                            return False
                        still_open = WindowManager.detect_blocking_popup(
                            self.sender.target_hwnd, self.sender.target_title)
                        if not still_open:
                            self.progress_updated.emit(
                                rows_processed, total_rows,
                                f"[Imprest] Auto-accepted LOV ‘{popup_title}’ → Tab+Enter")
                            return not self._is_stop_requested()

                        # Fallback: manual pause
                        self._pause_requested = True
                        self.popup_paused.emit(popup_title)
                        self.progress_updated.emit(
                            rows_processed, total_rows,
                            f"Paused: popup ‘{popup_title}’ could not be auto-dismissed. "
                            f"Dismiss it then click Resume.")
                        self._check_pause()
                        return not self._is_stop_requested()

                    ok = execute_row_for_loader(
                        self.sender, row_dict, self._is_stop_requested,
                        inter_action_delay=self.sender.speed_delay,
                        is_last_row=(row_idx == self.end_row),
                        popup_fn=_imprest_popup_fn)
                    row_had_activity = True
                    if not ok:
                        if not self._is_stop_requested():
                            self._handle_send_failure(
                                row_idx, 0, rows_processed, total_rows)
                        break

                elif self.load_mode == "per_row_paste":
                    self._check_pause()
                    paste_payload, ok = self._build_row_paste_payload(row_idx, row_data)
                    if not ok:
                        break

                    row_had_activity = bool(paste_payload.strip())
                    if row_had_activity:
                        paste_col_idx = 0
                        for col_idx in range(len(row_data)):
                            if self.selected_columns is not None and col_idx not in self.selected_columns:
                                continue
                            paste_col_idx = col_idx
                            break

                        paste_cell = ParsedCell(
                            cell_type=CellType.DATA,
                            raw_value=paste_payload,
                            data_text=paste_payload,
                        )
                        success = self._send_cell_with_retry(paste_cell)
                        self.cell_processed.emit(row_idx, paste_col_idx, success)
                        if not success:
                            self._handle_send_failure(row_idx, paste_col_idx, rows_processed, total_rows)
                            break
                else:
                    row_cells = []
                    for col_idx, cell_value in enumerate(row_data):
                        if self.selected_columns is not None and col_idx not in self.selected_columns:
                            continue
                        is_delay = col_idx in self.delay_columns
                        parsed = self._parse_cell_for_column(col_idx, cell_value, is_delay=is_delay)
                        parsed = self._apply_form_business_rules(row_idx, col_idx, parsed)
                        parsed = self._sanitize_form_data_cell(parsed)
                        if parsed.cell_type != CellType.EMPTY:
                            row_cells.append((col_idx, parsed))
                    row_had_activity = bool(row_cells)

                    data_positions = [
                        idx for idx, (_, parsed) in enumerate(row_cells)
                        if parsed.cell_type == CellType.DATA
                    ]
                    pending_tab_after_receipt = False

                    for i, (col_idx, parsed) in enumerate(row_cells):
                        if self._is_stop_requested():
                            break

                        if self._check_blocking_popup(rows_processed, total_rows):
                            if self._is_stop_requested():
                                break
                            self._check_pause()
                            if self._is_stop_requested():
                                break

                        self._check_pause()

                        # Receipt flow: after sending type-ahead 'r' in Type column,
                        # move to the next field just before next cell send.
                        # Do NOT send an extra TAB if the current cell is already an
                        # explicit tab keystroke — bank statement rows embed their own
                        # "tab" cells and firing on top produces a double-TAB that
                        # shifts all subsequent fields one column right.
                        _cur_is_tab = (
                            parsed.cell_type == CellType.KEYSTROKE
                            and parsed.key_actions
                            and parsed.key_actions[0].get("key", "").lower() == "tab"
                        )
                        if pending_tab_after_receipt and parsed.cell_type != CellType.EMPTY and not _cur_is_tab:
                            if self.sender.fast_send_row_mode:
                                self.sender._si_send_vk(0x09)  # VK_TAB
                                time.sleep(0.002)
                            else:
                                pyautogui.press('tab')
                                if not self._wait_after_ui_action(self.sender.speed_delay):
                                    break
                        pending_tab_after_receipt = False

                        success = self._send_cell_with_retry(parsed)
                        # In fast-send mode skip per-cell signals for successes to
                        # prevent the Qt event-queue backlog that freezes the overlay.
                        if not _is_fast or not success:
                            self.cell_processed.emit(row_idx, col_idx, success)
                        if not success:
                            self._handle_send_failure(row_idx, col_idx, rows_processed, total_rows)
                            break

                        raw_norm = str(parsed.raw_value or "").strip().lower()
                        receipt_token = False
                        if parsed.cell_type == CellType.KEYSTROKE:
                            if raw_norm in {"receipt", "r", r"\r"}:
                                receipt_token = True
                            elif any(
                                a.get("type") == "type" and str(a.get("text", "")).lower() == "r"
                                for a in parsed.key_actions
                            ):
                                receipt_token = True
                        type_col_hit = self._form_type_col is not None and col_idx == self._form_type_col
                        early_col_fallback = self._form_type_col is None and col_idx <= 2
                        if self.form_mode and receipt_token and (type_col_hit or early_col_fallback):
                            pending_tab_after_receipt = True

                        # The Code field is validated by Oracle Forms on blur. If we
                        # Tab away too quickly after a fast send, the last character in
                        # values like TRFD/TRFC can be lost and the LOV opens to resolve
                        # the now-partial code.
                        extra_settle = self._form_field_extra_settle(col_idx)
                        if self.sender.fast_send_row_mode and extra_settle > 0:
                            if not self._wait_after_ui_action(extra_settle):
                                break

                        # Auto-Tab only between plain data fields AND only when the
                        # immediately following cell is not already an explicit tab
                        # keystroke. Bank statement rows embed explicit "tab" cells
                        # between every field; firing auto-Tab on top of those produces
                        # a double-Tab that shifts all subsequent fields by one column.
                        next_is_tab = (
                            i + 1 < len(row_cells)
                            and row_cells[i + 1][1].cell_type == CellType.KEYSTROKE
                            and row_cells[i + 1][1].key_actions
                            and row_cells[i + 1][1].key_actions[0].get("key", "").lower() == "tab"
                        )
                        if i in data_positions and data_positions and i != data_positions[-1] and not next_is_tab:
                            if self.sender.fast_send_row_mode:
                                self.sender._si_send_vk(0x09)  # VK_TAB
                                time.sleep(0.002)
                            else:
                                pyautogui.press('tab')
                                if not self._wait_after_ui_action(self.sender.speed_delay):
                                    break

                if (
                    self.load_mode != "imprest_surrender"
                    and not self._stop_requested
                    and row_had_activity
                ):
                    if not self._perform_end_of_row_action(rows_processed, is_last_row=(row_idx == self.end_row)):
                        break

            else:
                # CELL MODE (original behavior)
                if not self.sender.activate_target():
                    self.loading_complete.emit(False, "Lost focus on target window - stopped.")
                    return

                for col_idx, cell_value in enumerate(row_data):
                    if self._is_stop_requested():
                        break

                    if self._check_blocking_popup(rows_processed, total_rows):
                        if self._is_stop_requested():
                            break
                        self._check_pause()
                        if self._is_stop_requested():
                            break

                    if self.selected_columns is not None and col_idx not in self.selected_columns:
                        continue

                    self._check_pause()

                    is_delay = col_idx in self.delay_columns
                    parsed = self._parse_cell_for_column(col_idx, cell_value, is_delay=is_delay)

                    if parsed.cell_type == CellType.EMPTY:
                        self.cell_processed.emit(row_idx, col_idx, True)
                        continue

                    row_had_activity = True

                    success = self._send_cell_with_retry(parsed)
                    self.cell_processed.emit(row_idx, col_idx, success)

                    if not success:
                        self._handle_send_failure(row_idx, col_idx, rows_processed, total_rows)
                        break

            if self._is_stop_requested():
                if row_had_activity:
                    partial_row_loaded = True
                break

            rows_processed += 1
            elapsed = self._format_elapsed(started_at)
            eta = self._format_eta(started_at, rows_processed, total_rows)
            _now = time.time()
            if not _is_fast or (_now - _last_overlay_update) >= _FAST_OVERLAY_INTERVAL:
                _last_overlay_update = _now
                self.progress_updated.emit(
                    rows_processed, total_rows,
                    f"Loaded {rows_processed}/{total_rows} row(s) | Elapsed: {elapsed} | ETA: {eta}"
                )

        # Done UI mode – in fast send, emit one final progress signal so the
        # overlay displays the true final row count before it is hidden.
        if _is_fast:
            elapsed = self._format_elapsed(started_at)
            self.progress_updated.emit(
                rows_processed, total_rows,
                f"Loaded {rows_processed}/{total_rows} row(s) | Elapsed: {elapsed} | ETA: --"
            )
        elapsed = self._format_elapsed(started_at)
        if self._stop_requested:
            stop_rows = min(total_rows, rows_processed + (1 if partial_row_loaded else 0))
            self.loading_complete.emit(
                False,
                self._build_stopped_message(stop_rows, total_rows, started_at)
            )
        elif errors > 0:
            self.loading_complete.emit(
                False,
                f"Loading completed with {errors} errors.\n"
                f"Rows: {rows_processed}/{total_rows}\n"
                f"Time: {elapsed}"
            )
        else:
            self.loading_complete.emit(
                True,
                f"Loading completed successfully!\n"
                f"Rows: {rows_processed}\n"
                f"Time: {elapsed}"
            )

    def _handle_send_failure(self, row_idx: int, col_idx: int, rows_processed: int, total_rows: int) -> None:
        """
        Called immediately after send_cell() returns False.
        Checks whether an IFMIS/Oracle popup caused the failure before
        labelling it a generic cell error.  Popup â†’ pause/stop per setting.
        Cell error (no popup) â†’ always stop.
        """
        if self._is_stop_requested():
            return
        # Bypass the 0.15 s throttle so the check runs right now.
        self._last_popup_check_at = 0.0
        if self._check_blocking_popup(rows_processed, total_rows):
            # Popup was the root cause.  Even if the user dismissed it
            # (pause mode), we cannot safely retry the failed cell â€” its
            # data is in an unknown state â€” so stop with a clear message.
            if not self._is_stop_requested():
                self._request_stop(
                    f"popup during send at R{row_idx + 1} C{col_idx + 1} "
                    f"â€” verify data and restart from that row"
                )
            return
        # No popup â€” genuine send failure.
        detail = (self.sender.last_error or "send failed").strip()
        self._request_stop(f"cell error at R{row_idx + 1} C{col_idx + 1}: {detail}")

    def _check_pause(self):
        """Check if paused and wait."""
        self._mutex.lock()
        while self._pause_requested and not self._is_stop_requested():
            self._wait_condition.wait(self._mutex, 30)
        self._mutex.unlock()

    def _check_blocking_popup(self, rows_processed: int, total_rows: int) -> bool:
        """
        Detect IFMIS/Oracle popup dialogs (LOVs, error dialogs, date pickers, etc.).
        Pauses the loader so the user can dismiss the popup and resume.
        Returns True if popup handling was triggered.
        """
        if not self._popup_auto_pause_enabled:
            return False

        now = time.time()
        if now - self._last_popup_check_at < self._fast_send_popup_check_interval():
            return False
        self._last_popup_check_at = now

        popup_title = WindowManager.detect_blocking_popup(
            self.sender.target_hwnd,
            self.sender.target_title,
        )
        if not popup_title:
            self._last_popup_title = ""
            return False

        if self._popup_stop_on_error:
            # Stop mode â€” for unsupervised runs.
            if popup_title != self._last_popup_title:
                self._last_popup_title = popup_title
                self._request_stop(f"popup detected: {popup_title}")
                self.progress_updated.emit(
                    rows_processed,
                    total_rows,
                    f"Stopped: popup '{popup_title}' detected."
                )
            return True

        # Pause mode â€” user dismisses popup and resumes manually.
        self._pause_requested = True
        if popup_title != self._last_popup_title:
            self._last_popup_title = popup_title
            self.popup_paused.emit(popup_title)
            self.progress_updated.emit(
                rows_processed,
                total_rows,
                f"Paused: detected popup '{popup_title}'. Dismiss the popup and click Resume."
            )
        # Wait until user resumes (or stops).
        self._check_pause()
        return True

    def _wait_for_step(self):
        """Wait for user step/resume while still allowing stop to interrupt."""
        self._mutex.lock()
        self._step_advance_requested = False
        while not self._step_advance_requested and not self._is_stop_requested():
            self._wait_condition.wait(self._mutex, 30)
        self._mutex.unlock()
