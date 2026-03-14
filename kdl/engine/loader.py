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
        "no": {"no", "no.", "line", "line no", "line number", "serial", "serial no"},
    }

    # Signals
    progress_updated = Signal(int, int, str)  # current_row, total_rows, status_msg
    cell_processed = Signal(int, int, bool)   # row, col, success
    loading_complete = Signal(bool, str)       # success, message
    row_started = Signal(int)                  # row_number
    step_waiting = Signal(int, int)            # row, col - waiting for user in step mode

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
        self._last_popup_check_at = 0.0
        self._last_popup_title = ""
        self._esc_was_down = False
        self._esc_guard_until = 0.0
        self._esc_hook = None
        self._esc_hook_backend = ""
        self._cell_send_retries = 2
        self._form_type_col = None
        self._form_no_col = None
        self._form_first_data_row = 0

    def configure(self, grid_data, start_row, end_row, target_hwnd, target_title,
                  speed_delay=0.1, wait_hourglass=False,
                  key_columns=None, selected_columns=None, delay_columns=None,
                  form_mode=False, load_mode="per_cell", end_of_row_action="none",
                  window_delay=0.1, db_settings=None):
        """Configure the loader before starting."""
        self.grid_data = grid_data
        self.start_row = start_row
        self.end_row = end_row
        # Step mode is disabled in this build.
        self._step_mode = False

        self.sender.set_target(target_hwnd, target_title)
        self.sender.set_speed(speed_delay)
        self.sender.set_window_delay(window_delay)
        self.sender.wait_for_hourglass = wait_hourglass
        self.sender.set_stop_checker(self._is_stop_requested)

        self.key_columns = set(key_columns) if key_columns else set()
        self.selected_columns = set(selected_columns) if selected_columns else None
        self.delay_columns = set(delay_columns) if delay_columns else set()

        # Form mode
        self.form_mode = form_mode
        self.load_mode = str(load_mode or "per_cell").strip().lower()
        self.end_of_row_action = end_of_row_action
        self.db_settings = db_settings or {}

        self._stop_requested = False
        self._pause_requested = False
        self._step_advance_requested = False
        self._stop_reason = ""
        self._last_popup_check_at = 0.0
        self._last_popup_title = ""
        self._esc_was_down = False
        self._esc_guard_until = 0.0
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

    def _check_esc_stop(self) -> bool:
        """Check if user pressed ESC (single press) to stop loading."""
        try:
            if time.time() < self._esc_guard_until:
                return False
            # VK_ESCAPE = 0x1B
            is_down = bool(ctypes.windll.user32.GetAsyncKeyState(0x1B) & 0x8000)
            if is_down and not self._esc_was_down:
                self._stop_requested = True
                if not self._stop_reason:
                    self._stop_reason = "ESC key"
                self._esc_was_down = True
                return True
            self._esc_was_down = is_down
        except Exception:
            pass
        return False

    def _request_esc_stop(self):
        """Stop the load from a global ESC hook callback."""
        if time.time() < self._esc_guard_until or self._stop_requested:
            return
        self._request_stop("ESC key")
        self._pause_requested = False
        self._mutex.lock()
        self._step_advance_requested = True
        self._wait_condition.wakeAll()
        self._mutex.unlock()

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

    def _format_elapsed(self, started_at: float) -> str:
        elapsed_sec = max(0, int(time.time() - started_at))
        mins = elapsed_sec // 60
        secs = elapsed_sec % 60
        return f"{mins} min(s), {secs} sec(s)"

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

    def run(self):
        """Main orchestrator for loading."""
        try:
            self._esc_guard_until = time.time() + 0.2
            try:
                self._esc_was_down = bool(ctypes.windll.user32.GetAsyncKeyState(0x1B) & 0x8000)
            except Exception:
                self._esc_was_down = False
            self._start_esc_listener()
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

    def _send_cell_with_retry(
        self, row_idx: int, col_idx: int, parsed: ParsedCell, rows_processed: int, total_rows: int
    ) -> bool:
        attempts = max(1, int(self._cell_send_retries) + 1)
        for attempt in range(1, attempts + 1):
            if self._is_stop_requested():
                return False

            success = self.sender.send_cell(parsed)
            if success:
                return True

            if self._is_stop_requested():
                return False

            if attempt < attempts:
                reason = (self.sender.last_error or "send failed").strip()
                self.progress_updated.emit(
                    rows_processed,
                    total_rows,
                    f"Retrying R{row_idx + 1} C{col_idx + 1} ({attempt}/{attempts - 1}) - {reason}"
                )
                self.sender.activate_target()
                if not self._interruptible_delay(max(0.05, self.sender.speed_delay)):
                    return False

        return False

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
            if self._form_no_col is None and name in self._FORM_HEADER_ALIASES["no"]:
                self._form_no_col = idx

        if self._form_type_col is not None or self._form_no_col is not None:
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
                    break

    def _apply_form_business_rules(self, row_idx: int, col_idx: int, parsed: ParsedCell) -> ParsedCell:
        """
        Business rules for Per Row table mode:
        - Type column:
          - Receipt/r => type 'r' (row loop tabs to next field before next send).
          - Payment/p => Tab to next field (leave app default type).
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
                    cell_type=CellType.KEYSTROKE,
                    raw_value=parsed.raw_value,
                    key_actions=[
                        {"type": "type", "text": "r"},
                    ],
                )
            if lowered in {"payment", "p"}:
                return ParsedCell(
                    cell_type=CellType.KEYSTROKE,
                    raw_value=parsed.raw_value,
                    key_actions=[{"type": "key", "key": "tab"}],
                )
            return parsed

        # Fallback: if type-column detection misses, still honor IFMIS type tokens
        # in early row columns where type is normally placed.
        if col_idx <= 2 and lowered in {"receipt", "r"}:
            return ParsedCell(
                cell_type=CellType.KEYSTROKE,
                raw_value=parsed.raw_value,
                key_actions=[
                    {"type": "type", "text": "r"},
                ],
            )
        if col_idx <= 2 and lowered in {"payment", "p"}:
            return ParsedCell(
                cell_type=CellType.KEYSTROKE,
                raw_value=parsed.raw_value,
                key_actions=[{"type": "key", "key": "tab"}],
            )

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

    def _perform_end_of_row_action(self, rows_processed: int) -> bool:
        eor_delay = max(0.0, float(self.sender.speed_delay))
        save_settle = max(0.1, float(self.sender.speed_delay))
        if self.end_of_row_action == "new_record":
            pyautogui.press('down')
            if not self._interruptible_delay(eor_delay):
                return False
        elif self.end_of_row_action == "new_record_save_50":
            completed_rows = rows_processed + 1
            if completed_rows % 50 == 0:
                pyautogui.hotkey('ctrl', 's')
                if not self._interruptible_delay(save_settle):
                    return False
            pyautogui.press('down')
            if not self._interruptible_delay(eor_delay):
                return False
        elif self.end_of_row_action == "save_proceed":
            pyautogui.hotkey('ctrl', 's')
            if not self._interruptible_delay(save_settle):
                return False
            pyautogui.press('down')
            if not self._interruptible_delay(eor_delay):
                return False
        elif self.end_of_row_action == "enter":
            pyautogui.press('enter')
            if not self._interruptible_delay(eor_delay):
                return False
        elif self.end_of_row_action == "tab":
            pyautogui.press('tab')
            if not self._interruptible_delay(self.sender.speed_delay):
                return False
        elif self.end_of_row_action == "none":
            pass

        if not self.sender._wait_if_hourglass():
            if not self._is_stop_requested():
                self._request_stop((self.sender.last_error or "target remained busy").strip())
            return False
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

        for row_idx in range(self.start_row, self.end_row + 1):
            if self._is_stop_requested():
                break

            self._check_pause()

            if row_idx >= len(self.grid_data):
                break

            row_data = self.grid_data[row_idx]
            row_had_activity = False
            self.row_started.emit(row_idx)
            elapsed = self._format_elapsed(started_at)
            self.progress_updated.emit(
                rows_processed, total_rows,
                f"Loading row {row_idx + 1}... ({rows_processed}/{total_rows}) [{elapsed}]"
            )

            if self._check_blocking_popup(rows_processed, total_rows):
                if self._is_stop_requested():
                    break
                self._check_pause()

            if self.form_mode:
                if not self.sender.activate_target():
                    self.loading_complete.emit(False, "Lost focus on target window - stopped.")
                    return

                if self.load_mode == "per_row_paste":
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
                        success = self._send_cell_with_retry(
                            row_idx, paste_col_idx, paste_cell, rows_processed, total_rows
                        )
                        self.cell_processed.emit(row_idx, paste_col_idx, success)
                        if not success:
                            if self._is_stop_requested():
                                break
                            errors += 1
                            detail = (self.sender.last_error or "send failed").strip()
                            self.progress_updated.emit(
                                rows_processed,
                                total_rows,
                                f"Row paste error at R{row_idx + 1}: {detail}"
                            )
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
                        if pending_tab_after_receipt and parsed.cell_type != CellType.EMPTY:
                            pyautogui.press('tab')
                            if not self._interruptible_delay(self.sender.speed_delay, wait_hourglass=True):
                                break
                            pending_tab_after_receipt = False

                        success = self._send_cell_with_retry(
                            row_idx, col_idx, parsed, rows_processed, total_rows
                        )
                        self.cell_processed.emit(row_idx, col_idx, success)
                        if not success:
                            if self._is_stop_requested():
                                break
                            errors += 1
                            detail = (self.sender.last_error or "send failed").strip()
                            self.progress_updated.emit(
                                rows_processed,
                                total_rows,
                                f"Cell error at R{row_idx + 1} C{col_idx + 1}: {detail}"
                            )

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

                        # Auto-Tab only between plain data fields.
                        if i in data_positions and data_positions and i != data_positions[-1]:
                            pyautogui.press('tab')
                            if not self._interruptible_delay(self.sender.speed_delay, wait_hourglass=True):
                                break

                if not self._stop_requested and row_had_activity:
                    if not self._perform_end_of_row_action(rows_processed):
                        break

            else:
                # CELL MODE (original behavior)
                if not self.sender.activate_target():
                    self.loading_complete.emit(False, "Lost focus on target window â€” stopped.")
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

                    success = self._send_cell_with_retry(
                        row_idx, col_idx, parsed, rows_processed, total_rows
                    )
                    self.cell_processed.emit(row_idx, col_idx, success)

                    if not success:
                        if self._is_stop_requested():
                            break
                        errors += 1
                        detail = (self.sender.last_error or "send failed").strip()
                        self.progress_updated.emit(
                            rows_processed,
                            total_rows,
                            f"Cell error at R{row_idx + 1} C{col_idx + 1}: {detail}"
                        )

            if self._is_stop_requested():
                if row_had_activity:
                    partial_row_loaded = True
                break

            rows_processed += 1
            elapsed = self._format_elapsed(started_at)
            self.progress_updated.emit(
                rows_processed, total_rows,
                f"Loaded {rows_processed}/{total_rows} row(s) [{elapsed}]"
            )

        # Done UI mode
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

    def _check_pause(self):
        """Check if paused and wait."""
        self._mutex.lock()
        while self._pause_requested and not self._is_stop_requested():
            self._wait_condition.wait(self._mutex, 30)
        self._mutex.unlock()

    def _check_blocking_popup(self, rows_processed: int, total_rows: int) -> bool:
        """
        Detect IFMIS/Oracle popup dialogs.
        Returns True if popup handling was triggered.
        """
        if not self._popup_auto_pause_enabled:
            return False

        now = time.time()
        if now - self._last_popup_check_at < 0.15:
            return False
        self._last_popup_check_at = now

        popup_title = WindowManager.detect_blocking_popup(
            self.sender.target_hwnd,
            self.sender.target_title,
        )
        if not popup_title:
            self._last_popup_title = ""
            return False

        # Always stop on blocking popup. Continuing key playback can interact with
        # the popup itself and cause unintended toggles/actions in Oracle forms.
        self._stop_requested = True
        if not self._stop_reason:
            self._stop_reason = f"Blocking popup detected: {popup_title}"
        if popup_title != self._last_popup_title:
            self._last_popup_title = popup_title
            self.progress_updated.emit(
                rows_processed,
                total_rows,
                f"Stopped: detected popup '{popup_title}'. Resolve popup and restart load."
            )
        return True

    def _wait_for_step(self):
        """Wait for user step/resume while still allowing stop to interrupt."""
        self._mutex.lock()
        self._step_advance_requested = False
        while not self._step_advance_requested and not self._is_stop_requested():
            self._wait_condition.wait(self._mutex, 30)
        self._mutex.unlock()

