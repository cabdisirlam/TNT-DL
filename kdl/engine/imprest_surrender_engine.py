"""
Imprest Surrender AP Loader — Engine
Converts IFMIS AP invoice rows from Excel into IFMIS keystrokes and loads them.

Excel format (Data_Entry sheet):
  Row 1: Title
  Row 2: Column headers
  Row 3: Format hints
  Row 4: Sample row (do not edit)
  Row 5+: Data rows (one row = one AP invoice)

Columns (A–K):
  Supplier_Num, Invoice_Date, Invoice_Num, Invoice_Amount, Description,
  Payment_Method, Terms_Date, GL_Date, Auth_Ref_No, Administrative_Code,
  Distribution_Account

Keystroke template (per row) — 72-cell DL grid (C1–C72):
  Requires "Use Alternate Method for processing Macros" in DL Load Settings.
  \\+{PGDN} = Shift+PageDown = Next Block (General→Lines, Lines→Distributions).
  \\{ENTER} = close modal / confirm.
  \\{BACKSPACE} = clear pre-filled field.
  Macro flow:
    General fields → \\{ENTER} (close modal) → \\+{PGDN} (→Lines block)
    → Tab to Amount → enter amount → \\+{PGDN} (→Distributions block)
    → Tab to Amount → enter amount → GL Date → Dist Account → Save → Down
"""

import time

from PySide6.QtCore import QThread, Signal


# ── Column definitions ────────────────────────────────────────────────────────

COLUMNS = [
    "Supplier_Num",
    "Invoice_Date",
    "Invoice_Num",
    "Invoice_Amount",
    "Description",
    "Payment_Method",
    "Terms_Date",
    "GL_Date",
    "Auth_Ref_No",
    "Administrative_Code",
    "Distribution_Account",
]

COLUMN_HINTS = {
    "Supplier_Num":        "e.g. 117711",
    "Invoice_Date":        "DD-MMM-YYYY",
    "Invoice_Num":         "e.g. SURR0001",
    "Invoice_Amount":      "e.g. 42,000.00",
    "Description":         "e.g. SURRENDER OF IMPREST",
    "Payment_Method":      "CHECK or ELECTRONIC",
    "Terms_Date":          "DD-MMM-YYYY",
    "GL_Date":             "DD-MMM-YYYY",
    "Auth_Ref_No":         "e.g. CFO",
    "Administrative_Code": "e.g. 5322000201",
    "Distribution_Account":"Full SCOA e.g. 0-5322-0000000000-00001001-...",
}

COLUMN_SAMPLE = {
    "Supplier_Num":        "117711",
    "Invoice_Date":        "30-JUN-2020",
    "Invoice_Num":         "SURR0001",
    "Invoice_Amount":      "42,000.00",
    "Description":         "SURRENDER OF IMPREST",
    "Payment_Method":      "CHECK",
    "Terms_Date":          "30-JUN-2020",
    "GL_Date":             "30-JUN-2020",
    "Auth_Ref_No":         "CFO",
    "Administrative_Code": "5322000201",
    "Distribution_Account":"0-5322-0000000000-00001001-0000000000-6580101-53100001-000",
}

# ── Keystroke grid row builder ────────────────────────────────────────────────

_T    = "{Tab}"
_BS   = "\\{BACKSPACE}"
_ENTER = "\\{ENTER}"
_PGDN  = "\\+{PGDN}"

# Column-index → value mapping for a 72-cell DataLoad grid row.
# Matches Keystroke_Updated.xlsx "Updated" sheet exactly (C1–C72, 0-indexed).
# REQUIRES: "Use Alternate Method for processing Macros" ticked in DL settings.
def build_keystroke_row(row: dict) -> list:
    """
    Convert one invoice row-dict (11 fields) into a 72-element list
    ready to be written into a DataLoad-style grid row.
    Fixed cells contain keystroke strings; data cells contain field values.
    """
    sup   = row.get("Supplier_Num", "")
    idate = row.get("Invoice_Date", "")
    inum  = row.get("Invoice_Num", "")
    amt   = row.get("Invoice_Amount", "")
    desc  = row.get("Description", "")
    pmeth = row.get("Payment_Method", "")
    gldt  = row.get("GL_Date", "")
    auth  = row.get("Auth_Ref_No", "")
    admc  = row.get("Administrative_Code", "")
    dist  = row.get("Distribution_Account", "")

    return [
        _T,          # C1
        _T,          # C2
        _BS,         # C3  \{BACKSPACE} — clear pre-filled field
        _T,          # C4
        "Standard",  # C5  Invoice Type (fixed)
        _T,          # C6
        _BS,         # C7  \{BACKSPACE} — clear pre-filled field
        _T,          # C8
        _BS,         # C9  \{BACKSPACE} — clear pre-filled field
        _T,          # C10
        sup,         # C11 Supplier_Num
        _T,          # C12
        "Provisional", # C13 Inv Type (fixed)
        _T,          # C14
        idate,       # C15 Invoice_Date
        _T,          # C16
        inum,        # C17 Invoice_Num
        _T,          # C18
        _T,          # C19
        amt,         # C20 Invoice_Amount
        _T,          # C21
        _T,          # C22
        _T,          # C23
        _T,          # C24
        _T,          # C25
        _T,          # C26
        _T,          # C27
        desc,        # C28 Description
        _T,          # C29
        _T,          # C30
        _T,          # C31
        "IMMEDIATE", # C32 Pay Terms (fixed)
        _T,          # C33
        pmeth,       # C34 Payment_Method
        _T,          # C35
        _T,          # C36
        _T,          # C37
        _T,          # C38
        _T,          # C39
        _T,          # C40
        _T,          # C41
        _T,          # C42
        _T,          # C43
        _T,          # C44
        _T,          # C45
        _T,          # C46
        _T,          # C47
        _T,          # C48
        _T,          # C49
        _T,          # C50
        _T,          # C51
        auth,        # C52 Auth_Ref_No
        _T,          # C53
        admc,        # C54 Administrative_Code
        _T,          # C55
        _ENTER,      # C56 \{ENTER} — close modal / confirm
        _PGDN,       # C57 \+{PGDN} — Next Block: General → Lines
        _T,          # C58
        _T,          # C59
        amt,         # C60 Line_Amount (= Invoice_Amount)
        _T,          # C61
        _PGDN,       # C62 \+{PGDN} — Next Block: Lines → Distributions
        _T,          # C63
        _T,          # C64
        amt,         # C65 Dist_Amount (= Invoice_Amount)
        _T,          # C66
        gldt,        # C67 GL_Date (distribution)
        _T,          # C68
        dist,        # C69 Distribution_Account
        _T,          # C70
        "\\*s",      # C71 \*s — Ctrl+S save
        "*dn",       # C72 *dn — move down to next row
    ]


# ── Template action sequence ──────────────────────────────────────────────────
# Action tuples:
#   ("tab", n)              press Tab n times via SendInput
#   ("key", name)           press a named key (must be in _SI_VK_MAP)
#   ("hotkey", mods, key)   modifier+key  e.g. (["alt"], "o")
#   ("delay", ms)           sleep ms milliseconds (interruptible)
#   ("field", col_name)     inject field value via SendInput unicode

TEMPLATE_ACTIONS = (
    ("tab",    2),
    ("key",    "backspace"),
    ("tab",    2),
    ("field",  "Supplier_Num"),
    ("tab",    1),
    ("hotkey", ["alt"], "o"),
    ("delay",  500),
    ("key",    "home"),
    ("tab",    1),
    ("field",  "Invoice_Date"),
    ("tab",    1),
    ("hotkey", ["alt"], "o"),
    ("delay",  500),
    ("field",  "Invoice_Num"),
    ("tab",    2),
    ("field",  "Invoice_Amount"),
    ("tab",    1),
    ("tab",    6),
    ("field",  "Description"),
    ("tab",    4),
    ("field",  "Payment_Method"),
    ("tab",    2),
    ("field",  "Terms_Date"),
    ("tab",    2),
    ("field",  "GL_Date"),
    ("tab",    17),
    ("field",  "Auth_Ref_No"),
    ("tab",    1),
    ("field",  "Administrative_Code"),
    ("tab",    1),
    ("hotkey", ["alt"], "o"),
    ("delay",  700),
    ("tab",    2),
    ("field",  "Invoice_Amount"),          # line amount
    ("tab",    1),
    ("hotkey", ["alt"], "d"),
    ("delay",  700),
    ("tab",    2),
    ("field",  "Invoice_Amount"),          # distribution amount
    ("tab",    1),
    ("field",  "GL_Date"),                 # distribution date
    ("tab",    1),
    ("field",  "Distribution_Account"),
    ("tab",    1),
    ("hotkey", ["ctrl"], "s"),
    ("delay",  1000),
)


# ── Excel I/O ─────────────────────────────────────────────────────────────────

def read_invoice_rows(filepath: str) -> tuple:
    """
    Read invoice rows from the Data_Entry sheet (or first sheet) of filepath.
    Data rows start at row 4 (rows 1–3 are title / headers / hints).
    Returns (rows: list[dict], error: str).
    """
    try:
        import openpyxl
    except ImportError:
        return [], "openpyxl is not installed."

    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
    except Exception as exc:
        return [], f"Cannot open file: {exc}"

    # Prefer a sheet named Data_Entry (or containing "data"/"entry"), else first sheet
    sheet_name = wb.sheetnames[0]
    for name in wb.sheetnames:
        if any(k in name.lower() for k in ("data", "entry", "invoice")):
            sheet_name = name
            break

    ws = wb[sheet_name]
    rows = []
    for row_values in ws.iter_rows(min_row=5, values_only=True):
        if all(v is None or str(v).strip() == "" for v in row_values):
            continue
        if len(row_values) < len(COLUMNS):
            continue
        row_dict = {
            col: (str(row_values[i]).strip() if row_values[i] is not None else "")
            for i, col in enumerate(COLUMNS)
        }
        # Skip rows where Supplier_Num is empty
        if not row_dict["Supplier_Num"]:
            continue
        rows.append(row_dict)

    return rows, ""


def export_template(filepath: str) -> str:
    """
    Write the Data_Entry Excel template to filepath.
    Returns an error string, or "" on success.
    """
    try:
        import openpyxl
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        return "openpyxl is not installed."

    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data_Entry"

        ncols = len(COLUMNS)

        # ── Shared style helpers ──────────────────────────────────────────────
        def _side(style="thin", color="BFBFBF"):
            return Side(style=style, color=color)

        def _border():
            s = _side()
            return Border(left=s, right=s, top=s, bottom=s)

        def _fill(color):
            return PatternFill(fill_type="solid", fgColor=color)

        BLUE      = "0070C0"   # header bg
        WHITE_TXT = "FFFFFF"
        HINT_BG   = "D6E4F0"   # light blue — hint row
        SAMPLE_BG = "EAF4FB"   # very light blue — sample row
        ROW_ODD   = "F2F2F2"   # alternating grey
        ROW_EVEN  = "FFFFFF"   # alternating white

        # ── Row 1: title bar ─────────────────────────────────────────────────
        ws.merge_cells(start_row=1, start_column=1,
                       end_row=1,   end_column=ncols)
        title = ws.cell(row=1, column=1,
                        value="IFMIS AP Invoice Loader  —  NT_DL  |  "
                              "Fill from row 5 only. One row = one invoice.")
        title.font      = Font(bold=True, size=11, color=WHITE_TXT)
        title.fill      = _fill(BLUE)
        title.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 18

        # ── Row 2: column headers ─────────────────────────────────────────────
        hdr_border_s = _side("medium", "FFFFFF")
        hdr_border   = Border(left=hdr_border_s, right=hdr_border_s,
                              top=hdr_border_s, bottom=hdr_border_s)
        for ci, col in enumerate(COLUMNS, start=1):
            cell = ws.cell(row=2, column=ci, value=col.replace("_", " "))
            cell.fill      = _fill(BLUE)
            cell.font      = Font(bold=True, color=WHITE_TXT, size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center",
                                       wrap_text=True)
            cell.border    = hdr_border
        ws.row_dimensions[2].height = 28

        # ── Row 3: format hints ───────────────────────────────────────────────
        for ci, col in enumerate(COLUMNS, start=1):
            cell = ws.cell(row=3, column=ci, value=COLUMN_HINTS.get(col, ""))
            cell.fill      = _fill(HINT_BG)
            cell.font      = Font(italic=True, color="1F5C8B", size=9)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border    = _border()
        ws.row_dimensions[3].height = 15

        # ── Row 4: sample data row ────────────────────────────────────────────
        for ci, col in enumerate(COLUMNS, start=1):
            cell = ws.cell(row=4, column=ci, value=COLUMN_SAMPLE.get(col, ""))
            cell.fill      = _fill(SAMPLE_BG)
            cell.font      = Font(italic=True, color="1F5C8B", size=9)
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.border    = _border()
        ws.row_dimensions[4].height = 15

        # ── Rows 5–104: data entry rows (alternating, no colour fill) ─────────
        for ri in range(5, 105):
            bg = ROW_ODD if ri % 2 == 1 else ROW_EVEN
            for ci in range(1, ncols + 1):
                cell = ws.cell(row=ri, column=ci)
                cell.fill      = _fill(bg)
                cell.font      = Font(size=10)
                cell.alignment = Alignment(horizontal="left", vertical="center")
                cell.border    = _border()
            ws.row_dimensions[ri].height = 15

        # ── Column widths ─────────────────────────────────────────────────────
        col_widths = [14, 14, 14, 14, 30, 16, 14, 14, 12, 18, 68]
        for ci, w in enumerate(col_widths, start=1):
            ws.column_dimensions[get_column_letter(ci)].width = w

        # ── Freeze panes below header rows ────────────────────────────────────
        ws.freeze_panes = "A5"

        wb.save(filepath)
        return ""
    except Exception as exc:
        return f"Export failed: {exc}"


# ── Helper ────────────────────────────────────────────────────────────────────

def build_row_summary(row: dict) -> str:
    """One-line human-readable summary of an invoice row."""
    desc = (row.get("Description") or "")[:32]
    return (f"{row.get('Invoice_Num', '?')} | "
            f"Supplier {row.get('Supplier_Num', '?')} | "
            f"Amt {row.get('Invoice_Amount', '?')} | {desc}")


def execute_row_for_loader(sender, row_dict: dict, is_stop_requested) -> bool:
    """
    Execute the AP invoice template for one row using an existing DataSender.
    Used by the main LoaderThread when load_mode == 'imprest_surrender'.
      sender            – a configured DataSender instance (use_fast_send=True)
      row_dict          – {col_name: value} for the 11 AP invoice columns
      is_stop_requested – callable() -> bool
    """
    from kdl.engine.data_sender import _SI_VK_MAP
    vk_tab = _SI_VK_MAP.get("tab", 0x09)

    for action in TEMPLATE_ACTIONS:
        if is_stop_requested():
            return False

        kind = action[0]

        if kind == "tab":
            for _ in range(action[1]):
                if not sender._si_send_vk(vk_tab):
                    return False

        elif kind == "key":
            vk = _SI_VK_MAP.get(action[1], 0)
            if vk and not sender._si_send_vk(vk):
                return False

        elif kind == "hotkey":
            if not sender._si_send_hotkey(action[1], action[2]):
                return False

        elif kind == "delay":
            deadline = time.monotonic() + action[1] / 1000.0
            while time.monotonic() < deadline:
                if is_stop_requested():
                    return False
                time.sleep(0.05)

        elif kind == "field":
            value = (row_dict.get(action[1]) or "").strip()
            if value and not sender._si_send_unicode(value):
                return False

    return True


# ── Background loader thread ──────────────────────────────────────────────────

class ImprestSurrenderThread(QThread):
    """Sends invoice keystrokes to IFMIS for each row using SendInput."""

    progress       = Signal(int, int)   # (current_row, total_rows)
    status_update  = Signal(str)        # status text
    finished_signal = Signal(int, str)  # (rows_loaded, message)
    stopped_signal  = Signal()

    def __init__(self, rows: list, target_hwnd, target_title: str, parent=None):
        super().__init__(parent)
        self._rows         = rows
        self._target_hwnd  = target_hwnd
        self._target_title = target_title
        self._stop         = False

    def request_stop(self):
        self._stop = True

    # ── run ──────────────────────────────────────────────────────────────────

    def run(self):
        from kdl.engine.data_sender import DataSender, _SI_VK_MAP

        sender = DataSender()
        sender.use_fast_send = True
        sender.set_target(self._target_hwnd, self._target_title)
        sender.set_stop_checker(lambda: self._stop)

        if not sender.activate_target():
            self.finished_signal.emit(
                0, f"Cannot activate target window: {self._target_title}")
            return

        total  = len(self._rows)
        loaded = 0

        for i, row in enumerate(self._rows):
            if self._stop:
                self.stopped_signal.emit()
                return

            inv = row.get("Invoice_Num") or f"row {i + 1}"
            self.status_update.emit(f"Loading {i + 1}/{total}: {inv} …")

            ok = execute_row_for_loader(sender, row, lambda: self._stop)
            if not ok:
                if self._stop:
                    self.stopped_signal.emit()
                    return
                self.finished_signal.emit(
                    loaded,
                    f"Error at row {i + 1} ({inv}): "
                    f"{sender.last_error or 'unknown error'}")
                return

            loaded += 1
            self.progress.emit(i + 1, total)

        self.finished_signal.emit(
            loaded, f"Done — {loaded} invoice(s) loaded into IFMIS.")

    # ── row executor ─────────────────────────────────────────────────────────

    def _execute_row(self, sender, row: dict, vk_map: dict) -> bool:
        vk_tab = vk_map.get("tab", 0x09)

        for action in TEMPLATE_ACTIONS:
            if self._stop:
                return False

            kind = action[0]

            if kind == "tab":
                for _ in range(action[1]):
                    if not sender._si_send_vk(vk_tab):
                        return False

            elif kind == "key":
                vk = vk_map.get(action[1], 0)
                if vk and not sender._si_send_vk(vk):
                    return False

            elif kind == "hotkey":
                if not sender._si_send_hotkey(action[1], action[2]):
                    return False

            elif kind == "delay":
                deadline = time.monotonic() + action[1] / 1000.0
                while time.monotonic() < deadline:
                    if self._stop:
                        return False
                    time.sleep(0.05)

            elif kind == "field":
                value = (row.get(action[1]) or "").strip()
                if value and not sender._si_send_unicode(value):
                    return False

        return True
