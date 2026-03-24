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

Keystroke template (per row) — 74-cell DL grid (C1–C74):
  Requires "Use Alternate Method for processing Macros" in DL Load Settings.
  \\+{PGDN} = Shift+PageDown = Next Block (General→Lines, Lines→Distributions).
  \\d500    = 500 ms delay (DL built-in) — inserted after each \\+{PGDN} for stability.
  \\{ENTER} = close modal / confirm.
  \\{BACKSPACE} = clear pre-filled field.
  Macro flow:
    General fields → \\{ENTER} (close modal) → \\+{PGDN} (→Lines block) → \\d500
    → Tab to Amount → enter amount → \\+{PGDN} (→Distributions block) → \\d500
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

_T       = "{Tab}"
_BS      = "\\{BACKSPACE}"
_ENTER   = "\\{ENTER}"
_ALT2ESC = "\\%2\\{ESC}"   # Alt+2 + Esc  → Lines block
_ALTD    = "\\%d"           # Alt+D        → Distributions block
_CTRLS   = "\\^s"           # Ctrl+S       → save
_CTRLF4  = "\\^{F4}"        # Ctrl+F4      → clear record
_ALTKEY  = "\\%"            # Alt alone    → activate menu
_DOWN    = "\\{DOWN}"       # Down arrow
_SHTAB   = "\\+{TAB}"       # Shift+Tab

# ── Imprest: 81-cell DataLoad grid row ────────────────────────────────────────
# Navigation: \%2\{ESC} (Alt+2+Esc) → Lines,  \%d (Alt+D) → Distributions.
# Post-save: \^{F4} + Alt + 4×Down + Enter + 3×Shift+Tab → ready for next row.
# REQUIRES: "Use Alternate Method for processing Macros" ticked in DL settings.
def build_keystroke_row(row: dict) -> list:
    """
    Convert one invoice row-dict into an 81-cell DataLoad grid row.
    Matches the full working macro exactly:
      General fields → Alt+2+Esc (Lines) → Tab×2 → Amount
      → Alt+D (Distributions) → Tab×2 → Amount → GL_Date → Account
      → Ctrl+S → Ctrl+F4 → Alt → Down×4 → Enter → Shift+Tab×3
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
        _T,           # C1
        _T,           # C2
        _BS,          # C3   clear pre-filled field
        _T,           # C4
        "Standard",   # C5   Invoice Type (fixed)
        _T,           # C6
        _BS,          # C7   clear pre-filled field
        _T,           # C8
        _BS,          # C9   clear pre-filled field
        _T,           # C10
        sup,          # C11  Supplier_Num
        _T,           # C12
        "Provisional",# C13  Invoice batch type (fixed)
        _T,           # C14
        idate,        # C15  Invoice_Date
        _T,           # C16
        inum,         # C17  Invoice_Num
        _T,           # C18
        _T,           # C19
        amt,          # C20  Invoice_Amount
        _T,           # C21
        _T,           # C22
        _T,           # C23
        _T,           # C24
        _T,           # C25
        _T,           # C26
        _T,           # C27
        desc,         # C28  Description
        _T,           # C29
        _T,           # C30
        _T,           # C31
        "IMMEDIATE",  # C32  Pay Terms (fixed)
        _T,           # C33
        pmeth,        # C34  Payment_Method
        _T,           # C35
        _T,           # C36
        _T,           # C37
        _T,           # C38
        _T,           # C39
        _T,           # C40
        _T,           # C41
        _T,           # C42
        _T,           # C43
        _T,           # C44
        _T,           # C45
        _T,           # C46
        _T,           # C47
        _T,           # C48
        _T,           # C49
        _T,           # C50
        _T,           # C51
        auth,         # C52  Auth_Ref_No
        _T,           # C53
        admc,         # C54  Administrative_Code
        _T,           # C55
        _ENTER,       # C56  \{ENTER} — close modal
        _ALT2ESC,     # C57  \%2\{ESC} — Alt+2+Esc → Lines block
        _T,           # C58
        _T,           # C59
        amt,          # C60  Line_Amount (= Invoice_Amount)
        _T,           # C61
        _ALTD,        # C62  \%d — Alt+D → Distributions block
        _T,           # C63
        _T,           # C64
        amt,          # C65  Dist_Amount (= Invoice_Amount)
        _T,           # C66
        gldt,         # C67  GL_Date
        _T,           # C68
        dist,         # C69  Distribution_Account
        _T,           # C70
        _CTRLS,       # C71  \^s — Ctrl+S save
        _CTRLF4,      # C72  \^{F4} — clear record
        _ALTKEY,      # C73  \% — Alt → activate menu
        _DOWN,        # C74  \{DOWN}
        _DOWN,        # C75  \{DOWN}
        _DOWN,        # C76  \{DOWN}
        _DOWN,        # C77  \{DOWN} — navigate to menu item
        _ENTER,       # C78  \{ENTER} — select menu item
        _SHTAB,       # C79  \+{TAB}
        _SHTAB,       # C80  \+{TAB}
        _SHTAB,       # C81  \+{TAB} — ready for next invoice
    ]


# ── Template action sequence ──────────────────────────────────────────────────
# Action tuples:
#   ("tab",    n)            press Tab n times via SendInput
#   ("key",    name)         press a named key (must be in _SI_VK_MAP)
#   ("hotkey", mods, key)    modifier+key  e.g. (["shift"], "pagedown")
#   ("delay",  ms)           sleep ms milliseconds (interruptible)
#   ("field",  col_name)     inject field value via SendInput unicode
#   ("text",   value)        type a fixed literal string via SendInput unicode

# Single imprest template — mirrors build_keystroke_row exactly.
# Navigation: Alt+2+Esc → Lines,  Alt+D → Distributions.
# Post-save: Ctrl+F4 + Alt + Down×4 + Enter + Shift+Tab×3 → ready for next.
# 500 ms delay after Alt+2+Esc for Lines block to open.
TEMPLATE_ACTIONS = (
    # ── General block ──────────────────────────────────────────────────────
    ("tab",    2),                           # C1–C2
    ("key",    "backspace"),                 # C3   clear pre-filled
    ("tab",    1),                           # C4
    ("text",   "Standard"),                  # C5   Invoice Type
    ("tab",    1),                           # C6
    ("key",    "backspace"),                 # C7   clear pre-filled
    ("tab",    1),                           # C8
    ("key",    "backspace"),                 # C9   clear pre-filled
    ("tab",    1),                           # C10
    ("field",  "Supplier_Num"),              # C11
    ("tab",    1),                           # C12
    ("text",   "Provisional"),               # C13  Invoice batch type
    ("tab",    1),                           # C14
    ("field",  "Invoice_Date"),              # C15
    ("tab",    1),                           # C16
    ("field",  "Invoice_Num"),               # C17
    ("tab",    2),                           # C18–C19
    ("field",  "Invoice_Amount"),            # C20
    ("tab",    7),                           # C21–C27
    ("field",  "Description"),               # C28
    ("tab",    3),                           # C29–C31
    ("text",   "IMMEDIATE"),                 # C32  Pay Terms (fixed)
    ("tab",    1),                           # C33
    ("text",   "CHECK"),                     # C34  Payment Method (fixed)
    ("tab",    17),                          # C35–C51
    ("field",  "Auth_Ref_No"),               # C52
    ("tab",    1),                           # C53
    ("field",  "Administrative_Code"),       # C54
    ("tab",    1),                           # C55
    ("key",    "enter"),                     # C56  close modal
    # ── Lines block (Alt+2+Esc) ────────────────────────────────────────────
    ("hotkey", ["alt"], "2"),                # C57a Alt+2 → Lines
    ("key",    "escape"),                    # C57b Esc
    ("delay",  500),                         # wait for Lines block to open
    ("tab",    2),                           # C58–C59
    ("field",  "Invoice_Amount"),            # C60  Line Amount
    ("tab",    1),                           # C61
    # ── Distributions block (Alt+D) ────────────────────────────────────────
    ("hotkey", ["alt"], "d"),                # C62  Alt+D → Distributions
    ("tab",    2),                           # C63–C64
    ("field",  "Invoice_Amount"),            # C65  Dist Amount
    ("tab",    1),                           # C66
    ("field",  "GL_Date"),                   # C67
    ("tab",    1),                           # C68
    ("field",  "Distribution_Account"),      # C69
    ("tab",    1),                           # C70
    # ── Save and advance to next invoice ──────────────────────────────────
    ("hotkey", ["ctrl"], "s"),               # C71  Ctrl+S save
    ("hotkey", ["ctrl"], "f4"),              # C72  Ctrl+F4 clear record
    ("key",    "alt"),                       # C73  Alt → activate menu
    ("key",    "down"),                      # C74  ↓
    ("key",    "down"),                      # C75  ↓
    ("key",    "down"),                      # C76  ↓
    ("key",    "down"),                      # C77  ↓ (4th menu item)
    ("key",    "enter"),                     # C78  select menu item
    ("hotkey", ["shift"], "tab"),            # C79  ⇧Tab
    ("hotkey", ["shift"], "tab"),            # C80  ⇧Tab
    ("hotkey", ["shift"], "tab"),            # C81  ⇧Tab → ready for next invoice
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


# ── IFMIS export importer ─────────────────────────────────────────────────────

# Columns that cannot be auto-mapped from the IFMIS export — user must fill them.
IFMIS_BLANK_COLS = {"Auth_Ref_No", "Administrative_Code", "Distribution_Account"}

# IFMIS header name fragment → our field name (case-insensitive contains match)
_IFMIS_COL_MAP = {
    "supplier num":    "Supplier_Num",
    "invoice date":    "Invoice_Date",
    "invoice num":     "Invoice_Num",
    "invoice amount":  "Invoice_Amount",
    "description":     "Description",
    "payment method":  "Payment_Method",
    "gl date":         "GL_Date",
    "terms date":      "Terms_Date",
}


def _fmt_ifmis_date(v) -> str:
    """Convert an IFMIS date value (datetime or 'YYYY-MM-DD …') to DD-MMM-YYYY."""
    if v is None:
        return ""
    from datetime import datetime as _dt
    if isinstance(v, _dt):
        return v.strftime("%d-%b-%Y").upper()
    s = str(v).strip()
    if not s or s.lower() == "none":
        return ""
    # "2022-06-08 00:00:00" or "2022-06-08"
    try:
        return _dt.strptime(s[:10], "%Y-%m-%d").strftime("%d-%b-%Y").upper()
    except ValueError:
        return s


def import_ifmis_export(filepath: str) -> tuple:
    """
    Read an IFMIS-exported AP invoice Excel and map to our 11-column format.
    Only rows where the 'Type' column equals 'Prepayment' are imported.

    Returns (rows: list[dict], skipped: int, error: str).
      rows    — list of dicts keyed by COLUMNS; blank fields = "".
      skipped — number of non-Prepayment rows ignored.
      error   — non-empty string if the file could not be read.
    """
    try:
        import openpyxl
    except ImportError:
        return [], 0, "openpyxl is not installed."

    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
    except Exception as exc:
        return [], 0, f"Cannot open file: {exc}"

    ws = wb.active

    # ── Build header index (0-based) ──────────────────────────────────────────
    headers = [str(ws.cell(1, c).value or "").strip().lower()
               for c in range(1, ws.max_column + 1)]

    def _find(fragment):
        for i, h in enumerate(headers):
            if fragment in h:
                return i
        return -1

    col_type = _find("type")
    field_cols = {field: _find(frag) for frag, field in _IFMIS_COL_MAP.items()}

    def _val(row_num, col_idx):
        if col_idx < 0:
            return ""
        v = ws.cell(row_num, col_idx + 1).value
        return "" if v is None else str(v).strip()

    def _date_val(row_num, col_idx):
        if col_idx < 0:
            return ""
        return _fmt_ifmis_date(ws.cell(row_num, col_idx + 1).value)

    def _amount_val(row_num, col_idx):
        if col_idx < 0:
            return ""
        v = ws.cell(row_num, col_idx + 1).value
        if v is None:
            return ""
        try:
            f = float(v)
            return f"{f:,.2f}" if f != int(f) else f"{int(f):,}"
        except (TypeError, ValueError):
            return str(v).strip()

    rows = []
    skipped = 0

    for r in range(2, ws.max_row + 1):
        # Filter: Prepayment only
        if col_type >= 0:
            t = _val(r, col_type).lower()
            if t and t != "prepayment":
                skipped += 1
                continue

        supplier = _val(r, field_cols.get("Supplier_Num", -1))
        if not supplier:
            continue   # skip blank rows

        pmeth = _val(r, field_cols.get("Payment_Method", -1)).upper()
        if pmeth in ("ELECTRONICS",):
            pmeth = "ELECTRONIC"

        amt_col = field_cols.get("Invoice_Amount", -1)

        row_dict = {
            "Supplier_Num":         supplier,
            "Invoice_Date":         _date_val(r, field_cols.get("Invoice_Date", -1)),
            "Invoice_Num":          _val(r,  field_cols.get("Invoice_Num", -1)),
            "Invoice_Amount":       _amount_val(r, amt_col),
            "Description":          _val(r,  field_cols.get("Description", -1)),
            "Payment_Method":       pmeth,
            "Terms_Date":           _date_val(r, field_cols.get("Terms_Date", -1)),
            "GL_Date":              _date_val(r, field_cols.get("GL_Date", -1)),
            "Auth_Ref_No":          "",   # must be filled by user
            "Administrative_Code":  "",   # must be filled by user
            "Distribution_Account": "",   # must be filled by user
        }
        rows.append(row_dict)

    return rows, skipped, ""


def export_prefilled_template(filepath: str, rows: list) -> str:
    """
    Export the 11-column template pre-filled with `rows` data.
    Blank columns (Auth_Ref_No, Administrative_Code, Distribution_Account)
    are highlighted orange to remind the user to fill them in.
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

        def _side(style="thin", color="BFBFBF"):
            return Side(style=style, color=color)

        def _border():
            s = _side()
            return Border(left=s, right=s, top=s, bottom=s)

        def _fill(color):
            return PatternFill(fill_type="solid", fgColor=color)

        BLUE      = "0070C0"
        WHITE_TXT = "FFFFFF"
        HINT_BG   = "D6E4F0"
        SAMPLE_BG = "EAF4FB"
        ROW_ODD   = "F2F2F2"
        ROW_EVEN  = "FFFFFF"
        NEED_FILL = "FFD966"   # amber — user must fill

        # Indices (0-based) of the blank/must-fill columns
        blank_ci = {ci for ci, col in enumerate(COLUMNS, start=1)
                    if col in IFMIS_BLANK_COLS}

        # ── Row 1: title ───────────────────────────────────────────────────────
        ws.merge_cells(start_row=1, start_column=1,
                       end_row=1,   end_column=ncols)
        title = ws.cell(row=1, column=1,
                        value="IFMIS AP Invoice Loader  —  NT_DL  |  "
                              "Fill AMBER cells then save. One row = one invoice.")
        title.font      = Font(bold=True, size=11, color=WHITE_TXT)
        title.fill      = _fill(BLUE)
        title.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 18

        # ── Row 2: headers ─────────────────────────────────────────────────────
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

        # ── Row 3: hints ───────────────────────────────────────────────────────
        for ci, col in enumerate(COLUMNS, start=1):
            bg = NEED_FILL if ci in blank_ci else HINT_BG
            cell = ws.cell(row=3, column=ci, value=COLUMN_HINTS.get(col, ""))
            cell.fill      = _fill(bg)
            cell.font      = Font(italic=True, color="1F5C8B", size=9)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border    = _border()
        ws.row_dimensions[3].height = 15

        # ── Row 4: sample ──────────────────────────────────────────────────────
        for ci, col in enumerate(COLUMNS, start=1):
            cell = ws.cell(row=4, column=ci, value=COLUMN_SAMPLE.get(col, ""))
            cell.fill      = _fill(SAMPLE_BG)
            cell.font      = Font(italic=True, color="1F5C8B", size=9)
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.border    = _border()
        ws.row_dimensions[4].height = 15

        # ── Rows 5+: pre-filled data ───────────────────────────────────────────
        for ri, row_dict in enumerate(rows, start=5):
            bg = ROW_ODD if ri % 2 == 1 else ROW_EVEN
            for ci, col in enumerate(COLUMNS, start=1):
                val = row_dict.get(col, "")
                cell_bg = NEED_FILL if (ci in blank_ci and not val) else bg
                cell = ws.cell(row=ri, column=ci, value=val if val else None)
                cell.fill      = _fill(cell_bg)
                cell.font      = Font(size=10)
                cell.alignment = Alignment(horizontal="left", vertical="center")
                cell.border    = _border()
            ws.row_dimensions[ri].height = 15

        # ── Blank rows after data (up to row 104) ─────────────────────────────
        last_data = 4 + len(rows)
        for ri in range(last_data + 1, max(last_data + 2, 105)):
            bg = ROW_ODD if ri % 2 == 1 else ROW_EVEN
            for ci in range(1, ncols + 1):
                cell_bg = NEED_FILL if ci in blank_ci else bg
                cell = ws.cell(row=ri, column=ci)
                cell.fill      = _fill(cell_bg)
                cell.font      = Font(size=10)
                cell.alignment = Alignment(horizontal="left", vertical="center")
                cell.border    = _border()
            ws.row_dimensions[ri].height = 15

        # ── Column widths ─────────────────────────────────────────────────────
        col_widths = [14, 14, 14, 14, 30, 16, 14, 14, 12, 18, 68]
        for ci, w in enumerate(col_widths, start=1):
            ws.column_dimensions[get_column_letter(ci)].width = w

        ws.freeze_panes = "A5"
        wb.save(filepath)
        return ""
    except Exception as exc:
        return f"Export failed: {exc}"


# ── DL keystroke export ───────────────────────────────────────────────────────

def _build_dl_keystroke_row(row: dict) -> list:
    """
    Build an 81-cell DataLoad-format (backslash-macro) row for direct use in DataLoad.
    Mirrors build_keystroke_row but uses \\{TAB} instead of {Tab} for DL compatibility.
    """
    _DT  = "\\{TAB}"
    _BS  = "\\{BACKSPACE}"
    _ENT = "\\{ENTER}"
    _A2E = "\\%2\\{ESC}"
    _AD  = "\\%d"
    _CS  = "\\^s"
    _CF4 = "\\^{F4}"
    _ALT = "\\%"
    _DN  = "\\{DOWN}"
    _SHT = "\\+{TAB}"

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
        # C1–C10
        _DT, _DT, _BS, _DT, "Standard", _DT, _BS, _DT, _BS, _DT,
        # C11–C20
        sup, _DT, "Provisional", _DT, idate, _DT, inum, _DT, _DT, amt,
        # C21–C30
        _DT, _DT, _DT, _DT, _DT, _DT, _DT, desc, _DT, _DT,
        # C31–C40
        _DT, "IMMEDIATE", _DT, pmeth, _DT, _DT, _DT, _DT, _DT, _DT,
        # C41–C50
        _DT, _DT, _DT, _DT, _DT, _DT, _DT, _DT, _DT, _DT,
        # C51–C60
        _DT, auth, _DT, admc, _DT, _ENT, _A2E, _DT, _DT, amt,
        # C61–C70
        _DT, _AD, _DT, _DT, amt, _DT, gldt, _DT, dist, _DT,
        # C71–C81
        _CS, _CF4, _ALT, _DN, _DN, _DN, _DN, _ENT, _SHT, _SHT, _SHT,
    ]


def export_keystroke_file(filepath: str, rows: list) -> str:
    """
    Export an 81-column DataLoad keystroke grid to filepath.
    Load directly in DataLoad (Per Cell mode, 'Use Alternate Method' ticked)
    as a fallback if SendInput mode fails.
    Returns an error string, or "" on success.
    """
    try:
        import openpyxl
        from openpyxl.styles import PatternFill, Font, Alignment
        from openpyxl.utils import get_column_letter
    except ImportError:
        return "openpyxl is not installed."

    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "DL_Keystrokes"

        BLUE      = "0070C0"
        WHITE_TXT = "FFFFFF"
        GREY_BG   = "F2F2F2"
        ncols     = 81

        # ── Row 1: title ──────────────────────────────────────────────────────
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
        title = ws.cell(
            row=1, column=1,
            value="NT_DL Imprest Surrender — DataLoad Keystroke Fallback  "
                  "| Load in Per Cell mode  |  'Use Alternate Method' must be ticked in DL settings")
        title.font      = Font(bold=True, size=10, color=WHITE_TXT)
        title.fill      = PatternFill(fill_type="solid", fgColor=BLUE)
        title.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 18

        # ── Row 2: column headers C1–C81 ─────────────────────────────────────
        for ci in range(1, ncols + 1):
            cell = ws.cell(row=2, column=ci, value=f"C{ci}")
            cell.font      = Font(bold=True, color=WHITE_TXT, size=9)
            cell.fill      = PatternFill(fill_type="solid", fgColor=BLUE)
            cell.alignment = Alignment(horizontal="center")
        ws.row_dimensions[2].height = 15

        # ── Rows 3+: one DL keystroke row per invoice ─────────────────────────
        for ri, row_dict in enumerate(rows, start=3):
            bg = GREY_BG if ri % 2 == 1 else "FFFFFF"
            ks_row = _build_dl_keystroke_row(row_dict)
            for ci, val in enumerate(ks_row, start=1):
                cell = ws.cell(row=ri, column=ci, value=val)
                cell.font      = Font(size=9)
                cell.fill      = PatternFill(fill_type="solid", fgColor=bg)
                cell.alignment = Alignment(horizontal="left")
            ws.row_dimensions[ri].height = 14

        # ── Column widths: narrow for macro cols, wider for data cols ─────────
        # Data is at C11 (sup), C15 (idate), C17 (inum), C20 (amt),
        # C28 (desc), C34 (pmeth), C52 (auth), C54 (admc), C60 (amt),
        # C65 (amt), C67 (gldt), C69 (dist)
        data_cols = {11: 12, 15: 12, 17: 12, 20: 12, 28: 28,
                     34: 12, 52: 8, 54: 14, 60: 10, 65: 10, 67: 12, 69: 52}
        for ci in range(1, ncols + 1):
            ws.column_dimensions[get_column_letter(ci)].width = (
                data_cols.get(ci, 10))

        ws.freeze_panes = "A3"
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


_INTER_ACTION_DELAY = 0.2   # 200 ms between actions — matches DataLoad's default cell timing


def _mid_row_popup_check(sender, popup_fn) -> bool:
    """
    Quick popup check for use inside execute_row_for_loader.
    Returns True to continue, False to abort the row.
    """
    if popup_fn is None:
        return True
    from kdl.window.window_manager import WindowManager
    popup = WindowManager.detect_blocking_popup(
        sender.target_hwnd, sender.target_title)
    if popup:
        return popup_fn(popup)
    return True


def execute_row_for_loader(sender, row_dict: dict, is_stop_requested,
                           actions=None, popup_fn=None) -> bool:
    """
    Execute the AP invoice template for one row using an existing DataSender.
    Used by the main LoaderThread when load_mode is 'imprest_surrender'.
      sender            – a configured DataSender instance (use_fast_send=True)
      row_dict          – {col_name: value} for the 11 AP invoice columns
      is_stop_requested – callable() -> bool
      actions           – action tuple sequence (default: TEMPLATE_ACTIONS)
      popup_fn          – optional callable(popup_title: str) -> bool
                          called when a blocking popup (e.g. Supplier Site LOV)
                          is detected mid-row; return True to continue after
                          the user dismisses it, False to abort the row.
    """
    if actions is None:
        actions = TEMPLATE_ACTIONS

    from kdl.engine.data_sender import _SI_VK_MAP
    vk_tab = _SI_VK_MAP.get("tab", 0x09)

    for action in actions:
        if is_stop_requested():
            return False

        kind = action[0]

        if kind == "tab":
            for _ in range(action[1]):
                if not sender._si_send_vk(vk_tab):
                    return False
                time.sleep(_INTER_ACTION_DELAY)

        elif kind == "key":
            vk = _SI_VK_MAP.get(action[1], 0)
            if vk and not sender._si_send_vk(vk):
                return False
            time.sleep(_INTER_ACTION_DELAY)

        elif kind == "hotkey":
            if not sender._si_send_hotkey(action[1], action[2]):
                return False
            time.sleep(_INTER_ACTION_DELAY)

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
            time.sleep(_INTER_ACTION_DELAY)
            if not _mid_row_popup_check(sender, popup_fn):
                return False

        elif kind == "text":
            if action[1] and not sender._si_send_unicode(action[1]):
                return False
            time.sleep(_INTER_ACTION_DELAY)
            if not _mid_row_popup_check(sender, popup_fn):
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
                    time.sleep(_INTER_ACTION_DELAY)

            elif kind == "key":
                vk = vk_map.get(action[1], 0)
                if vk and not sender._si_send_vk(vk):
                    return False
                time.sleep(_INTER_ACTION_DELAY)

            elif kind == "hotkey":
                if not sender._si_send_hotkey(action[1], action[2]):
                    return False
                time.sleep(_INTER_ACTION_DELAY)

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
                time.sleep(_INTER_ACTION_DELAY)

            elif kind == "text":
                if action[1] and not sender._si_send_unicode(action[1]):
                    return False
                time.sleep(_INTER_ACTION_DELAY)

        return True
