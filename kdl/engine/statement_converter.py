"""
Bank Statement Converter — Python port of Build_Statement_Output VBA macro.
Reads a bank statement sheet from an openpyxl workbook and writes
Output + Audit_Skipped sheets back into the same workbook.
"""

from datetime import date, datetime, timedelta
from typing import Optional, List, Dict
from dataclasses import dataclass, field

try:
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.worksheet.worksheet import Worksheet
    from openpyxl.styles import Font
except ImportError:
    openpyxl = None

# ── Constants ──────────────────────────────────────────────
DN_PREFIX_CELL = r"\*s"
OUTPUT_FIRST_ROW = 2
AUDIT_DETAIL_HEADER_ROW = 9
AUDIT_DETAIL_FIRST_ROW = 10


@dataclass
class ConversionResult:
    success: bool
    message: str
    output_rows: int = 0
    skipped_rows: int = 0
    total_debits: float = 0.0
    total_credits: float = 0.0
    opening_balance: Optional[float] = None
    closing_balance_stmt: Optional[float] = None
    closing_balance_calc: Optional[float] = None
    variance: Optional[float] = None
    output_data: List[List] = field(default_factory=list)   # rows for TNT DL grid
    audit_rows: List[List] = field(default_factory=list)    # skipped-row detail records
    saved_path: str = ""


# ── Number helpers ─────────────────────────────────────────

def _safe_double(v) -> float:
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(',', '')
    if len(s) >= 2 and s[0] == '(' and s[-1] == ')':
        s = '-' + s[1:-1]
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def _parse_number(v):
    """Return float or None."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(',', '')
    if s == '':
        return None
    if len(s) >= 2 and s[0] == '(' and s[-1] == ')':
        s = '-' + s[1:-1]
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


# ── Date helpers ───────────────────────────────────────────

_EXCEL_EPOCH = date(1899, 12, 30)   # Excel serial-date epoch (accounts for 1900 bug)


def _parse_date_cell(cell_value, cell_text: str = '') -> Optional[date]:
    """Parse date from cell value or text. Supports date objects, M/D/YYYY, D-MMM-YYYY,
    and raw Excel serial numbers (int/float) that openpyxl sometimes returns."""
    if isinstance(cell_value, datetime):
        return cell_value.date()
    if isinstance(cell_value, date):
        return cell_value

    # openpyxl occasionally returns the raw Excel serial integer for date cells
    if isinstance(cell_value, (int, float)) and not isinstance(cell_value, bool):
        serial = int(cell_value)
        if 1 <= serial <= 2958465:          # valid Excel date range (1900-01-01 to 9999-12-31)
            try:
                return _EXCEL_EPOCH + timedelta(days=serial)
            except (OverflowError, ValueError):
                pass

    # Try text-based parsing
    s = str(cell_value).strip() if cell_value is not None else ''
    if not s or s == '####':
        s = cell_text.strip()
    if not s:
        return None

    # Strip time portion if present
    if ' ' in s:
        s = s.split()[0]

    if '/' in s:
        result = _parse_mdy(s, '/')
        if result is not None:
            return result
    if '-' in s:
        result = _parse_dmon_y(s, '-')
        if result is not None:
            return result

    # Last resort: try Python's own ISO / flexible parser
    try:
        from datetime import datetime as _dt
        return _dt.fromisoformat(s).date()
    except (ValueError, TypeError):
        pass

    return None


def _parse_mdy(s: str, sep: str) -> Optional[date]:
    parts = s.strip().split(sep)
    if len(parts) != 3:
        return None
    try:
        mm, dd, yy = int(parts[0]), int(parts[1]), int(parts[2])
        if yy < 100:
            yy += 2000
        if not (1 <= mm <= 12 and 1 <= dd <= 31):
            return None
        return date(yy, mm, dd)
    except (ValueError, TypeError):
        return None


def _parse_dmon_y(s: str, sep: str) -> Optional[date]:
    parts = s.strip().split(sep)
    if len(parts) != 3:
        return None
    try:
        dd = int(parts[0])
        mm = _month_text_to_num(parts[1])
        yy = int(parts[2])
        if yy < 100:
            yy += 2000
        if not (1 <= mm <= 12 and 1 <= dd <= 31):
            return None
        return date(yy, mm, dd)
    except (ValueError, TypeError):
        return None


_MONTH_MAP = {
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
}


def _month_text_to_num(mon: str) -> int:
    return _MONTH_MAP.get(mon.strip().lower()[:3], 0)


# ── Doc number extraction ──────────────────────────────────

def _extract_doc_no_10(details: str) -> str:
    """Extract rightmost 10-digit run, strip leading zeros."""
    if not details:
        return ''
    best10 = ''
    run = ''
    for ch in str(details):
        if ch.isdigit():
            run += ch
        else:
            if len(run) >= 10:
                best10 = run[-10:]
            run = ''
    if len(run) >= 10:
        best10 = run[-10:]
    if not best10:
        return ''
    return best10.lstrip('0') or ''


# ── Header helpers ─────────────────────────────────────────

def _normalize_header(s: str) -> str:
    s = str(s).replace('\xa0', ' ').replace('\r', ' ').replace('\n', ' ').strip()
    while '  ' in s:
        s = s.replace('  ', ' ')
    return s.upper()


def _find_header_col(ws: Worksheet, hdr_row: int, header_text: str) -> int:
    """Return 1-based column index or 0 if not found."""
    want = _normalize_header(header_text)
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=hdr_row, column=col).value
        if val is not None and _normalize_header(str(val)) == want:
            return col
    return 0


def _find_txn_header_row(ws: Worksheet) -> int:
    """Auto-detect the header row (must have Date, Transaction Details, Debit, Credit)."""
    for r in range(1, min(30, ws.max_row + 1)):
        if (
            _find_header_col(ws, r, 'Date') > 0
            and _find_header_col(ws, r, 'Transaction Details') > 0
            and _find_header_col(ws, r, 'Debit') > 0
            and _find_header_col(ws, r, 'Credit') > 0
        ):
            return r
    return 0


# ── Output row builders (return plain lists, no worksheet I/O) ─────────────

def _fmt_date(d: date) -> str:
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
              'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    return f"{d.day:02d}-{months[d.month - 1]}-{d.year}"


def _make_payment_row(doc_no, dt: date, amt: float) -> list:
    return [
        'tab',
        'tab',
        'TRFD',
        'tab',
        str(doc_no) if doc_no else '',
        'tab',
        _fmt_date(dt),
        'tab',
        _fmt_date(dt),
        'tab',
        amt,
        'tab',
        DN_PREFIX_CELL,
        '*DN',
    ]


def _make_receipt_row(doc_no: str, dt: date, amt: float) -> list:
    return [
        'tab',
        '*DN',
        'r',
        'tab',
        'TRFC',
        'tab',
        doc_no,
        'tab',
        _fmt_date(dt),
        'tab',
        _fmt_date(dt),
        'tab',
        amt,
        'tab',
        DN_PREFIX_CELL,
        '*DN',
    ]


def _write_rows_to_sheet(ws_out: Worksheet, rows: list[list]):
    """Write keystroke-format rows to ws_out starting at row 2, leaving row 1 blank."""
    for i, row_vals in enumerate(rows):
        for c, val in enumerate(row_vals, 1):
            ws_out.cell(row=i + OUTPUT_FIRST_ROW, column=c, value=val)


# ── Audit writers ──────────────────────────────────────────

def _write_audit_header(ws_audit: Worksheet):
    headers = [
        'Row#', 'Date (Raw)', 'Reference', 'Transaction Type',
        'Transaction Details', 'Debit (Raw)', 'Credit (Raw)',
        'Debit (Parsed)', 'Credit (Parsed)', 'Skip Reason'
    ]
    for c, h in enumerate(headers, 1):
        cell = ws_audit.cell(row=AUDIT_DETAIL_HEADER_ROW, column=c, value=h)
        cell.font = Font(bold=True)


def _add_audit(ws_audit: Worksheet, audit_row: int, row_num: int,
               raw_date: str, raw_ref, txn_type: str, details: str,
               raw_debit, raw_credit, reason: str):
    vals = [row_num, raw_date, raw_ref, txn_type, details,
            raw_debit, raw_credit,
            _safe_double(raw_debit), _safe_double(raw_credit), reason]
    for c, v in enumerate(vals, 1):
        ws_audit.cell(row=audit_row, column=c, value=v)


def _write_audit_summary(ws_audit: Worksheet, opening, additions, outs,
                          closing_calc, closing_stmt, variance):
    labels = [
        'BANK RECONCILIATION SUMMARY',
        'Opening Balance',
        'Additions (Receipts/Credits Captured)',
        'Less: Outs (Payments/Debits Captured)',
        'Closing Balance (Calculated)',
        'Closing Balance (Statement)',
        'Closing Match (Statement - Calculated)',
    ]
    for i, lbl in enumerate(labels):
        c = ws_audit.cell(row=i + 1, column=1, value=lbl)
        if i == 0:
            c.font = Font(bold=True)

    if opening is None:
        ws_audit.cell(row=2, column=2, value='N/A (No Closing Balance column detected)')
        ws_audit.cell(row=3, column=2, value=additions)
        ws_audit.cell(row=4, column=2, value=outs)
    else:
        ws_audit.cell(row=2, column=2, value=opening)
        ws_audit.cell(row=3, column=2, value=additions)
        ws_audit.cell(row=4, column=2, value=outs)
        ws_audit.cell(row=5, column=2, value=closing_calc)
        ws_audit.cell(row=6, column=2, value=closing_stmt if closing_stmt is not None else 'N/A')
        ws_audit.cell(row=7, column=2, value=variance if variance is not None else 'N/A')


# ── Sheet helper ───────────────────────────────────────────

def _get_or_create_sheet(wb: Workbook, name: str) -> Worksheet:
    if name in wb.sheetnames:
        idx = wb.sheetnames.index(name)
        del wb[name]
        return wb.create_sheet(title=name, index=idx)
    return wb.create_sheet(title=name)


# ── Main converter ─────────────────────────────────────────

def convert_statement(wb: Workbook, sheet_name: str, skip_contra: bool = True) -> ConversionResult:
    """
    Run the full conversion on wb[sheet_name].
    Writes Output and Audit_Skipped sheets into wb.
    Returns ConversionResult with output_data for loading into the TNT DL grid.
    """
    if openpyxl is None:
        return ConversionResult(success=False, message='openpyxl not installed.')

    if sheet_name not in wb.sheetnames:
        return ConversionResult(success=False, message=f'Sheet "{sheet_name}" not found.')

    ws = wb[sheet_name]

    # 1) Find header row
    hdr_row = _find_txn_header_row(ws)
    if hdr_row == 0:
        return ConversionResult(
            success=False,
            message='Required columns not found (Date, Transaction Details, Debit, Credit).'
        )
    data_start = hdr_row + 1

    # 2) Map columns
    col_date = _find_header_col(ws, hdr_row, 'Date')
    col_ref = _find_header_col(ws, hdr_row, 'Transaction Reference')
    col_details = _find_header_col(ws, hdr_row, 'Transaction Details')
    col_type = _find_header_col(ws, hdr_row, 'Transaction Type')
    col_debit = _find_header_col(ws, hdr_row, 'Debit')
    col_credit = _find_header_col(ws, hdr_row, 'Credit')
    col_bal = _find_header_col(ws, hdr_row, 'Closing Balance')
    if col_bal == 0:
        col_bal = _find_header_col(ws, hdr_row, 'Balance')

    # 3) Last data row
    last_row = ws.max_row
    while last_row >= data_start and ws.cell(row=last_row, column=col_date).value is None:
        last_row -= 1
    if last_row < data_start:
        return ConversionResult(success=False, message='No data rows found.')

    # 4) Create output sheets
    ws_out = _get_or_create_sheet(wb, 'Output')
    ws_audit = _get_or_create_sheet(wb, 'Audit_Skipped')
    _write_audit_header(ws_audit)

    # 5) Opening / closing balance
    opening_bal = None
    closing_bal_stmt = None
    if col_bal > 0:
        for r in range(data_start, last_row + 1):
            v = _parse_number(ws.cell(row=r, column=col_bal).value)
            if v is not None:
                fb = v
                fd = _safe_double(ws.cell(row=r, column=col_debit).value)
                fc = _safe_double(ws.cell(row=r, column=col_credit).value)
                opening_bal = fb - fc + fd
                break
        for r in range(last_row, data_start - 1, -1):
            v = _parse_number(ws.cell(row=r, column=col_bal).value)
            if v is not None:
                closing_bal_stmt = v
                break

    # 6) Contra matching (only when skip_contra=True)
    skip_row: Dict[int, str] = {}
    if skip_contra:
        deb_map: Dict[str, List[int]] = {}
        cr_map: Dict[str, List[int]] = {}

        for r in range(data_start, last_row + 1):
            details = str(ws.cell(row=r, column=col_details).value or '')
            ref10 = _extract_doc_no_10(details)
            if not ref10:
                continue
            d = _safe_double(ws.cell(row=r, column=col_debit).value)
            c = _safe_double(ws.cell(row=r, column=col_credit).value)
            key = f'{ref10}|{amt:.2f}' if (amt := max(d, c)) > 0 else ''
            if not key:
                continue
            if d > 0 and c == 0:
                deb_map.setdefault(key, []).append(r)
            elif c > 0 and d == 0:
                cr_map.setdefault(key, []).append(r)

        for key in deb_map:
            if key in cr_map:
                pairs = min(len(deb_map[key]), len(cr_map[key]))
                for i in range(pairs):
                    skip_row[deb_map[key][i]] = 'CONTRA_MATCHED_DETAILS10_AMOUNT'
                    skip_row[cr_map[key][i]] = 'CONTRA_MATCHED_DETAILS10_AMOUNT'

    # 7) Build output rows in memory (no worksheet I/O yet)
    payment_rows: List[list] = []
    receipt_rows: List[list] = []
    audit_detail_rows: List[list] = []   # collected in memory for multi-sheet merge
    audit_row = AUDIT_DETAIL_FIRST_ROW
    skip_tally: Dict[str, int] = {}
    tot_deb = 0.0
    tot_cred = 0.0

    for r in range(data_start, last_row + 1):
        raw_date_val = ws.cell(row=r, column=col_date).value
        raw_date_text = str(raw_date_val or '')
        raw_ref = ws.cell(row=r, column=col_ref).value if col_ref else ''
        raw_debit = ws.cell(row=r, column=col_debit).value
        raw_credit = ws.cell(row=r, column=col_credit).value
        details = str(ws.cell(row=r, column=col_details).value or '')
        txn_type = str(ws.cell(row=r, column=col_type).value or '') if col_type else ''
        ref10 = _extract_doc_no_10(details)

        def _audit(reason: str, ref=raw_ref):
            nonlocal audit_row
            vals = [r, raw_date_text, ref, txn_type, details,
                    raw_debit, raw_credit,
                    _safe_double(raw_debit), _safe_double(raw_credit), reason]
            audit_detail_rows.append(vals)
            _add_audit(ws_audit, audit_row, r, raw_date_text, ref,
                       txn_type, details, raw_debit, raw_credit, reason)
            skip_tally[reason] = skip_tally.get(reason, 0) + 1
            audit_row += 1

        if r in skip_row:
            _audit(skip_row[r], ref=ref10)
            continue

        dt = _parse_date_cell(raw_date_val, raw_date_text)
        if dt is None:
            _audit('BLANK_OR_INVALID_DATE')
            continue

        d2 = _safe_double(raw_debit)
        c2 = _safe_double(raw_credit)

        if d2 > 0 and c2 > 0:
            _audit('BOTH_DEBIT_AND_CREDIT_PRESENT')
            continue

        if d2 == 0 and c2 == 0:
            _audit('ZERO_OR_NON_NUMERIC_AMOUNT')
            continue

        if d2 > 0 and c2 == 0:
            payment_rows.append(_make_payment_row(ref10, dt, d2))
            tot_deb += d2
        elif c2 > 0 and d2 == 0:
            doc_no_rec = str(raw_ref).strip() if raw_ref else ''
            if not doc_no_rec:
                doc_no_rec = 'BNK is Blank'
            receipt_rows.append(_make_receipt_row(doc_no_rec, dt, c2))
            tot_cred += c2

    # 8) Sort combined rows by date in Python, then write to worksheet once
    def _sort_key(row_vals: list) -> date:
        # Keystroke format: value date is column I (index 8), payment date fallback is G (index 6).
        v = row_vals[8] if len(row_vals) > 8 and row_vals[8] else row_vals[6]
        if isinstance(v, str):
            try:
                return _parse_dmon_y(v, '-') or date.min
            except Exception:
                return date.min
        if isinstance(v, (date, datetime)):
            return v.date() if isinstance(v, datetime) else v
        return date.min

    output_data = payment_rows + receipt_rows
    output_data.sort(key=_sort_key)
    _write_rows_to_sheet(ws_out, output_data)

    # 9) Audit summary
    closing_calc = None
    variance = None
    if opening_bal is not None:
        closing_calc = opening_bal + tot_cred - tot_deb
        if closing_bal_stmt is not None:
            variance = closing_bal_stmt - closing_calc

    _write_audit_summary(ws_audit, opening_bal, tot_cred, tot_deb,
                          closing_calc, closing_bal_stmt, variance)

    skipped = audit_row - AUDIT_DETAIL_FIRST_ROW

    msg_parts = [
        'Conversion complete.',
        f'Output rows: {len(output_data)}',
        f'Skipped rows: {skipped}',
        f'Total Debits: {tot_deb:,.2f}',
        f'Total Credits: {tot_cred:,.2f}',
    ]
    if opening_bal is not None:
        msg_parts.append(f'Opening Balance: {opening_bal:,.2f}')
    if closing_bal_stmt is not None:
        msg_parts.append(f'Statement Closing: {closing_bal_stmt:,.2f}')
    if closing_calc is not None:
        msg_parts.append(f'Calculated Closing: {closing_calc:,.2f}')
    if variance is not None:
        msg_parts.append(f'Variance: {variance:,.2f}')
    if skip_tally:
        msg_parts.append('\nSkip breakdown:')
        for reason, cnt in sorted(skip_tally.items(), key=lambda x: -x[1]):
            msg_parts.append(f'  {reason}: {cnt}')

    return ConversionResult(
        success=True,
        message='\n'.join(msg_parts),
        output_rows=len(output_data),
        skipped_rows=skipped,
        total_debits=tot_deb,
        total_credits=tot_cred,
        opening_balance=opening_bal,
        closing_balance_stmt=closing_bal_stmt,
        closing_balance_calc=closing_calc,
        variance=variance,
        output_data=output_data,
        audit_rows=audit_detail_rows,
    )
