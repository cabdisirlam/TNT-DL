"""
GOK IFMIS Notes to Financial Statements Engine (v7)
Pure Python / openpyxl implementation of the VBA macro.

Reads an IFMIS Notes worksheet (openpyxl Worksheet) and produces an
openpyxl Workbook with 5 sheets:
  Notes · Performance · Position · Net Assets · Cash Flow
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

try:
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    _HAS_OPENPYXL = True
except ImportError:
    _HAS_OPENPYXL = False

# ── Sheet names ────────────────────────────────────────────────────────────
SH_NOTES = "Notes"
SH_PERF  = "Performance"
SH_NA    = "Net Assets"
SH_POS   = "Position"
SH_CF    = "Cash Flow"

# ── Formatting constants ───────────────────────────────────────────────────
NUM_FMT = '#,##0.00;[Red](#,##0.00);"-"'
TNR     = "Times New Roman"
_BLUE   = "0070C0"
_WHITE  = "FFFFFF"
_GRAY   = "F2F2F2"
_GREEN  = "008000"
_RED_C  = "FF0000"


# ── Data classes ──────────────────────────────────────────────────────────
@dataclass
class NoteItem:
    desc: str = ""
    code: str = ""
    cur:  float = 0.0
    prev: float = 0.0


@dataclass
class ParsedNote:
    key:        str   = ""
    title:      str   = ""
    items:      list  = field(default_factory=list)
    total_cur:  float = 0.0
    total_prev: float = 0.0
    seq_num:    object = 0   # int or str e.g. "15(a)"


@dataclass
class ReportResult:
    success:  bool            = False
    message:  str             = ""
    workbook: Optional[object] = None   # openpyxl.Workbook when success=True


# ── Public entry point ─────────────────────────────────────────────────────
def generate_ifmis_report(source_ws) -> ReportResult:
    """
    Generate financial statements from a parsed IFMIS Notes worksheet.

    Args:
        source_ws: an openpyxl Worksheet containing the Notes data.

    Returns:
        ReportResult with success flag, message, and workbook.
    """
    if not _HAS_OPENPYXL:
        return ReportResult(
            False,
            "openpyxl is not installed.\n\nRun:  pip install openpyxl"
        )
    try:
        engine = _IFMISEngine(source_ws)
        engine.parse()
        wb = engine.build()
        return ReportResult(True, "Financial statements generated successfully.", wb)
    except Exception as exc:
        import traceback
        return ReportResult(False, f"Generation failed: {exc}\n\n{traceback.format_exc()}")


# ── Engine ─────────────────────────────────────────────────────────────────
class _IFMISEngine:
    def __init__(self, source_ws):
        self._ws     = source_ws
        self._notes: list[ParsedNote] = []
        self._entity      = ""
        self._cur_period  = ""
        self._prev_period = ""
        # Aggregates computed in _compute_aggregates
        self._surplus_cur   = 0.0
        self._surplus_prev  = 0.0
        self._net_assets_cur  = 0.0
        self._net_assets_prev = 0.0
        self._na_prior_opening = 0.0
        self._na_close_prev    = 0.0
        self._na_open_cur      = 0.0
        self._na_close_cur     = 0.0
        self._rte_cur          = 0.0   # Return to Exchequer (Note 26 current)
        # Revenue subtotals
        self._total_rev_cur  = 0.0
        self._total_rev_prev = 0.0
        self._total_exp_cur  = 0.0
        self._total_exp_prev = 0.0
        # Asset/liability subtotals
        self._total_ca_cur   = 0.0
        self._total_ca_prev  = 0.0
        self._total_nca_cur  = 0.0
        self._total_nca_prev = 0.0
        self._total_assets_cur   = 0.0
        self._total_assets_prev  = 0.0
        self._total_liab_cur  = 0.0
        self._total_liab_prev = 0.0
        # Position sheet rows for late-binding Accumulated Surplus
        self._pos_acc_row   = 0
        self._pos_total_row = 0

    # ── Parse ──────────────────────────────────────────────────────────────
    def parse(self):
        self._parse_header()
        self._parse_notes()
        self._combine_two("6",   "7",   "6_7",  "Transfers from Domestic and Foreign Partners")
        self._combine_two("22A", "22B", "CASH", "Cash and Cash Equivalents")
        self._make_ppe_cumulative()
        self._make_wc_note()
        self._assign_seq_nums()
        self._compute_aggregates()

    def _parse_header(self):
        for r in range(1, 11):
            raw = self._ws.cell(row=r, column=2).value
            t   = str(raw or "").replace("\xa0", " ")
            if "Entity:" in t:
                self._entity = t[t.index("Entity:") + 7:].strip()
            elif "Current Period:" in t:
                self._cur_period = t[t.index("Current Period:") + 15:].strip()
            elif "Compare With:" in t:
                self._prev_period = t[t.index("Compare With:") + 13:].strip()

    def _parse_notes(self):
        max_row  = self._ws.max_row or 0
        in_items = False
        ci       = -1
        for r in range(1, max_row + 1):
            raw = self._ws.cell(row=r, column=1).value
            c1  = str(raw or "").replace("\xa0", " ").strip()
            if not c1:
                continue

            nk, nt = _is_note_title(c1)
            if nk:
                self._notes.append(ParsedNote(key=nk, title=nt))
                ci       = len(self._notes) - 1
                in_items = False
                continue

            if c1 == "Item Description":
                in_items = True
                continue
            if c1 == "Kshs":
                continue
            if c1 == "TOTAL" and ci >= 0:
                self._notes[ci].total_cur  = _to_float(self._ws.cell(row=r, column=3).value)
                self._notes[ci].total_prev = _to_float(self._ws.cell(row=r, column=4).value)
                in_items = False
                continue
            if in_items and ci >= 0:
                cv = _to_float(self._ws.cell(row=r, column=3).value)
                pv = _to_float(self._ws.cell(row=r, column=4).value)
                if cv != 0 or pv != 0:
                    code = str(self._ws.cell(row=r, column=2).value or "")
                    self._notes[ci].items.append(NoteItem(desc=c1, code=code, cur=cv, prev=pv))

    def _combine_two(self, k1: str, k2: str, nk: str, nt: str):
        n1 = self._find(k1)
        n2 = self._find(k2)
        if n1 is None and n2 is None:
            return
        merged = ParsedNote(key=nk, title=nt)
        if n1:
            merged.items.extend(n1.items)
            merged.total_cur  += n1.total_cur
            merged.total_prev += n1.total_prev
            n1.key = f"SKIP_{k1}"
        if n2:
            merged.items.extend(n2.items)
            merged.total_cur  += n2.total_cur
            merged.total_prev += n2.total_prev
            n2.key = f"SKIP_{k2}"
        self._notes.append(merged)

    def _make_ppe_cumulative(self):
        n = self._find("18")
        if n is None or not n.items:
            return
        opening = NoteItem(
            desc="Opening Balance (Prior Period PPE)",
            code="",
            cur=n.total_prev,
            prev=0.0,
        )
        n.items.insert(0, opening)
        n.total_cur = n.total_prev + n.total_cur
        n.title = "Property, Plant and Equipment"

    def _make_wc_note(self):
        n23 = self._find("23")
        n24 = self._find("24")
        wc  = ParsedNote(key="WC", title="Changes in Working Capital")
        if n23 and (n23.total_cur != 0 or n23.total_prev != 0):
            item = NoteItem(
                desc="Movement in Receivables (Imprest & Clearance)",
                code="Note 23",
                cur=n23.total_prev - n23.total_cur,
                prev=0.0,
            )
            wc.items.append(item)
            wc.total_cur += item.cur
        if n24 and (n24.total_cur != 0 or n24.total_prev != 0):
            item = NoteItem(
                desc="Movement in Trade and Other Payables",
                code="Note 24",
                cur=n24.total_cur - n24.total_prev,
                prev=0.0,
            )
            wc.items.append(item)
            wc.total_cur += item.cur
        self._notes.append(wc)

    def _assign_seq_nums(self):
        order = [
            "4", "6_7", "3", "1", "2", "9", "8", "11",
            "12", "13", "14", "15", "16", "17", "19", "21",
            "20", "CASH", "23", "18", "24", "26", "WC",
        ]
        sq    = 6
        t_num = 0
        for k in order:
            n = self._find(k)
            if n is None:
                continue
            if not (n.total_cur != 0 or n.total_prev != 0):
                continue
            if not n.items:
                continue
            if k == "WC":
                n.seq_num = f"{t_num}(a)" if t_num > 0 else sq
                if not t_num:
                    sq += 1
            else:
                n.seq_num = sq
                if k == "15":
                    t_num = sq
                sq += 1

    def _compute_aggregates(self):
        # Revenue
        rev_ne_keys = ["4", "6_7", "3", "1", "2", "9", "8"]
        rev_ne_cur  = sum(self._nc(k) for k in rev_ne_keys)
        rev_ne_prev = sum(self._np(k) for k in rev_ne_keys)
        rev_ex_cur  = self._nc("11")
        rev_ex_prev = self._np("11")
        self._rev_ne_cur  = rev_ne_cur
        self._rev_ne_prev = rev_ne_prev
        self._rev_ex_cur  = rev_ex_cur
        self._rev_ex_prev = rev_ex_prev
        self._total_rev_cur  = rev_ne_cur + rev_ex_cur
        self._total_rev_prev = rev_ne_prev + rev_ex_prev

        # Expenses
        exp_keys = ["12", "13", "14", "15", "16", "17", "19", "21"]
        self._total_exp_cur  = sum(self._nc(k) for k in exp_keys)
        self._total_exp_prev = sum(self._np(k) for k in exp_keys)

        # Surplus / Deficit
        self._surplus_cur  = self._total_rev_cur  - self._total_exp_cur
        self._surplus_prev = self._total_rev_prev - self._total_exp_prev

        # Assets
        self._total_ca_cur   = self._nc("CASH") + self._nc("23")
        self._total_ca_prev  = self._np("CASH") + self._np("23")
        self._total_nca_cur  = self._nc("18")
        self._total_nca_prev = self._np("18")
        self._total_assets_cur  = self._total_ca_cur  + self._total_nca_cur
        self._total_assets_prev = self._total_ca_prev + self._total_nca_prev

        # Liabilities
        self._total_liab_cur  = self._nc("24")
        self._total_liab_prev = self._np("24")

        # Net Assets
        self._net_assets_cur  = self._total_assets_cur  - self._total_liab_cur
        self._net_assets_prev = self._total_assets_prev - self._total_liab_prev

        # Return to Exchequer (Note 26, current period only)
        self._rte_cur = self._nc("26")

        # Net Assets statement values
        self._na_prior_opening = self._net_assets_prev - self._surplus_prev
        self._na_close_prev    = self._net_assets_prev
        self._na_open_cur      = self._net_assets_prev   # = prior closing
        self._na_close_cur     = self._na_open_cur + self._surplus_cur - self._rte_cur

    # ── Build workbook ─────────────────────────────────────────────────────
    def build(self) -> object:
        wb = Workbook()
        wb.remove(wb.active)
        self._build_notes_sheet(wb)
        self._build_perf_sheet(wb)
        self._build_pos_sheet(wb)
        self._build_na_sheet(wb)
        self._build_cf_sheet(wb)
        return wb

    # ── Notes sheet ────────────────────────────────────────────────────────
    def _build_notes_sheet(self, wb):
        ws = wb.create_sheet(SH_NOTES)
        r  = _write_title(ws, 1, 4, "NOTES TO THE FINANCIAL STATEMENTS",
                          self._entity, self._cur_period)
        r += 1

        order = [
            "4", "6_7", "3", "1", "2", "9", "8", "11",
            "12", "13", "14", "15", "WC", "16", "17", "19", "21",
            "20", "CASH", "23", "18", "24", "26",
        ]
        for k in order:
            n = self._find(k)
            if n is None or not n.seq_num or not n.items:
                continue

            # Note title (merged A:D)
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
            c = ws.cell(row=r, column=1, value=f"Note {n.seq_num}. {n.title}")
            c.font = Font(name=TNR, size=12, bold=True)
            r += 1

            # Header row 1
            for col, val in enumerate(
                ["Item Description", "Item Code", "Current Period", "Previous Period"], 1
            ):
                _blue_hdr(ws.cell(row=r, column=col, value=val))
            r += 1

            # Header row 2 (Kshs)
            for col in range(1, 5):
                val = "Kshs" if col in (3, 4) else None
                _blue_hdr(ws.cell(row=r, column=col, value=val))
            r += 1

            # Items
            for i, item in enumerate(n.items):
                ws.cell(row=r, column=1, value=item.desc).font = Font(name=TNR, size=11)
                ws.cell(row=r, column=2, value=item.code).font = Font(name=TNR, size=11)
                _num(ws.cell(row=r, column=3), item.cur)
                _num(ws.cell(row=r, column=4), item.prev)
                if i % 2 == 0:
                    gf = PatternFill(start_color=_GRAY, end_color=_GRAY, fill_type="solid")
                    for c in range(1, 5):
                        ws.cell(row=r, column=c).fill = gf
                r += 1

            # Total row
            tc = sum(it.cur  for it in n.items)
            tp = sum(it.prev for it in n.items)
            ws.cell(row=r, column=1, value="TOTAL").font = Font(name=TNR, size=11, bold=True)
            _num(ws.cell(row=r, column=3), tc, bold=True)
            _num(ws.cell(row=r, column=4), tp, bold=True)
            for col in range(1, 5):
                ws.cell(row=r, column=col).border = Border(bottom=Side(style="thin"))
            r += 2

        ws.column_dimensions["A"].width = 60
        ws.column_dimensions["B"].width = 14
        ws.column_dimensions["C"].width = 22
        ws.column_dimensions["D"].width = 22

    # ── Performance sheet ──────────────────────────────────────────────────
    def _build_perf_sheet(self, wb):
        ws = wb.create_sheet(SH_PERF)
        r  = _write_title(ws, 2, 5, "STATEMENT OF FINANCIAL PERFORMANCE",
                          self._entity, self._cur_period)
        _blue_hdr(ws.cell(row=r, column=2))
        for col, val in [(3, "Notes"), (4, self._cur_period), (5, self._prev_period)]:
            _blue_hdr(ws.cell(row=r, column=col, value=val))
        r += 1
        for col, val in [(2, None), (3, None), (4, "Kshs"), (5, "Kshs")]:
            _blue_hdr(ws.cell(row=r, column=col, value=val))
        r += 1

        # Non-exchange revenue
        ws.cell(row=r, column=2,
                value="Revenue from non-exchange transactions").font = Font(name=TNR, size=11, bold=True)
        r += 1
        ne_items = [
            ("Transfers from Exchequer",                       "4"),
            ("Transfers from Domestic and Foreign Partners",   "6_7"),
            ("Proceeds from Domestic and Foreign Grants",      "3"),
            ("Tax Receipts",                                   "1"),
            ("Social Security Contributions",                  "2"),
            ("Reimbursements and Refunds",                     "9"),
            ("Proceeds from Sale of Assets",                   "8"),
        ]
        s = r
        for lbl, k in ne_items:
            if self._has(k):
                self._line(ws, r, lbl, k)
                r += 1
        if s <= r - 1:
            _total_line(ws, r, "Total non-exchange revenue",
                        self._rev_ne_cur, self._rev_ne_prev)
            r += 2

        # Exchange revenue
        ws.cell(row=r, column=2,
                value="Revenue from exchange transactions").font = Font(name=TNR, size=11, bold=True)
        r += 1
        s = r
        if self._has("11"):
            self._line(ws, r, "Other Receipts", "11")
            r += 1
        if s <= r - 1:
            _total_line(ws, r, "Total exchange revenue",
                        self._rev_ex_cur, self._rev_ex_prev)
            r += 2

        # Total Revenue
        ws.cell(row=r, column=2, value="Total Revenue").font = Font(name=TNR, size=11, bold=True)
        _num(ws.cell(row=r, column=4), self._total_rev_cur,  bold=True)
        _num(ws.cell(row=r, column=5), self._total_rev_prev, bold=True)
        for c in range(2, 6):
            ws.cell(row=r, column=c).border = Border(bottom=Side(style="thin"))
        r += 2

        # Expenses
        ws.cell(row=r, column=2,
                value="Expenses").font = Font(name=TNR, size=11, bold=True)
        r += 1
        exp_items = [
            ("Compensation of Employees",          "12"),
            ("Use of Goods and Services",           "13"),
            ("Subsidies",                           "14"),
            ("Transfers to Other Government Units", "15"),
            ("Other Grants and Transfers",          "16"),
            ("Social Security Benefits",            "17"),
            ("Finance Costs",                       "19"),
            ("Other Payments",                      "21"),
        ]
        s = r
        for lbl, k in exp_items:
            if self._has(k):
                self._line(ws, r, lbl, k)
                r += 1
        if s <= r - 1:
            _total_line(ws, r, "Total Expenses",
                        self._total_exp_cur, self._total_exp_prev)
            r += 2

        # Surplus / Deficit
        r += 1
        ws.cell(row=r, column=2,
                value="Surplus/(Deficit) for the Period").font = Font(name=TNR, size=11, bold=True)
        _num(ws.cell(row=r, column=4), self._surplus_cur,  bold=True)
        _num(ws.cell(row=r, column=5), self._surplus_prev, bold=True)
        for c in range(2, 6):
            ws.cell(row=r, column=c).border = Border(bottom=Side(style="double"))
        _finalize_sheet(ws)

    # ── Position sheet ─────────────────────────────────────────────────────
    def _build_pos_sheet(self, wb):
        ws = wb.create_sheet(SH_POS)
        r  = _write_title(ws, 2, 5, "STATEMENT OF FINANCIAL POSITION",
                          self._entity, self._cur_period)
        _blue_hdr(ws.cell(row=r, column=2))
        for col, val in [(3, "Notes"), (4, self._cur_period), (5, self._prev_period)]:
            _blue_hdr(ws.cell(row=r, column=col, value=val))
        r += 1
        for col, val in [(2, None), (3, None), (4, "Kshs"), (5, "Kshs")]:
            _blue_hdr(ws.cell(row=r, column=col, value=val))
        r += 1

        # ASSETS
        ws.cell(row=r, column=2, value="ASSETS").font = Font(name=TNR, size=11, bold=True)
        r += 1
        ws.cell(row=r, column=2, value="Current Assets").font = Font(name=TNR, size=11, bold=True)
        r += 1
        s = r
        if self._has("CASH"):
            self._line(ws, r, "Cash and Cash Equivalents",        "CASH")
            r += 1
        if self._has("23"):
            self._line(ws, r, "Receivables - Imprest & Clearance", "23")
            r += 1
        if s <= r - 1:
            _total_line(ws, r, "Total Current Assets",
                        self._total_ca_cur, self._total_ca_prev)
            r += 2

        ws.cell(row=r, column=2,
                value="Non-Current Assets").font = Font(name=TNR, size=11, bold=True)
        r += 1
        s = r
        if self._has("18"):
            self._line(ws, r, "Property, Plant and Equipment", "18")
            r += 1
        if s <= r - 1:
            _total_line(ws, r, "Total Non-Current Assets",
                        self._total_nca_cur, self._total_nca_prev)
            r += 2

        ws.cell(row=r, column=2, value="Total Assets (a)").font = Font(name=TNR, size=11, bold=True)
        _num(ws.cell(row=r, column=4), self._total_assets_cur,  bold=True)
        _num(ws.cell(row=r, column=5), self._total_assets_prev, bold=True)
        for c in range(2, 6):
            ws.cell(row=r, column=c).border = Border(bottom=Side(style="thin"))
        r += 2

        # LIABILITIES
        ws.cell(row=r, column=2, value="LIABILITIES").font = Font(name=TNR, size=11, bold=True)
        r += 1
        ws.cell(row=r, column=2,
                value="Current Liabilities").font = Font(name=TNR, size=11, bold=True)
        r += 1
        s = r
        if self._has("24"):
            self._line(ws, r, "Accounts Payable", "24")
            r += 1
        if s <= r - 1:
            _total_line(ws, r, "Total Current Liabilities",
                        self._total_liab_cur, self._total_liab_prev)
            r += 2

        ws.cell(row=r, column=2,
                value="Total Liabilities (b)").font = Font(name=TNR, size=11, bold=True)
        _num(ws.cell(row=r, column=4), self._total_liab_cur,  bold=True)
        _num(ws.cell(row=r, column=5), self._total_liab_prev, bold=True)
        for c in range(2, 6):
            ws.cell(row=r, column=c).border = Border(bottom=Side(style="thin"))
        r += 2

        # Net Assets
        ws.cell(row=r, column=2,
                value="Net Assets (a - b)").font = Font(name=TNR, size=11, bold=True)
        _num(ws.cell(row=r, column=4), self._net_assets_cur,  bold=True)
        _num(ws.cell(row=r, column=5), self._net_assets_prev, bold=True)
        for c in range(2, 6):
            ws.cell(row=r, column=c).border = Border(bottom=Side(style="double"))
        r += 2

        # Represented by
        ws.cell(row=r, column=2,
                value="Represented by:").font = Font(name=TNR, size=11, bold=True)
        r += 1
        ws.cell(row=r, column=2,
                value="Accumulated Surplus/(Deficit)").font = Font(name=TNR, size=11)
        # Filled later by _link_pos_equity
        self._pos_acc_row = r
        r += 1

        ws.cell(row=r, column=2,
                value="Total Net Assets").font = Font(name=TNR, size=11, bold=True)
        self._pos_total_row = r
        for c in range(2, 6):
            ws.cell(row=r, column=c).border = Border(bottom=Side(style="double"))
        _finalize_sheet(ws)

    # ── Net Assets sheet ───────────────────────────────────────────────────
    def _build_na_sheet(self, wb):
        ws = wb.create_sheet(SH_NA)
        r  = _write_title(ws, 1, 3, "STATEMENT OF CHANGES IN NET ASSETS",
                          self._entity, self._cur_period)
        _blue_hdr(ws.cell(row=r, column=1))
        _blue_hdr(ws.cell(row=r, column=2, value="Accumulated Surplus"))
        _blue_hdr(ws.cell(row=r, column=3, value="Total"))
        r += 1

        # ── Prior year section ──
        ws.cell(row=r, column=1,
                value="Balance as at 1st July").font = Font(name=TNR, size=11)
        _num(ws.cell(row=r, column=2), self._na_prior_opening, color=_GREEN)
        _num(ws.cell(row=r, column=3), self._na_prior_opening, color=_GREEN)
        r += 1

        ws.cell(row=r, column=1,
                value="Surplus/(Deficit) for the year").font = Font(name=TNR, size=11)
        _num(ws.cell(row=r, column=2), self._surplus_prev, color=_GREEN)
        _num(ws.cell(row=r, column=3), self._surplus_prev, color=_GREEN)
        r += 1

        ws.cell(row=r, column=1,
                value="As at June 30, 20xx").font = Font(name=TNR, size=11, bold=True)
        _num(ws.cell(row=r, column=2), self._na_close_prev, bold=True)
        _num(ws.cell(row=r, column=3), self._na_close_prev, bold=True)
        for c in range(1, 4):
            ws.cell(row=r, column=c).border = Border(bottom=Side(style="thin"))
        r += 2

        # ── Current year section ──
        ws.cell(row=r, column=1,
                value="As at July 1, 20xx").font = Font(name=TNR, size=11)
        _num(ws.cell(row=r, column=2), self._na_open_cur)
        _num(ws.cell(row=r, column=3), self._na_open_cur)
        r += 1

        ws.cell(row=r, column=1,
                value="Surplus/(Deficit) for the year").font = Font(name=TNR, size=11)
        _num(ws.cell(row=r, column=2), self._surplus_cur, color=_GREEN)
        _num(ws.cell(row=r, column=3), self._surplus_cur, color=_GREEN)
        r += 1

        # Return to Exchequer (Note 26, current period — reduces net assets)
        rte_val = -self._rte_cur
        ws.cell(row=r, column=1,
                value="Return to Exchequer").font = Font(name=TNR, size=11)
        _num(ws.cell(row=r, column=2), rte_val, color=_GREEN)
        _num(ws.cell(row=r, column=3), rte_val, color=_GREEN)
        r += 1

        ws.cell(row=r, column=1,
                value="As at June 30, 20xx").font = Font(name=TNR, size=11, bold=True)
        _num(ws.cell(row=r, column=2), self._na_close_cur, bold=True)
        _num(ws.cell(row=r, column=3), self._na_close_cur, bold=True)
        for c in range(1, 4):
            ws.cell(row=r, column=c).border = Border(bottom=Side(style="double"))

        ws.column_dimensions["A"].width = 55
        ws.column_dimensions["B"].width = 28
        ws.column_dimensions["C"].width = 22

        # Link back to Position sheet
        pos_ws = wb[SH_POS]
        _num(pos_ws.cell(row=self._pos_acc_row,   column=4), self._na_close_cur,  color=_GREEN)
        _num(pos_ws.cell(row=self._pos_acc_row,   column=5), self._na_close_prev, color=_GREEN)
        _num(pos_ws.cell(row=self._pos_total_row, column=4), self._na_close_cur,  bold=True)
        _num(pos_ws.cell(row=self._pos_total_row, column=5), self._na_close_prev, bold=True)

    # ── Cash Flow sheet ────────────────────────────────────────────────────
    def _build_cf_sheet(self, wb):
        ws = wb.create_sheet(SH_CF)
        r  = _write_title(ws, 2, 5, "STATEMENT OF CASH FLOWS",
                          self._entity, self._cur_period)
        _blue_hdr(ws.cell(row=r, column=2))
        for col, val in [(3, "Notes"), (4, self._cur_period), (5, self._prev_period)]:
            _blue_hdr(ws.cell(row=r, column=col, value=val))
        r += 1
        for col, val in [(2, None), (3, None), (4, "Kshs"), (5, "Kshs")]:
            _blue_hdr(ws.cell(row=r, column=col, value=val))
        r += 1

        # ── Operating activities ──
        ws.cell(row=r, column=2,
                value="Cash flows from operating activities").font = Font(name=TNR, size=11, bold=True)
        r += 1
        ws.cell(row=r, column=2,
                value="Receipts").font = Font(name=TNR, size=11, bold=True)
        r += 1

        rec_items = [
            ("Exchequer Releases",                             "4"),
            ("Transfers from Domestic and Foreign Partners",   "6_7"),
            ("Tax Receipts",                                   "1"),
            ("Social Security Contributions",                  "2"),
            ("Proceeds from Domestic and Foreign Grants",      "3"),
            ("Proceeds from Sale of Assets",                   "8"),
            ("Other Receipts",                                 "11"),
        ]
        rec_cur = rec_prev = 0.0
        for lbl, k in rec_items:
            if self._has(k):
                self._line(ws, r, lbl, k)
                rec_cur  += self._nc(k)
                rec_prev += self._np(k)
                r += 1
        if rec_cur != 0 or rec_prev != 0:
            _total_line(ws, r, "Total Receipts", rec_cur, rec_prev)
            r += 2

        ws.cell(row=r, column=2,
                value="Payments").font = Font(name=TNR, size=11, bold=True)
        r += 1
        pay_items = [
            ("Compensation of Employees",          "12"),
            ("Use of Goods and Services",           "13"),
            ("Transfers to Other Government Units", "15"),
            ("Other Grants and Transfers",          "16"),
            ("Social Security Benefits",            "17"),
        ]
        pay_cur = pay_prev = 0.0
        for lbl, k in pay_items:
            if self._has(k):
                self._line(ws, r, lbl, k, negate=True)
                pay_cur  += -self._nc(k)
                pay_prev += -self._np(k)
                r += 1
        if pay_cur != 0 or pay_prev != 0:
            _total_line(ws, r, "Total Payments", pay_cur, pay_prev)
            r += 1

        # Working Capital
        r += 1
        wc     = self._find("WC")
        wc_cur = wc.total_cur if wc else 0.0
        ws.cell(row=r, column=2,
                value="Changes in Working Capital").font = Font(name=TNR, size=11)
        if wc and wc.seq_num:
            ws.cell(row=r, column=3,
                    value=str(wc.seq_num)).font = Font(name=TNR, size=11)
        _num(ws.cell(row=r, column=4), wc_cur)
        _num(ws.cell(row=r, column=5), 0.0)
        r += 2

        # Net operating
        net_ops_cur  = rec_cur  + pay_cur  + wc_cur
        net_ops_prev = rec_prev + pay_prev
        ws.cell(row=r, column=2,
                value="Net cash flows from/(used in) operating activities"
                ).font = Font(name=TNR, size=11, bold=True)
        _num(ws.cell(row=r, column=4), net_ops_cur,  bold=True)
        _num(ws.cell(row=r, column=5), net_ops_prev, bold=True)
        for c in range(2, 6):
            ws.cell(row=r, column=c).border = Border(bottom=Side(style="thin"))
        r += 2

        # ── Investing activities ──
        ws.cell(row=r, column=2,
                value="Cash flows from investing activities").font = Font(name=TNR, size=11, bold=True)
        r += 1
        inv_cur = inv_prev = 0.0
        ppe = self._find("18")
        if ppe and ppe.items:
            # items[0] is the inserted opening balance; new purchases = total - opening
            opening_cur = ppe.items[0].cur
            acq_cur  = -(ppe.total_cur - opening_cur)
            acq_prev = -ppe.total_prev
            ws.cell(row=r, column=2,
                    value="Acquisition of Assets").font = Font(name=TNR, size=11)
            ws.cell(row=r, column=3,
                    value=str(ppe.seq_num) if ppe.seq_num else "").font = Font(name=TNR, size=11)
            _num(ws.cell(row=r, column=4), acq_cur)
            _num(ws.cell(row=r, column=5), acq_prev)
            inv_cur  = acq_cur
            inv_prev = acq_prev
            r += 1
        if inv_cur != 0 or inv_prev != 0:
            _total_line(ws, r,
                        "Net cash flows from/(used in) investing activities",
                        inv_cur, inv_prev)
            r += 2

        # ── Financing activities (Return to Exchequer — Note 26) ──
        fin_cur = fin_prev = 0.0
        n26 = self._find("26")
        if n26 and n26.items and self._nc("26") != 0:
            ws.cell(row=r, column=2,
                    value="Cash flows from financing activities").font = Font(name=TNR, size=11, bold=True)
            r += 1
            ws.cell(row=r, column=2,
                    value="Return to Exchequer").font = Font(name=TNR, size=11)
            ws.cell(row=r, column=3,
                    value=str(n26.seq_num) if n26.seq_num else "").font = Font(name=TNR, size=11)
            fin_cur = -self._nc("26")
            _num(ws.cell(row=r, column=4), fin_cur)
            _num(ws.cell(row=r, column=5), 0.0)
            r += 1
            _total_line(ws, r,
                        "Net cash flows from/(used in) financing activities",
                        fin_cur, 0.0)
            r += 2

        # Net change in cash
        net_chg_cur  = net_ops_cur  + inv_cur  + fin_cur
        net_chg_prev = net_ops_prev + inv_prev + fin_prev
        ws.cell(row=r, column=2,
                value="Net increase/(decrease) in cash and cash equivalents"
                ).font = Font(name=TNR, size=11, bold=True)
        _num(ws.cell(row=r, column=4), net_chg_cur,  bold=True)
        _num(ws.cell(row=r, column=5), net_chg_prev, bold=True)
        for c in range(2, 6):
            ws.cell(row=r, column=c).border = Border(bottom=Side(style="thin"))
        r += 2

        # Opening / Closing cash
        cash_sn    = str(self._find("CASH").seq_num) if self._find("CASH") else ""
        cash_prev  = self._np("CASH")
        cash_cur   = self._nc("CASH")
        cash_start_cur  = cash_prev
        cash_start_prev = cash_prev - net_chg_prev

        ws.cell(row=r, column=2,
                value="Cash and cash equivalents at start of period").font = Font(name=TNR, size=11)
        ws.cell(row=r, column=3, value=cash_sn).font = Font(name=TNR, size=11)
        _num(ws.cell(row=r, column=4), cash_start_cur)
        _num(ws.cell(row=r, column=5), cash_start_prev)
        r += 1

        cash_end_cur  = cash_start_cur  + net_chg_cur
        cash_end_prev = cash_start_prev + net_chg_prev
        ws.cell(row=r, column=2,
                value="Cash and cash equivalents at end of period"
                ).font = Font(name=TNR, size=11, bold=True)
        ws.cell(row=r, column=3, value=cash_sn).font = Font(name=TNR, size=11)
        _num(ws.cell(row=r, column=4), cash_end_cur,  bold=True)
        _num(ws.cell(row=r, column=5), cash_end_prev, bold=True)
        for c in range(2, 6):
            ws.cell(row=r, column=c).border = Border(bottom=Side(style="double"))
        r += 2

        # ── Controls ──
        def _red_ctrl(row, label, cur_v, prev_v):
            ws.cell(row=row, column=2,
                    value=label).font = Font(name=TNR, size=11, bold=True, color=_RED_C)
            _num(ws.cell(row=row, column=4), cur_v,  color=_RED_C)
            _num(ws.cell(row=row, column=5), prev_v, color=_RED_C)

        _red_ctrl(r, "CONTROL: Cash per Note", cash_cur, cash_prev)
        ws.cell(row=r, column=3, value=cash_sn).font = Font(name=TNR, size=11, color=_RED_C)
        r += 1
        _red_ctrl(r, "CONTROL: Difference (CF vs Note)",
                  cash_end_cur - cash_cur,
                  cash_end_prev - cash_prev)
        r += 2
        _red_ctrl(r, "CONTROL: Position Balance Check",
                  self._net_assets_cur  - self._na_close_cur,
                  self._net_assets_prev - self._na_close_prev)

        _finalize_sheet(ws)

    # ── Helpers ────────────────────────────────────────────────────────────
    def _find(self, key: str) -> Optional[ParsedNote]:
        for n in self._notes:
            if n.key == key:
                return n
        return None

    def _nc(self, key: str) -> float:
        n = self._find(key)
        return n.total_cur if n else 0.0

    def _np(self, key: str) -> float:
        n = self._find(key)
        return n.total_prev if n else 0.0

    def _has(self, key: str) -> bool:
        n = self._find(key)
        return n is not None and (n.total_cur != 0 or n.total_prev != 0)

    def _line(self, ws, r: int, label: str, key: str, negate: bool = False):
        """Write one line row in Performance / Position / CF sheets."""
        n    = self._find(key)
        sign = -1 if negate else 1
        cv   = (n.total_cur  if n else 0.0) * sign
        pv   = (n.total_prev if n else 0.0) * sign
        seq  = str(n.seq_num) if n and n.seq_num else ""
        ws.cell(row=r, column=2, value=label).font = Font(name=TNR, size=11)
        ws.cell(row=r, column=3, value=seq).font   = Font(name=TNR, size=11)
        _num(ws.cell(row=r, column=4), cv)
        _num(ws.cell(row=r, column=5), pv)


# ── Module-level formatting helpers ───────────────────────────────────────
def _write_title(ws, col_start: int, col_end: int, title: str,
                 entity: str, cur_period: str) -> int:
    """Write 3 title rows; return the next row number (4)."""
    rows = [entity, f"FOR THE PERIOD ENDED {cur_period}", title]
    for i, text in enumerate(rows, start=1):
        ws.merge_cells(start_row=i, start_column=col_start,
                       end_row=i,   end_column=col_end)
        c = ws.cell(row=i, column=col_start, value=text)
        c.font      = Font(name=TNR, bold=True, size=13 if i == 3 else 11)
        c.alignment = Alignment(horizontal="center")
    return 4


def _blue_hdr(cell):
    cell.font      = Font(name=TNR, bold=True, size=11, color=_WHITE)
    cell.fill      = PatternFill(start_color=_BLUE, end_color=_BLUE, fill_type="solid")
    cell.alignment = Alignment(horizontal="center", vertical="center")


def _num(cell, value: float, bold: bool = False, color: str = None):
    cell.value         = value
    cell.number_format = NUM_FMT
    kwargs = {"name": TNR, "size": 11, "bold": bold}
    if color:
        kwargs["color"] = color
    cell.font      = Font(**kwargs)
    cell.alignment = Alignment(horizontal="right")


def _total_line(ws, r: int, label: str, cur_val: float, prev_val: float):
    ws.cell(row=r, column=2, value=label).font = Font(name=TNR, size=11, bold=True)
    _num(ws.cell(row=r, column=4), cur_val,  bold=True)
    _num(ws.cell(row=r, column=5), prev_val, bold=True)
    for c in range(2, 6):
        ws.cell(row=r, column=c).border = Border(bottom=Side(style="thin"))


def _finalize_sheet(ws):
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 22


def _to_float(v) -> float:
    try:
        return float(v) if v is not None else 0.0
    except (ValueError, TypeError):
        return 0.0


def _is_note_title(txt: str):
    """
    Parse a note title line like '4. Transfers from Exchequer'
    Returns (key, title) or ('', '') if not a note title.
    """
    if not txt:
        return "", ""
    i = 0
    while i < len(txt) and txt[i].isdigit():
        i += 1
    if i == 0:
        return "", ""
    # Optional letter suffix (A / B)
    if i < len(txt) and txt[i].upper() in ("A", "B"):
        i += 1
    num  = txt[:i]
    rest = txt[i:]
    rest_stripped = rest.strip()
    if rest_stripped.startswith("."):
        title = rest_stripped[1:].strip()
    elif len(rest) >= 2 and rest[:2] == "  ":
        title = rest_stripped
    else:
        return "", ""
    if len(title) < 3:
        return "", ""
    return num, title
