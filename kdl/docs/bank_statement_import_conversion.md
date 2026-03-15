# Bank Statement Import & Conversion — Process Documentation

## Overview

The Bank Statement Converter reads an Excel workbook (.xlsx / .xlsm / .xls), parses one or more bank-statement sheets, and outputs two artefacts:

1. **Output sheet** — formatted rows ready to be loaded into the TNT DL grid.
2. **Audit\_Skipped sheet** — every row that was skipped, with the reason.

The converted workbook is saved back to the source file (or a `_converted.xlsx` copy for legacy `.xls` files), and the output data is optionally loaded straight into the grid.

---

## File Map

| File | Role |
|---|---|
| `dialogs/statement_converter_dialog.py` | PySide6 dialog — file picker, sheet selection, progress, save |
| `engine/statement_converter.py` | Pure-Python conversion logic (no UI dependency) |

---

## End-to-End Flow

```
User clicks "Browse…"
  │
  ▼
_SheetLoaderWorker (background thread)
  │  Fast path  : zipfile → xl/workbook.xml  (xlsx/xlsm, ~instant)
  │  Medium path: xlrd on_demand             (.xls)
  │  Slow path  : win32com → Excel           (legacy .xls, no xlrd)
  │  Fallback   : openpyxl read_only
  ▼
Sheet checkboxes populated in UI
  │  Auto-trigger if exactly 1 sheet is pre-selected
  ▼
User clicks "Convert"  (or auto-triggered)
  │
  ▼
_ConverterWorker (background thread)
  │  1. For .xls: COM-convert to tmp .xlsx via Excel, then proceed
  │  2. openpyxl.load_workbook(data_only=True)
  │  3. For each selected sheet → convert_statement()
  │     ├─ Detect header row (Date / Transaction Details / Debit / Credit)
  │     ├─ Map column positions
  │     ├─ Find last data row
  │     ├─ Derive opening & closing balances from "Closing Balance" column
  │     ├─ Contra matching (optional) — debit/credit pairs same ref+amount
  │     ├─ Build payment_rows + receipt_rows as Python lists (no worksheet writes yet)
  │     ├─ Sort combined list by date (in Python, zero extra I/O)
  │     ├─ Write output rows to Output sheet  (single pass)
  │     └─ Return ConversionResult(output_data, audit_rows, …)
  │  4. Multi-sheet: merge output_data + audit_rows across all results
  │     └─ Write merged Output & Audit_Skipped sheets once
  ▼
_on_worker_finished (UI thread)
  │  Display summary in Result text box
  │  Enable "Load Output into Grid" button
  ▼
User clicks "Load Output into Grid"
  │  Save workbook back to source file (or _converted.xlsx for .xls)
  │  Emit load_into_grid signal → grid receives output_data rows
  └─ Dialog closes
```

---

## Performance Design

### Sheet-name reading (the visible delay on "Browse")

| Approach | Cost | When used |
|---|---|---|
| `zipfile` → parse `xl/workbook.xml` | ~5 ms | `.xlsx` / `.xlsm` (primary) |
| `xlrd` `on_demand=True` | ~50 ms | `.xls` (primary) |
| `win32com` / Excel COM | 2–5 s | `.xls` fallback (xlrd fails) |
| `openpyxl` `read_only` | 200 ms–2 s | Universal fallback |

All sheet-name loading runs in `_SheetLoaderWorker` — the UI thread never blocks.

### Conversion I/O passes

Previously the converter did **4 worksheet passes** per sheet:

1. Write payment/receipt rows to `Output`
2. Read them back for sorting
3. Write sorted rows back to `Output`
4. Read again to collect `output_data`

Now it does **1 worksheet pass**:

1. Accumulate `payment_rows` + `receipt_rows` as plain Python lists
2. Sort in Python (zero I/O)
3. Write sorted list to `Output` once
4. `output_data` **is** the sorted list — no read-back needed

### Sheet clearing

`_get_or_create_sheet` previously iterated every cell to set `value = None`.
Now it deletes the sheet and recreates it at the same index — O(1) regardless of sheet size.

---

## Output Sheet Format (columns A–P)

### Payment row (debit)

| Col | Value |
|---|---|
| A | `tab` |
| B | `tab` |
| C | `trfd` |
| D | `tab` |
| E | Doc number (10-digit, stripped of leading zeros) |
| F | `tab` |
| G | Value date (`DD-Mon-YYYY`) |
| H | `tab` |
| I | Payment date (`DD-Mon-YYYY`) |
| J | `tab` |
| K | Amount |
| L | `tab` |
| M | `\*s` |
| N | `*dn` |
| O–P | _(blank)_ |

### Receipt row (credit)

| Col | Value |
|---|---|
| A | `tab` |
| B | `*dn` |
| C | `r` |
| D | `tab` |
| E | `trfc` |
| F | `tab` |
| G | BNK reference (raw) |
| H | `tab` |
| I | Value date (`DD-Mon-YYYY`) |
| J | `tab` |
| K | Value date (repeated) |
| L | `tab` |
| M | Amount |
| N | `tab` |
| O | `\*s` |
| P | `*dn` |

---

## Audit\_Skipped Sheet Format

### Summary block (rows 1–7)

| Row | Label | Col B |
|---|---|---|
| 1 | BANK RECONCILIATION SUMMARY | — |
| 2 | Opening Balance | value or "N/A" |
| 3 | Additions (Receipts) | total credits |
| 4 | Less: Outs (Payments) | total debits |
| 5 | Closing Balance (Calculated) | value |
| 6 | Closing Balance (Statement) | value or "N/A" |
| 7 | Closing Match (Variance) | value or "N/A" |

### Detail block (rows 9–N)

Row 9 is a bold header. Rows 10+ are skipped rows.

| Col | Header |
|---|---|
| A | Row# |
| B | Date (Raw) |
| C | Reference |
| D | Transaction Type |
| E | Transaction Details |
| F | Debit (Raw) |
| G | Credit (Raw) |
| H | Debit (Parsed) |
| I | Credit (Parsed) |
| J | Skip Reason |

### Skip reasons

| Code | Meaning |
|---|---|
| `CONTRA_MATCHED_DETAILS10_AMOUNT` | Debit/credit pair with matching 10-digit reference and amount |
| `BLANK_OR_INVALID_DATE` | Date cell is empty or unparseable |
| `BOTH_DEBIT_AND_CREDIT_PRESENT` | Row has non-zero values in both Debit and Credit |
| `ZERO_OR_NON_NUMERIC_AMOUNT` | Both Debit and Credit are zero or non-numeric |

---

## Supported Date Formats

The converter parses dates in any of these formats (cell value or text):

- Python `date` / `datetime` object (openpyxl native)
- Excel serial integer (e.g. `45000`)
- `MM/DD/YYYY` or `M/D/YY`
- `DD-Mon-YYYY` (e.g. `15-Mar-2024`)
- ISO 8601 (`YYYY-MM-DD`)

---

## Save Behaviour

| Source format | Save target |
|---|---|
| `.xlsx` / `.xlsm` | Overwrites source file |
| `.xls` | Saves as `<original_name>_converted.xlsx` (source kept) |

If the file is locked (open in Excel), the converter attempts to close it via COM before retrying the save.

---

## Dependencies

| Package | Purpose | Required |
|---|---|---|
| `openpyxl` | Read/write `.xlsx` / `.xlsm` | Yes |
| `xlrd` | Fast `.xls` sheet-name reading | Optional (recommended) |
| `win32com` / `pywin32` | `.xls` COM fallback + locked-file save | Windows only, optional |
| `pythoncom` | COM thread init for `.xls` conversion | Windows only, optional |
| `PySide6` | UI framework | Yes |
