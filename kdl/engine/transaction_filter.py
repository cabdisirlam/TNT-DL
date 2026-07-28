"""Standalone IFMIS transaction filtering and cleaning engine."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import os
from pathlib import Path
import re
import tempfile
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo

from kdl.tabular_import import load_workbook_from_source


OUTPUT_HEADERS = (
    "Source",
    "Category",
    "GL Date",
    "Event Class",
    "Transaction Number",
    "Line Description",
    "Debit",
    "Credit",
    "Count",
)

HEADER_ALIASES = {
    "Source": {"source", "journalsource"},
    "Category": {"category", "journalcategory"},
    "GL Date": {"gldate", "accountingdate", "effectivedate", "date"},
    "Event Class": {"eventclass", "eventtypeclass"},
    "Transaction Number": {
        "transactionnumber",
        "transactionno",
        "transactionnum",
        "documentnumber",
    },
    "Line Description": {
        "linedescription",
        "line",
        "description",
        "transactiondescription",
    },
    "Debit": {
        "debit",
        "debitamount",
        "entereddr",
        "accounteddr",
        "dr",
    },
    "Credit": {
        "credit",
        "creditamount",
        "enteredcr",
        "accountedcr",
        "cr",
    },
}

KNOWN_REPORT_LABELS = {
    "",
    "account",
    "beginningbalanceforperiod",
    "endofreport",
    "endingbalanceforperiod",
    "gokledger",
    "ledgername",
    "source",
    "subledgeraccounting",
}

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
TOTAL_FILL = PatternFill("solid", fgColor="D9EAF7")
WHITE_FONT = Font(color="FFFFFF", bold=True)
THIN_BLUE = Side(style="thin", color="9EBCD2")
AMOUNT_FORMAT = '#,##0.00;[Red]-#,##0.00'


@dataclass
class FilterResult:
    output_path: str
    source_sheet: str
    row_count: int
    debit_total: Decimal
    credit_total: Decimal
    message: str


def _normalise_header(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().lower())


def _display_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _decimal_amount(value: Any) -> Decimal | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float, Decimal)):
        try:
            return Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        except (InvalidOperation, ValueError):
            return None

    text = str(value).strip()
    if not text or text in {"-", "—"}:
        return None
    negative = text.startswith("(") and text.endswith(")")
    if negative:
        text = text[1:-1]
    text = (
        text.replace(",", "")
        .replace("KES", "")
        .replace("Ksh", "")
        .replace("KSH", "")
        .strip()
    )
    try:
        amount = Decimal(text)
    except InvalidOperation:
        return None
    if negative:
        amount = -amount
    return amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def _header_mapping(values: list[Any]) -> dict[str, int]:
    normalised = [_normalise_header(value) for value in values]
    mapping: dict[str, int] = {}
    for canonical, aliases in HEADER_ALIASES.items():
        for index, candidate in enumerate(normalised):
            if candidate in aliases:
                mapping[canonical] = index
                break
    return mapping


def _looks_like_data_row(values: list[Any]) -> bool:
    if len(values) < 8:
        return False
    has_transaction = bool(_display_text(values[4]) or _display_text(values[5]))
    has_amount = _decimal_amount(values[6]) is not None or _decimal_amount(values[7]) is not None
    return has_transaction and has_amount


def _locate_data(ws) -> tuple[int, dict[str, int]]:
    scan_rows = min(ws.max_row, 250)
    scan_cols = min(max(ws.max_column, 8), 100)
    for row_number in range(1, scan_rows + 1):
        values = [ws.cell(row_number, column).value for column in range(1, scan_cols + 1)]
        mapping = _header_mapping(values)
        if (
            {"Debit", "Credit"}.issubset(mapping)
            and ("Transaction Number" in mapping or "Line Description" in mapping)
            and len(mapping) >= 5
        ):
            return row_number + 1, mapping

    # Compatibility with the original VBA report whose transaction data begins
    # after a fixed banner block and always occupies A:H.
    for row_number in range(1, scan_rows + 1):
        values = [ws.cell(row_number, column).value for column in range(1, 9)]
        if _looks_like_data_row(values):
            return row_number, {
                name: index for index, name in enumerate(OUTPUT_HEADERS[:8])
            }

    raise ValueError(
        "Could not find the transaction columns. The source must contain Debit, "
        "Credit and either Transaction Number or Line Description columns, or use "
        "the original eight-column A:H layout."
    )


def _serialisable_value(value: Any) -> Any:
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, (datetime, date, int, float, str, bool)) or value is None:
        return value
    return str(value)


def _read_filtered_rows(ws) -> tuple[list[list[Any]], Decimal, Decimal]:
    data_start, mapping = _locate_data(ws)
    rows: list[list[Any]] = []
    debit_total = Decimal("0.00")
    credit_total = Decimal("0.00")

    for row_number in range(data_start, ws.max_row + 1):
        source_values = [
            ws.cell(row_number, mapping[header] + 1).value for header in OUTPUT_HEADERS[:8]
        ]
        if not any(_display_text(value) for value in source_values):
            continue
        if _normalise_header(source_values[0]) in KNOWN_REPORT_LABELS:
            continue
        if len(_header_mapping(source_values)) >= 5:
            continue

        has_transaction = bool(
            _display_text(source_values[4]) or _display_text(source_values[5])
        )
        debit = _decimal_amount(source_values[6])
        credit = _decimal_amount(source_values[7])
        has_context = any(_display_text(value) for value in source_values[:4])
        if not has_transaction and debit is None and credit is None and not has_context:
            continue

        if debit is not None:
            source_values[6] = float(debit)
            debit_total += debit
        if credit is not None:
            credit = -abs(credit) if credit != 0 else Decimal("0.00")
            source_values[7] = float(credit)
            credit_total += credit

        rows.append([_serialisable_value(value) for value in source_values])

    if not rows:
        raise ValueError("No transaction rows were found on the selected worksheet.")
    return rows, debit_total, credit_total


def _fit_columns(ws, rows: list[list[Any]]) -> None:
    for column_number, header in enumerate(OUTPUT_HEADERS, 1):
        values = [_display_text(row[column_number - 1]) for row in rows[:250]]
        width = max([len(header)] + [len(value) for value in values]) + 2
        if header == "Line Description":
            width = min(max(width, 28), 52)
        elif header in {"Debit", "Credit"}:
            width = min(max(width, 14), 18)
        elif header == "Count":
            width = 11
        else:
            width = min(max(width, 11), 24)
        ws.column_dimensions[ws.cell(1, column_number).column_letter].width = width


def _build_output_workbook(
    rows: list[list[Any]],
    source_path: str,
    source_sheet: str,
) -> Workbook:
    workbook = Workbook()
    ws = workbook.active
    ws.title = "Filtered_Data"
    ws.sheet_view.showGridLines = False

    for column, header in enumerate(OUTPUT_HEADERS, 1):
        cell = ws.cell(1, column, header)
        cell.fill = HEADER_FILL
        cell.font = WHITE_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=THIN_BLUE)
    ws.row_dimensions[1].height = 32

    data_last_row = len(rows) + 1
    for row_number, values in enumerate(rows, 2):
        for column, value in enumerate(values, 1):
            ws.cell(row_number, column, value)
        ws.cell(
            row_number,
            9,
            f'=COUNTIF($E$2:$E${data_last_row},E{row_number})',
        )

    total_row = data_last_row + 1
    ws.cell(total_row, 6, "Total")
    ws.cell(total_row, 7, f"=SUM(G2:G{data_last_row})")
    ws.cell(total_row, 8, f"=SUM(H2:H{data_last_row})")
    for cell in ws[total_row][:9]:
        cell.fill = TOTAL_FILL
        cell.font = Font(bold=True, color="1F4E78")

    for row_number in range(2, total_row + 1):
        ws.cell(row_number, 7).number_format = AMOUNT_FORMAT
        ws.cell(row_number, 8).number_format = AMOUNT_FORMAT
        ws.cell(row_number, 3).number_format = "dd-mmm-yyyy"

    table = Table(
        displayName="FilterDataTable",
        ref=f"A1:I{total_row}",
    )
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(table)
    _fit_columns(ws, [row + [None] for row in rows])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:I{total_row}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1

    workbook.properties.title = "IFMIS Transaction Filter"
    workbook.properties.subject = (
        f"Filtered transaction data from {os.path.basename(source_path)} [{source_sheet}]"
    )
    workbook.properties.creator = "NT DL"
    try:
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
    except AttributeError:
        pass
    return workbook


def suggest_filter_output(source_path: str) -> str:
    source = Path(source_path)
    return str(source.with_name(f"{source.stem}_Filtered.xlsx"))


def create_filtered_workbook(
    source_path: str,
    output_path: str,
    sheet_name: str,
) -> FilterResult:
    """Create a filtered workbook without modifying the selected source."""

    source_path = os.path.abspath(source_path)
    output_path = os.path.abspath(output_path)
    if not os.path.isfile(source_path):
        raise FileNotFoundError(f"Source workbook was not found: {source_path}")
    if os.path.normcase(source_path) == os.path.normcase(output_path):
        raise ValueError("The output path must be different from the source workbook.")

    source_wb = load_workbook_from_source(
        source_path,
        sheet_names=[sheet_name] if sheet_name else None,
        data_only=True,
        read_only=False,
        keep_links=False,
    )
    try:
        if sheet_name not in source_wb.sheetnames:
            raise ValueError(f"Worksheet '{sheet_name}' was not found in the source workbook.")
        rows, debit_total, credit_total = _read_filtered_rows(source_wb[sheet_name])
    finally:
        close = getattr(source_wb, "close", None)
        if callable(close):
            close()

    workbook = _build_output_workbook(rows, source_path, sheet_name)
    output_dir = os.path.dirname(output_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    temp_handle = tempfile.NamedTemporaryFile(
        prefix=".transaction_filter_",
        suffix=".xlsx",
        dir=output_dir or None,
        delete=False,
    )
    temp_path = temp_handle.name
    temp_handle.close()
    try:
        workbook.save(temp_path)
        os.replace(temp_path, output_path)
    finally:
        if os.path.exists(temp_path):
            os.unlink(temp_path)

    message = "\n".join(
        (
            "Filter Engine completed.",
            f"Source sheet: {sheet_name}",
            f"Transaction rows retained: {len(rows):,}",
            f"Debit total: {debit_total:,.2f}",
            f"Credit total: {credit_total:,.2f}",
            f"Difference: {debit_total + credit_total:,.2f}",
            "",
            "Count formulas were added using Transaction Number.",
            "The original workbook was not changed.",
        )
    )
    return FilterResult(
        output_path=output_path,
        source_sheet=sheet_name,
        row_count=len(rows),
        debit_total=debit_total,
        credit_total=credit_total,
        message=message,
    )
