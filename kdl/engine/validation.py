"""
NT_DL built-in data validation helpers.

Focused on IFMIS/Oracle transaction templates:
- required fields
- date format checks
- numeric amount checks
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import Dict, List, Optional, Sequence, Set


DATE_RE = re.compile(r"^\d{1,2}-[A-Za-z]{3}-\d{4}$")
VALID_TYPES = {"payment", "receipt"}


@dataclass
class ValidationIssue:
    severity: str  # "error" or "warning"
    row: int       # 0-based
    col: int       # 0-based
    message: str


def _normalize_header(text: str) -> str:
    value = str(text or "").strip().lower()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def _is_empty(value) -> bool:
    return str(value or "").strip() == ""


def _safe_cell(row: Sequence, col: Optional[int]) -> str:
    if col is None or col < 0 or col >= len(row):
        return ""
    return str(row[col] if row[col] is not None else "").strip()


def _detect_columns(headers: Sequence) -> Dict[str, int]:
    aliases = {
        "type": {"type", "transaction type", "trx type"},
        "code": {"code", "transaction code", "trx code"},
        "number": {"number", "reference", "reference number", "transaction number"},
        "transaction_date": {"transaction date", "trx date", "date"},
        "value_date": {"value date"},
        "amount": {"amount", "transaction amount"},
    }

    found: Dict[str, int] = {}
    normalized_headers = [_normalize_header(h) for h in headers]
    for idx, norm in enumerate(normalized_headers):
        if not norm:
            continue
        for key, names in aliases.items():
            if key in found:
                continue
            if norm in names:
                found[key] = idx
                continue
            # Contains-based fallback for slightly different header names.
            if key == "transaction_date" and "transaction" in norm and "date" in norm:
                found[key] = idx
            elif key == "value_date" and "value" in norm and "date" in norm:
                found[key] = idx
            elif key == "amount" and "amount" in norm:
                found[key] = idx
            elif key == "type" and norm.endswith(" type"):
                found[key] = idx
            elif key == "code" and norm.endswith(" code"):
                found[key] = idx
    return found


def validate_ifmis_data(
    grid_data: Sequence[Sequence],
    *,
    has_header_row: bool,
    start_row: int = 0,
    end_row: Optional[int] = None,
    selected_columns: Optional[Set[int]] = None,
) -> List[ValidationIssue]:
    """
    Validate IFMIS-like row data.
    Returns a list of row/col issues with severity and message.
    """
    if not grid_data:
        return []

    total_rows = len(grid_data)
    first_data_row = 1 if has_header_row else 0
    start = max(start_row, first_data_row)
    end = total_rows - 1 if end_row is None else min(end_row, total_rows - 1)
    if end < start:
        return []

    headers = grid_data[0] if has_header_row else []
    col_map = _detect_columns(headers) if has_header_row else {}
    issues: List[ValidationIssue] = []

    required_keys = ("type", "code", "transaction_date", "value_date", "amount")

    for row_idx in range(start, end + 1):
        row = grid_data[row_idx] if row_idx < total_rows else []
        if not row:
            continue

        # Skip completely empty rows.
        has_any_data = any(not _is_empty(v) for v in row)
        if not has_any_data:
            continue

        # Required fields.
        for key in required_keys:
            col_idx = col_map.get(key)
            if col_idx is None:
                continue
            if selected_columns is not None and col_idx not in selected_columns:
                continue
            if _is_empty(_safe_cell(row, col_idx)):
                issues.append(
                    ValidationIssue(
                        severity="error",
                        row=row_idx,
                        col=col_idx,
                        message=f"Missing required value: {key.replace('_', ' ')}",
                    )
                )

        # Type checks.
        type_col = col_map.get("type")
        if type_col is not None and (selected_columns is None or type_col in selected_columns):
            type_val = _safe_cell(row, type_col).lower()
            if type_val and type_val not in VALID_TYPES:
                issues.append(
                    ValidationIssue(
                        severity="warning",
                        row=row_idx,
                        col=type_col,
                        message="Type is uncommon (expected Payment or Receipt).",
                    )
                )

        # Date checks.
        for date_key in ("transaction_date", "value_date"):
            col_idx = col_map.get(date_key)
            if col_idx is None:
                continue
            if selected_columns is not None and col_idx not in selected_columns:
                continue
            date_text = _safe_cell(row, col_idx)
            if date_text and not DATE_RE.fullmatch(date_text):
                issues.append(
                    ValidationIssue(
                        severity="error",
                        row=row_idx,
                        col=col_idx,
                        message="Invalid date format (expected DD-MMM-YYYY).",
                    )
                )

        # Amount checks.
        amount_col = col_map.get("amount")
        if amount_col is not None and (selected_columns is None or amount_col in selected_columns):
            amount_text = _safe_cell(row, amount_col)
            if amount_text:
                normalized = amount_text.replace(",", "")
                try:
                    Decimal(normalized)
                except InvalidOperation:
                    issues.append(
                        ValidationIssue(
                            severity="error",
                            row=row_idx,
                            col=amount_col,
                            message="Amount is not numeric.",
                        )
                    )

    return issues
