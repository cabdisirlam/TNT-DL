"""
Imprest Old Date AP Loader — Engine

Identical to imprest_surrender_engine except it sends Enter after the first
invoice date field to dismiss Oracle IFMIS's "date in prior period" dialog box
before moving to the next field.  Use this mode when loading invoices with dates
that fall in a prior accounting period.
"""

# Re-export the full public API from the base engine so dialogs and loader can
# import from either engine interchangeably.
from kdl.engine.imprest_surrender_engine import (  # noqa: F401
    COLUMNS,
    COLUMN_HINTS,
    COLUMN_SAMPLE,
    IFMIS_BLANK_COLS,
    ImprestSurrenderThread,
    _INTER_ACTION_DELAY,
    _fmt_ifmis_date,
    _normalize_invoice_row,
    build_row_summary,
    export_prefilled_template,
    export_template,
    import_ifmis_export,
    read_invoice_rows,
)
from kdl.engine.imprest_surrender_engine import (
    _build_old_date_dl_keystroke_row as _base_build_old_date_dl_keystroke_row,
    execute_row_for_loader as _base_execute_row_for_loader,
    export_keystroke_sheet_to_workbook as _combined_export_keystroke_sheet_to_workbook,
)

# ── Modified action sequence ──────────────────────────────────────────────────
# Identical to TEMPLATE_ACTIONS in imprest_surrender_engine except that
# ("key", "enter") is inserted after the first invoice date field so that
# Oracle IFMIS's prior-period date confirmation dialog is dismissed automatically.

TEMPLATE_ACTIONS = (
    ("tab", 1),
    ("key", "backspace"),
    ("tab", 2),
    ("field", "Supplier_Num"),
    ("tab", 1),
    ("key", "enter"),
    ("text", "provisional"),
    ("tab", 1),
    ("field", "Invoice_Date"),
    ("tab", 1),
    ("key", "enter"),               # dismiss Oracle prior-period date dialog
    ("field", "Invoice_Num"),
    ("tab", 2),
    ("field", "Application_Amount"),
    ("tab", 7),
    ("field", "Description"),
    ("tab", 5),
    ("text", "CHECK"),
    ("tab", 17),
    ("field", "Auth_Ref_No"),
    ("tab", 1),
    ("field", "Administrative_Code"),
    ("tab", 2),
    ("key", "enter"),
    ("hotkey", ["ctrl"], "s"),
    ("delay", 300),
    ("hotkey", ["alt"], "2"),            # Alt+2 then Esc → Lines block
    ("key", "esc"),
    ("delay", 500),
    ("tab", 2),
    ("field", "Application_Amount"),
    ("tab", 1),
    ("hotkey", ["alt"], "d"),            # Alt+D → Distributions block
    ("tab", 2),
    ("field", "Application_Amount"),
    ("tab", 1),
    ("field", "GL_Date"),
    ("tab", 1),
    ("field", "Distribution_Account"),
    ("tab", 1),
    ("hotkey", ["ctrl"], "s"),
    ("delay", 500),
    ("hotkey", ["ctrl"], "f4"),
    ("delay", 500),
    ("hotkey", ["alt"], "c"),
    ("hotkey", ["alt"], "u"),
    ("hotkey", ["alt"], "k"),
    ("hotkey", ["alt"], "v"),
    ("key", "down"),
    ("key", "down"),
    ("key", "enter"),
    ("delay", 500),
    ("field", "Old_Imprest_No"),
    ("tab", 7),
    ("field", "Application_Amount"),
    ("tab", 2),
    ("key", "enter"),
    ("key", "space"),
    ("tab", 1),
    ("field", "Application_Amount"),
    ("tab", 1),
    ("field", "GL_Date"),
    ("tab", 1),
    ("hotkey", ["ctrl"], "s"),
    ("delay", 500),
    ("hotkey", ["ctrl"], "f4"),
    ("delay", 500),
    ("key", "alt"),
    ("delay", 250),
    ("key", "down"),
    ("key", "down"),
    ("key", "down"),
    ("key", "down"),
    ("key", "enter"),
    ("key", "down"),
)


def _build_dl_keystroke_row(row: dict) -> list:
    """Build old-date DataLoad row with Enter after the invoice date field."""
    return _base_build_old_date_dl_keystroke_row(row)


def export_keystroke_sheet_to_workbook(source_path: str, save_path: str, rows: list) -> str:
    """Export both normal imprest and old-date imprest keystrokes on one sheet."""
    return _combined_export_keystroke_sheet_to_workbook(source_path, save_path, rows)


def execute_row_for_loader(sender, row_dict: dict, is_stop_requested,
                           actions=None, popup_fn=None,
                           inter_action_delay=None, is_last_row=False) -> bool:
    """Thin wrapper — uses the old-date TEMPLATE_ACTIONS by default."""
    if actions is None:
        actions = TEMPLATE_ACTIONS
    return _base_execute_row_for_loader(
        sender, row_dict, is_stop_requested,
        actions=actions,
        popup_fn=popup_fn,
        inter_action_delay=inter_action_delay,
        is_last_row=is_last_row,
    )
