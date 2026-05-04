"""
Imprest Old Date AP Loader — Engine

Identical to imprest_surrender_engine but sends Enter after every date field
to dismiss Oracle IFMIS's "date in prior period" dialog box before moving to
the next field.  Use this mode when loading invoices with dates that fall in a
prior accounting period.
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
    export_keystroke_sheet_to_workbook,
    export_prefilled_template,
    export_template,
    import_ifmis_export,
    read_invoice_rows,
)
from kdl.engine.imprest_surrender_engine import (
    execute_row_for_loader as _base_execute_row_for_loader,
)

# ── Modified action sequence ──────────────────────────────────────────────────
# Identical to TEMPLATE_ACTIONS in imprest_surrender_engine except that
# ("key", "enter") is inserted after every date field so that Oracle IFMIS's
# "prior-period date" confirmation dialog is dismissed automatically.

TEMPLATE_ACTIONS = (
    ("tab", 2),
    ("key", "backspace"),
    ("tab", 1),
    ("text", "Standard"),
    ("tab", 1),
    ("key", "backspace"),
    ("tab", 1),
    ("key", "backspace"),
    ("tab", 1),
    ("field", "Supplier_Num"),
    ("tab", 1),
    ("key", "enter"),
    ("text", "Provisional"),
    ("tab", 1),
    ("field", "Invoice_Date"),
    ("tab", 1),
    ("key", "enter"),               # dismiss Oracle prior-period date dialog
    ("field", "Invoice_Num"),
    ("tab", 2),
    ("field", "Invoice_Amount"),
    ("tab", 7),
    ("field", "Description"),
    ("tab", 3),
    ("text", "IMMEDIATE"),
    ("tab", 1),
    ("text", "CHECK"),
    ("tab", 17),
    ("field", "Auth_Ref_No"),
    ("tab", 1),
    ("field", "Administrative_Code"),
    ("tab", 1),
    ("key", "enter"),
    ("hotkey", ["alt"], "2"),
    ("key", "escape"),
    ("delay", 500),
    ("tab", 2),
    ("field", "Application_Amount"),
    ("tab", 1),
    ("hotkey", ["alt"], "d"),
    ("tab", 2),
    ("field", "Application_Amount"),
    ("tab", 1),
    ("field", "GL_Date"),
    ("tab", 1),
    ("key", "enter"),               # dismiss Oracle prior-period date dialog
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
    ("key", "enter"),               # dismiss Oracle prior-period date dialog
    ("hotkey", ["ctrl"], "s"),
    ("delay", 500),
    ("hotkey", ["ctrl"], "f4"),
    ("delay", 700),
    ("key", "alt"),
    ("delay", 250),
    ("key", "down"),
    ("key", "down"),
    ("key", "down"),
    ("key", "down"),
    ("key", "enter"),
    ("key", "down"),
    ("delay", 350),
    ("hotkey", ["shift"], "tab"),
    ("hotkey", ["shift"], "tab"),
    ("hotkey", ["shift"], "tab"),
)


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
