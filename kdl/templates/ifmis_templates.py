"""
NT_DL IFMIS Templates - Transaction Lines
"""


def get_bank_statement_lines_template() -> dict:
    return {
        "name": "Transaction Lines (IFMIS)",
        "description": "Oracle IFMIS transaction lines. Use Cell Mode.",
        "headers": [],  # No column headers
        "key_columns": [0, 1, 2, 3, 5, 7, 9, 11, 13, 14, 15],
        "sample_data": [
            # PAYMENT: tab, tab, trfd, tab, 20592, tab, 02-Feb-2026, tab, 02-Feb-2026, tab, 1301800, tab, *S, *DN
            ["tab",     "",    "",  "tab",     "TRFD", "tab",     "20592", "tab",     "02-FEB-2026", "tab",     "02-FEB-2026", "tab",     "1301800", "tab",     "*S", "*DN"],
            # RECEIPT: tab, *DN, \r, tab, trfc, tab, 101, tab, 29-Jan-2026, tab, 29-Jan-2026, tab, 16300, tab, *S, *DN
            ["tab",     "*DN", "\\r", "tab",   "TRFC", "tab",     "101",   "tab",     "29-JAN-2026", "tab",     "29-JAN-2026", "tab",     "16300",   "tab",     "*S", "*DN"],
        ],
    }


def get_bank_statement_template() -> dict:
    return get_bank_statement_lines_template()


def get_all_templates() -> list:
    return [get_bank_statement_lines_template()]


def get_template_names() -> list:
    return [t["name"] for t in get_all_templates()]

