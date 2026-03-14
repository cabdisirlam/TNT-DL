r"""
KDL Keystroke Parser
Parses cell content to determine if it's data, a keystroke command, 
a shortcut, a mouse click, or a delay instruction.

Conventions (matching FormsDataLoader):
  - Data: Plain text -> copied to clipboard and pasted into target
  - Keystroke: Starts with \ -> parsed into key sequences
  - Command/Shortcut: Starts with * -> resolved via shortcut map
  - Mouse click: *MC(x,y) or *MR(x,y)
  - Delay: Cell in delay column -> milliseconds to wait
"""

import re
from enum import Enum, auto
from dataclasses import dataclass, field
from typing import List, Optional, Tuple


class CellType(Enum):
    """Type of cell content."""
    DATA = auto()          # Plain data to paste
    KEYSTROKE = auto()     # Keystroke sequence (starts with \)
    SHORTCUT = auto()      # Shortcut command (starts with *)
    MOUSE_LEFT = auto()    # Left mouse click *MC(x,y)
    MOUSE_RIGHT = auto()   # Right mouse click *MR(x,y)
    DELAY = auto()         # Delay in milliseconds
    EMPTY = auto()         # Empty cell


@dataclass
class ParsedCell:
    """Result of parsing a cell."""
    cell_type: CellType
    raw_value: str = ""
    # For DATA type
    data_text: str = ""
    # For KEYSTROKE type - list of key actions
    key_actions: List[dict] = field(default_factory=list)
    # For MOUSE clicks
    mouse_x: int = 0
    mouse_y: int = 0
    # For DELAY
    delay_ms: int = 0


# Mapping of special key names to pyautogui key names
SPECIAL_KEYS = {
    "BACKSPACE": "backspace", "BS": "backspace", "BKSP": "backspace",
    "BREAK": "pause",
    "CAPSLOCK": "capslock",
    "DELETE": "delete", "DEL": "delete",
    "DOWN": "down",
    "END": "end",
    "ENTER": "enter",
    "ESC": "escape",
    "HELP": "help",
    "HOME": "home",
    "INSERT": "insert", "INS": "insert",
    "LEFT": "left",
    "NUMLOCK": "numlock",
    "PGDN": "pagedown",
    "PGUP": "pageup",
    "PRTSC": "printscreen",
    "RIGHT": "right",
    "SCROLLLOCK": "scrolllock",
    "TAB": "tab",
    "UP": "up",
    "F1": "f1", "F2": "f2", "F3": "f3", "F4": "f4",
    "F5": "f5", "F6": "f6", "F7": "f7", "F8": "f8",
    "F9": "f9", "F10": "f10", "F11": "f11", "F12": "f12",
    "F13": "f13", "F14": "f14", "F15": "f15", "F16": "f16",
}

# Modifier key mapping
MODIFIER_MAP = {
    "+": "shift",
    "^": "ctrl",
    "%": "alt",
}

# Default IFMIS shortcuts
DEFAULT_SHORTCUTS = {
    "*SP": "\\%f%v",        # Save & Proceed
    "*SV": "\\^s",          # Save (Ctrl+S)
    "*S": "\\^s",           # Save shorthand
    "*NR": "\\%a",          # New Record (Alt+A)
    "*NX": "\\{DOWN}{HOME}",      # Next Record and jump to row start
    "*DN": "\\{DOWN}",            # Down one step (e.g. dropdown selection)
    "*UP": "\\{UP}",        # Up shorthand
    "*PV": "\\{UP}",        # Previous Record
    "*NB": "\\{TAB}",       # Next Block / Field
    "*CL": "\\{ESC}",       # Clear / Cancel
    "*EX": "\\%{F4}",       # Exit Form
    "*DL": "\\^{DELETE}",   # Delete Record
    "*QR": "\\{F11}",       # Enter Query
    "*EQ": "\\^{F11}",      # Execute Query
    "*CM": "\\^s",          # Commit (Save)
    "*DF": "\\{F6}",        # Duplicate Field
    "*DR": "\\{F5}",        # Duplicate Record
    "*LOV": "\\^l",         # List of Values
}

# Plain keyword shortcuts for easier typing in cells.
# Example: "tab" instead of "\{TAB}".
SIMPLE_KEYWORDS = {
    "TAB": "\\{TAB}",
    "ENTER": "\\{ENTER}",
    "DOWN": "\\{DOWN}",
    "DN": "\\{DOWN}",
    "UP": "\\{UP}",
    "LEFT": "\\{LEFT}",
    "RIGHT": "\\{RIGHT}",
}


class KeystrokeParser:
    """Parses cell content into actionable commands."""

    def __init__(self):
        self.shortcuts = dict(DEFAULT_SHORTCUTS)

    def parse_cell(self, value: str, is_delay_column: bool = False) -> ParsedCell:
        """Parse a single cell value and return a ParsedCell."""
        if value is None or str(value).strip() == "":
            return ParsedCell(cell_type=CellType.EMPTY)

        value = str(value).strip()

        # Force literal data with leading apostrophe, e.g. '*dn or '\{TAB}
        if value.startswith("'"):
            literal = value[1:]
            if literal.strip() == "":
                return ParsedCell(cell_type=CellType.EMPTY, raw_value=value)
            return ParsedCell(
                cell_type=CellType.DATA,
                raw_value=value,
                data_text=literal,
            )

        # If this is a delay column, treat numeric values as delays
        if is_delay_column:
            try:
                delay_ms = int(float(value))
                return ParsedCell(
                    cell_type=CellType.DELAY,
                    raw_value=value,
                    delay_ms=delay_ms,
                )
            except ValueError:
                pass  # Not a number, parse normally

        # Check for mouse click commands
        mc_match = re.match(r'^\*MC\((\d+),(\d+)\)$', value, re.IGNORECASE)
        if mc_match:
            return ParsedCell(
                cell_type=CellType.MOUSE_LEFT,
                raw_value=value,
                mouse_x=int(mc_match.group(1)),
                mouse_y=int(mc_match.group(2)),
            )

        mr_match = re.match(r'^\*MR\((\d+),(\d+)\)$', value, re.IGNORECASE)
        if mr_match:
            return ParsedCell(
                cell_type=CellType.MOUSE_RIGHT,
                raw_value=value,
                mouse_x=int(mr_match.group(1)),
                mouse_y=int(mr_match.group(2)),
            )

        # Check for shortcut commands (starts with *)
        if value.startswith("*"):
            shortcut_key = value.upper()
            if shortcut_key in self.shortcuts:
                resolved = self.shortcuts[shortcut_key]
                return self._parse_keystroke(resolved, value)
            else:
                return ParsedCell(
                    cell_type=CellType.DATA,
                    raw_value=value,
                    data_text=value,
                )

        # Accept escaped shortcuts such as \*s (common in some loaders)
        if value.startswith("\\*"):
            shortcut_key = value[1:].upper()
            if shortcut_key in self.shortcuts:
                resolved = self.shortcuts[shortcut_key]
                return self._parse_keystroke(resolved, value)

        # Check for keystroke (starts with \)
        if value.startswith("\\"):
            return self._parse_keystroke(value, value)

        # Simple plain-text keywords like "tab", "enter", "down"
        keyword = value.upper()
        if keyword in SIMPLE_KEYWORDS:
            return self._parse_keystroke(SIMPLE_KEYWORDS[keyword], value)

        # Plain data
        return ParsedCell(
            cell_type=CellType.DATA,
            raw_value=value,
            data_text=value,
        )

    def _parse_keystroke(self, keystroke_str: str, raw_value: str) -> ParsedCell:
        """Parse a keystroke string like \\{TAB} or \\%f into key actions."""
        result = ParsedCell(
            cell_type=CellType.KEYSTROKE,
            raw_value=raw_value,
        )

        # Remove leading backslash
        seq = keystroke_str[1:] if keystroke_str.startswith("\\") else keystroke_str

        i = 0
        while i < len(seq):
            char = seq[i]

            # Modifier keys: +, ^, %
            if char in MODIFIER_MAP:
                modifier = MODIFIER_MAP[char]
                i += 1
                if i < len(seq):
                    # Check if next is a group (parentheses)
                    if seq[i] == "(":
                        # Find closing paren
                        end = seq.find(")", i)
                        if end != -1:
                            group = seq[i + 1:end]
                            for ch in group:
                                result.key_actions.append({
                                    "type": "hotkey",
                                    "modifiers": [modifier],
                                    "key": ch.lower(),
                                })
                            i = end + 1
                        else:
                            # Unmatched ( — treat rest as data
                            result.key_actions.append({
                                "type": "type",
                                "text": seq[i:],
                            })
                            i = len(seq)
                    elif seq[i] == "{":
                        # Modifier + special key like %{F4}
                        end = seq.find("}", i)
                        if end != -1:
                            key_spec = seq[i + 1:end]
                            key_name, count = self._parse_key_spec(key_spec)
                            for _ in range(count):
                                result.key_actions.append({
                                    "type": "hotkey",
                                    "modifiers": [modifier],
                                    "key": key_name,
                                })
                            i = end + 1
                        else:
                            # Unmatched { — treat rest as data
                            result.key_actions.append({
                                "type": "type",
                                "text": seq[i:],
                            })
                            i = len(seq)
                    else:
                        # Modifier + single character
                        result.key_actions.append({
                            "type": "hotkey",
                            "modifiers": [modifier],
                            "key": seq[i].lower(),
                        })
                        i += 1

            # Special key in braces: {TAB}, {ENTER}, {LEFT 5}
            elif char == "{":
                end = seq.find("}", i)
                if end != -1:
                    key_spec = seq[i + 1:end]
                    key_name, count = self._parse_key_spec(key_spec)
                    for _ in range(count):
                        result.key_actions.append({
                            "type": "key",
                            "key": key_name,
                        })
                    i = end + 1
                else:
                    # Malformed - treat rest as data
                    result.key_actions.append({
                        "type": "type",
                        "text": seq[i:],
                    })
                    break

            # Tilde ~ means ENTER
            elif char == "~":
                result.key_actions.append({
                    "type": "key",
                    "key": "enter",
                })
                i += 1

            # Regular character - type it
            else:
                result.key_actions.append({
                    "type": "type",
                    "text": char,
                })
                i += 1

        return result

    def _parse_key_spec(self, spec: str) -> Tuple[str, int]:
        """Parse a key spec like 'TAB' or 'LEFT 5' into (key_name, count)."""
        parts = spec.strip().split(" ", 1)
        key_str = parts[0].upper()
        count = 1
        if len(parts) > 1:
            try:
                count = int(parts[1])
            except ValueError:
                count = 1

        key_name = SPECIAL_KEYS.get(key_str, key_str.lower())
        return key_name, count

    def update_shortcut(self, shortcut: str, keystroke: str):
        """Add or update a shortcut mapping."""
        self.shortcuts[shortcut.upper()] = keystroke

    def remove_shortcut(self, shortcut: str):
        """Remove a shortcut."""
        self.shortcuts.pop(shortcut.upper(), None)

    def get_shortcuts(self) -> dict:
        """Return a copy of all shortcuts."""
        return dict(self.shortcuts)
