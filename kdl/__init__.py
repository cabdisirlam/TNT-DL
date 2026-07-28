"""NT DL Multipurpose Tool."""

import os

__version__ = "1.4.24"
__edition__ = os.environ.get("NT_DL_EDITION", "full").strip().lower()
IMPREST_ENABLED = __edition__ != "standard"
__app_name__ = "NT_DL_Full" if IMPREST_ENABLED else "NT_DL_Standard"
__display_name__ = "NT DL Full" if IMPREST_ENABLED else "NT DL Standard"
