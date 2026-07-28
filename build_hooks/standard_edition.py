"""PyInstaller runtime hook that selects the general-user edition."""

import os

os.environ["NT_DL_EDITION"] = "standard"
