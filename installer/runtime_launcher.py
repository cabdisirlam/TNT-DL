import os
import subprocess
import sys
import tempfile
from pathlib import Path


APP_EXE_NAME = "NT_DL_app.exe"
RUNTIME_TEMP_DIR_NAME = "N"


def _base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _runtime_temp_dir() -> Path:
    local_appdata = os.environ.get("LOCALAPPDATA")
    if local_appdata:
        return Path(local_appdata) / RUNTIME_TEMP_DIR_NAME
    return Path(tempfile.gettempdir()) / RUNTIME_TEMP_DIR_NAME


def _show_error(message: str) -> None:
    try:
        import ctypes

        ctypes.windll.user32.MessageBoxW(None, message, "NT_DL", 0x10)
    except Exception:
        pass


def main() -> int:
    base_dir = _base_dir()
    app_exe = base_dir / APP_EXE_NAME
    if not app_exe.exists():
        _show_error(f"Missing {APP_EXE_NAME} next to NT_DL.exe.")
        return 1

    runtime_temp_dir = _runtime_temp_dir()
    try:
        runtime_temp_dir.mkdir(parents=True, exist_ok=True)
    except Exception as exc:
        _show_error(f"Failed to prepare runtime temp folder:\n{runtime_temp_dir}\n\n{exc}")
        return 1

    env = os.environ.copy()
    env["TEMP"] = str(runtime_temp_dir)
    env["TMP"] = str(runtime_temp_dir)

    try:
        subprocess.Popen(
            [str(app_exe), *sys.argv[1:]],
            cwd=str(base_dir),
            env=env,
        )
        return 0
    except Exception as exc:
        _show_error(f"Failed to launch {APP_EXE_NAME}.\n\n{exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
