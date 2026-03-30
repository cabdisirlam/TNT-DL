import ctypes
import os
import shutil
import subprocess
import sys
import tempfile
import traceback
from pathlib import Path


APP_NAME = "NT_DL"
TITLE = "NT_DL Setup"
PAYLOAD_FILES = ("NT_DL_payload.dat", "NT_DL_app.exe", "install.cmd", "uninstall.cmd", "kdl_a.ico")
BOOT_LOG = Path(os.environ.get("TEMP", str(Path.home()))) / "NT_DL_setup_bootstrap.log"


def _log(message: str) -> None:
    try:
        with BOOT_LOG.open("a", encoding="utf-8") as fh:
            fh.write(message + "\n")
    except Exception:
        pass


def _message(text: str, is_error: bool = False) -> None:
    if os.environ.get("NT_DL_SETUP_SILENT", "").strip() == "1":
        return
    icon = 0x10 if is_error else 0x40
    ctypes.windll.user32.MessageBoxW(0, text, TITLE, icon)


def _resource_dir() -> Path:
    if hasattr(sys, "_MEIPASS"):
        return Path(getattr(sys, "_MEIPASS"))
    return Path(__file__).resolve().parent


def _run() -> int:
    _log("=== setup bootstrap start ===")
    src_dir = _resource_dir()
    _log(f"resource_dir={src_dir}")
    temp_dir = Path(tempfile.mkdtemp(prefix="nt_dl_setup_"))
    _log(f"work_temp_dir={temp_dir}")
    try:
        for name in PAYLOAD_FILES:
            src = src_dir / name
            if not src.exists():
                _log(f"missing payload file: {src}")
                _message(f"Installer payload missing: {name}", is_error=True)
                return 2
            dest_name = "NT_DL.exe" if name == "NT_DL_payload.dat" else name
            shutil.copy2(src, temp_dir / dest_name)
            _log(f"copied {name} -> {dest_name}")

        result = subprocess.run(
            ["cmd", "/c", "install.cmd", "/quiet"],
            cwd=str(temp_dir),
            check=False,
        )
        _log(f"install.cmd exit={result.returncode}")
        if result.returncode != 0:
            log_file = Path(os.environ.get("TEMP", str(Path.home()))) / "NT_DL_install.log"
            _log(f"install failed; expected log={log_file}")
            _message(
                f"{APP_NAME} installation failed (code {result.returncode}).\n"
                f"Check log: {log_file}",
                is_error=True,
            )
            return result.returncode

        install_exe = Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / APP_NAME / f"{APP_NAME}.exe"
        _log(f"installed_exe={install_exe} exists={install_exe.exists()}")
        if install_exe.exists():
            subprocess.Popen([str(install_exe)], cwd=str(install_exe.parent))
            _log("launched installed app")

        _message(f"{APP_NAME} installed successfully.")
        _log("setup bootstrap success")
        return 0
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
        _log("temp cleanup done")


if __name__ == "__main__":
    try:
        raise SystemExit(_run())
    except SystemExit:
        raise
    except Exception:
        _log("unhandled exception:")
        _log(traceback.format_exc())
        _message("NT_DL setup failed unexpectedly. Check NT_DL_setup_bootstrap.log in %TEMP%.", is_error=True)
        raise SystemExit(99)
