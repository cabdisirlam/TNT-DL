import os
import subprocess
import sys
from pathlib import Path


LOG_FILE = Path(os.environ.get("TEMP", str(Path.home()))) / "NT_DL_iexpress_launcher.log"
APP_EXE = Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "NT_DL" / "NT_DL.exe"


def _log(message: str) -> None:
    try:
        with LOG_FILE.open("a", encoding="utf-8") as fh:
            fh.write(message + "\n")
    except Exception:
        pass


def _work_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def main() -> int:
    work_dir = _work_dir()
    _log("=== launcher start ===")
    _log(f"work_dir={work_dir}")

    cmd = ["cmd", "/c", "install.cmd", "/quiet"]
    _log(f"running: {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=str(work_dir), check=False)
    _log(f"install_exit={result.returncode}")

    if result.returncode == 0 and APP_EXE.exists():
        try:
            subprocess.Popen([str(APP_EXE)], cwd=str(APP_EXE.parent))
            _log("launched installed app")
        except Exception as exc:
            _log(f"launch_failed={exc!r}")

    _log("=== launcher end ===")
    return int(result.returncode)


if __name__ == "__main__":
    raise SystemExit(main())
