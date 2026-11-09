from __future__ import annotations

import importlib.util
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
APP_NAME = "Comments_from_reports"


def main() -> int:
    if importlib.util.find_spec("PyInstaller") is None:
        print(
            "PyInstaller is not installed. Install build dependencies with:\n"
            f'  "{sys.executable}" -m pip install -r "{ROOT / "requirements-build.txt"}"',
            file=sys.stderr,
        )
        return 1

    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--name",
        APP_NAME,
        "--distpath",
        str(ROOT / "dist"),
        "--workpath",
        str(ROOT / "build" / "work"),
        "--specpath",
        str(ROOT / "build" / "spec"),
        str(ROOT / "main.py"),
    ]
    try:
        result = subprocess.run(command, cwd=ROOT, check=False)
    except OSError as error:
        print(f"Could not start PyInstaller: {error}", file=sys.stderr)
        return 1
    return result.returncode


if __name__ == "__main__":
    raise SystemExit(main())
