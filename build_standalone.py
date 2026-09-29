#!/usr/bin/env python3
"""Build a single-file executable for the current operating system."""

from __future__ import annotations

import importlib.util
import os
import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent
PROGRAM_NAME = "excel_sheet_merger"


def main() -> None:
    if importlib.util.find_spec("PyInstaller") is None:
        raise SystemExit(
            "缺少构建依赖，请先运行：\n"
            "python -m pip install -r requirements-build.txt"
        )

    build_root = PROJECT_ROOT / "build"
    build_root.mkdir(exist_ok=True)
    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--console",
        "--collect-all",
        "openpyxl",
        "--collect-all",
        "xlrd",
        "--name",
        PROGRAM_NAME,
        "--distpath",
        str(PROJECT_ROOT / "dist"),
        "--workpath",
        str(build_root / "pyinstaller"),
        "--specpath",
        str(build_root),
        str(PROJECT_ROOT / "merge_workbooks.py"),
    ]
    subprocess.run(command, cwd=PROJECT_ROOT, check=True)

    suffix = ".exe" if os.name == "nt" else ""
    executable = PROJECT_ROOT / "dist" / f"{PROGRAM_NAME}{suffix}"
    print(f"构建完成：{executable}")


if __name__ == "__main__":
    main()
