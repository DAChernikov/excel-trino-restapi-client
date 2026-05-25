from __future__ import annotations

import argparse
import shutil
import subprocess
from pathlib import Path

EXPECTED_INSTALLER = Path("install") / "setup.exe"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Build the Windows installer using Inno Setup after PyInstaller artifacts are ready."
    )
    parser.add_argument(
        "--iscc",
        default="ISCC.exe",
        help="Path to Inno Setup compiler ISCC.exe.",
    )
    parser.add_argument(
        "--script",
        type=Path,
        default=Path("installer") / "windows" / "trino-excel-client.iss",
        help="Path to the Inno Setup .iss script.",
    )
    return parser


def main() -> None:
    args = build_parser().parse_args()
    required = [
        Path("dist") / "trino-excel-client-cmd.exe",
        Path("dist") / "trino-excel-client-gui.exe",
    ]
    missing = [path for path in required if not path.exists()]
    if missing:
        raise SystemExit(
            "Missing PyInstaller artifacts. Run first: python scripts\\build_windows_exe.py --clean\n"
            + "\n".join(str(path) for path in missing)
        )

    compiler = shutil.which(args.iscc) or args.iscc
    command = [compiler, str(args.script)]
    print(">", " ".join(command))
    subprocess.run(command, check=True)
    if not EXPECTED_INSTALLER.exists():
        raise SystemExit(f"Installer was not created: {EXPECTED_INSTALLER}")
    print(f"Built installer: {EXPECTED_INSTALLER}")


if __name__ == "__main__":
    main()
