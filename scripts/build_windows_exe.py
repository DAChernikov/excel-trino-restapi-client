from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path

DEFAULT_ICON = Path("assets") / "app_icon.ico"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Build a Windows one-file executable for trino-excel-client using PyInstaller."
    )
    parser.add_argument(
        "--name",
        default="trino-excel-client",
        help='Base executable name without suffix. Default: "trino-excel-client".',
    )
    parser.add_argument(
        "--dist-dir",
        type=Path,
        default=Path("dist"),
        help="PyInstaller dist directory.",
    )
    parser.add_argument(
        "--work-dir",
        type=Path,
        default=Path("build") / "pyinstaller",
        help="PyInstaller work directory.",
    )
    parser.add_argument(
        "--clean",
        action="store_true",
        help="Ask PyInstaller to clean temporary files before building.",
    )
    parser.add_argument(
        "--target",
        choices=("all", "cli", "gui"),
        default="all",
        help="Which executable to build. Default: all.",
    )
    parser.add_argument(
        "--icon",
        type=Path,
        default=DEFAULT_ICON,
        help="Path to .ico file used for Windows executables.",
    )
    return parser


def _run_pyinstaller(args, *, name: str, entrypoint: str, windowed: bool) -> Path:
    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--onefile",
        "--name",
        name,
        "--distpath",
        str(args.dist_dir),
        "--workpath",
        str(args.work_dir / name),
        "--paths",
        "src",
        entrypoint,
    ]
    if args.icon.exists():
        command[-1:-1] = ["--icon", str(args.icon)]
        command[-1:-1] = ["--add-data", f"{args.icon}{os.pathsep}assets"]
    if windowed:
        command.insert(4, "--windowed")
    if args.clean:
        command.insert(3, "--clean")

    print(">", " ".join(command))
    subprocess.run(command, check=True)
    output = args.dist_dir / (name + ".exe")
    print(f"Built executable: {output}")
    return output


def main() -> None:
    args = build_parser().parse_args()
    built: list[Path] = []
    if args.target in {"all", "cli"}:
        built.append(
            _run_pyinstaller(
                args,
                name=args.name + "-cmd",
                entrypoint="scripts/pyinstaller_cli_entry.py",
                windowed=False,
            )
        )
    if args.target in {"all", "gui"}:
        built.append(
            _run_pyinstaller(
                args,
                name=args.name + "-gui",
                entrypoint="scripts/pyinstaller_gui_entry.py",
                windowed=True,
            )
        )

    print("Built artifacts:")
    for artifact in built:
        print(f"  {artifact}")


if __name__ == "__main__":
    main()
