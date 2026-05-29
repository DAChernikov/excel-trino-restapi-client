from __future__ import annotations

import argparse
import platform
from pathlib import Path

from .excel_com import create_workbook as create_workbook_com
from .excel_com import install_client as install_client_com
from .openpyxl_backend import create_workbook_ui, install_client_ui_file
from .validation import validate_project_contract


BACKENDS = ("auto", "com", "openpyxl")


def _auto_backend() -> str:
    return "com" if platform.system() == "Windows" else "openpyxl"


def _effective_backend(backend: str) -> str:
    return _auto_backend() if backend == "auto" else backend


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="trino-excel-client",
        description="Generate or install an Excel Power Query Trino REST API client.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    create_parser = subparsers.add_parser(
        "create",
        help="Create a new Excel workbook with embedded Trino REST client.",
    )
    create_parser.add_argument(
        "--output",
        required=True,
        type=Path,
        help="Path to output .xlsx or .xlsm file.",
    )
    create_parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite output file if it already exists.",
    )
    create_parser.add_argument(
        "--visible",
        action="store_true",
        help="Show Excel during generation. Only used by the COM backend.",
    )
    create_parser.add_argument(
        "--backend",
        choices=BACKENDS,
        default="auto",
        help="Generation backend. auto uses COM on Windows and openpyxl elsewhere.",
    )
    create_parser.add_argument(
        "--sheet-prefix",
        default="Trino",
        help='Prefix for client sheets. Default: "Trino".',
    )
    create_parser.add_argument(
        "--advanced",
        action="store_true",
        help="Create an advanced .xlsm client with a VBA button for exporting the current SQL to a named result sheet. COM backend only.",
    )

    install_parser = subparsers.add_parser(
        "install",
        help="Install the Trino REST client into an existing workbook.",
    )
    install_parser.add_argument(
        "--workbook",
        required=True,
        type=Path,
        help="Existing .xlsx or .xlsm workbook to modify.",
    )
    install_parser.add_argument(
        "--output",
        type=Path,
        help="Optional output path. If omitted, the workbook is updated in place.",
    )
    install_parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite --output file if it already exists.",
    )
    install_parser.add_argument(
        "--visible",
        action="store_true",
        help="Show Excel during installation. Only used by the COM backend.",
    )
    install_parser.add_argument(
        "--backend",
        choices=BACKENDS,
        default="auto",
        help="Installation backend. auto uses COM on Windows and openpyxl elsewhere.",
    )
    install_parser.add_argument(
        "--sheet-prefix",
        default="Trino",
        help='Prefix for client sheets. Default: "Trino".',
    )
    install_parser.add_argument(
        "--advanced",
        action="store_true",
        help="Install the advanced .xlsm client with a VBA button for exporting the current SQL to a named result sheet. COM backend only.",
    )

    subparsers.add_parser(
        "validate",
        help="Validate the cross-platform project contract without Excel.",
    )
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    if args.command == "create":
        backend = _effective_backend(args.backend)
        if args.advanced and backend != "com":
            parser.error("--advanced requires --backend com and Windows desktop Excel")
        if backend == "com":
            output_path = create_workbook_com(
                output_path=args.output,
                overwrite=args.overwrite,
                visible=args.visible,
                sheet_prefix=args.sheet_prefix,
                advanced=args.advanced,
            )
            if args.advanced:
                print(f"Created advanced macro-enabled workbook with Excel COM, Power Query, and VBA: {output_path}")
            else:
                print(f"Created workbook with Excel COM and Power Query: {output_path}")
        else:
            output_path = create_workbook_ui(
                output_path=args.output,
                overwrite=args.overwrite,
                sheet_prefix=args.sheet_prefix,
            )
            print(f"Created cross-platform UI workbook without embedded Power Query: {output_path}")
        return

    if args.command == "install":
        backend = _effective_backend(args.backend)
        if args.advanced and backend != "com":
            parser.error("--advanced requires --backend com and Windows desktop Excel")
        if backend == "com":
            output_path = install_client_com(
                workbook_path=args.workbook,
                output_path=args.output,
                overwrite=args.overwrite,
                visible=args.visible,
                sheet_prefix=args.sheet_prefix,
                advanced=args.advanced,
            )
            if args.advanced:
                print(f"Installed advanced Trino client with Excel COM, Power Query, and VBA: {output_path}")
            else:
                print(f"Installed Trino client with Excel COM and Power Query: {output_path}")
        else:
            output_path = install_client_ui_file(
                workbook_path=args.workbook,
                output_path=args.output,
                overwrite=args.overwrite,
                sheet_prefix=args.sheet_prefix,
            )
            print(f"Installed cross-platform UI sheets without embedded Power Query: {output_path}")
        return

    if args.command == "validate":
        errors = validate_project_contract()
        if errors:
            for error in errors:
                print(f"ERROR: {error}")
            raise SystemExit(1)
        print("Project contract is valid.")
        return

    parser.error(f"Unknown command: {args.command}")


if __name__ == "__main__":
    main()
