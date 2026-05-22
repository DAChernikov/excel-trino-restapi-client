from trino_excel_client.cli import build_parser


def test_parser_accepts_create_openpyxl_backend() -> None:
    parser = build_parser()
    args = parser.parse_args(["create", "--output", "client.xlsx", "--backend", "openpyxl"])
    assert args.command == "create"
    assert args.backend == "openpyxl"


def test_parser_accepts_install_com_backend() -> None:
    parser = build_parser()
    args = parser.parse_args(["install", "--workbook", "book.xlsx", "--backend", "com"])
    assert args.command == "install"
    assert args.backend == "com"


def test_parser_accepts_validate() -> None:
    parser = build_parser()
    args = parser.parse_args(["validate"])
    assert args.command == "validate"


def test_parser_accepts_gui_dry_run() -> None:
    parser = build_parser()
    args = parser.parse_args(["gui", "--dry-run"])
    assert args.command == "gui"
    assert args.dry_run is True
