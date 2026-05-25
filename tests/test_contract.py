from trino_excel_client.m_code import M_QUERIES
from trino_excel_client.template import CONFIG_TABLE, REQUIRED_QUERY_NAMES, client_sheet_names
from trino_excel_client.validation import validate_project_contract


def _read_text(path: str) -> str:
    from pathlib import Path

    return Path(path).read_text(encoding="utf-8")


def test_required_m_queries_exist() -> None:
    assert tuple(M_QUERIES) == REQUIRED_QUERY_NAMES


def test_project_contract_is_valid() -> None:
    assert validate_project_contract() == []


def test_m_code_contains_bigdata_row_limit_controls() -> None:
    function_formula = M_QUERIES["fnTrinoRestQuery"]
    result_formula = M_QUERIES["TrinoResult"]

    assert "MaxResultRows" in function_formula
    assert "ResultOverflowBehavior" in function_formula
    assert "DefaultQueryLimitRows" in function_formula
    assert "100000" in function_formula
    assert "ApplyDefaultQueryLimit" in function_formula
    assert "EffectiveSqlText" in function_formula
    assert "Trino result row limit exceeded" in function_formula
    assert "Table.TransformColumnTypes" in function_formula
    assert "TrinoTypeToPowerQueryType" in function_formula
    assert "max_result_rows" in result_formula
    assert "result_overflow_behavior" in result_formula
    assert "default_query_limit_rows" in result_formula
    assert "apply_default_query_limit" in result_formula
    assert "TrinoResultSchema" not in M_QUERIES


def test_default_source_name_is_official_excel_client_name() -> None:
    assert "Microsoft Excel PowerQuery client" in M_QUERIES["fnTrinoRestQuery"]


def test_sheet_names_are_excel_safe() -> None:
    names = client_sheet_names("Very Long Trino Client Prefix")
    assert len(names.config) <= 31
    assert len(names.query) <= 31
    assert len(names.help) <= 31


def test_windows_installer_outputs_repo_install_setup_exe() -> None:
    iss = _read_text("installer/windows/trino-excel-client.iss")
    build_script = _read_text("scripts/build_windows_installer.py")

    assert "OutputDir=..\\..\\install" in iss
    assert "OutputBaseFilename=setup" in iss
    assert 'SetupIconFile={#MyAppIcon}' in iss
    assert 'UninstallDisplayIcon={app}\\{#MyAppExeName}' in iss
    assert 'IconFilename: "{app}\\{#MyAppExeName}"' in iss
    assert "DisableDirPage=no" in iss
    assert "UsePreviousAppDir=no" not in iss
    assert "CloseApplications=yes" in iss
    assert "RestartApplications=no" in iss
    assert 'Name: "gui"; Description: "Только GUI"' in iss
    assert 'Name: "full"; Description: "GUI и CLI"' in iss
    assert 'Name: "cli"; Description: "Only CLI"' not in iss
    assert "Name: \"gui\"" in iss
    assert "Name: \"cli\"" in iss
    assert "Flags: fixed" in iss
    assert "Tasks: startmenuicon" in iss
    assert 'Name: "startmenuicon"' in iss
    assert 'Name: "desktopicon"' in iss
    assert "[UninstallRun]" in iss
    assert "taskkill /IM {#MyAppExeName}" in iss
    assert "taskkill /IM {#MyCmdExeName}" in iss
    assert "[UninstallDelete]" in iss
    assert 'Name: "{app}\\build"' in iss
    assert "install\") / \"setup.exe\"" in build_script


def test_refresh_policy_is_not_user_configurable() -> None:
    config_parameters = {row[0] for row in CONFIG_TABLE.rows}
    assert "refresh_on_file_open" not in config_parameters
    assert "refresh_period" not in config_parameters
    assert "background_query" not in config_parameters
    assert "refresh_with_refresh_all" not in config_parameters


def test_pyinstaller_uses_package_safe_entrypoints() -> None:
    build_script = _read_text("scripts/build_windows_exe.py")
    cli_entry = _read_text("scripts/pyinstaller_cli_entry.py")
    gui_entry = _read_text("scripts/pyinstaller_gui_entry.py")

    assert 'entrypoint="scripts/pyinstaller_cli_entry.py"' in build_script
    assert 'entrypoint="scripts/pyinstaller_gui_entry.py"' in build_script
    assert 'entrypoint="src/trino_excel_client/cli.py"' not in build_script
    assert 'entrypoint="src/trino_excel_client/gui_app.py"' not in build_script
    assert 'DEFAULT_ICON = Path("assets") / "app_icon.ico"' in build_script
    assert '["--icon", str(args.icon)]' in build_script
    assert "--add-data" in build_script
    assert "os.pathsep" in build_script
    assert "from trino_excel_client.cli import main" in cli_entry
    assert "from trino_excel_client.gui_app import main" in gui_entry


def test_gui_sets_runtime_window_icon() -> None:
    gui_app = _read_text("src/trino_excel_client/gui_app.py")

    assert "SetCurrentProcessExplicitAppUserModelID" in gui_app
    assert "_resource_path(\"assets/app_icon.ico\")" in gui_app
    assert "self.iconbitmap(default=str(icon_path))" in gui_app


def test_app_icon_asset_exists() -> None:
    from pathlib import Path

    icon_path = Path("assets") / "app_icon.ico"
    icon_script = _read_text("scripts/build_app_icon.py")

    assert icon_path.exists()
    assert icon_path.stat().st_size > 1024
    assert icon_path.read_bytes()[:4] == b"\x00\x00\x01\x00"
    assert "--source" in icon_script
    assert "assets/app_icon.ico" in icon_script


def test_gui_defaults_save_workbooks_to_desktop() -> None:
    gui_app = _read_text("src/trino_excel_client/gui_app.py")

    assert 'self.create_output = tk.StringVar(value=str(_desktop_path("Trino_REST_Client.xlsx")))' in gui_app
    assert 'self.install_output = tk.StringVar(value=str(_desktop_path("Workbook_With_Trino.xlsx")))' in gui_app
    assert 'Path("build") / "Trino_REST_Client.xlsx"' not in gui_app
    assert 'Path("build") / "Workbook_With_Trino.xlsx"' not in gui_app
