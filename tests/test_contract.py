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


def test_refresh_policy_is_not_user_configurable() -> None:
    config_parameters = {row[0] for row in CONFIG_TABLE.rows}
    assert "refresh_on_file_open" not in config_parameters
    assert "refresh_period" not in config_parameters
    assert "background_query" not in config_parameters
    assert "refresh_with_refresh_all" not in config_parameters


def test_no_application_packaging_layer_remains() -> None:
    from pathlib import Path

    removed_paths = [
        ".github/workflows/build-windows-installer.yml",
        "assets/app_icon.ico",
        "install/setup.exe",
        "installer/windows/trino-excel-client.iss",
        "scripts/build_windows_exe.py",
        "scripts/build_windows_installer.py",
        "src/trino_excel_client/gui_app.py",
    ]

    for path in removed_paths:
        assert not Path(path).exists(), path


def test_template_build_script_targets_templates_directory() -> None:
    script = _read_text("scripts/build_template.py")
    readme = _read_text("README.md")

    assert 'default=Path("templates") / "Trino_REST_Client.xlsx"' in script
    assert "create_workbook_com" in script
    assert "GitHub Actions для сборки шаблона намеренно не используется" in readme
