from trino_excel_client.m_code import M_QUERIES
from trino_excel_client.template import REQUIRED_QUERY_NAMES, client_sheet_names
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
    assert "Name: \"gui\"" in iss
    assert "Name: \"cli\"" in iss
    assert "install\") / \"setup.exe\"" in build_script
