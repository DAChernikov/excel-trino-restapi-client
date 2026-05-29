from trino_excel_client.m_code import M_QUERIES
from trino_excel_client.template import CONFIG_TABLE, REQUIRED_QUERY_NAMES, client_sheet_names
from trino_excel_client.validation import validate_project_contract


def test_required_m_queries_exist() -> None:
    assert tuple(M_QUERIES) == REQUIRED_QUERY_NAMES


def test_project_contract_is_valid() -> None:
    assert validate_project_contract() == []


def test_m_code_contains_bigdata_row_limit_controls() -> None:
    function_formula = M_QUERIES["fnTrinoRestQuery"]
    result_formula = M_QUERIES["TrinoResult"]

    assert "ResultLimitRows" in function_formula
    assert "1000000" in function_formula
    assert "EffectiveSqlText" in function_formula
    assert "Trino result row limit exceeded" in function_formula
    assert "Table.TransformColumnTypes" in function_formula
    assert "TrinoTypeToPowerQueryType" in function_formula
    assert "result_limit_rows" in result_formula
    assert "TrinoResultSchema" not in M_QUERIES


def test_m_code_maps_timestamp_before_time() -> None:
    function_formula = M_QUERIES["fnTrinoRestQuery"]
    timestamp_position = function_formula.index('Text.StartsWith(NormalizedType, "timestamp")')
    time_position = function_formula.index('Text.StartsWith(NormalizedType, "time")')

    assert timestamp_position < time_position
    assert "type datetimezone" in function_formula
    assert "type datetime" in function_formula
    assert "type time" in function_formula


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
