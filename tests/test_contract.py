from trino_excel_client.advanced_vba import ADVANCED_VBA_CODE
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
    assert 'BaseType = "date" then\n                    type date' in function_formula
    assert 'Text.StartsWith(NormalizedType, "timestamp") and Text.Contains(NormalizedType, "with time zone") then\n                    type datetimezone' in function_formula
    assert 'Text.StartsWith(NormalizedType, "timestamp") then\n                    type datetime' in function_formula
    assert 'Text.StartsWith(NormalizedType, "time") then\n                    type time' in function_formula
    assert "result_limit_rows" in result_formula
    assert "TrinoResultSchema" in M_QUERIES
    assert "ReturnSchemaOnly = true" in M_QUERIES["TrinoResultSchema"]
    assert "excel_trino_client_schema limit 0" in function_formula
    assert "TrinoTypeToExcelFormat" in function_formula


def test_m_code_parses_timestamp_as_datetime_before_time() -> None:
    function_formula = M_QUERIES["fnTrinoRestQuery"]
    timestamp_position = function_formula.index('Text.StartsWith(NormalizedType, "timestamp")')
    time_position = function_formula.index('Text.StartsWith(NormalizedType, "time")')

    assert timestamp_position < time_position
    assert "ParseTrinoTimestamp" in function_formula
    assert "ParseTrinoTimestampZone" in function_formula
    assert "type datetimezone" in function_formula
    assert "type datetime" in function_formula


def test_m_code_parses_all_trino_temporal_types_explicitly() -> None:
    function_formula = M_QUERIES["fnTrinoRestQuery"]

    assert "ParseTrinoDate" in function_formula
    assert "ParseTrinoTimestamp" in function_formula
    assert "ParseTrinoTimestampZone" in function_formula
    assert "ParseTrinoTime" in function_formula
    assert "DateColumnTransforms" in function_formula
    assert "TimestampColumnTransforms" in function_formula
    assert "TimestampZoneColumnTransforms" in function_formula
    assert "TimeColumnTransforms" in function_formula
    assert "type date" in function_formula
    assert "type time" in function_formula


def test_m_code_does_not_parse_timestamp_as_time() -> None:
    function_formula = M_QUERIES["fnTrinoRestQuery"]

    assert 'Text.StartsWith(_{2}, "time") and not Text.StartsWith(_{2}, "timestamp")' in function_formula


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


def test_advanced_vba_blocks_are_balanced_and_avoid_line_continuations() -> None:
    stack: list[tuple[str, int, str]] = []

    for line_number, line in enumerate(ADVANCED_VBA_CODE.splitlines(), 1):
        stripped = line.strip().lower()
        if stripped.startswith(("public sub ", "private sub ")):
            stack.append(("sub", line_number, line.strip()))
        elif stripped.startswith(("public function ", "private function ")):
            stack.append(("function", line_number, line.strip()))
        elif stripped == "end sub":
            assert stack and stack[-1][0] == "sub", (line_number, stack[-1] if stack else None)
            stack.pop()
        elif stripped == "end function":
            assert stack and stack[-1][0] == "function", (line_number, stack[-1] if stack else None)
            stack.pop()

        assert not line.rstrip().endswith("_"), line_number

    assert stack == []
