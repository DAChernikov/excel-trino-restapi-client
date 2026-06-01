from __future__ import annotations

from pathlib import Path
from tempfile import NamedTemporaryFile

from .advanced_vba import ADVANCED_VBA_CODE, ADVANCED_VBA_MODULE_NAME
from .m_code import M_QUERIES
from .openpyxl_backend import (
    create_advanced_workbook_ui,
    create_workbook_ui,
    install_advanced_client_ui_file,
    install_client_ui_file,
)
from .ooxml import patch_standard_temporal_formats
from .template import client_sheet_names

MSO_SHAPE_ROUNDED_RECTANGLE = 5
XL_SRC_EXTERNAL = 0
XL_YES = 1
XL_CMD_SQL = 2
XL_OPENXML_WORKBOOK = 51
XL_OPENXML_WORKBOOK_MACRO_ENABLED = 52
XL_OVERWRITE_CELLS = 0
XL_EXPRESSION = 2
XL_SHEET_HIDDEN = 0

SCHEMA_QUERY_NAME = "TrinoResultSchema"
SCHEMA_TABLE_NAME = "tblTrinoResultSchema"
SCHEMA_COLUMN_NAME_RANGE_NAME = "TrinoSchemaColumnName"
SCHEMA_EXCEL_FORMAT_RANGE_NAME = "TrinoSchemaExcelFormat"
RESULT_TEMPORAL_FORMAT_RANGE = "A6:XFD1048576"


def _require_excel_modules():
    try:
        import pythoncom  # type: ignore
        import win32com.client as win32  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "Excel automation requires Windows, desktop Microsoft Excel, and pywin32. "
            "Install on Windows with: pip install -e ."
        ) from exc

    return pythoncom, win32


def _workbook_file_format(path: Path) -> int:
    if path.suffix.lower() == ".xlsm":
        return XL_OPENXML_WORKBOOK_MACRO_ENABLED

    return XL_OPENXML_WORKBOOK


def _delete_query_if_exists(wb, query_name: str) -> None:
    try:
        wb.Queries(query_name).Delete()
    except Exception:
        pass


def _delete_name_if_exists(wb, name: str) -> None:
    try:
        wb.Names(name).Delete()
    except Exception:
        pass


def _quote_excel_sheet_name(name: str) -> str:
    return "'" + name.replace("'", "''") + "'"


def _add_power_queries(wb, include_schema: bool = True) -> None:
    for query_name, formula in M_QUERIES.items():
        if not include_schema and query_name == SCHEMA_QUERY_NAME:
            continue
        _delete_query_if_exists(wb, query_name)
        wb.Queries.Add(
            Name=query_name,
            Formula=formula.strip(),
            Description="Trino REST API client query",
        )


def _safe_set_attr(obj, attr_name: str, value) -> None:
    try:
        setattr(obj, attr_name, value)
    except Exception:
        pass


def _disable_automatic_connection_refresh(connection) -> None:
    for attr_name, value in (
        ("RefreshWithRefreshAll", False),
        ("RefreshOnFileOpen", False),
    ):
        _safe_set_attr(connection, attr_name, value)

    for child_name in ("OLEDBConnection", "ODBCConnection"):
        try:
            child = getattr(connection, child_name)
        except Exception:
            continue
        for attr_name, value in (
            ("BackgroundQuery", False),
            ("RefreshOnFileOpen", False),
            ("RefreshPeriod", 0),
            ("EnableRefresh", True),
        ):
            _safe_set_attr(child, attr_name, value)


def _enable_manual_result_refresh(connection) -> None:
    _disable_automatic_connection_refresh(connection)
    _safe_set_attr(connection, "RefreshWithRefreshAll", True)


def _enforce_manual_refresh_only(wb, refresh_connections=None) -> None:
    refresh_connections = tuple(refresh_connections or ())
    refresh_connection_names = set()
    for connection in refresh_connections:
        try:
            refresh_connection_names.add(connection.Name)
        except Exception:
            pass

    try:
        connection_count = wb.Connections.Count
    except Exception:
        return

    for idx in range(1, connection_count + 1):
        try:
            connection = wb.Connections(idx)
        except Exception:
            continue

        try:
            connection_name = connection.Name
        except Exception:
            connection_name = ""

        is_refresh_connection = any(
            connection is refresh_connection for refresh_connection in refresh_connections
        ) or (
            connection_name != "" and connection_name in refresh_connection_names
        )
        if is_refresh_connection:
            _enable_manual_result_refresh(connection)
        else:
            _disable_automatic_connection_refresh(connection)


def _delete_result_table_if_exists(result_ws) -> None:
    try:
        result_ws.ListObjects("tblTrinoResult").Delete()
    except Exception:
        pass


def _delete_schema_table_if_exists(schema_ws) -> None:
    try:
        schema_ws.ListObjects(SCHEMA_TABLE_NAME).Delete()
    except Exception:
        pass


def _create_power_query_table(ws, *, table_name: str, query_name: str, destination_address: str):
    destination = ws.Range(destination_address)
    connection_string = (
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;"
        "Data Source=$Workbook$;"
        f"Location={query_name};"
        'Extended Properties=""'
    )

    list_object = ws.ListObjects.Add(
        XL_SRC_EXTERNAL,
        connection_string,
        None,
        XL_YES,
        destination,
    )

    list_object.Name = table_name
    try:
        list_object.TableStyle = "TableStyleMedium4"
    except Exception:
        pass

    query_table = list_object.QueryTable
    query_table.CommandType = XL_CMD_SQL
    query_table.CommandText = f"SELECT * FROM [{query_name}]"
    preserve_formatting = table_name != "tblTrinoResult"
    preserve_column_info = table_name != "tblTrinoResult"
    for attr_name, value in (
        ("BackgroundQuery", False),
        ("RefreshOnFileOpen", False),
        ("RefreshPeriod", 0),
        ("EnableRefresh", True),
        ("SaveData", True),
        ("PreserveFormatting", preserve_formatting),
        ("PreserveColumnInfo", preserve_column_info),
        ("AdjustColumnWidth", True),
        ("RefreshStyle", XL_OVERWRITE_CELLS),
    ):
        _safe_set_attr(query_table, attr_name, value)

    try:
        return query_table.WorkbookConnection
    except Exception:
        return None


def _write_result_table_creation_note(result_ws, exc: Exception) -> None:
    message = (
        "Power Query-запросы добавлены, но автоматическое создание таблицы результата "
        "не сработало. Используйте: Данные -> Запросы и подключения -> "
        "TrinoResult -> Загрузить в. "
        f"Детали: {exc}"
    )
    note = result_ws.Range("A5")
    note.Value = message

    try:
        note.WrapText = True
    except Exception:
        pass

    try:
        note.Interior.Color = 14277081
    except Exception:
        pass


def _try_create_result_table(wb, sheet_prefix: str = "Trino"):
    result_sheet_name = client_sheet_names(sheet_prefix).result
    result_ws = wb.Worksheets(result_sheet_name)

    _delete_result_table_if_exists(result_ws)

    result_ws.Range("A5:Z500").Clear()

    return _create_power_query_table(
        result_ws,
        table_name="tblTrinoResult",
        query_name="TrinoResult",
        destination_address="A5",
    )


def _try_create_schema_table(wb, sheet_prefix: str = "Trino"):
    schema_sheet_name = client_sheet_names(sheet_prefix).schema
    schema_ws = wb.Worksheets(schema_sheet_name)

    _delete_schema_table_if_exists(schema_ws)
    schema_ws.Range("A1:Z500").Clear()

    schema_connection = _create_power_query_table(
        schema_ws,
        table_name=SCHEMA_TABLE_NAME,
        query_name=SCHEMA_QUERY_NAME,
        destination_address="A1",
    )

    try:
        schema_ws.Visible = XL_SHEET_HIDDEN
    except Exception:
        pass

    return schema_connection


def _add_schema_format_names(wb, sheet_prefix: str) -> None:
    schema_sheet_name = _quote_excel_sheet_name(client_sheet_names(sheet_prefix).schema)
    for name in (SCHEMA_COLUMN_NAME_RANGE_NAME, SCHEMA_EXCEL_FORMAT_RANGE_NAME):
        _delete_name_if_exists(wb, name)

    try:
        wb.Names.Add(Name=SCHEMA_COLUMN_NAME_RANGE_NAME, RefersTo=f"={schema_sheet_name}!$A:$A")
    except Exception:
        pass

    try:
        wb.Names.Add(Name=SCHEMA_EXCEL_FORMAT_RANGE_NAME, RefersTo=f"={schema_sheet_name}!$C:$C")
    except Exception:
        pass


def _localize_excel_formulas(ws, formulas: tuple[str, ...]) -> tuple[str, ...]:
    """
    Conditional formatting formulas are locale-sensitive in Excel COM.
    Use A6, the top-left result data cell, so relative header references like A$5
    stay aligned with the conditional formatting range that also starts at A6.
    """
    probe = None
    localized: list[str] = []

    try:
        probe = ws.Range("A6")
        for formula in formulas:
            try:
                probe.Formula = formula
                localized.append(str(probe.FormulaLocal))
            except Exception:
                localized.append(formula)
        return tuple(localized)
    finally:
        if probe is not None:
            try:
                probe.Clear()
            except Exception:
                pass


def _add_temporal_number_format_rule(target_range, formula: str, number_format: str) -> None:
    try:
        condition = target_range.FormatConditions.Add(Type=XL_EXPRESSION, Formula1=formula)
        condition.NumberFormat = number_format
        condition.StopIfTrue = False
    except Exception:
        pass


def _apply_temporal_result_formats(result_ws) -> None:
    target_range = result_ws.Range(RESULT_TEMPORAL_FORMAT_RANGE)
    try:
        target_range.FormatConditions.Delete()
    except Exception:
        pass

    date_formula = (
        f'=IFERROR(INDEX({SCHEMA_EXCEL_FORMAT_RANGE_NAME},'
        f'MATCH(A$5,{SCHEMA_COLUMN_NAME_RANGE_NAME},0))="date",FALSE)'
    )
    datetime_formula = (
        f'=IFERROR(OR('
        f'INDEX({SCHEMA_EXCEL_FORMAT_RANGE_NAME},MATCH(A$5,{SCHEMA_COLUMN_NAME_RANGE_NAME},0))="datetime",'
        f'INDEX({SCHEMA_EXCEL_FORMAT_RANGE_NAME},MATCH(A$5,{SCHEMA_COLUMN_NAME_RANGE_NAME},0))="datetimezone"'
        f'),FALSE)'
    )
    time_formula = (
        f'=IFERROR(INDEX({SCHEMA_EXCEL_FORMAT_RANGE_NAME},'
        f'MATCH(A$5,{SCHEMA_COLUMN_NAME_RANGE_NAME},0))="time",FALSE)'
    )
    localized_formulas = _localize_excel_formulas(
        result_ws,
        (date_formula, datetime_formula, time_formula),
    )

    _add_temporal_number_format_rule(target_range, localized_formulas[0], "dd.mm.yyyy")
    _add_temporal_number_format_rule(target_range, localized_formulas[1], "dd.mm.yyyy hh:mm")
    _add_temporal_number_format_rule(target_range, localized_formulas[2], "hh:mm:ss")


def install_power_query_into_workbook(wb, sheet_prefix: str = "Trino") -> None:
    _add_power_queries(wb, include_schema=True)

    refresh_connections = []
    try:
        schema_connection = _try_create_schema_table(wb, sheet_prefix=sheet_prefix)
        if schema_connection is not None:
            refresh_connections.append(schema_connection)
        _add_schema_format_names(wb, sheet_prefix=sheet_prefix)

        result_connection = _try_create_result_table(wb, sheet_prefix=sheet_prefix)
        if result_connection is not None:
            refresh_connections.append(result_connection)

        result_ws = wb.Worksheets(client_sheet_names(sheet_prefix).result)
        _apply_temporal_result_formats(result_ws)
    except Exception as exc:
        result_ws = wb.Worksheets(client_sheet_names(sheet_prefix).result)
        _write_result_table_creation_note(result_ws, exc)
    finally:
        _enforce_manual_refresh_only(wb, refresh_connections=refresh_connections)


def _delete_vba_module_if_exists(wb, module_name: str) -> None:
    try:
        components = wb.VBProject.VBComponents
    except Exception as exc:
        raise RuntimeError(
            "Advanced .xlsm generation requires Excel VBA project access. "
            "Enable 'Trust access to the VBA project object model' in Excel Trust Center."
        ) from exc

    try:
        components.Remove(components(module_name))
    except Exception:
        pass


def _install_advanced_vba_module(wb) -> None:
    _delete_vba_module_if_exists(wb, ADVANCED_VBA_MODULE_NAME)
    component = wb.VBProject.VBComponents.Add(1)
    component.Name = ADVANCED_VBA_MODULE_NAME
    component.CodeModule.AddFromString(ADVANCED_VBA_CODE.strip())


def _delete_advanced_button_if_exists(query_ws) -> None:
    try:
        for idx in range(query_ws.Shapes.Count, 0, -1):
            shape = query_ws.Shapes(idx)
            try:
                if shape.Name == "btnTrinoRunQueryToSheet":
                    shape.Delete()
            except Exception:
                pass
    except Exception:
        pass


def _install_advanced_button(wb, sheet_prefix: str, macro_workbook_name: str | None = None) -> None:
    query_ws = wb.Worksheets(client_sheet_names(sheet_prefix).query)
    _delete_advanced_button_if_exists(query_ws)
    workbook_name = macro_workbook_name or wb.Name
    macro_name = f"'{workbook_name}'!{ADVANCED_VBA_MODULE_NAME}.TrinoRunQueryToSheet"

    try:
        anchor = query_ws.Range("A15")
        button = query_ws.Shapes.AddShape(
            MSO_SHAPE_ROUNDED_RECTANGLE,
            anchor.Left,
            anchor.Top,
            220,
            32,
        )
        button.Name = "btnTrinoRunQueryToSheet"
        button.OnAction = macro_name
        button.TextFrame.Characters().Text = "Выгрузить в лист"
    except Exception:
        note = query_ws.Range("A15")
        note.Value = "Макрос TrinoRunQueryToSheet добавлен. Запустите его через Alt+F8, если кнопка не создана."
        try:
            note.WrapText = True
        except Exception:
            pass


def install_advanced_client_automation(
    wb,
    sheet_prefix: str = "Trino",
    macro_workbook_name: str | None = None,
) -> None:
    _add_power_queries(wb, include_schema=False)
    _enforce_manual_refresh_only(wb)
    _install_advanced_vba_module(wb)
    _install_advanced_button(wb, sheet_prefix=sheet_prefix, macro_workbook_name=macro_workbook_name)


def _open_workbook_and_install_power_query(
    workbook_path: Path,
    *,
    visible: bool,
    sheet_prefix: str,
    save_as_path: Path | None = None,
    advanced: bool = False,
) -> Path:
    pythoncom, win32 = _require_excel_modules()
    pythoncom.CoInitialize()

    excel = None
    wb = None
    final_path = save_as_path.resolve() if save_as_path is not None else workbook_path.resolve()

    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = visible
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(str(workbook_path.resolve()))
        if advanced:
            install_advanced_client_automation(
                wb,
                sheet_prefix=sheet_prefix,
                macro_workbook_name=final_path.name,
            )
        else:
            install_power_query_into_workbook(wb, sheet_prefix=sheet_prefix)

        if save_as_path is not None:
            final_path.parent.mkdir(parents=True, exist_ok=True)
            wb.SaveAs(str(final_path), FileFormat=_workbook_file_format(final_path))
        else:
            wb.Save()

        wb.Close(SaveChanges=True)
        wb = None
        if not advanced and final_path.suffix.lower() == ".xlsx":
            patch_standard_temporal_formats(final_path, sheet_prefix=sheet_prefix)
        return final_path

    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass

        if excel is not None:
            try:
                excel.DisplayAlerts = True
                excel.Quit()
            except Exception:
                pass

        pythoncom.CoUninitialize()


def _temporary_xlsx_path(output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with NamedTemporaryFile(
        prefix=f".{output_path.stem}.",
        suffix=".xlsx",
        dir=output_path.parent,
        delete=False,
    ) as temp_file:
        return Path(temp_file.name)


def create_workbook(
    output_path: Path,
    overwrite: bool = False,
    visible: bool = False,
    sheet_prefix: str = "Trino",
    advanced: bool = False,
) -> Path:
    output_path = output_path.resolve()

    if output_path.exists() and not overwrite:
        raise FileExistsError(f"File already exists: {output_path}")

    if advanced and output_path.suffix.lower() != ".xlsm":
        raise ValueError("Advanced Trino client must be saved as .xlsm.")

    if output_path.suffix.lower() == ".xlsm":
        if output_path.exists():
            output_path.unlink()

        temp_path = _temporary_xlsx_path(output_path)
        try:
            create_ui = create_advanced_workbook_ui if advanced else create_workbook_ui
            create_ui(
                output_path=temp_path,
                overwrite=True,
                sheet_prefix=sheet_prefix,
            )
            return _open_workbook_and_install_power_query(
                temp_path,
                visible=visible,
                sheet_prefix=sheet_prefix,
                save_as_path=output_path,
                advanced=advanced,
            )
        finally:
            try:
                temp_path.unlink()
            except FileNotFoundError:
                pass

    create_workbook_ui(
        output_path=output_path,
        overwrite=overwrite,
        sheet_prefix=sheet_prefix,
    )
    return _open_workbook_and_install_power_query(
        output_path,
        visible=visible,
        sheet_prefix=sheet_prefix,
        advanced=advanced,
    )


def create_advanced_workbook(
    output_path: Path,
    overwrite: bool = False,
    visible: bool = False,
    sheet_prefix: str = "Trino",
) -> Path:
    return create_workbook(
        output_path=output_path,
        overwrite=overwrite,
        visible=visible,
        sheet_prefix=sheet_prefix,
        advanced=True,
    )


def install_client(
    workbook_path: Path,
    output_path: Path | None = None,
    overwrite: bool = False,
    visible: bool = False,
    sheet_prefix: str = "Trino",
    advanced: bool = False,
) -> Path:
    if advanced:
        final_path_candidate = output_path if output_path is not None else workbook_path
        if final_path_candidate.suffix.lower() != ".xlsm":
            raise ValueError("Advanced Trino client must be installed into or saved as .xlsm.")

    install_ui = install_advanced_client_ui_file if advanced else install_client_ui_file
    final_path = install_ui(
        workbook_path=workbook_path,
        output_path=output_path,
        overwrite=overwrite,
        sheet_prefix=sheet_prefix,
    )
    return _open_workbook_and_install_power_query(
        final_path,
        visible=visible,
        sheet_prefix=sheet_prefix,
        advanced=advanced,
    )


def install_advanced_client(
    workbook_path: Path,
    output_path: Path | None = None,
    overwrite: bool = False,
    visible: bool = False,
    sheet_prefix: str = "Trino",
) -> Path:
    return install_client(
        workbook_path=workbook_path,
        output_path=output_path,
        overwrite=overwrite,
        visible=visible,
        sheet_prefix=sheet_prefix,
        advanced=True,
    )
