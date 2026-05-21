from __future__ import annotations

from pathlib import Path
from typing import Iterable, Sequence

from .m_code import M_QUERIES
from .template import (
    CONFIG_TABLE,
    EXTRA_CREDENTIALS_TABLE,
    HELP_TABLE,
    SQL_TABLE,
    client_sheet_names,
    sheet_name,
)

XL_SRC_RANGE = 1
XL_SRC_EXTERNAL = 0
XL_YES = 1
XL_CMD_SQL = 2
XL_OPENXML_WORKBOOK = 51
XL_OPENXML_WORKBOOK_MACRO_ENABLED = 52
XL_VALIDATE_LIST = 3
XL_VALID_ALERT_STOP = 1
XL_BETWEEN = 1
XL_CENTER = -4108
XL_LEFT = -4131
XL_TOP = -4160


def _rgb(red: int, green: int, blue: int) -> int:
    return red + (green * 256) + (blue * 65536)


COLORS = {
    "navy": _rgb(31, 51, 73),
    "blue": _rgb(46, 117, 182),
    "light_blue": _rgb(221, 235, 247),
    "green": _rgb(112, 173, 71),
    "light_green": _rgb(226, 239, 218),
    "orange": _rgb(244, 176, 132),
    "light_orange": _rgb(252, 228, 214),
    "gray": _rgb(242, 242, 242),
    "dark_gray": _rgb(89, 89, 89),
    "white": _rgb(255, 255, 255),
}


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


def _matrix(values: Iterable[Iterable[object]]) -> tuple[tuple[object, ...], ...]:
    return tuple(tuple(row) for row in values)


def _worksheet_exists(wb, name: str) -> bool:
    try:
        wb.Worksheets(name)
        return True
    except Exception:
        return False


def _ensure_sheet(wb, name: str, after=None):
    if _worksheet_exists(wb, name):
        return wb.Worksheets(name)
    if after is None:
        ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    else:
        ws = wb.Worksheets.Add(After=after)
    ws.Name = name
    return ws


def _delete_extra_default_sheets(wb, keep_names: set[str]) -> None:
    for idx in range(wb.Worksheets.Count, 0, -1):
        ws = wb.Worksheets(idx)
        if ws.Name not in keep_names and wb.Worksheets.Count > len(keep_names):
            ws.Delete()


def _delete_query_if_exists(wb, query_name: str) -> None:
    try:
        wb.Queries(query_name).Delete()
    except Exception:
        pass


def _delete_table_if_exists(wb, table_name: str) -> None:
    for ws in wb.Worksheets:
        try:
            ws.ListObjects(table_name).Delete()
        except Exception:
            pass


def _write_table(
    ws,
    start_cell: str,
    table_name: str,
    headers: Sequence[str],
    rows: Sequence[Sequence[object]],
    style: str = "TableStyleMedium2",
):
    data = [headers, *rows]
    top_left = ws.Range(start_cell)
    rng = top_left.Resize(len(data), len(headers))
    rng.Value = _matrix(data)
    table = ws.ListObjects.Add(XL_SRC_RANGE, rng, None, XL_YES)
    table.Name = table_name
    table.TableStyle = style
    return table


def _style_title(ws, address: str, text: str, width: int = 5) -> None:
    cell = ws.Range(address)
    cell.Value = text
    title_range = cell.Resize(1, width)
    title_range.Merge()
    title_range.Interior.Color = COLORS["navy"]
    title_range.Font.Color = COLORS["white"]
    title_range.Font.Bold = True
    title_range.Font.Size = 16
    title_range.HorizontalAlignment = XL_LEFT
    title_range.VerticalAlignment = XL_CENTER
    title_range.RowHeight = 32


def _section_header(ws, address: str, text: str, width: int = 5) -> None:
    cell = ws.Range(address)
    cell.Value = text
    rng = cell.Resize(1, width)
    rng.Merge()
    rng.Interior.Color = COLORS["light_blue"]
    rng.Font.Color = COLORS["navy"]
    rng.Font.Bold = True
    rng.RowHeight = 24


def _note(ws, address: str, text: str, width: int = 5, fill: str = "gray") -> None:
    cell = ws.Range(address)
    cell.Value = text
    rng = cell.Resize(1, width)
    rng.Merge()
    rng.Interior.Color = COLORS[fill]
    rng.Font.Color = COLORS["dark_gray"]
    rng.WrapText = True
    rng.VerticalAlignment = XL_TOP
    rng.RowHeight = 46


def _apply_table_defaults(table) -> None:
    table.Range.Font.Name = "Aptos"
    table.HeaderRowRange.Font.Bold = True
    table.HeaderRowRange.WrapText = True
    table.DataBodyRange.WrapText = True
    table.Range.VerticalAlignment = XL_TOP


def _add_validation_list(rng, values: Sequence[str]) -> None:
    rng.Validation.Delete()
    rng.Validation.Add(
        Type=XL_VALIDATE_LIST,
        AlertStyle=XL_VALID_ALERT_STOP,
        Operator=XL_BETWEEN,
        Formula1=",".join(values),
    )
    rng.Validation.IgnoreBlank = True
    rng.Validation.InCellDropdown = True


def _prepare_sheet(ws) -> None:
    ws.Cells.Clear()
    for idx in range(ws.ListObjects.Count, 0, -1):
        ws.ListObjects(idx).Delete()
    ws.Cells.Font.Name = "Aptos"
    ws.Cells.Font.Size = 10
    ws.Tab.Color = COLORS["blue"]


def _format_config_sheet(ws) -> None:
    ws.Columns("A:A").ColumnWidth = 28
    ws.Columns("B:B").ColumnWidth = 34
    ws.Columns("C:C").ColumnWidth = 13
    ws.Columns("D:D").ColumnWidth = 62
    ws.Columns("E:E").ColumnWidth = 2
    ws.Rows("1:1").RowHeight = 32
    ws.Range("A:Z").VerticalAlignment = XL_TOP


def _format_query_sheet(ws) -> None:
    ws.Columns("A:A").ColumnWidth = 150
    ws.Columns("B:D").ColumnWidth = 18
    ws.Rows("4:5").RowHeight = 24
    ws.Range("A:A").WrapText = False
    ws.Range("A:Z").VerticalAlignment = XL_TOP


def _format_help_sheet(ws) -> None:
    ws.Columns("A:A").ColumnWidth = 6
    ws.Columns("B:B").ColumnWidth = 38
    ws.Columns("C:C").ColumnWidth = 96
    ws.Range("A:Z").VerticalAlignment = XL_TOP
    ws.Range("B:C").WrapText = True


def _add_config_sheet(ws) -> None:
    _prepare_sheet(ws)
    _style_title(ws, "A1", "Trino REST API Client - Config", width=4)
    _note(
        ws,
        "A2",
        "Заполните колонку 'значение'. Парольные ячейки только визуально скрыты, однако не защищают ваши учетные данные, будьте осторожны.",
        width=4,
        fill="light_orange",
    )

    _section_header(ws, "A4", "Параметры подключения", width=4)
    config_table = _write_table(
        ws,
        "A5",
        CONFIG_TABLE.name,
        list(CONFIG_TABLE.headers),
        [list(row) for row in CONFIG_TABLE.rows],
        style=CONFIG_TABLE.style,
    )
    _apply_table_defaults(config_table)
    config_table.TableStyle = "TableStyleMedium2"
    config_table.ListColumns(2).DataBodyRange.Rows(3).NumberFormat = "General"
    config_table.ListColumns(2).DataBodyRange.Rows(4).NumberFormat = ";;;"
    config_table.ListColumns(3).DataBodyRange.HorizontalAlignment = XL_CENTER

    _section_header(ws, "A19", "Extra credentials для источников данных", width=5)
    _note(
        ws,
        "A20",
        "Используйте эту таблицу для заполнения extraCredentials параметров Trino, например gp-user/gp-password. Активные строки должны быть = TRUE для использования.",
        width=5,
        fill="light_green",
    )
    extra_table = _write_table(
        ws,
        "A22",
        EXTRA_CREDENTIALS_TABLE.name,
        list(EXTRA_CREDENTIALS_TABLE.headers),
        [list(row) for row in EXTRA_CREDENTIALS_TABLE.rows],
        style=EXTRA_CREDENTIALS_TABLE.style,
    )
    _apply_table_defaults(extra_table)
    _add_validation_list(extra_table.ListColumns(1).DataBodyRange, ["TRUE", "FALSE"])
    extra_table.ListColumns(5).DataBodyRange.NumberFormat = ";;;"
    ws.Range("A22:F22").Interior.Color = COLORS["green"]
    ws.Range("A22:F22").Font.Color = COLORS["white"]

    _format_config_sheet(ws)
    ws.Range("A5:D5").AutoFilter()
    ws.Activate()
    ws.Range("A5").Select()
    ws.Application.ActiveWindow.FreezePanes = True


def _add_query_sheet(ws) -> None:
    _prepare_sheet(ws)
    _style_title(ws, "A1", "Trino REST API Client - Query", width=4)
    _note(
        ws,
        "A2",
        "Введите SQL в одну широкую строку ниже. После заполнения Config нажмите Данные -> Обновить все.",
        width=4,
        fill="light_blue",
    )
    sql_table = _write_table(
        ws,
        "A4",
        SQL_TABLE.name,
        list(SQL_TABLE.headers),
        [list(row) for row in SQL_TABLE.rows],
        style=SQL_TABLE.style,
    )
    _apply_table_defaults(sql_table)
    sql_table.DataBodyRange.Font.Name = "Consolas"
    sql_table.DataBodyRange.Font.Size = 10
    sql_table.DataBodyRange.WrapText = False

    _section_header(ws, "A8", "Результат", width=4)
    ws.Range("A9").Value = "Результат Power Query TrinoResult будет загружен ниже после обновления."
    ws.Range("A9").Font.Italic = True
    ws.Range("A9").Interior.Color = COLORS["gray"]
    _format_query_sheet(ws)
    ws.Activate()
    ws.Range("A5").Select()
    ws.Application.ActiveWindow.FreezePanes = True


def _add_help_sheet(ws) -> None:
    _prepare_sheet(ws)
    _style_title(ws, "A1", "Trino REST API Client - Help", width=3)
    table = _write_table(
        ws,
        "A4",
        HELP_TABLE.name,
        list(HELP_TABLE.headers),
        [list(row) for row in HELP_TABLE.rows],
        style=HELP_TABLE.style,
    )
    _apply_table_defaults(table)
    table.ListColumns("step").DataBodyRange.HorizontalAlignment = XL_CENTER
    ws.Range("A4:C4").Interior.Color = COLORS["navy"]
    ws.Range("A4:C4").Font.Color = COLORS["white"]
    _format_help_sheet(ws)


def _add_power_queries(wb) -> None:
    for query_name, formula in M_QUERIES.items():
        _delete_query_if_exists(wb, query_name)
        wb.Queries.Add(
            Name=query_name,
            Formula=formula.strip(),
            Description="Trino REST API client query",
        )


def _try_create_result_table(query_ws) -> None:
    destination = query_ws.Range("A10")
    query_ws.Range("A10:Z500").Clear()
    connection_string = (
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;"
        "Data Source=$Workbook$;"
        "Location=TrinoResult;"
        'Extended Properties=""'
    )
    list_object = query_ws.ListObjects.Add(
        XL_SRC_EXTERNAL,
        connection_string,
        None,
        XL_YES,
        destination,
    )
    list_object.Name = "tblTrinoResult"
    query_table = list_object.QueryTable
    query_table.CommandType = XL_CMD_SQL
    query_table.CommandText = "SELECT * FROM [TrinoResult]"
    query_table.BackgroundQuery = False
    query_table.RefreshOnFileOpen = False


def _workbook_file_format(path: Path) -> int:
    if path.suffix.lower() == ".xlsm":
        return XL_OPENXML_WORKBOOK_MACRO_ENABLED
    return XL_OPENXML_WORKBOOK


def install_client_into_workbook(wb, sheet_prefix: str = "Trino") -> None:
    sheet_names = client_sheet_names(sheet_prefix)
    config_name = sheet_names.config
    query_name = sheet_names.query
    help_name = sheet_names.help

    _delete_table_if_exists(wb, "tblTrinoConfig")
    _delete_table_if_exists(wb, "tblExtraCredentials")
    _delete_table_if_exists(wb, "tblTrinoSql")
    _delete_table_if_exists(wb, "tblTrinoResult")
    _delete_table_if_exists(wb, "tblTrinoHelp")

    config_ws = _ensure_sheet(wb, config_name)
    query_ws = _ensure_sheet(wb, query_name, after=config_ws)
    help_ws = _ensure_sheet(wb, help_name, after=query_ws)

    _add_config_sheet(config_ws)
    _add_query_sheet(query_ws)
    _add_help_sheet(help_ws)
    _add_power_queries(wb)

    try:
        _try_create_result_table(query_ws)
    except Exception as exc:
        query_ws.Range("A10").Value = (
            "Power Query-запросы добавлены, но автоматическое создание таблицы результата не сработало. "
            f"Используйте Данные -> Запросы и подключения -> TrinoResult -> Загрузить в. Детали: {exc}"
        )
        query_ws.Range("A10").WrapText = True
        query_ws.Range("A10").Interior.Color = COLORS["light_orange"]


def create_workbook(
    output_path: Path,
    overwrite: bool = False,
    visible: bool = False,
    sheet_prefix: str = "Trino",
) -> Path:
    output_path = output_path.resolve()
    if output_path.exists() and not overwrite:
        raise FileExistsError(f"File already exists: {output_path}")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    pythoncom, win32 = _require_excel_modules()
    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = visible
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Add()
        install_client_into_workbook(wb, sheet_prefix=sheet_prefix)
        keep_names = {
            sheet_name(sheet_prefix, "Config"),
            sheet_name(sheet_prefix, "Query"),
            sheet_name(sheet_prefix, "Help"),
        }
        _delete_extra_default_sheets(wb, keep_names)
        wb.SaveAs(str(output_path), FileFormat=_workbook_file_format(output_path))
        wb.Close(SaveChanges=True)
        return output_path
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            excel.DisplayAlerts = True
            excel.Quit()
        pythoncom.CoUninitialize()


def install_client(
    workbook_path: Path,
    output_path: Path | None = None,
    overwrite: bool = False,
    visible: bool = False,
    sheet_prefix: str = "Trino",
) -> Path:
    workbook_path = workbook_path.resolve()
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook does not exist: {workbook_path}")

    final_path = (output_path.resolve() if output_path is not None else workbook_path)
    if output_path is not None and final_path.exists() and final_path != workbook_path and not overwrite:
        raise FileExistsError(f"File already exists: {final_path}")

    pythoncom, win32 = _require_excel_modules()
    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = visible
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(workbook_path))
        install_client_into_workbook(wb, sheet_prefix=sheet_prefix)
        if output_path is None or final_path == workbook_path:
            wb.Save()
        else:
            final_path.parent.mkdir(parents=True, exist_ok=True)
            wb.SaveAs(str(final_path), FileFormat=_workbook_file_format(final_path))
        wb.Close(SaveChanges=True)
        return final_path
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            excel.DisplayAlerts = True
            excel.Quit()
        pythoncom.CoUninitialize()
