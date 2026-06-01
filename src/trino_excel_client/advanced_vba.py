from __future__ import annotations

from .m_code import M_QUERIES

ADVANCED_VBA_MODULE_NAME = "TrinoAdvancedClient"

ADVANCED_VBA_CODE = r'''
Option Explicit

Private Const TARGET_TABLE_NAME As String = "tblTrinoAdvancedTarget"
Private Const SQL_TABLE_NAME As String = "tblTrinoSql"
Private Const CONFIG_TABLE_NAME As String = "tblTrinoConfig"
Private Const EXTRA_CREDENTIALS_TABLE_NAME As String = "tblExtraCredentials"

Public Sub TrinoRunQueryToSheet()
    On Error GoTo Fail

    Dim targetSheetName As String
    Dim sqlText As String
    Dim queryName As String
    Dim tableName As String
    Dim resultWs As Worksheet
    Dim resultTable As ListObject

    targetSheetName = Trim(ReadAdvancedValue("target_sheet"))
    If targetSheetName = "" Then
        Err.Raise vbObjectError + 5100, "TrinoRunQueryToSheet", "Укажите лист результата в параметре target_sheet."
    End If

    ValidateWorksheetName targetSheetName

    sqlText = ReadSqlText()
    If Trim(sqlText) = "" Then
        Err.Raise vbObjectError + 5101, "TrinoRunQueryToSheet", "Введите SQL-запрос в таблицу tblTrinoSql."
    End If

    queryName = "TrinoResult_" & StableToken(targetSheetName)
    tableName = "tblTrinoResult_" & StableToken(targetSheetName)

    Set resultWs = EnsureWorksheet(targetSheetName)
    PrepareResultWorksheet resultWs
    UpsertResultQuery queryName, sqlText
    Set resultTable = CreateResultTable(resultWs, queryName, tableName)

    resultTable.QueryTable.Refresh BackgroundQuery:=False
    MsgBox "Результат выгружен на лист '" & targetSheetName & "'.", vbInformation, "Trino Excel Client"
    Exit Sub

Fail:
    MsgBox BuildFriendlyErrorMessage(Err.Description), vbCritical, "Trino Excel Client"
End Sub

Private Function BuildFriendlyErrorMessage(ByVal details As String) As String
    Dim message As String
    Dim hasFirewallError As Boolean

    message = "Не удалось выполнить выгрузку Trino." & vbCrLf & details

    If InStr(1, details, "ResourceAccessForbidden", vbTextCompare) > 0 Then
        message = message & vbCrLf & vbCrLf
        message = message & "Excel заблокировал доступ Power Query к Trino host на этом ПК." & vbCrLf
        message = message & "Откройте: Данные -> Получить данные -> Параметры источника данных." & vbCrLf
        message = message & "Очистите старые разрешения для Trino host и при следующем запросе выберите Anonymous."
    End If

    hasFirewallError = InStr(1, details, "Formula.Firewall", vbTextCompare) > 0
    hasFirewallError = hasFirewallError Or InStr(1, details, "ссылается на другие запросы", vbTextCompare) > 0
    If hasFirewallError Then
        message = message & vbCrLf & vbCrLf
        message = message & "Excel применил Power Query Privacy Firewall на этом ПК." & vbCrLf
        message = message & "Пересоберите advanced-шаблон актуальной версией клиента. В актуальной версии параметры подключения встраиваются в запрос как значения, чтобы не смешивать Excel-таблицы и Web.Contents в одном Power Query partition."
    End If

    BuildFriendlyErrorMessage = message
End Function

Private Function ReadAdvancedValue(ByVal key As String) As String
    ReadAdvancedValue = ReadKeyValue(TARGET_TABLE_NAME, key)
End Function

Private Function ReadConfigValue(ByVal key As String) As String
    ReadConfigValue = ReadKeyValue(CONFIG_TABLE_NAME, key)
End Function

Private Function ReadKeyValue(ByVal tableName As String, ByVal key As String) As String
    Dim lo As ListObject
    Dim row As ListRow
    Dim currentKey As String

    Set lo = FindListObject(tableName)
    For Each row In lo.ListRows
        currentKey = LCase(Trim(CStr(row.Range.Cells(1, 1).Value)))
        If currentKey = LCase(key) Then
            ReadKeyValue = CStr(row.Range.Cells(1, 2).Value)
            Exit Function
        End If
    Next row

    ReadKeyValue = ""
End Function

Private Function ReadSqlText() As String
    Dim lo As ListObject
    Dim row As ListRow
    Dim parts As Collection
    Dim value As String
    Dim idx As Long
    Dim result As String

    Set lo = FindListObject(SQL_TABLE_NAME)
    Set parts = New Collection

    For Each row In lo.ListRows
        value = CStr(row.Range.Cells(1, 1).Value)
        If Trim(value) <> "" Then
            parts.Add value
        End If
    Next row

    For idx = 1 To parts.Count
        If idx > 1 Then result = result & vbLf
        result = result & CStr(parts(idx))
    Next idx

    ReadSqlText = result
End Function

Private Function FindListObject(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = tableName Then
                Set FindListObject = lo
                Exit Function
            End If
        Next lo
    Next ws

    Err.Raise vbObjectError + 5102, "FindListObject", "Не найдена таблица " & tableName & "."
End Function

Private Sub ValidateWorksheetName(ByVal sheetName As String)
    Dim badChars As Variant
    Dim item As Variant

    If Len(sheetName) > 31 Then
        Err.Raise vbObjectError + 5103, "ValidateWorksheetName", "Название листа не должно быть длиннее 31 символа."
    End If

    badChars = Array(":", "\", "/", "?", "*", "[", "]")
    For Each item In badChars
        If InStr(1, sheetName, CStr(item), vbBinaryCompare) > 0 Then
            Err.Raise vbObjectError + 5104, "ValidateWorksheetName", "Название листа содержит запрещенный символ: " & CStr(item)
        End If
    Next item
End Sub

Private Function EnsureWorksheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            Set EnsureWorksheet = ws
            Exit Function
        End If
    Next ws

    Set EnsureWorksheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    EnsureWorksheet.Name = sheetName
End Function

Private Sub UpsertResultQuery(ByVal queryName As String, ByVal sqlText As String)
    On Error Resume Next
    ThisWorkbook.Queries(queryName).Delete
    On Error GoTo 0

    ThisWorkbook.Queries.Add Name:=queryName, Formula:=BuildResultFormula(sqlText), Description:="Trino advanced sheet result query"
End Sub

Private Function BuildResultFormula(ByVal sqlText As String) As String
    Dim formula As String
    Dim baseUrl As String
    Dim trinoUser As String
    Dim trinoPassword As String
    Dim timeZoneName As String
    Dim catalogName As String
    Dim schemaName As String
    Dim requestTimeoutMinutes As String
    Dim resultLimitRows As String

    baseUrl = RequireConfigValue("trino_base_url")
    trinoUser = RequireConfigValue("trino_user")
    trinoPassword = RequireConfigValue("trino_password")
    timeZoneName = ConfigValueOrDefault("time_zone", "Europe/Moscow")
    catalogName = ReadConfigValue("trino_catalog")
    schemaName = ReadConfigValue("trino_schema")
    requestTimeoutMinutes = ConfigValueOrDefault("request_timeout_minutes", "5")
    resultLimitRows = ConfigValueOrDefault("result_limit_rows", "1000000")

    formula = "let" & vbCrLf
    formula = formula & "    TrinoRestQuery = " & TrinoRestFunctionFormula() & "," & vbCrLf
    formula = formula & "    SqlText = " & MTextLiteral(sqlText) & "," & vbCrLf
    formula = formula & "    ExtraCredentials = " & BuildExtraCredentialsTableLiteral() & "," & vbCrLf
    formula = formula & "    Result = TrinoRestQuery(" & vbCrLf
    formula = formula & "        " & MTextLiteral(baseUrl) & "," & vbCrLf
    formula = formula & "        " & MTextLiteral(trinoUser) & "," & vbCrLf
    formula = formula & "        " & MTextLiteral(trinoPassword) & "," & vbCrLf
    formula = formula & "        SqlText," & vbCrLf
    formula = formula & "        ExtraCredentials," & vbCrLf
    formula = formula & "        [" & vbCrLf
    formula = formula & "            TimeZone = " & MTextLiteral(timeZoneName) & "," & vbCrLf
    formula = formula & "            Catalog = " & MTextLiteral(catalogName) & "," & vbCrLf
    formula = formula & "            Schema = " & MTextLiteral(schemaName) & "," & vbCrLf
    formula = formula & "            RequestTimeoutMinutes = " & MTextLiteral(requestTimeoutMinutes) & "," & vbCrLf
    formula = formula & "            ResultLimitRows = " & MTextLiteral(resultLimitRows) & vbCrLf
    formula = formula & "        ]" & vbCrLf
    formula = formula & "    )" & vbCrLf
    formula = formula & "in" & vbCrLf
    formula = formula & "    Result"

    BuildResultFormula = formula
End Function

Private Function RequireConfigValue(ByVal key As String) As String
    Dim value As String

    value = Trim(ReadConfigValue(key))
    If value = "" Or UCase(value) = "CHANGE_ME" Then
        Err.Raise vbObjectError + 5105, "RequireConfigValue", "Заполните обязательный параметр Config: " & key & "."
    End If

    RequireConfigValue = value
End Function

Private Function ConfigValueOrDefault(ByVal key As String, ByVal defaultValue As String) As String
    Dim value As String

    value = Trim(ReadConfigValue(key))
    If value = "" Then
        ConfigValueOrDefault = defaultValue
    Else
        ConfigValueOrDefault = value
    End If
End Function

Private Function BuildExtraCredentialsTableLiteral() As String
    Dim lo As ListObject
    Dim row As ListRow
    Dim parts As Collection
    Dim rowsText As String
    Dim idx As Long

    Set lo = FindListObject(EXTRA_CREDENTIALS_TABLE_NAME)
    Set parts = New Collection

    For Each row In lo.ListRows
        If IsEnabledValue(row.Range.Cells(1, 1).Value) Then
            AddCredentialPair parts, CStr(row.Range.Cells(1, 2).Value), CStr(row.Range.Cells(1, 3).Value)
            AddCredentialPair parts, CStr(row.Range.Cells(1, 4).Value), CStr(row.Range.Cells(1, 5).Value)
        End If
    Next row

    If parts.Count = 0 Then
        BuildExtraCredentialsTableLiteral = "#table({""credential_name"", ""credential_value""}, {})"
        Exit Function
    End If

    For idx = 1 To parts.Count
        If idx > 1 Then rowsText = rowsText & ", "
        rowsText = rowsText & CStr(parts(idx))
    Next idx

    BuildExtraCredentialsTableLiteral = "#table({""credential_name"", ""credential_value""}, {" & rowsText & "})"
End Function

Private Sub AddCredentialPair(ByRef parts As Collection, ByVal credentialName As String, ByVal credentialValue As String)
    credentialName = Trim(credentialName)
    If credentialName <> "" And credentialValue <> "" And UCase(credentialValue) <> "CHANGE_ME" Then
        parts.Add "{" & MTextLiteral(credentialName) & ", " & MTextLiteral(credentialValue) & "}"
    End If
End Sub

Private Function IsEnabledValue(ByVal value As Variant) As Boolean
    Dim textValue As String

    textValue = LCase(Trim(CStr(value)))
    IsEnabledValue = (textValue = "true" Or textValue = "1" Or textValue = "yes" Or textValue = "y" Or textValue = "да" Or textValue = "истина")
End Function

Private Function MTextLiteral(ByVal value As String) As String
    value = Replace(value, """", """""")
    value = Replace(value, vbCrLf, "#(lf)")
    value = Replace(value, vbCr, "#(lf)")
    value = Replace(value, vbLf, "#(lf)")
    MTextLiteral = """" & value & """"
End Function

Private Sub PrepareResultWorksheet(ByVal resultWs As Worksheet)
    DeleteManagedResultTablesOnSheet resultWs
End Sub

Private Sub DeleteManagedResultTablesOnSheet(ByVal resultWs As Worksheet)
    Dim idx As Long
    Dim lo As ListObject
    Dim tableAddress As String

    For idx = resultWs.ListObjects.Count To 1 Step -1
        Set lo = resultWs.ListObjects(idx)
        If IsManagedResultTableName(lo.Name) Then
            tableAddress = lo.Range.Address
            lo.Delete
            resultWs.Range(tableAddress).Clear
        End If
    Next idx
End Sub

Private Function IsManagedResultTableName(ByVal tableName As String) As Boolean
    IsManagedResultTableName = (Left$(tableName, Len("tblTrinoResult")) = "tblTrinoResult")
End Function

Private Function CreateResultTable(ByVal resultWs As Worksheet, ByVal queryName As String, ByVal tableName As String) As ListObject
    Dim connectionString As String
    Dim lo As ListObject

    resultWs.Range("A1").Value = "Trino Result - " & resultWs.Name
    resultWs.Range("A1").Font.Bold = True
    resultWs.Range("A3").Value = "Результат ниже перезаписывается при следующей выгрузке в этот лист."
    resultWs.Range("A3").Font.Italic = True

    connectionString = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & queryName & ";Extended Properties="""""
    Set lo = resultWs.ListObjects.Add(0, connectionString, False, 1, resultWs.Range("A5"))
    lo.Name = tableName
    On Error Resume Next
    lo.TableStyle = "TableStyleMedium4"
    On Error GoTo 0

    With lo.QueryTable
        .CommandType = 2
        .CommandText = "SELECT * FROM [" & queryName & "]"
        .BackgroundQuery = False
        .RefreshOnFileOpen = False
        .RefreshPeriod = 0
        .EnableRefresh = True
        .SaveData = True
        .PreserveFormatting = True
        .PreserveColumnInfo = True
        .AdjustColumnWidth = True
        .RefreshStyle = 0
    End With

    On Error Resume Next
    lo.QueryTable.WorkbookConnection.RefreshWithRefreshAll = False
    lo.QueryTable.WorkbookConnection.RefreshOnFileOpen = False
    lo.QueryTable.WorkbookConnection.OLEDBConnection.BackgroundQuery = False
    lo.QueryTable.WorkbookConnection.OLEDBConnection.RefreshOnFileOpen = False
    lo.QueryTable.WorkbookConnection.OLEDBConnection.RefreshPeriod = 0
    lo.QueryTable.WorkbookConnection.OLEDBConnection.EnableRefresh = True
    On Error GoTo 0

    Set CreateResultTable = lo
End Function

Private Function StableToken(ByVal value As String) As String
    Dim idx As Long
    Dim code As Long
    Dim hash As Double

    hash = 5381
    For idx = 1 To Len(value)
        code = AscW(Mid$(value, idx, 1))
        If code < 0 Then code = code + 65536
        hash = hash * 33 + code
        hash = hash - Int(hash / 2147483647#) * 2147483647#
    Next idx

    StableToken = "S" & Hex(CLng(hash))
End Function
'''


def _vba_string_literal(value: str) -> str:
    return '"' + value.replace('"', '""') + '"'


def _vba_function_returning_text(function_name: str, text: str) -> str:
    lines = [
        f"Private Function {function_name}() As String",
        "    Dim formula As String",
        '    formula = ""',
        "",
    ]

    for source_line in text.strip().splitlines():
        lines.append(f"    formula = formula & {_vba_string_literal(source_line)} & vbCrLf")

    lines.extend(
        [
            "",
            f"    {function_name} = formula",
            "End Function",
        ]
    )
    return "\n".join(lines)


ADVANCED_VBA_CODE = (
    ADVANCED_VBA_CODE.strip()
    + "\n\n"
    + _vba_function_returning_text("TrinoRestFunctionFormula", M_QUERIES["fnTrinoRestQuery"])
    + "\n"
)
