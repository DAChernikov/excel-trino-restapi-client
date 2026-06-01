from __future__ import annotations

ADVANCED_VBA_MODULE_NAME = "TrinoAdvancedClient"

ADVANCED_VBA_CODE = r'''
Option Explicit

Private Const TARGET_TABLE_NAME As String = "tblTrinoAdvancedTarget"
Private Const SQL_TABLE_NAME As String = "tblTrinoSql"

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

    message = "Не удалось выполнить выгрузку Trino." & vbCrLf & details

    If InStr(1, details, "ResourceAccessForbidden", vbTextCompare) > 0 Then
        message = message & vbCrLf & vbCrLf
        message = message & "Excel заблокировал доступ Power Query к Trino host на этом ПК." & vbCrLf
        message = message & "Откройте: Данные -> Получить данные -> Параметры источника данных." & vbCrLf
        message = message & "Очистите старые разрешения для Trino host и при следующем запросе выберите Anonymous."
    End If

    BuildFriendlyErrorMessage = message
End Function

Private Function ReadAdvancedValue(ByVal key As String) As String
    Dim lo As ListObject
    Dim row As ListRow
    Dim currentKey As String

    Set lo = FindListObject(TARGET_TABLE_NAME)
    For Each row In lo.ListRows
        currentKey = LCase(Trim(CStr(row.Range.Cells(1, 1).Value)))
        If currentKey = LCase(key) Then
            ReadAdvancedValue = CStr(row.Range.Cells(1, 2).Value)
            Exit Function
        End If
    Next row

    ReadAdvancedValue = ""
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

    formula = "let" & vbCrLf
    formula = formula & "    Config = qConfig," & vbCrLf
    formula = formula & "    SqlText = " & MTextLiteral(sqlText) & "," & vbCrLf
    formula = formula & "    RequiredFields = {""trino_base_url"", ""trino_user"", ""trino_password""}," & vbCrLf
    formula = formula & "    MissingFields = List.Select(RequiredFields, each not Record.HasFields(Config, _) or Text.Trim(Text.From(Record.Field(Config, _))) = """")," & vbCrLf
    formula = formula & "    Result =" & vbCrLf
    formula = formula & "        if List.Count(MissingFields) > 0 then" & vbCrLf
    formula = formula & "            error Error.Record(""Trino configuration error"", ""Fill required Config parameters: "" & Text.Combine(MissingFields, "", ""), MissingFields)" & vbCrLf
    formula = formula & "        else" & vbCrLf
    formula = formula & "            fnTrinoRestQuery(" & vbCrLf
    formula = formula & "                Config[trino_base_url]," & vbCrLf
    formula = formula & "                Config[trino_user]," & vbCrLf
    formula = formula & "                Config[trino_password]," & vbCrLf
    formula = formula & "                SqlText," & vbCrLf
    formula = formula & "                qExtraCredentials," & vbCrLf
    formula = formula & "                [" & vbCrLf
    formula = formula & "                    TimeZone = try Config[time_zone] otherwise ""Europe/Moscow""," & vbCrLf
    formula = formula & "                    Catalog = try Config[trino_catalog] otherwise """"," & vbCrLf
    formula = formula & "                    Schema = try Config[trino_schema] otherwise """"," & vbCrLf
    formula = formula & "                    RequestTimeoutMinutes = try Config[request_timeout_minutes] otherwise 5," & vbCrLf
    formula = formula & "                    ResultLimitRows = try Config[result_limit_rows] otherwise 1000000" & vbCrLf
    formula = formula & "                ]" & vbCrLf
    formula = formula & "            )" & vbCrLf
    formula = formula & "in" & vbCrLf
    formula = formula & "    Result"

    BuildResultFormula = formula
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
