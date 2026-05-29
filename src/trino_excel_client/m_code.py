from __future__ import annotations

M_QUERIES: dict[str, str] = {
    "qConfig": r'''
let
    Source = Excel.CurrentWorkbook(){[Name="tblTrinoConfig"]}[Content],
    Columns = Table.ColumnNames(Source),
    Selected = Table.SelectColumns(Source, {Columns{0}, Columns{1}}),
    Canonical = Table.RenameColumns(Selected, {{Columns{0}, "parameter"}, {Columns{1}, "value"}}),
    Textified = Table.TransformColumns(
        Canonical,
        {
            {"parameter", each if _ = null then "" else Text.Trim(Text.From(_)), type text},
            {"value", each if _ = null then "" else Text.From(_), type text}
        }
    ),
    NonEmpty = Table.SelectRows(Textified, each [parameter] <> ""),
    NameValue = Table.RenameColumns(
        Table.SelectColumns(NonEmpty, {"parameter", "value"}),
        {{"parameter", "Name"}, {"value", "Value"}}
    ),
    ConfigRecord = Record.FromTable(NameValue)
in
    ConfigRecord
''',
    "qSqlText": r'''
let
    Source = Excel.CurrentWorkbook(){[Name="tblTrinoSql"]}[Content],
    Columns = Table.ColumnNames(Source),
    Selected = Table.SelectColumns(Source, {Columns{0}}),
    Canonical = Table.RenameColumns(Selected, {{Columns{0}, "sql_text"}}),
    Textified = Table.TransformColumns(
        Canonical,
        {{"sql_text", each if _ = null then "" else Text.From(_), type text}}
    ),
    NonEmpty = Table.SelectRows(Textified, each Text.Trim([sql_text]) <> ""),
    SqlText = Text.Combine(NonEmpty[sql_text], "#(lf)")
in
    SqlText
''',
    "qExtraCredentials": r'''
let
    Source = Excel.CurrentWorkbook(){[Name="tblExtraCredentials"]}[Content],
    Columns = Table.ColumnNames(Source),
    Selected = Table.SelectColumns(Source, List.FirstN(Columns, 5)),
    Canonical = Table.RenameColumns(
        Selected,
        {
            {Columns{0}, "enabled"},
            {Columns{1}, "credential_login_type"},
            {Columns{2}, "credential_login"},
            {Columns{3}, "credential_password_type"},
            {Columns{4}, "credential_password"}
        }
    ),
    IsEnabled = (value as any) as logical =>
        let
            TextValue = Text.Lower(Text.Trim(Text.From(value)))
        in
            List.Contains({"true", "1", "yes", "y", "да", "истина"}, TextValue),
    Textified = Table.TransformColumns(
        Canonical,
        {
            {"enabled", each if _ = null then "" else Text.From(_), type text},
            {"credential_login_type", each if _ = null then "" else Text.Trim(Text.From(_)), type text},
            {"credential_login", each if _ = null then "" else Text.From(_), type text},
            {"credential_password_type", each if _ = null then "" else Text.Trim(Text.From(_)), type text},
            {"credential_password", each if _ = null then "" else Text.From(_), type text}
        }
    ),
    EnabledRows = Table.SelectRows(Textified, each IsEnabled([enabled])),
    Pairs = List.Combine(
        List.Transform(
            Table.ToRecords(EnabledRows),
            each
                let
                    LoginType = Text.Trim(_[credential_login_type]),
                    LoginValue = _[credential_login],
                    PasswordType = Text.Trim(_[credential_password_type]),
                    PasswordValue = _[credential_password],
                    LoginRecord =
                        if LoginType <> "" and LoginValue <> "" then
                            {[credential_name = LoginType, credential_value = LoginValue]}
                        else
                            {},
                    PasswordRecord =
                        if PasswordType <> "" and PasswordValue <> "" then
                            {[credential_name = PasswordType, credential_value = PasswordValue]}
                        else
                            {}
                in
                    List.Combine({LoginRecord, PasswordRecord})
        )
    ),
    Result =
        if List.Count(Pairs) = 0 then
            #table({"credential_name", "credential_value"}, {})
        else
            Table.FromRecords(Pairs)
in
    Result
''',
    "fnTrinoRestQuery": r'''
(
    trinoBaseUrl as text,
    trinoUser as text,
    trinoPassword as text,
    sqlText as text,
    optional extraCredentialsTable as nullable table,
    optional options as nullable record
) as table =>
let
    GetOption = (fieldName as text, defaultValue as any) as any =>
        if options <> null and Record.HasFields(options, fieldName) then
            Record.Field(options, fieldName)
        else
            defaultValue,
    ToNumber = (value as any, defaultValue as number) as number =>
        try Number.From(value) otherwise defaultValue,
    ToText = (value as any, defaultValue as text) as text =>
        let
            TextValue = try Text.From(value) otherwise defaultValue
        in
            if Text.Trim(TextValue) = "" then defaultValue else TextValue,
    OptionalHeader = (headerName as text, headerValue as any) as record =>
        let
            TextValue = try Text.Trim(Text.From(headerValue)) otherwise ""
        in
            if TextValue = "" then [] else Record.FromList({TextValue}, {headerName}),
    NormalizedBaseUrl =
        if Text.EndsWith(trinoBaseUrl, "/") then
            Text.Start(trinoBaseUrl, Text.Length(trinoBaseUrl) - 1)
        else
            trinoBaseUrl,
    SourceName = "Microsoft Excel PowerQuery client",
    TimeZoneName = ToText(GetOption("TimeZone", "Europe/Moscow"), "Europe/Moscow"),
    CatalogName = try GetOption("Catalog", "") otherwise "",
    SchemaName = try GetOption("Schema", "") otherwise "",
    RequestTimeoutMinutes = ToNumber(GetOption("RequestTimeoutMinutes", 5), 5),
    PollingDelaySeconds = 0.2,
    MaxRetryCount = 5,
    RetryBaseDelaySeconds = 0.5,
    ResultLimitRows = Number.RoundDown(ToNumber(GetOption("ResultLimitRows", 1000000), 1000000)),
    NormalizeSqlForLimit = (query as text) as text =>
        let
            Trimmed = Text.Trim(query),
            WithoutTrailingSemicolon =
                if Text.EndsWith(Trimmed, ";") then
                    Text.Start(Trimmed, Text.Length(Trimmed) - 1)
                else
                    Trimmed
        in
            WithoutTrailingSemicolon,
    EffectiveSqlText =
        if ResultLimitRows > 0 then
            "select * from ("
                & "#(lf)"
                & NormalizeSqlForLimit(sqlText)
                & "#(lf)"
                & ") as excel_trino_client_query limit "
                & Text.From(ResultLimitRows)
        else
            sqlText,
    RequestTimeout = #duration(0, 0, RequestTimeoutMinutes, 0),
    PollingDelay = #duration(0, 0, 0, PollingDelaySeconds),
    AuthHeader =
        "Basic "
        & Binary.ToText(
            Text.ToBinary(trinoUser & ":" & trinoPassword, TextEncoding.Utf8),
            BinaryEncoding.Base64
        ),
    ExtraCredentialRows =
        if extraCredentialsTable = null then
            {}
        else
            Table.ToRecords(extraCredentialsTable),
    ExtraCredentialHeader =
        if List.Count(ExtraCredentialRows) = 0 then
            null
        else
            Text.Combine(
                List.Transform(
                    ExtraCredentialRows,
                    each Text.From(_[credential_name])
                        & "="
                        & Uri.EscapeDataString(Text.From(_[credential_value]))
                ),
                ","
            ),
    ExtraCredentialsHeaderRecord =
        if ExtraCredentialHeader = null or Text.Trim(ExtraCredentialHeader) = "" then
            []
        else
            [#"X-Trino-Extra-Credential" = ExtraCredentialHeader],
    OptionalContextHeaders = Record.Combine({
        OptionalHeader("X-Trino-Catalog", CatalogName),
        OptionalHeader("X-Trino-Schema", SchemaName)
    }),
    PostHeaders = Record.Combine({
        [
            #"Authorization" = AuthHeader,
            #"X-Trino-User" = trinoUser,
            #"X-Trino-Source" = SourceName,
            #"X-Trino-Time-Zone" = TimeZoneName,
            #"Content-Type" = "text/plain"
        ],
        OptionalContextHeaders,
        ExtraCredentialsHeaderRecord
    }),
    GetHeaders = Record.Combine({
        [
            #"Authorization" = AuthHeader,
            #"X-Trino-User" = trinoUser,
            #"X-Trino-Source" = SourceName
        ],
        ExtraCredentialsHeaderRecord
    }),
    ManualStatuses = {400, 401, 403, 404, 410, 429, 500, 502, 503, 504},
    RetryableStatuses = {429, 502, 503, 504},
    WebRequestWithRetry =
        (baseUrl as text, requestOptions as record, optional attempt as nullable number) as record =>
        let
            CurrentAttempt = if attempt = null then 0 else attempt,
            EffectiveOptions = Record.Combine({
                requestOptions,
                [
                    IsRetry = CurrentAttempt > 0,
                    ManualStatusHandling = ManualStatuses
                ]
            }),
            RawResponseAttempt = try Web.Contents(baseUrl, EffectiveOptions),
            RawResponse =
                if RawResponseAttempt[HasError] then
                    error Error.Record(
                        "Trino connection lost",
                        "Потерялась связь с Trino во время выполнения или получения результата. Попробуйте обновить запрос еще раз. Если ошибка повторяется, уменьшите объем результата: добавьте LIMIT, фильтры, выберите меньше колонок или используйте агрегаты.",
                        [
                            reason = try RawResponseAttempt[Error][Message] otherwise "Power Query не получил ответ от Trino",
                            result_limit_rows = ResultLimitRows
                        ]
                    )
                else
                    RawResponseAttempt[Value],
            Status = try Value.Metadata(RawResponse)[Response.Status] otherwise 200,
            BufferedBody = Binary.Buffer(RawResponse),
            ShouldRetry =
                List.Contains(RetryableStatuses, Status)
                and CurrentAttempt < MaxRetryCount,
            Delay =
                #duration(
                    0,
                    0,
                    0,
                    RetryBaseDelaySeconds * Number.Power(2, CurrentAttempt)
                ),
            Result =
                if ShouldRetry then
                    Function.InvokeAfter(
                        () => @WebRequestWithRetry(baseUrl, requestOptions, CurrentAttempt + 1),
                        Delay
                    )
                else
                    [Status = Status, Body = BufferedBody]
        in
            Result,
    ParseResponse = (Response as record) as record =>
        let
            Status = Response[Status],
            Body = Response[Body],
            BodyText = try Text.FromBinary(Body, TextEncoding.Utf8) otherwise "",
            BodyJson =
                if Status = 200 then
                    Json.Document(Body)
                else
                    error Error.Record(
                        "Trino HTTP error",
                        "HTTP status: " & Text.From(Status),
                        BodyText
                    ),
            CheckedJson =
                if Record.HasFields(BodyJson, "error") then
                    error Error.Record(
                        "Trino query error",
                        BodyJson[error][message],
                        BodyJson[error]
                    )
                else
                    BodyJson
        in
            CheckedJson,
    PostStatement = () as record =>
        let
            Response =
                WebRequestWithRetry(
                    NormalizedBaseUrl,
                    [
                        RelativePath = "v1/statement",
                        Headers = PostHeaders,
                        Content = Text.ToBinary(EffectiveSqlText, TextEncoding.Utf8),
                        Timeout = RequestTimeout
                    ]
                )
        in
            ParseResponse(Response),
    RelativePathFromNextUri = (nextUri as text) as text =>
        let
            Parts = Uri.Parts(nextUri),
            PathParts = List.Select(Text.Split(Parts[Path], "/"), each _ <> ""),
            RelativePath = Text.Combine(PathParts, "/")
        in
            RelativePath,
    GetNextPage = (nextUri as text) as record =>
        let
            RelativePath = RelativePathFromNextUri(nextUri),
            Response =
                WebRequestWithRetry(
                    NormalizedBaseUrl,
                    [
                        RelativePath = RelativePath,
                        Headers = GetHeaders,
                        Timeout = RequestTimeout
                    ]
                )
        in
            ParseResponse(Response),
    PageRowCount = (page as nullable record) as number =>
        if page <> null and Record.HasFields(page, "data") then
            List.Count(page[data])
        else
            0,
    FirstPage = PostStatement(),
    FirstRowCount = PageRowCount(FirstPage),
    FirstOverflow = ResultLimitRows > 0 and FirstRowCount > ResultLimitRows,
    GeneratedStates =
        List.Generate(
            () => [Page = FirstPage, RowCount = FirstRowCount, Overflow = FirstOverflow],
            each [Page] <> null,
            each
                let
                    CurrentPage = [Page],
                    CurrentRowCount = [RowCount],
                    CurrentOverflow = [Overflow],
                    HasNext = CurrentPage <> null and Record.HasFields(CurrentPage, "nextUri"),
                    ShouldFetchNext = HasNext and not CurrentOverflow,
                    NextPage =
                        if ShouldFetchNext then
                            Function.InvokeAfter(
                                () => GetNextPage(CurrentPage[nextUri]),
                                PollingDelay
                            )
                        else
                            null,
                    NextRowCount = PageRowCount(NextPage),
                    TotalRowCount = CurrentRowCount + NextRowCount,
                    NextOverflow = ResultLimitRows > 0 and TotalRowCount > ResultLimitRows
                in
                    [Page = NextPage, RowCount = TotalRowCount, Overflow = NextOverflow],
            each _
        ),
    States = List.Buffer(GeneratedStates),
    Pages = List.Buffer(List.Transform(States, each [Page])),
    TotalRowsSeen =
        if List.Count(States) = 0 then
            0
        else
            List.Last(States)[RowCount],
    IsOverflow = List.AnyTrue(List.Transform(States, each [Overflow])),
    PageWithColumns =
        List.First(
            List.Select(Pages, each Record.HasFields(_, "columns")),
            null
        ),
    ColumnNames =
        if PageWithColumns <> null then
            List.Transform(PageWithColumns[columns], each _[name])
        else
            {},
    NormalizeTrinoTemporalText = (value as any) as nullable text =>
        let
            TextValue = if value = null then null else Text.Trim(Text.From(value)),
            Normalized =
                if TextValue = null or TextValue = "" then
                    null
                else if Text.Length(TextValue) > 10 and Text.Range(TextValue, 10, 1) = " " then
                    Text.Start(TextValue, 10) & "T" & Text.Trim(Text.Range(TextValue, 10))
                else
                    TextValue
        in
            Normalized,
    ParseTrinoTimestamp = (value as any) as nullable datetime =>
        let
            Normalized = NormalizeTrinoTemporalText(value),
            WithoutZone =
                if Normalized = null then
                    null
                else if Text.Contains(Normalized, " ") then
                    Text.BeforeDelimiter(Normalized, " ")
                else
                    Normalized,
            Parsed =
                if WithoutZone = null then
                    null
                else
                    try DateTime.FromText(WithoutZone) otherwise try DateTime.From(value) otherwise null
        in
            Parsed,
    ParseTrinoTimestampZone = (value as any) as nullable datetimezone =>
        let
            Normalized = NormalizeTrinoTemporalText(value),
            UtcNormalized =
                if Normalized = null then
                    null
                else if Text.EndsWith(Normalized, " UTC") then
                    Text.BeforeDelimiter(Normalized, " UTC") & "Z"
                else
                    Normalized,
            Parsed =
                if UtcNormalized = null then
                    null
                else
                    try DateTimeZone.FromText(UtcNormalized)
                    otherwise try DateTime.AddZone(ParseTrinoTimestamp(value), 0)
                    otherwise null
        in
            Parsed,
    ParseTrinoDate = (value as any) as nullable date =>
        let
            Normalized = NormalizeTrinoTemporalText(value),
            DateText =
                if Normalized = null then
                    null
                else if Text.Contains(Normalized, "T") then
                    Text.BeforeDelimiter(Normalized, "T")
                else if Text.Contains(Normalized, " ") then
                    Text.BeforeDelimiter(Normalized, " ")
                else
                    Normalized,
            Parsed =
                if DateText = null then
                    null
                else
                    try Date.FromText(DateText) otherwise try Date.From(value) otherwise null
        in
            Parsed,
    ParseTrinoTime = (value as any) as nullable time =>
        let
            TextValue = if value = null then null else Text.Trim(Text.From(value)),
            TimeText =
                if TextValue = null or TextValue = "" then
                    null
                else if Text.Contains(TextValue, "T") then
                    Text.AfterDelimiter(TextValue, "T")
                else if Text.Length(TextValue) > 10 and Text.Range(TextValue, 10, 1) = " " then
                    Text.Trim(Text.Range(TextValue, 10))
                else
                    TextValue,
            WithoutZone =
                if TimeText = null then
                    null
                else if Text.Contains(TimeText, " ") then
                    Text.BeforeDelimiter(TimeText, " ")
                else
                    TimeText,
            Parsed =
                if WithoutZone = null then
                    null
                else
                    try Time.FromText(WithoutZone) otherwise try Time.From(value) otherwise null
        in
            Parsed,
    TrinoTypeToPowerQueryType = (trinoType as nullable text) as any =>
        let
            NormalizedType = if trinoType = null then "" else Text.Lower(Text.Trim(trinoType)),
            BaseType = if Text.Contains(NormalizedType, "(") then Text.BeforeDelimiter(NormalizedType, "(") else NormalizedType,
            ResultType =
                if List.Contains({"tinyint", "smallint", "integer", "bigint"}, BaseType) then
                    Int64.Type
                else if List.Contains({"real", "double", "decimal"}, BaseType) then
                    type number
                else if BaseType = "boolean" then
                    type logical
                else if BaseType = "date" then
                    null
                else if Text.StartsWith(NormalizedType, "timestamp") then
                    null
                else if Text.StartsWith(NormalizedType, "time") then
                    null
                else if List.Contains({"varchar", "char", "json", "uuid", "ipaddress"}, BaseType) then
                    type text
                else
                    null
        in
            ResultType,
    ColumnTypePairs =
        if PageWithColumns <> null then
            List.Transform(
                PageWithColumns[columns],
                each
                    let
                        TrinoType = try _[type] otherwise null,
                        NormalizedType = if TrinoType = null then "" else Text.Lower(Text.Trim(Text.From(TrinoType)))
                    in
                        {_[name], TrinoTypeToPowerQueryType(TrinoType), NormalizedType}
            )
        else
            {},
    AllRows =
        List.Combine(
            List.Transform(
                Pages,
                each if Record.HasFields(_, "data") then _[data] else {}
            )
        ),
    Rows =
        if ResultLimitRows > 0 then
            List.FirstN(AllRows, ResultLimitRows)
        else
            AllRows,
    OverflowMessage =
        "Trino returned more rows than result_limit_rows="
        & Text.From(ResultLimitRows)
        & ". Excel cannot safely load more than about one million rows. "
        & "Reduce the result size with LIMIT, filters, fewer columns or aggregations, then refresh the query again. "
        & "Rows seen before stop: "
        & Text.From(TotalRowsSeen),
    Result =
        if IsOverflow then
            error Error.Record(
                "Trino result row limit exceeded",
                OverflowMessage,
                [
                    result_limit_rows = ResultLimitRows,
                    rows_seen_before_stop = TotalRowsSeen
                ]
            )
        else if List.Count(ColumnNames) > 0 then
            let
                RawTable = Table.FromRows(Rows, ColumnNames),
                DateColumnTransforms =
                    List.Transform(
                        List.Select(ColumnTypePairs, each _{2} = "date"),
                        (columnSpec) => {columnSpec{0}, each ParseTrinoDate(_), type date}
                    ),
                TimestampZoneColumnTransforms =
                    List.Transform(
                        List.Select(
                            ColumnTypePairs,
                            each Text.StartsWith(_{2}, "timestamp") and Text.Contains(_{2}, "with time zone")
                        ),
                        (columnSpec) => {columnSpec{0}, each ParseTrinoTimestampZone(_), type datetimezone}
                    ),
                TimestampColumnTransforms =
                    List.Transform(
                        List.Select(
                            ColumnTypePairs,
                            each Text.StartsWith(_{2}, "timestamp") and not Text.Contains(_{2}, "with time zone")
                        ),
                        (columnSpec) => {columnSpec{0}, each ParseTrinoTimestamp(_), type datetime}
                    ),
                TimeColumnTransforms =
                    List.Transform(
                        List.Select(
                            ColumnTypePairs,
                            each Text.StartsWith(_{2}, "time")
                        ),
                        (columnSpec) => {columnSpec{0}, each ParseTrinoTime(_), type time}
                    ),
                TemporalColumnTransforms = List.Combine({
                    DateColumnTransforms,
                    TimestampZoneColumnTransforms,
                    TimestampColumnTransforms,
                    TimeColumnTransforms
                }),
                TemporalTypedTable =
                    if List.Count(TemporalColumnTransforms) = 0 then
                        RawTable
                    else
                        Table.TransformColumns(
                            RawTable,
                            TemporalColumnTransforms,
                            null,
                            MissingField.Ignore
                        ),
                TypedColumns =
                    List.Transform(
                        List.Select(ColumnTypePairs, each _{1} <> null),
                        each {_{0}, _{1}}
                    ),
                TypedTable =
                    if List.Count(TypedColumns) = 0 then
                        TemporalTypedTable
                    else
                        Table.TransformColumnTypes(TemporalTypedTable, TypedColumns, "en-US")
            in
                TypedTable
        else
            #table({}, {})
in
    Result
''',
    "TrinoResult": r'''
let
    Config = qConfig,
    RequiredFields = {"trino_base_url", "trino_user", "trino_password"},
    MissingFields = List.Select(
        RequiredFields,
        each not Record.HasFields(Config, _) or Text.Trim(Text.From(Record.Field(Config, _))) = ""
    ),
    Result =
        if List.Count(MissingFields) > 0 then
            error Error.Record(
                "Trino configuration error",
                "Fill required Config parameters: " & Text.Combine(MissingFields, ", "),
                MissingFields
            )
        else if Text.Trim(qSqlText) = "" then
            error Error.Record(
                "Trino SQL error",
                "Fill SQL text in tblTrinoSql.",
                null
            )
        else
            fnTrinoRestQuery(
                Config[trino_base_url],
                Config[trino_user],
                Config[trino_password],
                qSqlText,
                qExtraCredentials,
                [
                    TimeZone = try Config[time_zone] otherwise "Europe/Moscow",
                    Catalog = try Config[trino_catalog] otherwise "",
                    Schema = try Config[trino_schema] otherwise "",
                    RequestTimeoutMinutes = try Config[request_timeout_minutes] otherwise 5,
                    ResultLimitRows = try Config[result_limit_rows] otherwise 1000000
                ]
            )
in
    Result
''',
}
