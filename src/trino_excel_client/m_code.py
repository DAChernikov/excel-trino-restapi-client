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
    SourceName = ToText(GetOption("SourceName", "excel_power_query"), "excel_power_query"),
    TimeZoneName = ToText(GetOption("TimeZone", "Europe/Moscow"), "Europe/Moscow"),
    CatalogName = try GetOption("Catalog", "") otherwise "",
    SchemaName = try GetOption("Schema", "") otherwise "",
    RequestTimeoutMinutes = ToNumber(GetOption("RequestTimeoutMinutes", 5), 5),
    PollingDelaySeconds = ToNumber(GetOption("PollingDelaySeconds", 0.2), 0.2),
    MaxRetryCount = Number.RoundDown(ToNumber(GetOption("MaxRetryCount", 5), 5)),
    RetryBaseDelaySeconds = ToNumber(GetOption("RetryBaseDelaySeconds", 0.5), 0.5),
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
            RawResponse = Web.Contents(baseUrl, EffectiveOptions),
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
                        Content = Text.ToBinary(sqlText, TextEncoding.Utf8),
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
    FirstPage = PostStatement(),
    GeneratedPages =
        List.Generate(
            () => FirstPage,
            each _ <> null,
            each
                if Record.HasFields(_, "nextUri") then
                    Function.InvokeAfter(
                        () => GetNextPage(_[nextUri]),
                        PollingDelay
                    )
                else
                    null,
            each _
        ),
    Pages = List.Buffer(GeneratedPages),
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
    Rows =
        List.Combine(
            List.Transform(
                Pages,
                each if Record.HasFields(_, "data") then _[data] else {}
            )
        ),
    Result =
        if List.Count(ColumnNames) > 0 then
            Table.FromRows(Rows, ColumnNames)
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
                    SourceName = try Config[source_name] otherwise "excel_power_query",
                    TimeZone = try Config[time_zone] otherwise "Europe/Moscow",
                    Catalog = try Config[trino_catalog] otherwise "",
                    Schema = try Config[trino_schema] otherwise "",
                    RequestTimeoutMinutes = try Config[request_timeout_minutes] otherwise 5,
                    PollingDelaySeconds = try Config[polling_delay_seconds] otherwise 0.2,
                    MaxRetryCount = try Config[max_retry_count] otherwise 5,
                    RetryBaseDelaySeconds = try Config[retry_base_delay_seconds] otherwise 0.5
                ]
            )
in
    Result
''',
}
