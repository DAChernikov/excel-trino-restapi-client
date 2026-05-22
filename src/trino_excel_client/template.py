from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class TableSpec:
    name: str
    headers: tuple[str, ...]
    rows: tuple[tuple[object, ...], ...]
    style: str


@dataclass(frozen=True)
class ClientSheetNames:
    config: str
    query: str
    result: str
    help: str

    @property
    def all(self) -> tuple[str, str, str, str]:
        return (self.config, self.query, self.result, self.help)


CONFIG_TABLE = TableSpec(
    name="tblTrinoConfig",
    headers=("Параметр", "Значение", "Обязателен", "Комментарий"),
    rows=(
        ("trino_base_url", "https://trino.example.com", "да", "URL координатора Trino без хвоста /v1/statement"),
        ("trino_user", "CHANGE_ME", "да", "Пользователь Basic Auth и заголовок X-Trino-User"),
        ("trino_password", "CHANGE_ME", "да", "Пароль Basic Auth"),
        ("trino_catalog", "", "нет", "Необязательный заголовок X-Trino-Catalog"),
        ("trino_schema", "", "нет", "Необязательный заголовок X-Trino-Schema"),
        ("source_name", "Microsoft Excel PowerQuery client", "нет", "Значение заголовка X-Trino-Source"),
        ("time_zone", "Europe/Moscow", "нет", "Значение заголовка X-Trino-Time-Zone"),
        ("request_timeout_minutes", "5", "нет", "Таймаут одного HTTP-запроса в минутах"),
        ("polling_delay_seconds", "0.2", "нет", "Пауза между запросами Trino nextUri"),
        ("max_retry_count", "5", "нет", "Количество повторов для 429/502/503/504"),
        ("retry_base_delay_seconds", "0.5", "нет", "Базовая задержка экспоненциального retry"),
        ("default_query_limit_rows", "100000", "нет", "Внешний LIMIT, который по умолчанию добавляется к SQL"),
        ("apply_default_query_limit", "TRUE", "нет", "TRUE добавляет default_query_limit_rows к SQL как защиту от больших выгрузок"),
        ("max_result_rows", "100000", "нет", "Максимум строк, которые разрешено загрузить в Excel"),
        ("result_overflow_behavior", "error", "нет", "Что делать при превышении лимита: error или truncate"),
    ),
    style="TableStyleMedium2",
)

EXTRA_CREDENTIALS_TABLE = TableSpec(
    name="tblExtraCredentials",
    headers=(
        "Включено",
        "Тип логина",
        "Логин",
        "Тип пароля",
        "Пароль",
        "Комментарий",
    ),
    rows=(
        ("TRUE", "gp-user", "CHANGE_ME", "gp-password", "CHANGE_ME", "Пример для Greenplum"),
        ("FALSE", "", "", "", "", "Запасная строка"),
        ("FALSE", "", "", "", "", "Запасная строка"),
    ),
    style="TableStyleMedium4",
)

SQL_TABLE = TableSpec(
    name="tblTrinoSql",
    headers=("SQL запрос",),
    rows=(
        ("select * from your_catalog.your_schema.your_table limit 100",),
    ),
    style="TableStyleMedium9",
)

HELP_TABLE = TableSpec(
    name="tblTrinoHelp",
    headers=("Шаг", "Раздел", "Подсказка"),
    rows=(
        ("1", "Заполнить Config", "Откройте лист Trino Config и заполните trino_base_url, trino_user, trino_password. Catalog/schema можно оставить пустыми."),
        ("2", "Заполнить extra credentials", "Если источник данных требует delegated credentials, включите строку через TRUE и заполните пары тип/значение."),
        ("3", "Написать SQL", "Откройте лист Trino Query и введите SQL в таблицу tblTrinoSql. Результат загружается отдельно на лист Trino Result."),
        ("4", "Обновить данные", "Нажмите Данные -> Обновить все. При первом обращении Excel может запросить настройки доступа к Trino host."),
        ("5", "Окно credentials", "В Power Query credential dialog обычно нужно выбрать Anonymous, потому что Basic Auth передается M-кодом через HTTP headers."),
        ("6", "Default LIMIT", "По умолчанию M-код добавляет внешний LIMIT из default_query_limit_rows. Это можно отключить через apply_default_query_limit = FALSE."),
        ("7", "Лимит строк", "Параметр max_result_rows защищает Excel от случайной выгрузки миллионов строк. Для BigData используйте LIMIT, агрегаты или выгрузку во внешнее хранилище."),
        ("8", "Безопасность", "Скрытие парольных ячеек не является шифрованием. Храните книгу так, как если бы внутри были реальные учетные данные."),
    ),
    style="TableStyleMedium2",
)

REQUIRED_QUERY_NAMES = (
    "qConfig",
    "qSqlText",
    "qExtraCredentials",
    "fnTrinoRestQuery",
    "TrinoResult",
    "TrinoResultSchema",
)

REQUIRED_TABLE_NAMES = (
    CONFIG_TABLE.name,
    EXTRA_CREDENTIALS_TABLE.name,
    SQL_TABLE.name,
)


def sheet_name(prefix: str, suffix: str) -> str:
    base = f"{prefix.strip()} {suffix}".strip()
    return base[:31]


def client_sheet_names(sheet_prefix: str = "Trino") -> ClientSheetNames:
    return ClientSheetNames(
        config=sheet_name(sheet_prefix, "Config"),
        query=sheet_name(sheet_prefix, "Query"),
        result=sheet_name(sheet_prefix, "Result"),
        help=sheet_name(sheet_prefix, "Help"),
    )
