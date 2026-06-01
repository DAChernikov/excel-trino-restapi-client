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
        ("trino_base_url", "https://trino.mlops.isb", "да", "URL координатора Trino без хвоста /v1/statement"),
        ("trino_user", "CHANGE_ME", "да", "Пользователь Basic Auth и заголовок X-Trino-User"),
        ("trino_password", "CHANGE_ME", "да", "Пароль Basic Auth"),
        ("trino_catalog", "", "нет", "Необязательный заголовок X-Trino-Catalog"),
        ("trino_schema", "", "нет", "Необязательный заголовок X-Trino-Schema"),
        ("time_zone", "Europe/Moscow", "нет", "Значение заголовка X-Trino-Time-Zone"),
        ("request_timeout_minutes", "5", "нет", "Таймаут одного HTTP-запроса в минутах"),
        ("result_limit_rows", "1000000", "нет", "Единый LIMIT для SQL и максимум строк, разрешенных для загрузки в Excel"),
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

ADVANCED_TARGET_TABLE = TableSpec(
    name="tblTrinoAdvancedTarget",
    headers=("Параметр", "Значение", "Комментарий"),
    rows=(
        ("target_sheet", "Trino Result", "Лист, куда будет создана или обновлена таблица результата"),
    ),
    style="TableStyleMedium4",
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
        ("5a", "Если доступ запрещен", "Если появляется ResourceAccessForbiddenException, откройте Данные -> Получить данные -> Параметры источника данных, очистите старые разрешения для Trino host и выберите Anonymous при следующем обновлении."),
        ("6", "Лимит строк", "Параметр result_limit_rows по умолчанию равен 1000000. M-код добавляет внешний LIMIT и не дает загрузить в Excel больше этого количества строк."),
        ("7", "Большие результаты", "Если связь с Trino оборвалась во время выдачи результата, повторите запрос и уменьшите результат через LIMIT, фильтры, выбор колонок или агрегаты."),
        ("8", "Advanced .xlsm", "В расширенном шаблоне укажите лист результата на Trino Query и нажмите кнопку Выгрузить в лист. Не используйте Данные -> Обновить все как основной сценарий advanced-выгрузки."),
        ("9", "Безопасность", "Скрытие парольных ячеек не является шифрованием. Храните книгу так, как если бы внутри были реальные учетные данные."),
    ),
    style="TableStyleMedium2",
)

REQUIRED_QUERY_NAMES = (
    "qConfig",
    "qSqlText",
    "qExtraCredentials",
    "fnTrinoRestQuery",
    "TrinoResult",
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
