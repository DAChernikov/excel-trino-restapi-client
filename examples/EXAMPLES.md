# Examples

## Windows: создать рабочую книгу

```bat
trino-excel-client create --backend com --output .\build\Trino_REST_Client.xlsx --overwrite --visible
```

## Windows: встроить клиент в существующую книгу

```bat
trino-excel-client install --backend com --workbook .\input\Analytics.xlsx --output .\build\Analytics_With_Trino.xlsx --overwrite --visible
```

## macOS/Linux: посмотреть UI-шаблон

```bash
trino-excel-client create --backend openpyxl --output ./build/Trino_UI_Preview.xlsx --overwrite
```

## macOS/Linux: открыть GUI в dry-run

```bash
trino-excel-client gui --dry-run
```

## Пример безопасного SQL для Excel

```sql
select
    article_id,
    article_name,
    created_at
from your_catalog.your_schema.your_table
where created_at >= date '2026-01-01'
limit 10000
```

Даже если пользователь забудет `limit`, шаблон по умолчанию добавит внешний `LIMIT 100000`.

## Пример агрегированного SQL

```sql
select
    date_trunc('day', created_at) as day,
    count(*) as rows_count
from your_catalog.your_schema.your_table
where created_at >= current_date - interval '30' day
group by 1
order by 1
```

Для Excel лучше загружать агрегаты, а не сырые BigData-таблицы.

## Пример extraCredentials

| Включено | Тип логина | Логин | Тип пароля | Пароль |
| --- | --- | --- | --- | --- |
| TRUE | gp-user | your_login | gp-password | your_password |

## Windows smoke-test

```bat
python scripts\windows_smoke_test.py --visible
```

## Сборка Windows GUI exe

```bat
pip install -e ".[package]"
python scripts\build_windows_exe.py --clean --target gui
```

## Сборка обоих Windows приложений

```bat
python scripts\build_windows_exe.py --clean
```

Результат:

```text
dist\trino-excel-client-cmd.exe
dist\trino-excel-client-gui.exe
```

`trino-excel-client-gui.exe` - основное приложение с интерфейсом.

`trino-excel-client-cmd.exe` - консольная версия для скриптов.
