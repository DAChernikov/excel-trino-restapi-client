# Trino Excel Client

Python-библиотека для генерации Excel-шаблона с Power Query клиентом к Trino REST API.

Проект больше не содержит отдельное приложение, установщик, GUI или GitHub Actions. Основной артефакт для пользователей - готовый Excel-файл в `templates/`, который разработчик собирает локально на Windows с установленным Microsoft Excel и коммитит в репозиторий.

## Оглавление

- [Для пользователей](#для-пользователей)
  - [Что скачать](#что-скачать)
  - [Как работать в Excel](#как-работать-в-excel)
  - [Лимиты и BigData](#лимиты-и-bigdata)
  - [Безопасность](#безопасность)
- [Для разработчиков](#для-разработчиков)
  - [Требования](#требования)
  - [Установка](#установка)
  - [Сборка готового шаблона](#сборка-готового-шаблона)
  - [Python API](#python-api)
  - [Встраивание в существующую книгу](#встраивание-в-существующую-книгу)
  - [Проверки](#проверки)
  - [Структура проекта](#структура-проекта)

## Для пользователей

### Что скачать

Скачайте готовый шаблон:

```text
templates/Trino_REST_Client.xlsx
```

Если файла нет или он устарел, попросите разработчика пересобрать шаблон на Windows через Excel COM и закоммитить новый `.xlsx` в `templates/`.

### Как работать в Excel

1. Откройте `Trino_REST_Client.xlsx`.
2. На листе `Trino Config` заполните:
   - `trino_base_url`
   - `trino_user`
   - `trino_password`
   - при необходимости `trino_catalog` и `trino_schema`
3. Если источник требует delegated credentials, заполните `tblExtraCredentials`.
4. На листе `Trino Query` введите SQL в таблицу `tblTrinoSql`.
5. Нажмите `Данные -> Обновить все`.
6. Результат появится на листе `Trino Result`.

При первом обращении Excel может спросить credentials/privacy level для Trino host. Обычно нужно выбрать `Anonymous`, потому что Basic Auth передается M-кодом через HTTP header.

Клиент настроен так, чтобы не отправлять запросы в Trino при открытии книги или при редактировании SQL. Запрос должен уходить только после пользовательского refresh-действия в Excel.

### Лимиты и BigData

Excel не подходит для прямой загрузки миллионов или миллиардов строк. В шаблоне есть защитные параметры:

| Параметр | По умолчанию | Что делает |
| --- | --- | --- |
| `default_query_limit_rows` | `100000` | M-код добавляет внешний `LIMIT` вокруг SQL |
| `apply_default_query_limit` | `TRUE` | Включает или отключает внешний `LIMIT` |
| `max_result_rows` | `100000` | Жесткий лимит строк, разрешенных для загрузки в Excel |
| `result_overflow_behavior` | `error` | `error` останавливает загрузку, `truncate` загружает первые строки |

По умолчанию SQL отправляется в Trino так:

```sql
select *
from (
    <ваш SQL>
) as excel_trino_client_query
limit 100000
```

Рекомендации:

- выбирайте только нужные колонки;
- используйте фильтры и агрегаты;
- не запускайте `select *` по большим таблицам без понимания объема результата;
- для больших выгрузок используйте внешние export pipeline, parquet/csv или BI/ETL-инструменты.

### Безопасность

Парольные ячейки визуально скрыты форматом Excel, но это не шифрование. Книга содержит учетные данные, если они были заполнены.

Не пересылайте заполненные credentials без необходимости и храните такие книги как файлы с чувствительными данными.

## Для разработчиков

### Требования

Для полноценной сборки шаблона с Power Query:

- Windows
- Desktop Microsoft Excel 2016+ с Power Query
- Python 3.10+
- `pywin32`

На macOS/Linux можно собрать только UI-preview через `openpyxl`; Power Query COM-встраивание доступно только в desktop Excel на Windows.

### Установка

Из репозитория:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -e ".[test]"
```

Windows PowerShell / cmd:

```bat
python -m venv .venv
.venv\Scripts\activate
pip install -e ".[test]"
```

После установки доступна команда:

```bash
trino-excel-client --help
```

### Сборка готового шаблона

На Windows с установленным Excel:

```bat
trino-excel-client create --backend com --output templates\Trino_REST_Client.xlsx --overwrite --visible
```

Без видимого Excel:

```bat
trino-excel-client create --backend com --output templates\Trino_REST_Client.xlsx --overwrite
```

Эта команда:

1. создает стабильный UI через `openpyxl`;
2. открывает книгу через Excel COM;
3. добавляет Power Query M-код;
4. создает таблицу результата на листе `Trino Result`;
5. сохраняет готовый шаблон.

После ручной проверки шаблон нужно закоммитить:

```bat
git add templates\Trino_REST_Client.xlsx
git commit -m "Update Trino Excel template"
```

GitHub Actions для сборки шаблона намеренно не используется: полноценное COM-встраивание Power Query требует desktop Microsoft Excel, поэтому актуальный шаблон собирается разработчиком локально на Windows.

### Python API

Собрать полноценный шаблон через Windows Excel COM:

```python
from pathlib import Path

from trino_excel_client import create_workbook

create_workbook(
    output_path=Path("templates/Trino_REST_Client.xlsx"),
    overwrite=True,
    visible=False,
)
```

Собрать UI-only preview без Power Query:

```python
from pathlib import Path

from trino_excel_client import create_workbook_ui

create_workbook_ui(
    output_path=Path("build/Trino_UI_Preview.xlsx"),
    overwrite=True,
)
```

### Встраивание в существующую книгу

Через CLI:

```bat
trino-excel-client install --backend com --workbook .\reports\Workbook.xlsx --output .\build\Workbook_With_Trino.xlsx --overwrite
```

Через Python API:

```python
from pathlib import Path

from trino_excel_client import install_client

install_client(
    workbook_path=Path("reports/Workbook.xlsx"),
    output_path=Path("build/Workbook_With_Trino.xlsx"),
    overwrite=True,
)
```

### Проверки

```bash
trino-excel-client validate
pytest
```

После сборки шаблона откройте `templates\Trino_REST_Client.xlsx`, заполните тестовый `Config`, выполните небольшой SQL и нажмите `Данные -> Обновить все` несколько раз. Лист `Trino Query` не должен изменяться, результат должен обновляться на `Trino Result`.

### Структура проекта

```text
src/trino_excel_client/openpyxl_backend.py  # UI Excel-шаблона
src/trino_excel_client/excel_com.py         # Power Query через Windows Excel COM
src/trino_excel_client/m_code.py            # M-код запросов
src/trino_excel_client/template.py          # таблицы, подписи, дефолты
src/trino_excel_client/cli.py               # CLI для разработчика
templates/                                  # готовые Excel-шаблоны для пользователей
```
