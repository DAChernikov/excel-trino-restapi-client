# Trino Excel Client

Python-генератор Excel-клиента для подключения к Trino REST API через Power Query M-код.

Цель проекта: дать пользователям Excel готовую книгу или встраиваемый модуль, где подключение к Trino настраивается через понятные таблицы, SQL пишется прямо в Excel, а результат загружается через `Данные -> Обновить все`.

## Архитектура проекта

Проект разделен на два уровня:

- Cross-platform уровень: модель шаблона, M-код, валидация контракта, тесты и `openpyxl` backend для создания UI-книги на macOS/Linux/Windows.
- Windows Excel уровень: COM backend через `pywin32`, который открывает уже созданную `openpyxl`-книгу и встраивает реальные Power Query-запросы через Excel API.

Единый источник UI - `openpyxl_backend.py`. COM не рисует интерфейс, не форматирует листы и не создает config/query/help таблицы. Он делает только то, чего не умеет `openpyxl`: `Workbook.Queries.Add(...)`, подключение `Microsoft.Mashup.OleDb.1` и попытку создать таблицу результата `tblTrinoResult`.

Это значит, что `trino-excel-client create --backend openpyxl ...` и `trino-excel-client create --backend com ...` используют один и тот же UI. Во втором случае в книгу дополнительно добавляются Power Query-запросы.

## Что входит в MVP

- Генерация новой книги Excel с готовым клиентом.
- Установка клиента в существующую книгу без удаления пользовательских листов.
- Cross-platform UI preview backend для macOS/Linux через `openpyxl`.
- Windows COM backend для настоящего встраивания Power Query в тот же UI-шаблон.
- Форматированные листы:
  - `Trino Config` - параметры подключения и extra credentials.
  - `Trino Query` - SQL-запрос и таблица результата.
  - `Trino Help` - краткая инструкция для пользователя.
- Power Query-запросы:
  - `qConfig`
  - `qSqlText`
  - `qExtraCredentials`
  - `fnTrinoRestQuery`
  - `TrinoResult`
- Поддержка Basic Auth через HTTP header.
- Поддержка `X-Trino-Extra-Credential`.
- Optional headers `X-Trino-Catalog` и `X-Trino-Schema`.
- Retry для `429`, `502`, `503`, `504`.

## Требования

Для разработки UI и тестов на macOS/Linux:

- Python 3.10+.
- `openpyxl`.
- `pytest`.

Для полной генерации рабочей книги с Power Query:

- Windows.
- Desktop Microsoft Excel 2016+ с Power Query.
- Python 3.10+.
- `pywin32`.
- Доступ к Trino REST API.
- Basic Auth для Trino.
- Если источник требует delegated credentials, нужны extra credentials, например `gp-user` и `gp-password`.

> Важно: генерация использует Excel COM через `pywin32`, поэтому запускать команды нужно на Windows-машине с установленным Excel.

## Установка для разработки

macOS/Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -e ".[test]"
```

Windows:

```bat
python -m venv .venv
.venv\Scripts\activate
pip install -e ".[test]"
```

## Тестирование на macOS/Linux

Проверить контракт проекта без Excel:

```bash
trino-excel-client validate
```

Запустить тесты:

```bash
pytest
```

Создать UI-preview книгу без Power Query:

```bash
trino-excel-client create --backend openpyxl --output ./build/Trino_UI_Preview.xlsx --overwrite
```

В эту книгу будут добавлены листы и таблицы для проверки пользовательского интерфейса. Power Query в ней не будет, потому что `openpyxl` не умеет создавать Excel Power Query объекты.

## Тестирование на Windows

Проверить контракт проекта:

```bat
trino-excel-client validate
```

Запустить unit-тесты:

```bat
pytest
```

Создать настоящую книгу с Power Query через Excel COM:

```bat
trino-excel-client create --backend com --output .\build\Trino_REST_Client.xlsx --overwrite
```

Эта команда сначала создает UI-книгу через `openpyxl`, затем открывает ее в desktop Excel через COM и добавляет Power Query-запросы.

Для отладки с видимым Excel:

```bat
trino-excel-client create --backend com --output .\build\Trino_REST_Client.xlsx --overwrite --visible
```

`--backend auto` использует `com` на Windows и `openpyxl` на остальных ОС.

## Встроить клиент в существующую книгу

Сохранить копию существующей книги с добавленным клиентом на Windows:

```bat
trino-excel-client install --backend com --workbook .\reports\MyWorkbook.xlsx --output .\build\MyWorkbook_With_Trino.xlsx --overwrite
```

Эта команда сначала добавляет UI-листы через `openpyxl`, затем открывает итоговую книгу через COM и встраивает Power Query.

Сделать cross-platform UI-preview установку на macOS/Linux:

```bash
trino-excel-client install --backend openpyxl --workbook ./reports/MyWorkbook.xlsx --output ./build/MyWorkbook_With_Trino_UI.xlsx --overwrite
```

Обновить существующую книгу на месте на Windows:

```bat
trino-excel-client install --backend com --workbook .\reports\MyWorkbook.xlsx
```

Если в книге уже есть листы с похожими именами, можно поменять префикс:

```bat
trino-excel-client install --backend com --workbook .\reports\MyWorkbook.xlsx --sheet-prefix "Data Client"
```

## Как пользоваться в Excel

1. Откройте книгу.
2. На листе `Trino Config` заполните:
   - `trino_base_url`
   - `trino_user`
   - `trino_password`
   - при необходимости `trino_catalog` и `trino_schema`
3. В таблице `tblExtraCredentials` включите нужные строки через `enabled = TRUE`.
4. На листе `Trino Query` напишите SQL в таблице `tblTrinoSql`.
5. Нажмите `Данные -> Обновить все`.

При первом обращении Excel может спросить уровень доступа к Trino host. Обычно нужно выбрать `Anonymous`, потому что Basic Auth передается самим M-кодом через HTTP header.

## Ручная отправка SQL через Trino REST API

Если генератор недоступен или нужно быстро проверить идею в обычной книге, клиент можно собрать вручную через Power Query. Логика та же: в книге создаются три Excel-таблицы с параметрами, а в Power Query вставляется большой M-код из [src/trino_excel_client/m_code.py](/Users/danielchernikov/Documents/Codex/2026-05-21/excel-trino-restapi-m-ui-excel/src/trino_excel_client/m_code.py).

### 1. Создать таблицу параметров

На листе `Trino Config` создайте Excel Table с именем `tblTrinoConfig`. Заголовки могут быть русскими, потому что M-код читает первые две колонки по позиции:

| Параметр | Значение | Обязателен | Комментарий |
| --- | --- | --- | --- |
| trino_base_url | https://trino.example.com | да | URL координатора Trino без /v1/statement |
| trino_user | CHANGE_ME | да | Пользователь Trino |
| trino_password | CHANGE_ME | да | Пароль Trino |
| trino_catalog |  | нет | Catalog, если нужен |
| trino_schema |  | нет | Schema, если нужна |
| source_name | excel_power_query | нет | X-Trino-Source |
| time_zone | Europe/Moscow | нет | X-Trino-Time-Zone |
| request_timeout_minutes | 5 | нет | Таймаут HTTP-запроса |
| polling_delay_seconds | 0.2 | нет | Пауза между nextUri |
| max_retry_count | 5 | нет | Retry для 429/502/503/504 |
| retry_base_delay_seconds | 0.5 | нет | Базовая retry-задержка |

### 2. Создать таблицу extraCredentials

Создайте Excel Table с именем `tblExtraCredentials`. M-код читает первые пять колонок по позиции:

| Включено | Тип логина | Логин | Тип пароля | Пароль | Комментарий |
| --- | --- | --- | --- | --- | --- |
| TRUE | gp-user | your_login | gp-password | your_password | Пример Greenplum |

Если extraCredentials не нужны, оставьте строки выключенными через `FALSE`.

### 3. Создать таблицу SQL

Создайте Excel Table с именем `tblTrinoSql`. Нужна минимум одна колонка:

| SQL запрос |
| --- |
| select * from your_catalog.your_schema.your_table limit 100 |

Для большого SQL можно вставить весь запрос в одну ячейку. Можно также разбить SQL на несколько строк таблицы: `qSqlText` склеит непустые строки через перенос строки.

### 4. Вставить M-код в Power Query

В Excel откройте `Данные -> Получить данные -> Из других источников -> Пустой запрос`, затем `Advanced Editor`.

Создайте пять запросов с точными именами:

- `qConfig`
- `qSqlText`
- `qExtraCredentials`
- `fnTrinoRestQuery`
- `TrinoResult`

Для каждого запроса вставьте соответствующее значение из словаря `M_QUERIES` в [m_code.py](/Users/danielchernikov/Documents/Codex/2026-05-21/excel-trino-restapi-m-ui-excel/src/trino_excel_client/m_code.py). Вставлять нужно содержимое строки без внешних Python-кавычек `r''' ... '''`.

### 5. Загрузить результат

После создания запроса `TrinoResult` выберите `Close & Load To...` и загрузите результат как таблицу на лист `Trino Query`.

При первом обращении Excel может спросить credentials для Trino host. Обычно нужно выбрать `Anonymous`, потому что Basic Auth уходит внутри M-кода через HTTP header `Authorization`.

Заполняемые параметры:

- `trino_base_url`: базовый URL Trino coordinator, например `https://trino.example.com`.
- `trino_user`: пользователь Basic Auth.
- `trino_password`: пароль Basic Auth.
- `trino_catalog`: опционально, попадет в `X-Trino-Catalog`.
- `trino_schema`: опционально, попадет в `X-Trino-Schema`.
- `source_name`: значение `X-Trino-Source`.
- `time_zone`: значение `X-Trino-Time-Zone`.
- `request_timeout_minutes`: таймаут одного HTTP-запроса.
- `polling_delay_seconds`: пауза между запросами `nextUri`.
- `max_retry_count`: количество повторов для временных HTTP-ошибок.
- `retry_base_delay_seconds`: базовая задержка exponential backoff.

## Extra credentials

Таблица `tblExtraCredentials`:

| enabled | credential_login_type | credential_login | credential_password_type | credential_password | comment |
| --- | --- | --- | --- | --- | --- |
| TRUE | gp-user | your_login | gp-password | your_password | Greenplum example |

M-код преобразует включенные строки в header:

```text
X-Trino-Extra-Credential: gp-user=your_login,gp-password=your_password
```

## Безопасность

Парольные ячейки визуально скрываются Excel-форматом `;;;`, но это не шифрование. Учетные данные остаются внутри файла. Для рабочих сценариев рекомендуется:

- хранить книгу в защищенном месте;
- не рассылать файлы с реальными паролями;
- рассмотреть отдельный сценарий без сохранения паролей в книге, если политика безопасности это требует.

## Ограничения MVP

- `openpyxl` backend нужен для разработки, тестов и UI-preview. Он не встраивает Power Query.
- `com` backend не отвечает за UI. Если внешний вид нужно менять, править нужно `openpyxl_backend.py` и тесты, а не `excel_com.py`.
- Автоматическое создание output table `tblTrinoResult` зависит от версии Excel и провайдера `Microsoft.Mashup.OleDb.1`. Если этот шаг не сработает, Power Query-запросы все равно будут добавлены, а `TrinoResult` можно загрузить вручную через `Queries & Connections -> Load To`.
- M-код пока возвращает значения Trino без явного приведения типов из `columns[type]`; это хороший следующий шаг после проверки базового обмена.

## Ближайшие улучшения

- Отдельный безопасный режим без хранения паролей в книге.
- Поддержка Bearer/JWT/OAuth, если Trino настроен не только на Basic Auth.
- UI-лист с выбором catalog/schema из справочника.
- Более умная типизация результата по Trino metadata.
- Пакетный режим установки клиента в набор книг.
