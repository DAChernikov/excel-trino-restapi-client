# Trino Excel Client

Trino Excel Client помогает пользователям Excel подключаться к Trino через REST API и Power Query без ручной сборки M-кода. Приложение создает новую Excel-книгу или добавляет Trino-клиент в уже существующую книгу.

## Что пользователь получает

В книге появляются отдельные листы:

- `Trino Config` - настройки подключения, лимиты и extraCredentials.
- `Trino Query` - поле для SQL-запроса.
- `Trino Result` - результат запроса.
- `Trino Help` - короткие подсказки внутри книги.

В Windows-версии дополнительно встраиваются Power Query-запросы для обращения к Trino REST API.

## Поддерживаемые платформы

### Windows

Windows - основная платформа для рабочей книги с Power Query.

Нужно:

- Desktop Microsoft Excel 2016+ с Power Query.
- Доступ к Trino REST API.
- Учетные данные Basic Auth для Trino.
- При необходимости extraCredentials для источников, например `gp-user` / `gp-password`.

Рекомендуемый пользовательский вход - GUI-приложение `Trino Excel Client`.

## Установка на Windows

Основной пользовательский сценарий - готовый installer из репозитория:

```text
install\setup.exe
```

Как установить:

1. Откройте `install\setup.exe`.
2. Выберите папку установки.
3. Выберите состав установки:
   - `GUI and CLI` - графический интерфейс и консольная утилита.
   - `Only GUI` - только графический интерфейс.
4. Выберите, нужны ли ярлыки в Start Menu и на рабочем столе.
5. Нажмите `Install` и дождитесь завершения progress bar.
6. На финальном экране нажмите `Finish`.

После установки GUI доступен как `trino-excel-client-gui.exe` в папке установки. Если при установке были выбраны ярлыки, приложение также появится в Start Menu или на рабочем столе.

CLI, если он установлен, доступен как `trino-excel-client-cmd.exe` в папке установки:

```bat
trino-excel-client-cmd.exe create --backend com --output .\build\Trino_REST_Client.xlsx --overwrite
```

### macOS / Linux / Unix

На Unix-платформах можно:

- создавать UI-preview книгу без Power Query;
- проверять внешний вид шаблона;
- запускать GUI в dry-run режиме;
- тестировать сценарии создания/встраивания без обращения к Excel COM.

Power Query-встраивание через COM доступно только на Windows desktop Excel.

```bash
trino-excel-client create --backend openpyxl --output ./build/Trino_UI_Preview.xlsx --overwrite
```

GUI dry-run:

```bash
trino-excel-client gui --dry-run
```

## Установка из исходников для разработки

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

Проверка:

```bash
trino-excel-client validate
pytest
```

## Создать новую книгу

Windows, рабочая книга с Power Query:

```bat
trino-excel-client-cmd.exe create --backend com --output .\build\Trino_REST_Client.xlsx --overwrite
```

Windows с видимым Excel для диагностики:

```bat
trino-excel-client-cmd.exe create --backend com --output .\build\Trino_REST_Client.xlsx --overwrite --visible
```

macOS/Linux, только UI-preview:

```bash
trino-excel-client create --backend openpyxl --output ./build/Trino_UI_Preview.xlsx --overwrite
```

## Встроить клиент в существующую книгу

Windows, сохранить копию с Trino-клиентом:

```bat
trino-excel-client-cmd.exe install --backend com --workbook .\reports\Workbook.xlsx --output .\build\Workbook_With_Trino.xlsx --overwrite
```

Windows, обновить книгу на месте:

```bat
trino-excel-client-cmd.exe install --backend com --workbook .\reports\Workbook.xlsx
```

macOS/Linux, добавить только UI-листы:

```bash
trino-excel-client install --backend openpyxl --workbook ./reports/Workbook.xlsx --output ./build/Workbook_With_Trino_UI.xlsx --overwrite
```

При встраивании существующие пользовательские листы сохраняются. Trino-клиент добавляет свои листы отдельно. Лист `Trino Query` не используется для результата, поэтому повторное обновление данных не должно стирать SQL-ввод.

## Как работать в Excel

1. Откройте созданную книгу.
2. На листе `Trino Config` заполните:
   - `trino_base_url`
   - `trino_user`
   - `trino_password`
   - если нужно: `trino_catalog`, `trino_schema`
3. Если нужны delegated credentials, заполните `tblExtraCredentials`.
4. На листе `Trino Query` введите SQL в таблицу `tblTrinoSql`.
5. Нажмите `Данные -> Обновить все`.
6. Результат появится на листе `Trino Result`.

При первом обращении Excel может спросить credentials/privacy level для Trino host. Обычно нужно выбрать `Anonymous`, потому что Basic Auth передается M-кодом через HTTP header.

Клиент настроен так, чтобы не отправлять запросы в Trino при открытии книги или при редактировании SQL. Запрос должен уходить только после пользовательского refresh-действия в Excel, например `Данные -> Обновить все`.

## Лимиты и BigData

Excel не подходит для прямой загрузки миллионов или миллиардов строк. Поэтому в шаблоне есть два защитных механизма.

| Параметр | По умолчанию | Что делает |
| --- | --- | --- |
| `default_query_limit_rows` | `100000` | M-код добавляет внешний `LIMIT` вокруг SQL |
| `apply_default_query_limit` | `TRUE` | Включает или отключает внешний `LIMIT` |
| `max_result_rows` | `100000` | Жесткий лимит строк, разрешенных для загрузки в Excel |
| `result_overflow_behavior` | `error` | `error` останавливает загрузку, `truncate` загружает первые строки |

По умолчанию ваш SQL отправляется в Trino так:

```sql
select *
from (
    <ваш SQL>
) as excel_trino_client_query
limit 100000
```

Если нужен полный контроль над SQL, поставьте:

```text
apply_default_query_limit = FALSE
```

Оставьте `max_result_rows` включенным. Это последний защитный слой от случайной выгрузки слишком большого результата.

Практические рекомендации:

- Для Excel используйте фильтры, агрегаты и выбор только нужных колонок.
- Не делайте `select *` по большим таблицам без понимания объема результата.
- Для миллионов строк используйте выгрузку во внешнее хранилище, parquet/csv/export pipeline или BI/ETL-инструменты.

## Extra credentials

Пример `tblExtraCredentials`:

| Включено | Тип логина | Логин | Тип пароля | Пароль |
| --- | --- | --- | --- | --- |
| TRUE | gp-user | your_login | gp-password | your_password |

M-код отправляет это как:

```text
X-Trino-Extra-Credential: gp-user=your_login,gp-password=your_password
```

## Schema preview и типы данных

Power Query использует metadata Trino `columns[type]`, чтобы назначать типы результата:

- `tinyint/smallint/integer/bigint` -> `Int64`
- `real/double/decimal` -> `number`
- `boolean` -> `logical`
- `date` -> `date`
- `timestamp` -> `datetime`
- `timestamp with time zone` -> `datetimezone`
- `varchar/char/json/uuid/ipaddress` -> `text`

Отдельный schema-запрос в шаблон не добавляется, потому что просмотр схемы через Power Query тоже выполняет обращение к Trino. Это сделано намеренно: книга не должна запускать дополнительные запросы сама по себе.

## Безопасность

Парольные ячейки визуально скрыты форматом Excel `;;;`, но это не шифрование. Книга все равно содержит учетные данные.

Рекомендации:

- не рассылайте книги с реальными паролями;
- храните файлы с заполненными credentials в защищенном месте;
- если политика безопасности запрещает хранить пароль в Excel, используйте отдельный процесс выдачи временных credentials.

## Два приложения после установки

После установки в папке приложения будут два исполняемых файла.

### `trino-excel-client-gui.exe`

Основное приложение для пользователя. Его стоит запускать через ярлык `Trino Excel Client`.

Навигация:

1. Откройте приложение.
2. На верхней панели выберите:
   - `Dry-run режим`, если хотите посмотреть действия без изменения файлов;
   - `Показывать Excel при COM-операциях`, если нужно видеть Excel во время генерации;
   - `Перезаписывать файлы`, если выходной файл можно заменить.
3. При необходимости поменяйте `Префикс листов`. По умолчанию это `Trino`.
4. Вкладка `Создать шаблон`:
   - выберите путь для новой книги;
   - выберите backend:
     - `auto` - COM на Windows, openpyxl на Unix;
     - `com` - рабочая книга с Power Query через Windows Excel;
     - `openpyxl` - только UI-preview;
   - нажмите `Создать книгу`.
5. Вкладка `Встроить в книгу`:
   - выберите исходную Excel-книгу;
   - выберите путь результата;
   - выберите backend;
   - нажмите `Встроить клиент`.
6. Вкладка `Подсказки` содержит краткие ограничения по Power Query, BigData и платформам.
7. Внизу приложения отображается `Журнал операций`.

На macOS/Linux используйте GUI в dry-run или с backend `openpyxl`. Полноценное встраивание Power Query работает только на Windows с desktop Excel.

Запуск GUI из исходников:

```bat
trino-excel-client-gui
```

Dry-run:

```bash
trino-excel-client gui --dry-run
```

### `trino-excel-client-cmd.exe`

Консольная версия для cmd/PowerShell, скриптов и автоматизации.

Команды:

```bat
trino-excel-client-cmd.exe create --backend com --output .\build\Trino_REST_Client.xlsx --overwrite
```

Создать рабочую книгу на Windows с Power Query.

```bat
trino-excel-client-cmd.exe create --backend com --output .\build\Trino_REST_Client.xlsx --overwrite --visible
```

То же самое, но Excel будет видимым во время генерации.

```bat
trino-excel-client-cmd.exe create --backend openpyxl --output .\build\Trino_UI_Preview.xlsx --overwrite
```

Создать UI-preview книгу без Power Query.

```bat
trino-excel-client-cmd.exe install --backend com --workbook .\reports\Workbook.xlsx --output .\build\Workbook_With_Trino.xlsx --overwrite
```

Встроить Trino-клиент в существующую книгу и сохранить копию.

```bat
trino-excel-client-cmd.exe install --backend com --workbook .\reports\Workbook.xlsx
```

Встроить Trino-клиент в существующую книгу на месте.

```bat
trino-excel-client-cmd.exe validate
```

Проверить внутренний контракт проекта.

```bat
trino-excel-client-cmd.exe gui --dry-run
```

Открыть GUI в dry-run режиме.

## Windows smoke-test

Перед релизом на Windows:

```bat
python scripts\windows_smoke_test.py --visible
```

На macOS/Linux можно проверить UI-часть:

```bash
python scripts/windows_smoke_test.py --skip-com
```

Ручная проверка на Windows:

1. Откройте `build\smoke\Trino_REST_Client_COM.xlsx`.
2. Заполните `Trino Config`.
3. Введите небольшой SQL, например `select * from ... limit 10`.
4. Нажмите `Данные -> Обновить все` несколько раз.
5. Проверьте, что `Trino Query` не меняется, а `Trino Result` обновляется.

## Сборка Windows exe и installer

Установить packaging-зависимости:

```bat
pip install -e ".[package]"
```

Собрать CLI и GUI exe:

```bat
python scripts\build_windows_exe.py --clean
```

Ожидаемые файлы:

```text
dist\trino-excel-client-cmd.exe
dist\trino-excel-client-gui.exe
```

Собрать installer через Inno Setup:

```bat
python scripts\build_windows_installer.py
```

Installer ожидает, что установлен Inno Setup compiler `ISCC.exe`. Готовый установщик будет создан как `install\setup.exe`.

Чтобы пользователи после клонирования репозитория сразу видели `install\setup.exe`, файл нужно коммитить в репозиторий как релизный артефакт. GitHub Actions настроен так, чтобы при push в `master` пересобрать installer и закоммитить новый `install\setup.exe`, если он изменился. Вручную это можно запустить через `Actions -> Build Windows setup.exe -> Run workflow -> commit_installer=true`.

В setup.exe доступны варианты установки:

- `GUI and CLI` - установить графическое приложение и консольную утилиту.
- `Only GUI` - установить только `trino-excel-client-gui.exe`.

GUI всегда является основным компонентом. CLI устанавливается только как дополнительная опция. Ярлыки в Start Menu и на рабочем столе выбираются отдельно во время установки.

Оба exe являются self-contained PyInstaller-артефактами. После установки папку приложения можно перенести или скопировать, сами exe продолжат работать, но ярлыки Windows будут указывать на старое место до переустановки или ручного обновления ярлыков.
