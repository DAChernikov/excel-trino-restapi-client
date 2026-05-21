# Examples

## Создание UI-preview на macOS/Linux

```bash
mkdir -p build
trino-excel-client create --backend openpyxl --output ./build/Trino_UI_Preview.xlsx --overwrite
```

## Создание рабочего шаблона на Windows

```bat
mkdir build
trino-excel-client create --backend com --output .\build\Trino_REST_Client.xlsx --overwrite --visible
```

## Установка в существующую книгу на Windows

```bat
trino-excel-client install --backend com --workbook .\input\Analytics.xlsx --output .\build\Analytics_With_Trino.xlsx --overwrite --visible
```

## UI-preview установка на macOS/Linux

```bash
trino-excel-client install --backend openpyxl --workbook ./input/Analytics.xlsx --output ./build/Analytics_With_Trino_UI.xlsx --overwrite
```

## Пример SQL

```sql
select *
from greenplum_private_connection_test_gp.dm.dm_bs_article_finplat
limit 1
```

В Excel этот SQL можно положить в `tblTrinoSql` построчно:

| sql_text |
| --- |
| select * |
| from greenplum_private_connection_test_gp.dm.dm_bs_article_finplat |
| limit 1 |

## Пример extra credentials

| enabled | credential_login_type | credential_login | credential_password_type | credential_password |
| --- | --- | --- | --- | --- |
| TRUE | gp-user | your_login | gp-password | your_password |
