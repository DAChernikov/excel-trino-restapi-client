from __future__ import annotations

from .m_code import M_QUERIES
from .template import REQUIRED_QUERY_NAMES, REQUIRED_TABLE_NAMES


def validate_project_contract() -> list[str]:
    errors: list[str] = []

    missing_queries = [name for name in REQUIRED_QUERY_NAMES if name not in M_QUERIES]
    if missing_queries:
        errors.append(f"Missing M queries: {', '.join(missing_queries)}")

    for query_name, formula in M_QUERIES.items():
        stripped = formula.strip()
        if not stripped.startswith("let") and not stripped.startswith("("):
            errors.append(f"M query {query_name} does not look like a valid M expression")
        if "#(cr)" in stripped:
            errors.append(f"M query {query_name} contains CR markers; use #(lf) for SQL line joins")

    required_table_refs = {
        "tblTrinoConfig": "qConfig",
        "tblTrinoSql": "qSqlText",
        "tblExtraCredentials": "qExtraCredentials",
    }
    for table_name, query_name in required_table_refs.items():
        formula = M_QUERIES.get(query_name, "")
        if table_name not in formula:
            errors.append(f"M query {query_name} does not reference required table {table_name}")

    return errors
