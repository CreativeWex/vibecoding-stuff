"""Heuristic table-name mapping hints (difflib) between method SQL and user schema."""

from __future__ import annotations

import difflib
import logging
import re
from typing import Any

log = logging.getLogger(__name__)

_ID_RE = re.compile(
    r"(?i)\b(?:FROM|JOIN|INTO|UPDATE|TABLE)\s+(?:ONLY\s+)?(?:(\w+)\.)?(\w+)\b"
)


def _identifiers_from_sql(sql: str) -> set[str]:
    names: set[str] = set()
    for m in _ID_RE.finditer(sql):
        schema, name = m.group(1), m.group(2)
        if name and name.upper() not in _KEYWORDS:
            names.add(name)
    return names


_KEYWORDS = frozenset(
    """
    SELECT INSERT UPDATE DELETE CREATE ALTER DROP TABLE VIEW INDEX FROM JOIN INNER LEFT RIGHT
    FULL OUTER ON WHERE GROUP BY ORDER HAVING LIMIT OFFSET VALUES SET INTO ONLY AS DISTINCT
    ALL UNION CASE WHEN THEN ELSE END BEGIN COMMIT ROLLBACK SAVEPOINT
    """.split()
)


def build_hints(
    method_statements: list[dict[str, Any]], schema_tables: dict[str, Any]
) -> dict[str, Any]:
    ident: set[str] = set()
    for st in method_statements:
        ident |= _identifiers_from_sql(st.get("sql", ""))

    schema_names = list(schema_tables.keys())
    suggestions: dict[str, list[dict[str, Any]]] = {}

    for name in sorted(ident):
        matches = difflib.get_close_matches(
            name, schema_names, n=5, cutoff=0.35
        )
        scored: list[dict[str, Any]] = []
        for m in matches:
            score = difflib.SequenceMatcher(None, name.lower(), m.lower()).ratio()
            scored.append({"schema_table": m, "score": round(score, 4)})
        scored.sort(key=lambda x: -x["score"])
        if name in schema_names:
            suggestions[name] = [{"schema_table": name, "score": 1.0}]
        elif scored:
            suggestions[name] = scored[:3]
        log.debug("Hints for %r: %s", name, suggestions.get(name))

    return {
        "identifiers": sorted(ident),
        "suggestions": suggestions,
    }
