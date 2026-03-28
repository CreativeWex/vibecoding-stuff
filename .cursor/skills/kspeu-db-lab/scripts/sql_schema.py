"""Parse PostgreSQL dump/script into schema.json structure using pglast."""

from __future__ import annotations

import logging
import re
from typing import Any

from pglast import parse_sql
from pglast.ast import (
    AlterTableStmt,
    AlterTableCmd,
    ColumnDef,
    Constraint,
    CreateFunctionStmt,
    CreateStmt,
    CreateTrigStmt,
    IndexStmt,
    String,
    TypeName,
    ViewStmt,
)
from pglast.enums import AlterTableType, ConstrType

log = logging.getLogger(__name__)


def _strip_copy_sections(sql: str) -> str:
    lines = sql.splitlines()
    out: list[str] = []
    i = 0
    while i < len(lines):
        if re.match(r"^\s*COPY\s+", lines[i], re.IGNORECASE):
            i += 1
            while i < len(lines) and not re.match(r"^\s*\\\.\s*$", lines[i]):
                i += 1
            if i < len(lines):
                i += 1
            continue
        out.append(lines[i])
        i += 1
    return "\n".join(out)


def _type_to_str(tn: TypeName | None) -> str:
    if tn is None:
        return ""
    parts: list[str] = []
    for node in tn.names:
        if isinstance(node, String):
            parts.append(node.sval)
        else:
            parts.append(str(node))
    s = ".".join(parts) if parts else ""
    if s.startswith("pg_catalog."):
        return s[len("pg_catalog.") :]
    return s


def _constr_labels(constraints: tuple[Any, ...] | list[Any] | None) -> list[str]:
    labels: list[str] = []

    for c in constraints or ():
        if not isinstance(c, Constraint):
            continue
        ct = c.contype
        if ct == ConstrType.CONSTR_PRIMARY:
            labels.append("PRIMARY KEY")
        elif ct == ConstrType.CONSTR_UNIQUE:
            labels.append("UNIQUE")
        elif ct == ConstrType.CONSTR_FOREIGN:
            labels.append("FOREIGN KEY")
        elif ct == ConstrType.CONSTR_CHECK:
            labels.append("CHECK")
        elif ct == ConstrType.CONSTR_NOTNULL:
            labels.append("NOT NULL")
    return labels


def _column_from_def(cd: ColumnDef) -> dict[str, Any]:
    constr = _constr_labels(cd.constraints)
    not_null = any(
        isinstance(c, Constraint) and c.contype == ConstrType.CONSTR_NOTNULL
        for c in (cd.constraints or ())
    )
    return {
        "name": cd.colname,
        "type": _type_to_str(cd.typeName),
        "nullable": not not_null,
        "constraints": [x for x in constr if x != "NOT NULL"],
    }


def _fk_from_column_constraint(
    table: str, col: str, c: Constraint
) -> dict[str, Any] | None:
    if c.contype != ConstrType.CONSTR_FOREIGN:
        return None
    pktable = c.pktable.relname if c.pktable else ""
    pk_attrs = [s.sval for s in c.pk_attrs] if c.pk_attrs else []
    return {
        "table": table,
        "columns": [col],
        "ref_table": pktable,
        "ref_columns": pk_attrs,
    }


def _process_create_table(stmt: CreateStmt) -> tuple[str, dict[str, Any], list[dict[str, Any]]]:
    name = stmt.relation.relname
    cols: list[dict[str, Any]] = []
    fks: list[dict[str, Any]] = []
    for elt in stmt.tableElts or ():
        if isinstance(elt, ColumnDef):
            cols.append(_column_from_def(elt))
            for c in elt.constraints or ():
                if isinstance(c, Constraint):
                    fk = _fk_from_column_constraint(name, elt.colname or "", c)
                    if fk:
                        fks.append(fk)
        elif isinstance(elt, Constraint):
            if elt.contype == ConstrType.CONSTR_FOREIGN and elt.fk_attrs and elt.pktable:
                fk = {
                    "table": name,
                    "columns": [s.sval for s in elt.fk_attrs],
                    "ref_table": elt.pktable.relname,
                    "ref_columns": [s.sval for s in elt.pk_attrs] if elt.pk_attrs else [],
                }
                fks.append(fk)
    return name, {"columns": cols}, fks


def _process_index(stmt: IndexStmt) -> dict[str, Any]:
    rel = stmt.relation.relname if stmt.relation else ""
    cols: list[str] = []
    for p in stmt.indexParams or ():
        if getattr(p, "name", None):
            cols.append(p.name)
    return {
        "name": stmt.idxname or "",
        "table": rel,
        "columns": cols,
    }


def _process_view(stmt: ViewStmt) -> dict[str, Any]:
    name = stmt.view.relname if stmt.view else ""
    return {"name": name, "kind": "view"}


def _process_function(stmt: CreateFunctionStmt) -> dict[str, Any]:
    name = stmt.funcname[-1].sval if stmt.funcname else ""
    return {"name": name, "args": "..."}


def _process_trigger(stmt: CreateTrigStmt) -> dict[str, Any]:
    rel = stmt.relation.relname if stmt.relation else ""
    return {"name": stmt.trigname or "", "table": rel}


def _split_sql_statements_naive(sql: str) -> list[str]:
    """Semicolon split with single-quote and dollar-quote awareness (best-effort)."""
    parts: list[str] = []
    buf: list[str] = []
    i = 0
    n = len(sql)
    in_squote = False
    dollar_delim: str | None = None

    while i < n:
        ch = sql[i]
        if dollar_delim is not None:
            if sql.startswith(dollar_delim, i):
                buf.append(dollar_delim)
                i += len(dollar_delim)
                dollar_delim = None
                continue
            buf.append(ch)
            i += 1
            continue
        if ch == "'" and not in_squote:
            in_squote = True
            buf.append(ch)
            i += 1
            continue
        if in_squote:
            buf.append(ch)
            if ch == "'" and (i + 1 >= n or sql[i + 1] != "'"):
                in_squote = False
            elif ch == "'" and i + 1 < n and sql[i + 1] == "'":
                buf.append("'")
                i += 2
                continue
            i += 1
            continue
        if ch == "$" and i + 1 < n:
            m = re.match(r"\$([a-zA-Z_]*)\$", sql[i:])
            if m:
                delim = m.group(0)
                dollar_delim = delim
                buf.append(delim)
                i += len(delim)
                continue
        if ch == ";":
            chunk = "".join(buf).strip()
            if chunk:
                parts.append(chunk)
            buf = []
            i += 1
            continue
        buf.append(ch)
        i += 1
    tail = "".join(buf).strip()
    if tail:
        parts.append(tail)
    return parts


def extract_schema(sql_text: str) -> dict[str, Any]:
    if not sql_text or not sql_text.strip():
        raise ValueError("SQL file is empty")

    cleaned = _strip_copy_sections(sql_text)
    cleaned = re.sub(r"^\s*\\[^\n]*$", "", cleaned, flags=re.MULTILINE)

    tables: dict[str, dict[str, Any]] = {}
    foreign_keys: list[dict[str, Any]] = []
    indexes: list[dict[str, Any]] = []
    functions: list[dict[str, Any]] = []
    triggers: list[dict[str, Any]] = []
    views: list[dict[str, Any]] = []

    raw_list: tuple[Any, ...] = ()
    try:
        raw_list = parse_sql(cleaned)
        log.info("Parsed SQL as single unit: %s raw statements", len(raw_list))
    except Exception as e:
        log.warning("Full-file parse failed (%s); trying chunked parse", e)
        chunks = _split_sql_statements_naive(cleaned)
        merged: list[Any] = []
        for idx, chunk in enumerate(chunks):
            if not chunk.strip() or chunk.strip().upper().startswith("SET "):
                continue
            try:
                merged.extend(parse_sql(chunk + ";"))
            except Exception as e2:
                log.debug("Chunk %s parse skip: %s — %s", idx, chunk[:80], e2)
        raw_list = tuple(merged)
        log.info("Chunked parse yielded %s raw statements", len(raw_list))

    for raw in raw_list:
        stmt = raw.stmt
        try:
            if isinstance(stmt, CreateStmt):
                tname, tdef, fks = _process_create_table(stmt)
                tables[tname] = tdef
                foreign_keys.extend(fks)
            elif isinstance(stmt, IndexStmt):
                indexes.append(_process_index(stmt))
            elif isinstance(stmt, ViewStmt):
                views.append(_process_view(stmt))
            elif isinstance(stmt, CreateFunctionStmt):
                functions.append(_process_function(stmt))
            elif isinstance(stmt, CreateTrigStmt):
                triggers.append(_process_trigger(stmt))
            elif isinstance(stmt, AlterTableStmt):
                rel = stmt.relation.relname if stmt.relation else None
                if not rel:
                    continue
                if rel not in tables:
                    tables[rel] = {"columns": []}
                for cmd in stmt.cmds or ():
                    if not isinstance(cmd, AlterTableCmd):
                        continue
                    def_ = cmd.def_
                    if cmd.subtype == AlterTableType.AT_AddColumn and isinstance(
                        def_, ColumnDef
                    ):
                        tables[rel]["columns"].append(_column_from_def(def_))
                        for c in def_.constraints or ():
                            if isinstance(c, Constraint):
                                fk = _fk_from_column_constraint(
                                    rel, def_.colname or "", c
                                )
                                if fk:
                                    foreign_keys.append(fk)
                    elif cmd.subtype == AlterTableType.AT_AddConstraint and isinstance(
                        def_, Constraint
                    ):
                        if def_.contype == ConstrType.CONSTR_FOREIGN and def_.pktable:
                            foreign_keys.append(
                                {
                                    "table": rel,
                                    "columns": [s.sval for s in (def_.fk_attrs or ())],
                                    "ref_table": def_.pktable.relname,
                                    "ref_columns": [
                                        s.sval for s in (def_.pk_attrs or ())
                                    ],
                                }
                            )
            else:
                log.debug("Unhandled statement type: %s", type(stmt).__name__)
        except Exception as ex:
            log.warning("Statement handler error (%s): %s", type(stmt).__name__, ex)

    if not tables:
        raise ValueError(
            "No tables found after parsing. Check that the file contains PostgreSQL "
            "CREATE TABLE statements and valid syntax."
        )

    return {
        "tables": tables,
        "foreign_keys": foreign_keys,
        "indexes": indexes,
        "functions": functions,
        "triggers": triggers,
        "views": views,
    }
