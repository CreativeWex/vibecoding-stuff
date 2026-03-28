"""Extract SQL-like fragments from methodology PDF plain text (linear order)."""

from __future__ import annotations

import logging
import re
from typing import Any

log = logging.getLogger(__name__)

# Start of a SQL statement (PostgreSQL-oriented)
_SQL_START = re.compile(
    r"(?is)^\s*("
    r"SELECT\b|INSERT\b|UPDATE\b|DELETE\b|CREATE\b|ALTER\b|DROP\b|WITH\b|"
    r"TRUNCATE\b|GRANT\b|REVOKE\b|BEGIN\b|COMMIT\b|ROLLBACK\b|SAVEPOINT\b"
    r")\s+",
)

# Section / chapter hints (previous non-empty line)
_SECTION_LINE = re.compile(
    r"^(?:"
    r"\d+(?:\.\d+)*[\.\)]\s+.{3,120}"
    r"|глава\s+\d+.{0,100}"
    r"|[A-ZА-ЯЁ][A-ZА-ЯЁ0-9\s\-]{2,50}$"
    r")",
    re.IGNORECASE | re.UNICODE,
)


def _current_section_hint(lines: list[str], start_idx: int) -> str:
    for j in range(start_idx - 1, max(-1, start_idx - 12), -1):
        if j < 0:
            break
        s = lines[j].strip()
        if not s:
            continue
        if _SECTION_LINE.match(s):
            return s[:200]
    return ""


def _extract_fenced_blocks(text: str) -> list[tuple[int, str, str]]:
    """Return list of (char_offset, block_content, 'fenced_block')."""
    found: list[tuple[int, str, str]] = []
    for m in re.finditer(r"(?m)^\s{2,}(\S.*(?:\n\s{2,}\S.*)*)", text):
        block = m.group(1)
        if not _SQL_START.match(block.strip()):
            continue
        if len(block.strip()) < 8:
            continue
        found.append((m.start(), block.strip(), "fenced_block"))
    return found


def _statements_from_lines(text: str) -> list[tuple[int, str, str]]:
    lines = text.splitlines()
    results: list[tuple[int, str, str]] = []
    char_offset = 0
    i = 0
    n = len(lines)

    while i < n:
        line = lines[i]
        stripped = line.strip()
        joined_start = "\n".join(lines[max(0, i - 1) : i + 2])
        if not (_SQL_START.match(stripped) or _SQL_START.search(joined_start)):
            char_offset += len(line) + 1
            i += 1
            continue

        start_i = i
        buf = [line]
        if ";" not in line:
            i += 1
            while i < n:
                buf.append(lines[i])
                if ";" in lines[i]:
                    break
                if i - start_i > 80:
                    break
                i += 1

        chunk = "\n".join(buf).strip()
        if ";" in chunk and len(chunk) >= 8:
            pos = text.find(chunk[: min(40, len(chunk))], char_offset)
            if pos < 0:
                pos = char_offset
            results.append((pos, chunk, "keyword_scan"))

        for row in buf:
            char_offset += len(row) + 1
        i = start_i + len(buf)

    return results


def _split_compound_sql(sql: str) -> list[str]:
    """Split 'SELECT ...; CREATE ...;' from one PDF line into separate entries."""
    s = sql.strip()
    if not s:
        return []
    pieces: list[str] = []
    for m in re.finditer(r"[^;]+;", s):
        chunk = m.group(0).strip()
        if chunk and _SQL_START.match(chunk):
            pieces.append(chunk)
    if len(pieces) > 1:
        return pieces
    if pieces:
        return pieces
    return [s] if s else []


def extract_method_statements(pdf_text: str) -> dict[str, Any]:
    if not pdf_text or not pdf_text.strip():
        raise ValueError("Methodology text is empty")

    candidates: list[tuple[int, str, str]] = []
    candidates.extend(_extract_fenced_blocks(pdf_text))
    candidates.extend(_statements_from_lines(pdf_text))

    candidates.sort(key=lambda x: x[0])

    seen: set[str] = set()
    statements: list[dict[str, Any]] = []
    idx = 0
    for char_off, sql, source in candidates:
        split_sqls = _split_compound_sql(sql)
        for part in split_sqls:
            key = re.sub(r"\s+", " ", part.strip())[:500]
            if key in seen:
                continue
            seen.add(key)
            hint = ""
            line_no = pdf_text[:char_off].count("\n")
            pre_lines = pdf_text[:char_off].splitlines()[-8:]
            for pl in reversed(pre_lines):
                t = pl.strip()
                if t and _SECTION_LINE.match(t):
                    hint = t[:200]
                    break

            statements.append(
                {
                    "index": idx,
                    "char_offset": char_off,
                    "line_approx": line_no + 1,
                    "section_hint": hint,
                    "sql": part.strip(),
                    "source": source,
                }
            )
            idx += 1

    log.info("Extracted %s SQL fragments from methodology text", len(statements))
    return {
        "statements": statements,
        "pdf_text_length": len(pdf_text),
    }
