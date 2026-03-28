#!/usr/bin/env python3
"""
Extract schema from user .sql and SQL fragments from methodology PDF.
Writes schema.json, method_statements.json, hints.json to --out-dir (async I/O + threaded parse).
"""

from __future__ import annotations

import argparse
import asyncio
import json
import logging
import sys
from pathlib import Path

_SCRIPT_DIR = Path(__file__).resolve().parent
if str(_SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPT_DIR))

import aiofiles

from hints_build import build_hints
from logging_config import setup_logging
from method_extract import extract_method_statements
from pdf_extract import extract_pdf_text
from sql_schema import extract_schema

log = logging.getLogger(__name__)


def _write_json(path: Path, obj: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(obj, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    log.info("Wrote %s", path)


async def run_pipeline(pdf_path: Path, sql_path: Path, out_dir: Path) -> None:
    async with aiofiles.open(pdf_path, "rb") as f:
        pdf_bytes = await f.read()
    log.debug("Read PDF %s bytes", len(pdf_bytes))

    async with aiofiles.open(sql_path, "r", encoding="utf-8", errors="replace") as f:
        sql_text = await f.read()
    log.debug("Read SQL %s characters", len(sql_text))

    pdf_text = await asyncio.to_thread(extract_pdf_text, pdf_bytes)
    if not pdf_text.strip():
        raise ValueError(
            "PDF text extraction yielded empty content. Try another PDF export; OCR is not supported."
        )

    schema = await asyncio.to_thread(extract_schema, sql_text)
    method = await asyncio.to_thread(extract_method_statements, pdf_text)
    hints = await asyncio.to_thread(
        build_hints, method["statements"], schema["tables"]
    )

    await asyncio.to_thread(_write_json, out_dir / "schema.json", schema)
    await asyncio.to_thread(_write_json, out_dir / "method_statements.json", method)
    await asyncio.to_thread(_write_json, out_dir / "hints.json", hints)


async def _async_main(pdf: Path, sql: Path, out_dir: Path) -> None:
    await run_pipeline(pdf, sql, out_dir)


def main() -> int:
    ap = argparse.ArgumentParser(
        description="KSPEU DB LAB: extract schema + methodology SQL hints"
    )
    ap.add_argument("--pdf", type=Path, required=True, help="Path to methodology .pdf")
    ap.add_argument("--sql", type=Path, required=True, help="Path to user .sql dump")
    ap.add_argument(
        "--out-dir",
        type=Path,
        required=True,
        help="Output directory for JSON artifacts",
    )
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args()

    setup_logging(args.verbose)

    if not args.pdf.is_file():
        log.error("PDF not found: %s", args.pdf)
        return 1
    if not args.sql.is_file():
        log.error("SQL not found: %s", args.sql)
        return 1

    try:
        asyncio.run(_async_main(args.pdf, args.sql, args.out_dir))
    except ValueError as e:
        log.error("%s", e)
        return 1
    except Exception as e:
        log.exception("Pipeline failed: %s", e)
        return 1

    log.info("Done. Outputs in %s", args.out_dir.resolve())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
