"""Extract plain text from PDF using pypdf."""

from __future__ import annotations

import logging
from io import BytesIO

from pypdf import PdfReader

log = logging.getLogger(__name__)


def extract_pdf_text(data: bytes) -> str:
    if not data:
        raise ValueError("PDF data is empty")
    reader = PdfReader(BytesIO(data))
    if reader.is_encrypted:
        try:
            reader.decrypt("")
        except Exception as e:
            log.warning("Encrypted PDF, decrypt failed: %s", e)
            raise ValueError("PDF is encrypted; cannot decrypt without password") from e
    parts: list[str] = []
    for i, page in enumerate(reader.pages):
        try:
            t = page.extract_text() or ""
        except Exception as e:
            log.error("Page %s extract failed: %s", i + 1, e)
            raise ValueError(f"Failed to extract text from page {i + 1}") from e
        parts.append(t)
        log.debug("Page %s: %s characters", i + 1, len(t))
    full = "\n\n".join(parts)
    log.info("PDF total extracted length: %s characters", len(full))
    return full
