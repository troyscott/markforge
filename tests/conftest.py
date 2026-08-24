from __future__ import annotations

from pathlib import Path

import pymupdf
import pytest


@pytest.fixture
def make_pdf():
    def factory(
        path: Path,
        pages: int,
        headings: dict[int, str] | None = None,
        toc: list[list[object]] | None = None,
    ) -> Path:
        headings = headings or {}
        with pymupdf.open() as document:
            for page_number in range(1, pages + 1):
                page = document.new_page()
                if page_number in headings:
                    page.insert_text((72, 80), headings[page_number], fontsize=24)
                page.insert_text((72, 140), f"Body text for page {page_number}", fontsize=11)
            if toc:
                document.set_toc(toc)
            document.save(path)
        return path
    return factory
