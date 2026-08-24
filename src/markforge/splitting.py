from __future__ import annotations

import re
import unicodedata
from collections import Counter
from pathlib import Path

import pymupdf

from .models import Segment, SplitMode


HEADING_RE = re.compile(
    r"^(?:part\s+(?:[ivxlcdm]+|\d+)|chapter\s+\d+|appendix(?:\s+[a-z0-9]+)?|preface|index)\b",
    re.IGNORECASE,
)


def slugify(value: str, fallback: str = "segment") -> str:
    normalized = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode()
    slug = re.sub(r"[^a-z0-9]+", "-", normalized.lower()).strip("-")
    return slug or fallback


def _dedupe_filenames(titles: list[str]) -> list[str]:
    counts: Counter[str] = Counter()
    result: list[str] = []
    for index, title in enumerate(titles, start=1):
        base = slugify(title, f"segment-{index:02d}")
        counts[base] += 1
        suffix = f"-{counts[base]}" if counts[base] > 1 else ""
        result.append(f"{index:02d}-{base}{suffix}.md")
    return result


def _segments_from_boundaries(
    boundaries: list[tuple[int, str, int]], total_pages: int
) -> list[Segment]:
    cleaned: list[tuple[int, str, int]] = []
    seen_pages: set[int] = set()
    for page, title, level in sorted(boundaries, key=lambda item: item[0]):
        page = max(1, min(page, total_pages))
        if page in seen_pages:
            continue
        cleaned.append((page, title.strip() or f"Pages {page}", level))
        seen_pages.add(page)

    if not cleaned:
        return []
    if cleaned[0][0] > 1:
        cleaned.insert(0, (1, "Front Matter", 1))

    titles = [title for _, title, _ in cleaned]
    filenames = _dedupe_filenames(titles)
    segments: list[Segment] = []
    for index, ((page, title, level), filename) in enumerate(zip(cleaned, filenames), start=1):
        end_page = cleaned[index][0] - 1 if index < len(cleaned) else total_pages
        if end_page < page:
            continue
        segments.append(
            Segment(
                id=f"segment-{index:03d}",
                title=title,
                start_page=page,
                end_page=end_page,
                markdown_path=filename,
                level=level,
            )
        )
    return segments


def _bookmark_boundaries(toc: list[list[object]], toc_level: int | None) -> list[tuple[int, str, int]]:
    entries: list[tuple[int, str, int]] = []
    for row in toc:
        if len(row) < 3:
            continue
        try:
            level, title, page = int(row[0]), str(row[1]).strip(), int(row[2])
        except (TypeError, ValueError):
            continue
        if page < 1 or not title:
            continue
        entries.append((page, title, level))

    if toc_level is not None:
        return [entry for entry in entries if entry[2] == toc_level]

    named = [entry for entry in entries if HEADING_RE.match(entry[1])]
    if len(named) >= 2:
        return named

    by_level = Counter(level for _, _, level in entries)
    viable = [level for level, count in by_level.items() if count >= 2]
    if not viable:
        return []
    selected = min(viable)
    return [entry for entry in entries if entry[2] == selected]


def detect_heading_boundaries(document: pymupdf.Document) -> list[tuple[int, str, int]]:
    boundaries: list[tuple[int, str, int]] = []
    for page_index in range(len(document)):
        page = document[page_index]
        page_height = page.rect.height or 1
        candidates: list[tuple[float, float, str]] = []
        data = page.get_text("dict")
        for block in data.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                text = "".join(span.get("text", "") for span in line.get("spans", [])).strip()
                if not text or not HEADING_RE.match(text):
                    continue
                y0 = float(line.get("bbox", (0, 0, 0, 0))[1])
                size = max((float(span.get("size", 0)) for span in line.get("spans", [])), default=0)
                if y0 <= page_height * 0.45:
                    candidates.append((size, -y0, text))
        if candidates:
            _, _, title = max(candidates)
            boundaries.append((page_index + 1, title, 1))
    return boundaries if len(boundaries) >= 2 else []


def page_segments(total_pages: int, pages_per_segment: int) -> list[Segment]:
    boundaries = [
        (page, f"Pages {page}-{min(page + pages_per_segment - 1, total_pages)}", 1)
        for page in range(1, total_pages + 1, pages_per_segment)
    ]
    return _segments_from_boundaries(boundaries, total_pages)


def propose_segments(
    document: pymupdf.Document,
    split: SplitMode = "chapters",
    output_pages: int = 25,
    toc_level: int | None = None,
) -> tuple[list[Segment], str]:
    total_pages = len(document)
    if total_pages == 0:
        return [], "empty"
    if split == "single":
        return [Segment("segment-001", Path(document.name or "document").stem, 1, total_pages, "01-document.md")], "single"
    if split == "pages":
        return page_segments(total_pages, output_pages), "pages"

    toc = document.get_toc(simple=True)
    boundaries = _bookmark_boundaries(toc, toc_level)
    strategy = "bookmarks"
    if len(boundaries) < 2:
        boundaries = detect_heading_boundaries(document)
        strategy = "headings"
    segments = _segments_from_boundaries(boundaries, total_pages)
    if len(segments) < 2:
        return page_segments(total_pages, output_pages), "pages-fallback"
    return segments, strategy


def validate_page_coverage(segments: list[Segment], total_pages: int) -> bool:
    pages = [page for segment in segments for page in range(segment.start_page, segment.end_page + 1)]
    return pages == list(range(1, total_pages + 1))
