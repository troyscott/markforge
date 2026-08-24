from __future__ import annotations

import re
import unicodedata
from collections import Counter
from pathlib import Path
from statistics import median

import pymupdf

from .models import Segment, SplitMode


HEADING_RE = re.compile(
    r"^(?:part\s+(?:[ivxlcdm]+|\d+)|chapter\s+(?:[ivxlcdm]+|\d+)|appendix(?:\s+[a-z0-9]+)?|preface|index)\b",
    re.IGNORECASE,
)
CHAPTER_NUMBER_RE = re.compile(r"^([ivxlcdmhy]+|\d+)[.:]?$", re.IGNORECASE)
INDEX_RE = re.compile(r"^index[.:]?$", re.IGNORECASE)


def _edit_distance(left: str, right: str) -> int:
    previous = list(range(len(right) + 1))
    for left_index, left_character in enumerate(left, start=1):
        current = [left_index]
        for right_index, right_character in enumerate(right, start=1):
            current.append(
                min(
                    current[-1] + 1,
                    previous[right_index] + 1,
                    previous[right_index - 1] + (left_character != right_character),
                )
            )
        previous = current
    return previous[-1]


def _chapter_token(text: str) -> str | None:
    words = text.rstrip(".:").split(maxsplit=1)
    if not words or _edit_distance(words[0].casefold(), "chapter") > 2:
        return None
    if len(words) == 1:
        return ""
    match = CHAPTER_NUMBER_RE.fullmatch(words[1])
    return match.group(1) if match else None


def _roman_to_int(token: str) -> int | None:
    values = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100, "D": 500, "M": 1000}
    if not token or any(character not in values for character in token):
        return None
    total = 0
    for index, character in enumerate(token):
        value = values[character]
        total += -value if index + 1 < len(token) and value < values[token[index + 1]] else value
    return total if _int_to_roman(total) == token else None


def _int_to_roman(value: int) -> str:
    numerals = (
        (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
        (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
        (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
    )
    result: list[str] = []
    for amount, numeral in numerals:
        count, value = divmod(value, amount)
        result.extend([numeral] * count)
    return "".join(result)


def _normalize_ocr_chapter_sequence(
    boundaries: list[tuple[int, str, int]],
) -> list[tuple[int, str, int]]:
    result: list[tuple[int, str, int]] = []
    previous: int | None = None
    for page, title, level in boundaries:
        match = re.fullmatch(r"Chapter\s+([A-Za-z0-9]+)", title, re.IGNORECASE)
        if not match:
            result.append((page, title, level))
            continue

        token = match.group(1).upper()
        value = int(token) if token.isdigit() else _roman_to_int(token)
        if previous is not None and not token.isdigit():
            expected = previous + 1
            expected_token = _int_to_roman(expected)
            # Preserve valid modest jumps (a source may omit chapters), but repair
            # invalid OCR or implausible values when the expected numeral is close.
            implausible = value is None or value <= previous or value > previous + 5
            if implausible and _edit_distance(token, expected_token) <= 2:
                value = expected
                token = expected_token
        if value is not None:
            previous = value
            token = str(value) if match.group(1).isdigit() else _int_to_roman(value)
        result.append((page, f"Chapter {token}", level))
    return result


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
        lines: list[tuple[float, float, str]] = []
        span_sizes: list[float] = []
        candidates: list[tuple[float, float, str]] = []
        data = page.get_text("dict")
        for block in data.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                spans = line.get("spans", [])
                text = "".join(span.get("text", "") for span in spans).strip()
                sizes = [float(span.get("size", 0)) for span in spans if float(span.get("size", 0)) > 0]
                span_sizes.extend(sizes)
                if not text:
                    continue
                y0 = float(line.get("bbox", (0, 0, 0, 0))[1])
                size = max(sizes, default=0)
                if y0 <= page_height * 0.45:
                    lines.append((y0, size, text))

        typical_size = median(span_sizes) if span_sizes else 0
        minimum_heading_size = max(11.5, typical_size * 1.25)
        lines.sort(key=lambda item: item[0])
        for index, (y0, size, text) in enumerate(lines):
            if size < minimum_heading_size:
                continue

            chapter_token = _chapter_token(text)
            if chapter_token:
                candidates.append((size, -y0, f"Chapter {chapter_token}"))
                continue

            if chapter_token == "":
                for next_y, _, next_text in lines[index + 1:index + 3]:
                    if next_y - y0 > max(36, size * 2):
                        break
                    number = CHAPTER_NUMBER_RE.fullmatch(next_text)
                    if number:
                        candidates.append((size, -y0, f"Chapter {number.group(1).rstrip('.')}"))
                        break
                continue

            if INDEX_RE.fullmatch(text):
                candidates.append((size, -y0, "Index"))
                continue

            if HEADING_RE.match(text):
                candidates.append((size, -y0, text))
        if candidates:
            _, _, title = max(candidates)
            boundaries.append((page_index + 1, title, 1))

    deduplicated: list[tuple[int, str, int]] = []
    seen_index = False
    for boundary in boundaries:
        if boundary[1].casefold() == "index":
            if seen_index:
                continue
            seen_index = True
        deduplicated.append(boundary)
    normalized = _normalize_ocr_chapter_sequence(deduplicated)
    return normalized if len(normalized) >= 2 else []


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
