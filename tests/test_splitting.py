from __future__ import annotations

import pymupdf

from markforge.splitting import propose_segments, slugify, validate_page_coverage


def test_bookmarks_create_chapters_and_front_matter(tmp_path, make_pdf):
    source = make_pdf(
        tmp_path / "book.pdf",
        12,
        toc=[[1, "Chapter 1 Overview", 3], [1, "Chapter 2 Storage", 8]],
    )
    with pymupdf.open(source) as document:
        segments, strategy = propose_segments(document)
    assert strategy == "bookmarks"
    assert [(item.title, item.start_page, item.end_page) for item in segments] == [
        ("Front Matter", 1, 2),
        ("Chapter 1 Overview", 3, 7),
        ("Chapter 2 Storage", 8, 12),
    ]
    assert validate_page_coverage(segments, 12)


def test_detected_headings_are_used_without_bookmarks(tmp_path, make_pdf):
    source = make_pdf(
        tmp_path / "headings.pdf",
        10,
        headings={1: "Chapter 1 Start", 6: "Chapter 2 Continue"},
    )
    with pymupdf.open(source) as document:
        segments, strategy = propose_segments(document)
    assert strategy == "headings"
    assert [(item.start_page, item.end_page) for item in segments] == [(1, 5), (6, 10)]


def test_fixed_page_fallback_covers_every_page(tmp_path, make_pdf):
    source = make_pdf(tmp_path / "plain.pdf", 11)
    with pymupdf.open(source) as document:
        segments, strategy = propose_segments(document, output_pages=4)
    assert strategy == "pages-fallback"
    assert [(item.start_page, item.end_page) for item in segments] == [(1, 4), (5, 8), (9, 11)]
    assert validate_page_coverage(segments, 11)


def test_duplicate_titles_get_unique_filenames(tmp_path, make_pdf):
    source = make_pdf(
        tmp_path / "duplicate.pdf",
        8,
        toc=[[1, "Chapter 1", 1], [1, "Chapter 1", 5]],
    )
    with pymupdf.open(source) as document:
        segments, _ = propose_segments(document)
    assert [item.markdown_path for item in segments] == ["01-chapter-1.md", "02-chapter-1-2.md"]


def test_slugify_is_portable():
    assert slugify("  OneLake: Security & Governance  ") == "onelake-security-governance"
