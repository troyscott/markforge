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


def test_ocr_roman_chapters_across_lines_ignore_repeated_index_headers(tmp_path):
    source = tmp_path / "historical-scan.pdf"
    with pymupdf.open() as document:
        for page_number in range(1, 13):
            page = document.new_page()
            page.insert_text((72, 140), f"Body text for page {page_number}", fontsize=10)
            if page_number == 2:
                page.insert_text((72, 75), "CHAPTEE", fontsize=20)
                page.insert_text((72, 100), "I.", fontsize=12)
            elif page_number == 6:
                page.insert_text((72, 75), "CHAPTER", fontsize=20)
                page.insert_text((72, 100), "II.", fontsize=12)
            elif page_number == 10:
                page.insert_text((72, 75), "INDEX.", fontsize=20)
            elif page_number in {11, 12}:
                page.insert_text((72, 40), "INDEX.", fontsize=10)
        document.save(source)

    with pymupdf.open(source) as document:
        segments, strategy = propose_segments(document)

    assert strategy == "headings"
    assert [(item.title, item.start_page, item.end_page) for item in segments] == [
        ("Front Matter", 1, 1),
        ("Chapter I", 2, 5),
        ("Chapter II", 6, 9),
        ("Index", 10, 12),
    ]
    assert validate_page_coverage(segments, 12)


def test_damaged_chapter_words_and_roman_numerals_follow_sequence(tmp_path):
    source = tmp_path / "damaged-ocr.pdf"
    noisy_headings = [
        "CHAPTEE I.",
        "CHAPTER II.",
        "CHAPTEE HI.",
        "CHAPTER IV.",
        "CHAPTEE V.",
        "CHAPTEE VI.",
        "CHAPTEB VII.",
        "GHAPTEE VIIL.",
        "CHAPTEE IX.",
        "CHAPTEE X.",
        "CHAPTEK XL.",
        "CHAPTER XII.",
        "CHAPTEK XIII.",
        "CHAPTEE XIY.",
    ]
    with pymupdf.open() as document:
        for page_number, heading in enumerate(noisy_headings, start=1):
            page = document.new_page()
            page.insert_text((72, 75), heading, fontsize=20)
            page.insert_text((72, 140), "Body text", fontsize=10)
        document.save(source)

    with pymupdf.open(source) as document:
        segments, strategy = propose_segments(document)

    assert strategy == "headings"
    assert [item.title for item in segments] == [
        f"Chapter {_roman}" for _roman in
        ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII", "XIII", "XIV"]
    ]
    assert validate_page_coverage(segments, len(noisy_headings))


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
