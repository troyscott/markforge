from __future__ import annotations

import json
from pathlib import Path

import pymupdf

from markforge.engine import ConversionEngine, cleanup_stale_temporary_directories
from markforge.models import ConversionOptions, ExtractionResult, Segment


class FakeImage:
    def save(self, destination: Path) -> None:
        destination.write_bytes(b"image")


class FakeExtractor:
    def __init__(self, device: str, calls: list[tuple[str, int]]) -> None:
        self.device = device
        self.name = "fake"
        self.calls = calls

    def extract(self, source: Path) -> ExtractionResult:
        with pymupdf.open(source) as document:
            pages = len(document)
        self.calls.append((self.device, pages))
        return ExtractionResult(
            markdown=f"Extracted {pages} pages with image.png",
            images={"image.png": FakeImage()},
        )


class SplittingExtractor(FakeExtractor):
    def extract(self, source: Path) -> ExtractionResult:
        with pymupdf.open(source) as document:
            pages = len(document)
        self.calls.append((self.device, pages))
        if pages > 5:
            raise RuntimeError("simulated memory pressure")
        if self.device != "cpu":
            raise RuntimeError("simulated accelerator failure")
        return ExtractionResult(markdown=f"CPU extracted {pages} pages")


def test_conversion_writes_manifest_images_combined_and_resumes(tmp_path, make_pdf):
    source = make_pdf(
        tmp_path / "book.pdf",
        12,
        toc=[[1, "Chapter 1", 1], [1, "Chapter 2", 8]],
    )
    calls: list[tuple[str, int]] = []
    factory = lambda kind, device: FakeExtractor(device, calls)
    options = ConversionOptions(
        output_dir=tmp_path / "out",
        processing_pages=5,
        device="cpu",
        combined=True,
    )
    engine = ConversionEngine(extractor_factory=factory)
    manifest_path = engine.convert(source, options)[0]
    manifest = json.loads(manifest_path.read_text())
    assert [item["status"] for item in manifest["segments"]] == ["success", "success"]
    assert manifest["source"]["page_count"] == 12
    assert manifest["run"]["processing_pages"] == 5
    assert manifest["run"]["device"] == "cpu"
    assert len(calls) == 3
    assert (manifest_path.parent / "book-combined.md").exists()
    for item in manifest["segments"]:
        markdown = (manifest_path.parent / item["markdown_path"]).read_text()
        assert "images/book/" in markdown

    engine.convert(source, options)
    resumed = json.loads(manifest_path.read_text())
    assert [item["status"] for item in resumed["segments"]] == ["skipped", "skipped"]
    assert len(calls) == 3


def test_failed_large_range_splits_and_falls_back_to_cpu(tmp_path, make_pdf):
    source = make_pdf(tmp_path / "large.pdf", 10)
    calls: list[tuple[str, int]] = []
    factory = lambda kind, device: SplittingExtractor(device, calls)
    options = ConversionOptions(
        output_dir=tmp_path / "out",
        split="single",
        processing_pages=10,
        device="mps",
    )
    manifest_path = ConversionEngine(extractor_factory=factory).convert(source, options)[0]
    manifest = json.loads(manifest_path.read_text())
    assert manifest["segments"][0]["status"] == "success"
    assert ("mps", 10) in calls
    assert calls.count(("cpu", 5)) == 2


def test_cleanup_only_removes_owned_temp_directories(tmp_path):
    owned = tmp_path / "doc" / ".markforge-abcd"
    unrelated = tmp_path / "doc" / "keep"
    owned.mkdir(parents=True)
    unrelated.mkdir()
    removed = cleanup_stale_temporary_directories(tmp_path)
    assert removed == [owned]
    assert not owned.exists()
    assert unrelated.exists()


def test_empty_image_references_are_removed(tmp_path):
    result = ExtractionResult(
        markdown="Before ![useful alt]() after ![]()",
        images={},
    )
    segment = Segment("segment-001", "Chapter", 1, 1, "01-chapter.md")

    markdown, image_count, warnings = ConversionEngine._save_images(
        result, tmp_path / "images", tmp_path, segment
    )

    assert markdown == "Before useful alt after "
    assert image_count == 0
    assert warnings == ["Removed 2 empty image reference(s) emitted by the extractor"]


def test_document_html_is_removed_and_unbalanced_inline_tags_warn(tmp_path):
    result = ExtractionResult(
        markdown="Before <body><p>converted text</p></body></html></sup>",
        images={},
    )
    segment = Segment("segment-001", "Chapter", 1, 1, "01-chapter.md")

    markdown, image_count, warnings = ConversionEngine._save_images(
        result, tmp_path / "images", tmp_path, segment
    )

    assert markdown == "Before <p>converted text</p></sup>"
    assert image_count == 0
    assert warnings == [
        "Removed 3 document-level HTML tag(s) emitted by the extractor",
        "Unbalanced <sup> tags remain in extracted Markdown; manual review required",
    ]
