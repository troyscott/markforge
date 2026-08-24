from __future__ import annotations

import gc
import re
import shutil
import tempfile
import threading
from dataclasses import replace
from pathlib import Path
from typing import Callable

import pymupdf

from .extractors import (
    MarkerPdfExtractor,
    NativePdfExtractor,
    OfficeExtractor,
    clear_device_cache,
    detect_device,
)
from .manifest import (
    build_manifest,
    configuration_fingerprint,
    read_manifest,
    sha256_file,
    write_json_atomic,
    write_text_atomic,
)
from .models import (
    ConversionOptions,
    ExtractionResult,
    InspectionReport,
    ProgressCallback,
    Segment,
    SegmentResult,
)
from .splitting import propose_segments, slugify, validate_page_coverage


SUPPORTED_EXTENSIONS = {".pdf", ".docx", ".pptx", ".xlsx", ".txt"}


class ConversionCancelled(RuntimeError):
    pass


class ConversionEngine:
    def __init__(
        self,
        extractor_factory: Callable[[str, str], object] | None = None,
        progress: ProgressCallback | None = None,
        cancel_event: threading.Event | None = None,
    ) -> None:
        self.extractor_factory = extractor_factory or self._default_extractor_factory
        self.progress = progress or (lambda event: None)
        self.cancel_event = cancel_event or threading.Event()
        self._extractors: dict[tuple[str, str], object] = {}

    @staticmethod
    def _default_extractor_factory(kind: str, device: str):
        if kind == "marker":
            return MarkerPdfExtractor(device)
        if kind == "native":
            return NativePdfExtractor()
        if kind == "office":
            return OfficeExtractor()
        raise ValueError(f"Unknown extractor: {kind}")

    def _get_extractor(self, kind: str, device: str):
        key = (kind, device)
        if key not in self._extractors:
            self._extractors[key] = self.extractor_factory(kind, device)
        return self._extractors[key]

    def _check_cancelled(self) -> None:
        if self.cancel_event.is_set():
            raise ConversionCancelled("Conversion cancelled")

    def inspect(self, source: Path, options: ConversionOptions | None = None) -> InspectionReport:
        source = Path(source)
        options = options or ConversionOptions(output_dir=Path("out"))
        options.validate()
        if source.suffix.lower() != ".pdf":
            raise ValueError("inspect currently supports PDF files")
        with pymupdf.open(source) as document:
            encrypted = bool(document.needs_pass)
            if encrypted:
                raise ValueError("The PDF is encrypted and requires a password")
            segments, _ = propose_segments(
                document, options.split, options.output_pages, options.toc_level
            )
            text_pages = sum(bool(page.get_text("text").strip()) for page in document)
            return InspectionReport(
                source=str(source),
                page_count=len(document),
                text_pages=text_pages,
                bookmarks=document.get_toc(simple=True),
                proposed_segments=segments,
                selected_device=detect_device() if options.device == "auto" else options.device,
                encrypted=encrypted,
            )

    def convert(self, source: Path, options: ConversionOptions) -> list[Path]:
        source = Path(source)
        options.validate()
        if not source.exists():
            raise FileNotFoundError(source)
        if source.is_file():
            return [self._convert_file(source, options, Path(source.stem))]

        outputs: list[Path] = []
        for path in sorted(item for item in source.rglob("*") if item.is_file()):
            if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
                continue
            relative = path.relative_to(source)
            outputs.append(self._convert_file(path, options, relative.parent / path.stem))
        return outputs

    def _convert_file(self, source: Path, options: ConversionOptions, relative_output: Path) -> Path:
        self._check_cancelled()
        document_output = options.output_dir / relative_output
        document_output.mkdir(parents=True, exist_ok=True)
        self.progress({"type": "document_started", "source": str(source)})
        if source.suffix.lower() == ".pdf":
            manifest_path = self._convert_pdf(source, document_output, options)
        else:
            manifest_path = self._convert_non_pdf(source, document_output, options)
        self.progress({"type": "document_completed", "source": str(source), "manifest": str(manifest_path)})
        return manifest_path

    def _convert_pdf(self, source: Path, output: Path, options: ConversionOptions) -> Path:
        with pymupdf.open(source) as document:
            if document.needs_pass:
                raise ValueError(f"Encrypted PDF is not supported: {source}")
            page_count = len(document)
            segments, split_strategy = propose_segments(
                document, options.split, options.output_pages, options.toc_level
            )
            if not validate_page_coverage(segments, page_count):
                raise RuntimeError("Proposed segments do not cover every PDF page exactly once")

        source_hash = sha256_file(source)
        selected_device = (
            "cpu"
            if options.extractor == "native"
            else detect_device() if options.device == "auto" else options.device
        )
        fingerprint = configuration_fingerprint(
            replace(options, device=selected_device), split_strategy
        )
        manifest_path = output / "manifest.json"
        previous = read_manifest(manifest_path)
        previous_segments = {
            item.get("id"): item for item in (previous or {}).get("segments", [])
        }
        can_resume = bool(
            previous
            and previous.get("source", {}).get("sha256") == source_hash
            and previous.get("run", {}).get("configuration_fingerprint") == fingerprint
        )
        results: list[SegmentResult] = []
        image_root = output / "images" / slugify(source.stem, "document")
        image_root.mkdir(parents=True, exist_ok=True)

        with tempfile.TemporaryDirectory(prefix=".markforge-", dir=output) as temp_name:
            temp_dir = Path(temp_name)
            for index, segment in enumerate(segments, start=1):
                self._check_cancelled()
                prior = previous_segments.get(segment.id)
                output_path = output / segment.markdown_path
                if (
                    can_resume
                    and not options.force
                    and prior
                    and prior.get("status") in {"success", "skipped"}
                    and output_path.exists()
                ):
                    result = SegmentResult(
                        segment=segment,
                        status="skipped",
                        character_count=int(prior.get("character_count", 0)),
                        image_count=int(prior.get("image_count", 0)),
                        warnings=list(prior.get("warnings", [])),
                    )
                    results.append(result)
                    self._write_manifest_progress(
                        manifest_path, source, source_hash, page_count, options,
                        selected_device, split_strategy, fingerprint, results,
                    )
                    continue

                self.progress({
                    "type": "segment_started",
                    "index": index,
                    "total": len(segments),
                    "title": segment.title,
                    "start_page": segment.start_page,
                    "end_page": segment.end_page,
                })
                try:
                    extracted = self._extract_segment(
                        source, temp_dir=temp_dir,
                        segment=segment, options=options, device=selected_device,
                    )
                    markdown, image_count, image_warnings = self._save_images(
                        extracted, image_root, output, segment
                    )
                    content = (
                        f"# {segment.title}\n\n"
                        f"<!-- Source: {source.name}; PDF pages {segment.start_page}-{segment.end_page} -->\n\n"
                        f"{markdown.strip()}\n"
                    )
                    warnings = list(image_warnings)
                    if not markdown.strip():
                        warnings.append("No text was extracted")
                    write_text_atomic(output_path, content)
                    result = SegmentResult(
                        segment=segment,
                        status="success",
                        character_count=len(markdown),
                        image_count=image_count,
                        warnings=warnings,
                    )
                except ConversionCancelled:
                    result = SegmentResult(segment=segment, status="cancelled", error="Conversion cancelled")
                    results.append(result)
                    self._write_manifest_progress(
                        manifest_path, source, source_hash, page_count, options,
                        selected_device, split_strategy, fingerprint, results,
                    )
                    raise
                except Exception as exc:
                    result = SegmentResult(segment=segment, status="failed", error=str(exc))

                results.append(result)
                self._write_manifest_progress(
                    manifest_path, source, source_hash, page_count, options,
                    selected_device, split_strategy, fingerprint, results,
                )
                self.progress({"type": "segment_completed", "title": segment.title, "status": result.status})

        if options.combined:
            successful = [item for item in results if item.status in {"success", "skipped"}]
            combined = "\n\n".join(
                (output / item.segment.markdown_path).read_text(encoding="utf-8").strip()
                for item in successful
                if (output / item.segment.markdown_path).exists()
            )
            write_text_atomic(output / f"{slugify(source.stem)}-combined.md", combined + "\n")
        return manifest_path

    def _write_manifest_progress(
        self, manifest_path: Path, source: Path, source_hash: str, page_count: int,
        options: ConversionOptions, device: str, strategy: str, fingerprint: str,
        results: list[SegmentResult],
    ) -> None:
        manifest = build_manifest(
            source, source_hash, page_count, options, device, strategy, fingerprint, results
        )
        write_json_atomic(manifest_path, manifest)

    def _extract_segment(
        self, source: Path, temp_dir: Path, segment: Segment,
        options: ConversionOptions, device: str,
    ) -> ExtractionResult:
        parts: list[str] = []
        images: dict[str, object] = {}
        for start_page in range(segment.start_page, segment.end_page + 1, options.processing_pages):
            end_page = min(start_page + options.processing_pages - 1, segment.end_page)
            result = self._extract_range_resilient(
                source, temp_dir, start_page, end_page, options.extractor, device
            )
            parts.append(result.markdown)
            images.update(result.images)
        return ExtractionResult(markdown="\n\n".join(parts), images=images)

    def _extract_range_resilient(
        self, source: Path, temp_dir: Path, start_page: int, end_page: int,
        extractor_kind: str, device: str,
    ) -> ExtractionResult:
        self._check_cancelled()
        try:
            chunk_path = self._write_pdf_range(source, temp_dir, start_page, end_page)
            extractor = self._get_extractor(extractor_kind, device)
            result = extractor.extract(chunk_path)
            return self._namespace_images(result, start_page, end_page)
        except ConversionCancelled:
            raise
        except Exception:
            length = end_page - start_page + 1
            if length > 5:
                midpoint = start_page + (length // 2) - 1
                left = self._extract_range_resilient(
                    source, temp_dir, start_page, midpoint, extractor_kind, device
                )
                right = self._extract_range_resilient(
                    source, temp_dir, midpoint + 1, end_page, extractor_kind, device
                )
                return ExtractionResult(
                    markdown=f"{left.markdown}\n\n{right.markdown}",
                    images={**left.images, **right.images},
                )
            if device != "cpu":
                chunk_path = self._write_pdf_range(source, temp_dir, start_page, end_page)
                result = self._get_extractor(extractor_kind, "cpu").extract(chunk_path)
                return self._namespace_images(result, start_page, end_page)
            raise
        finally:
            gc.collect()
            clear_device_cache(device)

    @staticmethod
    def _write_pdf_range(source: Path, temp_dir: Path, start_page: int, end_page: int) -> Path:
        destination = temp_dir / f"pages-{start_page:04d}-{end_page:04d}.pdf"
        if destination.exists():
            return destination
        with pymupdf.open(source) as document, pymupdf.open() as chunk:
            chunk.insert_pdf(document, from_page=start_page - 1, to_page=end_page - 1)
            chunk.save(destination)
        return destination

    @staticmethod
    def _namespace_images(result: ExtractionResult, start_page: int, end_page: int) -> ExtractionResult:
        renamed: dict[str, object] = {}
        markdown = result.markdown
        for index, (name, image) in enumerate(result.images.items(), start=1):
            safe_name = Path(str(name)).name
            new_name = f"p{start_page:04d}-{end_page:04d}-{index:03d}-{safe_name}"
            markdown = markdown.replace(str(name), new_name)
            renamed[new_name] = image
        return ExtractionResult(markdown=markdown, images=renamed, metadata=result.metadata)

    @staticmethod
    def _save_images(
        result: ExtractionResult, image_root: Path, document_output: Path, segment: Segment
    ) -> tuple[str, int, list[str]]:
        markdown = result.markdown
        count = 0
        for name, image in result.images.items():
            destination = image_root / Path(name).name
            image.save(destination)
            relative = destination.relative_to(document_output).as_posix()
            markdown = markdown.replace(str(name), relative)
            count += 1
        empty_images = re.findall(r"!\[([^\]]*)\]\(\s*\)", markdown)
        if empty_images:
            markdown = re.sub(
                r"!\[([^\]]*)\]\(\s*\)",
                lambda match: match.group(1).strip(),
                markdown,
            )
        warnings = (
            [f"Removed {len(empty_images)} empty image reference(s) emitted by the extractor"]
            if empty_images
            else []
        )
        return markdown, count, warnings

    def _convert_non_pdf(self, source: Path, output: Path, options: ConversionOptions) -> Path:
        self._check_cancelled()
        if source.suffix.lower() == ".txt":
            markdown = source.read_text(encoding="utf-8")
            extractor_name = "text"
        else:
            markdown = self._get_extractor("office", "cpu").extract(source).markdown
            extractor_name = "markitdown"
        segment = Segment("segment-001", source.stem, 1, 1, f"01-{slugify(source.stem)}.md")
        write_text_atomic(output / segment.markdown_path, f"# {source.stem}\n\n{markdown.strip()}\n")
        local_options = replace(options, extractor=extractor_name, split="single")
        source_hash = sha256_file(source)
        fingerprint = configuration_fingerprint(replace(local_options, device="cpu"), "single")
        result = SegmentResult(segment, "success", len(markdown), 0)
        manifest = build_manifest(
            source, source_hash, None, local_options, "cpu", "single", fingerprint, [result]
        )
        manifest_path = output / "manifest.json"
        write_json_atomic(manifest_path, manifest)
        return manifest_path


def cleanup_stale_temporary_directories(output: Path) -> list[Path]:
    removed: list[Path] = []
    if not output.exists():
        return removed
    for path in output.rglob(".markforge-*"):
        if path.is_dir():
            shutil.rmtree(path)
            removed.append(path)
    return removed
