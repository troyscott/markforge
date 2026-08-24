from __future__ import annotations

from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Callable, Literal


SplitMode = Literal["chapters", "pages", "single"]
Device = Literal["auto", "mps", "cuda", "cpu"]


@dataclass(frozen=True)
class Segment:
    id: str
    title: str
    start_page: int
    end_page: int
    markdown_path: str
    level: int = 1

    @property
    def page_count(self) -> int:
        return self.end_page - self.start_page + 1


@dataclass
class ConversionOptions:
    output_dir: Path
    split: SplitMode = "chapters"
    processing_pages: int = 20
    output_pages: int = 25
    toc_level: int | None = None
    device: Device = "auto"
    combined: bool = False
    force: bool = False
    extractor: str = "marker"

    def validate(self) -> None:
        if self.processing_pages < 5:
            raise ValueError("processing_pages must be at least 5")
        if self.output_pages < 1:
            raise ValueError("output_pages must be at least 1")
        if self.toc_level is not None and self.toc_level < 1:
            raise ValueError("toc_level must be auto or a positive integer")


@dataclass
class ExtractionResult:
    markdown: str
    images: dict[str, Any] = field(default_factory=dict)
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass
class SegmentResult:
    segment: Segment
    status: Literal["success", "failed", "skipped", "cancelled"]
    character_count: int = 0
    image_count: int = 0
    warnings: list[str] = field(default_factory=list)
    error: str | None = None

    def to_dict(self) -> dict[str, Any]:
        value = asdict(self)
        value.update(value.pop("segment"))
        return value


@dataclass
class InspectionReport:
    source: str
    page_count: int
    text_pages: int
    bookmarks: list[list[Any]]
    proposed_segments: list[Segment]
    selected_device: str
    encrypted: bool = False

    def to_dict(self) -> dict[str, Any]:
        return {
            "source": self.source,
            "page_count": self.page_count,
            "text_pages": self.text_pages,
            "bookmarks": self.bookmarks,
            "proposed_segments": [asdict(item) for item in self.proposed_segments],
            "selected_device": self.selected_device,
            "encrypted": self.encrypted,
        }


ProgressCallback = Callable[[dict[str, Any]], None]
