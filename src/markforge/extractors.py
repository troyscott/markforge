from __future__ import annotations

import os
from pathlib import Path
from typing import Protocol

import pymupdf

from .models import ExtractionResult


class Extractor(Protocol):
    name: str
    device: str

    def extract(self, source: Path) -> ExtractionResult: ...


def detect_device() -> str:
    try:
        import torch

        if torch.cuda.is_available():
            return "cuda"
        if torch.backends.mps.is_available():
            return "mps"
    except (ImportError, AttributeError):
        pass
    return "cpu"


def clear_device_cache(device: str) -> None:
    try:
        import torch

        if device == "cuda" and torch.cuda.is_available():
            torch.cuda.empty_cache()
        elif device == "mps" and torch.backends.mps.is_available():
            torch.mps.empty_cache()
    except (ImportError, AttributeError, RuntimeError):
        pass


class MarkerPdfExtractor:
    name = "marker"

    def __init__(self, device: str = "auto") -> None:
        self.device = detect_device() if device == "auto" else device
        self._converter = None

    def _load(self):
        if self._converter is None:
            os.environ["TORCH_DEVICE"] = self.device
            from marker.converters.pdf import PdfConverter
            from marker.models import create_model_dict
            from marker.settings import settings

            settings.TORCH_DEVICE = self.device
            self._converter = PdfConverter(artifact_dict=create_model_dict())
        return self._converter

    def extract(self, source: Path) -> ExtractionResult:
        from marker.output import text_from_rendered

        rendered = self._load()(str(source))
        text, metadata, images = text_from_rendered(rendered)
        return ExtractionResult(
            markdown=text or "",
            images=images or {},
            metadata=metadata if isinstance(metadata, dict) else {"marker_metadata": metadata},
        )


class NativePdfExtractor:
    """Fast embedded-text extractor used as an explicit local fallback."""

    name = "native"

    def __init__(self, device: str = "cpu") -> None:
        self.device = "cpu"

    def extract(self, source: Path) -> ExtractionResult:
        with pymupdf.open(source) as document:
            pages = [page.get_text("text").strip() for page in document]
        return ExtractionResult(markdown="\n\n".join(page for page in pages if page))


class OfficeExtractor:
    name = "markitdown"
    device = "cpu"

    def __init__(self) -> None:
        self._converter = None

    def extract(self, source: Path) -> ExtractionResult:
        if self._converter is None:
            from markitdown import MarkItDown

            self._converter = MarkItDown(enable_plugins=False)
        result = self._converter.convert(str(source))
        return ExtractionResult(markdown=result.text_content or "")
