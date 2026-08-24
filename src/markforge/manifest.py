from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
from typing import Any

from . import __version__
from .models import ConversionOptions, SegmentResult


SCHEMA_VERSION = "1.0"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def configuration_fingerprint(options: ConversionOptions, split_strategy: str) -> str:
    value = {
        "split": options.split,
        "processing_pages": options.processing_pages,
        "output_pages": options.output_pages,
        "toc_level": options.toc_level,
        "device": options.device,
        "extractor": options.extractor,
        "split_strategy": split_strategy,
        "markforge_version": __version__,
    }
    payload = json.dumps(value, sort_keys=True, separators=(",", ":")).encode()
    return hashlib.sha256(payload).hexdigest()


def read_manifest(path: Path) -> dict[str, Any] | None:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return None


def write_json_atomic(path: Path, value: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.tmp")
    temporary.write_text(json.dumps(value, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    os.replace(temporary, path)


def write_text_atomic(path: Path, value: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.tmp")
    temporary.write_text(value, encoding="utf-8")
    os.replace(temporary, path)


def build_manifest(
    source: Path,
    source_hash: str,
    page_count: int | None,
    options: ConversionOptions,
    selected_device: str,
    split_strategy: str,
    fingerprint: str,
    results: list[SegmentResult],
) -> dict[str, Any]:
    return {
        "schema_version": SCHEMA_VERSION,
        "markforge_version": __version__,
        "source": {
            "filename": source.name,
            "sha256": source_hash,
            "page_count": page_count,
        },
        "run": {
            "extractor": options.extractor,
            "device": selected_device,
            "requested_device": options.device,
            "split": options.split,
            "split_strategy": split_strategy,
            "processing_pages": options.processing_pages,
            "output_pages": options.output_pages,
            "toc_level": options.toc_level,
            "configuration_fingerprint": fingerprint,
        },
        "segments": [result.to_dict() for result in results],
    }
