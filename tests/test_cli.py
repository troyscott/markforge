from __future__ import annotations

import json

import pytest

from markforge.cli import build_parser, main, parse_toc_level


def test_convert_defaults():
    args = build_parser().parse_args(["convert", "book.pdf", "--output", "out"])
    assert args.split == "chapters"
    assert args.processing_pages == 20
    assert args.output_pages == 25
    assert args.device == "auto"
    assert not args.combined


def test_toc_level_parser():
    assert parse_toc_level("auto") is None
    assert parse_toc_level("2") == 2
    with pytest.raises(Exception):
        parse_toc_level("0")


def test_convert_returns_nonzero_when_manifest_has_failed_segment(monkeypatch, tmp_path):
    manifest_path = tmp_path / "manifest.json"
    manifest_path.write_text(json.dumps({"segments": [{"status": "failed"}]}))
    monkeypatch.setattr(
        "markforge.cli.ConversionEngine.convert",
        lambda self, source, options: [manifest_path],
    )

    assert main(["convert", "book.pdf", "--output", str(tmp_path)]) == 2
