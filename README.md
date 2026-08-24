# MarkForge

MarkForge is a local, cross-platform document-to-Markdown converter. It uses Marker for high-fidelity PDF layout, tables, equations, OCR, and images, and MarkItDown for Word, PowerPoint, and Excel files.

Large PDFs can be exported as NotebookLM-friendly chapter files while smaller internal page chunks keep GPU and unified-memory use predictable. Documents stay on your computer: MarkForge has no cloud extraction, telemetry, or API-key requirement.

## Highlights

- Chapter discovery from PDF bookmarks or visible chapter headings
- Fixed-page fallback when a PDF has no usable structure
- Separate logical chapters and memory-safe processing chunks
- Apple Metal (MPS), NVIDIA CUDA, and CPU support
- Extracted images with portable relative Markdown links
- Per-document `manifest.json` with page coverage, hashes, status, and errors
- Resume, forced reprocessing, adaptive chunk splitting, and CPU fallback
- CLI and CustomTkinter desktop interfaces backed by one conversion engine
- DOCX, PPTX, XLSX, TXT, and recursive folder conversion

## Requirements

- macOS, Windows, or Linux
- [`uv`](https://docs.astral.sh/uv/)
- Sufficient free disk space for local model downloads and converted images

MarkForge pins Python 3.12. Do not use the macOS system Python or Micromamba for this project.

## Setup with uv

The distribution is named `markforge-docs`; the installed commands and Python
package remain `markforge`. Install the current GitHub prerelease directly:

```shell
uv tool install https://github.com/troyscott/markforge/releases/download/v0.2.1/markforge_docs-0.2.1-py3-none-any.whl
markforge --help
```

For development from a repository checkout:

```shell
git clone https://github.com/troyscott/markforge.git
cd markforge
uv python install 3.12
uv sync --all-extras
```

The first Marker conversion downloads its local models. Later runs reuse the model cache.

Marker 2 uses a local Surya inference server when a page needs OCR or layout
recovery. On macOS and CPU-only Linux, install the local `llama.cpp` backend:

```shell
brew install llama.cpp
```

Clean digital PDFs may convert without starting this backend, but installing it
prevents difficult pages, equations, and scanned content from failing midway.
No cloud OCR service or API key is used.

Verify the environment:

```shell
uv run markforge --help
uv run pytest
uv lock --check
```

Use `uv run --locked` in repeatable or automated workflows:

```shell
uv sync --locked --all-extras
uv run --locked pytest
```

## Inspect before converting

Inspection does not load Marker models or convert the PDF. It reports page count, selectable-text coverage, bookmarks, proposed chapter files, and the detected device.

```shell
uv run markforge inspect /path/to/book.pdf
```

For an unusual bookmark hierarchy:

```shell
uv run markforge inspect /path/to/book.pdf --toc-level 2
```

## Convert

```shell
uv run markforge convert /path/to/book.pdf --output /path/to/markdown
```

Defaults:

- Chapter-oriented output
- Bookmarks, then detected headings, then 25-page output groups
- 20-page internal processing chunks
- Automatic MPS, CUDA, or CPU selection
- Resume enabled
- No combined full-book file

Useful options:

```shell
# Always create 30-page Markdown files
uv run markforge convert book.pdf --output out --split pages --output-pages 30

# Create chapter files plus an optional combined copy
uv run markforge convert book.pdf --output out --combined

# Force CPU or reprocess completed chapters
uv run markforge convert book.pdf --output out --device cpu --force

# Fast embedded-text conversion when layout/OCR is unnecessary
uv run markforge convert book.pdf --output out --extractor native

# Recursively convert a folder
uv run markforge convert documents --output converted
```

Every source receives its own output directory. A PDF named `fabric-book.pdf` produces:

```text
out/
  fabric-book/
    manifest.json
    01-front-matter.md
    02-chapter-1-introduction.md
    03-chapter-2-storage.md
    images/
      fabric-book/
```

Physical PDF pages in the manifest are one-based. Long chapters remain one Markdown file even when MarkForge processes them in several smaller PDF chunks.

## Desktop application

```shell
uv run markforge gui
```

The desktop interface supports file or folder input, structure preview, device selection, chapter/page/single output, progress, optional combined output, and cancellation between processing chunks.

Keyboard shortcuts keep the primary workflow available without relying on the mouse:

| Action | macOS | Windows and Linux |
| --- | --- | --- |
| Preview structure | `Command-P` | `Ctrl-P` |
| Start conversion | `Command-R` | `Ctrl-R` |

## Device behavior

Marker chooses the device in this order: CUDA, Apple MPS, then CPU. Override it with `--device` when diagnosing a conversion.

### macOS / Apple Silicon

- Use the uv-managed arm64 Python 3.12 environment.
- Install `llama.cpp` with Homebrew so Marker can start its local OCR server.
- Keep processing at the 20-page default initially.
- If an MPS operation fails, MarkForge reduces the chunk size and retries the smallest failed chunk on CPU.

### Windows / NVIDIA

- Install current NVIDIA drivers before `uv sync`.
- Confirm detection with `uv run markforge inspect sample.pdf`.
- MarkForge preserves CUDA support, but a release should only claim CUDA verification after a real NVIDIA smoke test.

### CPU

CPU conversion is supported but slower. Use `--device cpu` when acceleration is unavailable or for troubleshooting. Marker OCR on CPU also requires the `llama-server` binary supplied by `llama.cpp`.

## Resume and safety

- Source documents are never modified or deleted.
- Temporary PDFs live in MarkForge-owned per-run directories under the output document.
- Markdown and manifests are replaced atomically.
- A command exits non-zero when any segment fails, while preserving successful output and failure details for retry.
- Successful chapters are skipped when the source hash and conversion configuration match.
- Changed inputs or settings invalidate the previous run.
- `cleanup.py OUTPUT` removes only stale `.markforge-*` temporary directories.

Keep copyrighted source documents and converted output private. The repository ignores `in/`, `out/`, local environments, caches, temporary pieces, PDFs, and Office documents.

## Dependency maintenance

Use uv rather than editing the lockfile:

```shell
uv add PACKAGE
uv remove PACKAGE
uv lock --upgrade-package PACKAGE
uv lock --check
```

Commit both `pyproject.toml` and `uv.lock` after a reviewed dependency change.

## Development

```shell
uv sync --all-extras
uv run pytest
uv run markforge inspect tests/fixtures/example.pdf
```

Routine CI uses synthetic PDFs and mocked extractors so it does not download production models. Local integration testing validates Marker on real hardware.

## License

MIT
