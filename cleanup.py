"""Remove only stale MarkForge-owned temporary run directories."""

import argparse
from pathlib import Path

from markforge.engine import cleanup_stale_temporary_directories


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("output", nargs="?", type=Path, default=Path("out"))
    args = parser.parse_args()
    removed = cleanup_stale_temporary_directories(args.output)
    noun = "directory" if len(removed) == 1 else "directories"
    print(f"Removed {len(removed)} MarkForge temporary {noun}.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
