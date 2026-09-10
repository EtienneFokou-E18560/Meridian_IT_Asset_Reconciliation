#!/usr/bin/env python3
"""Fail-closed deliverable gate: require Yanou_IT_Asset_Reconciliation.xlsx by name and OOXML type.

Does not pass on aggregate output size alone.
"""
from __future__ import annotations

import argparse
import sys
import zipfile
from pathlib import Path

REQUIRED_NAME = "Yanou_IT_Asset_Reconciliation.xlsx"
OOXML_MARKERS = ("[Content_Types].xml", "xl/workbook.xml")


def is_ooxml_xlsx(path: Path) -> bool:
    if path.suffix.lower() != ".xlsx":
        return False
    try:
        with path.open("rb") as f:
            if f.read(2) != b"PK":
                return False
        with zipfile.ZipFile(path, "r") as zf:
            names = set(zf.namelist())
            return all(m in names for m in OOXML_MARKERS)
    except (OSError, zipfile.BadZipFile):
        return False


def find_required(outputs: Path) -> Path | None:
    direct = outputs / REQUIRED_NAME
    if direct.is_file():
        return direct
    # Single zip that contains the required workbook at top level only
    zips = sorted(
        p for p in outputs.iterdir() if p.is_file() and p.suffix.lower() == ".zip"
    )
    if len(zips) == 1:
        try:
            with zipfile.ZipFile(zips[0], "r") as zf:
                if REQUIRED_NAME in zf.namelist():
                    extracted = outputs / "_check_outputs_extracted" / REQUIRED_NAME
                    extracted.parent.mkdir(parents=True, exist_ok=True)
                    extracted.write_bytes(zf.read(REQUIRED_NAME))
                    return extracted
        except (OSError, zipfile.BadZipFile):
            return None
    return None


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "outputs",
        nargs="?",
        default=".",
        type=Path,
        help="Directory to search for the graded deliverable (default: cwd)",
    )
    args = parser.parse_args()
    outputs = args.outputs.resolve()
    if not outputs.is_dir():
        print(f"FAIL: outputs directory not found: {outputs}", file=sys.stderr)
        return 1

    found = find_required(outputs)
    if found is None:
        print(
            f"FAIL: required deliverable {REQUIRED_NAME!r} not found under {outputs}",
            file=sys.stderr,
        )
        return 1
    if not is_ooxml_xlsx(found):
        print(
            f"FAIL: {found.name} is not a valid OOXML .xlsx workbook",
            file=sys.stderr,
        )
        return 1

    print(f"OK: {found} ({found.stat().st_size} bytes, OOXML)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
