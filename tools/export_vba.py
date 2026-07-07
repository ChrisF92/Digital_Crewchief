#!/usr/bin/env python3
"""Export VBA source from the Digital Crewchief workbook into vba/ as .txt files."""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

from pyopenvba import ExcelFile
from pyopenvba.vba import VBAModuleKind

REPO_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_WORKBOOK = REPO_ROOT / "20260603-Digital Crewchief_BETA_MASTER.xlsm"
VBA_ROOT = REPO_ROOT / "vba"
MODULES_DIR = VBA_ROOT / "Modules"
EXCEL_OBJECTS_DIR = VBA_ROOT / "ExcelObjects"
FORMS_DIR = VBA_ROOT / "Forms"
TEXT_EXTENSION = ".txt"

FORM_NAMES = frozenset(
    {
        "frmAircraftManager",
        "frmPlannerSettings",
        "frmTaskRules",
    }
)


def module_destination(module_name: str, kind: VBAModuleKind) -> Path:
    if module_name in FORM_NAMES:
        return FORMS_DIR / f"{module_name}{TEXT_EXTENSION}"
    if kind == VBAModuleKind.standard:
        return MODULES_DIR / f"{module_name}{TEXT_EXTENSION}"
    return EXCEL_OBJECTS_DIR / f"{module_name}{TEXT_EXTENSION}"


def export_workbook(workbook_path: Path, *, clean: bool = True) -> list[Path]:
    if clean:
        for directory in (MODULES_DIR, EXCEL_OBJECTS_DIR, FORMS_DIR):
            if directory.exists():
                shutil.rmtree(directory)
            directory.mkdir(parents=True, exist_ok=True)
    else:
        for directory in (MODULES_DIR, EXCEL_OBJECTS_DIR, FORMS_DIR):
            directory.mkdir(parents=True, exist_ok=True)

    written: list[Path] = []
    with ExcelFile(workbook_path) as workbook:
        for module in workbook.vba_project().modules:
            target = module_destination(module.name, module.kind)
            text = module.source.replace("\r\n", "\n").replace("\r", "\n")
            data = text.replace("\n", "\r\n").encode("utf-8")
            target.write_bytes(data)
            written.append(target)

    return written


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "workbook",
        nargs="?",
        default=str(DEFAULT_WORKBOOK),
        help="Path to the macro-enabled workbook (.xlsm)",
    )
    parser.add_argument(
        "--no-clean",
        action="store_true",
        help="Do not remove existing files before export",
    )
    args = parser.parse_args(argv)

    workbook_path = Path(args.workbook)
    if not workbook_path.is_file():
        print(f"Workbook not found: {workbook_path}", file=sys.stderr)
        return 1

    written = export_workbook(workbook_path, clean=not args.no_clean)
    print(f"Exported {len(written)} VBA components to {VBA_ROOT}")
    for path in sorted(written):
        print(f"  {path.relative_to(REPO_ROOT)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
