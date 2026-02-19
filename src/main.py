from __future__ import annotations

import argparse
import csv
import json
import zipfile
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable, List

from openpyxl import load_workbook

EXCEL_EXTENSIONS = {".xls", ".xlsx", ".xlsm"}
VOLATILE_FUNCTIONS = {
    "NOW(",
    "TODAY(",
    "RAND(",
    "RANDBETWEEN(",
    "OFFSET(",
    "INDIRECT(",
    "CELL(",
    "INFO(",
}


@dataclass
class FileReport:
    path: str
    extension: str
    size_bytes: int
    modified_utc: str
    has_macro: bool
    vba_module_count: int
    sheet_count: int
    formula_count: int
    volatile_formula_count: int
    named_range_count: int
    external_link_count: int
    risk_score: int
    notes: str


@dataclass
class ScanReport:
    generated_utc: str
    input_root: str
    file_count: int
    files: List[FileReport]


def find_excel_files(root: Path) -> Iterable[Path]:
    for file_path in root.rglob("*"):
        if file_path.is_file() and file_path.suffix.lower() in EXCEL_EXTENSIONS:
            yield file_path


def contains_vba_project(file_path: Path) -> bool:
    if file_path.suffix.lower() == ".xlsm":
        return True
    if file_path.suffix.lower() != ".xlsx":
        return False

    try:
        with zipfile.ZipFile(file_path, "r") as zf:
            names = set(zf.namelist())
        return "xl/vbaProject.bin" in names
    except zipfile.BadZipFile:
        return False


def try_count_vba_modules(file_path: Path) -> int:
    try:
        from oletools.olevba import VBA_Parser  # type: ignore
    except Exception:
        return -1

    try:
        parser = VBA_Parser(str(file_path))
        if not parser.detect_vba_macros():
            return 0
        module_count = 0
        for _ in parser.extract_macros():
            module_count += 1
        parser.close()
        return module_count
    except Exception:
        return -1


def count_external_links(workbook) -> int:
    links = getattr(workbook, "_external_links", None)
    if links is None:
        return 0
    return len(links)


def analyze_openxml_file(file_path: Path, has_macro: bool, vba_module_count: int) -> FileReport:
    workbook = load_workbook(filename=file_path, data_only=False, keep_vba=True)
    try:
        formula_count = 0
        volatile_formula_count = 0
        for sheet in workbook.worksheets:
            for row in sheet.iter_rows(values_only=False):
                for cell in row:
                    value = cell.value
                    if isinstance(value, str) and value.startswith("="):
                        formula_count += 1
                        upper_formula = value.upper()
                        if any(fn in upper_formula for fn in VOLATILE_FUNCTIONS):
                            volatile_formula_count += 1

        named_range_count = len(workbook.defined_names)
        external_link_count = count_external_links(workbook)
    finally:
        workbook.close()

    risk_score = 0
    notes = []
    if has_macro:
        risk_score += 40
        notes.append("Macro enabled")
    if volatile_formula_count > 0:
        risk_score += 15
        notes.append("Volatile formulas detected")
    if external_link_count > 0:
        risk_score += 20
        notes.append("External links detected")
    if formula_count > 10000:
        risk_score += 10
        notes.append("Large formula count")

    stat = file_path.stat()
    modified = datetime.fromtimestamp(stat.st_mtime, tz=timezone.utc).isoformat()

    return FileReport(
        path=str(file_path),
        extension=file_path.suffix.lower(),
        size_bytes=stat.st_size,
        modified_utc=modified,
        has_macro=has_macro,
        vba_module_count=vba_module_count,
        sheet_count=len(workbook.sheetnames),
        formula_count=formula_count,
        volatile_formula_count=volatile_formula_count,
        named_range_count=named_range_count,
        external_link_count=external_link_count,
        risk_score=min(risk_score, 100),
        notes="; ".join(notes) if notes else "Low risk",
    )


def analyze_legacy_xls(file_path: Path, has_macro: bool, vba_module_count: int) -> FileReport:
    stat = file_path.stat()
    modified = datetime.fromtimestamp(stat.st_mtime, tz=timezone.utc).isoformat()

    risk_score = 30
    notes = "Legacy .xls format; convert to .xlsx for deep analysis"
    if has_macro:
        risk_score += 40
        notes += "; Macro likely present"

    return FileReport(
        path=str(file_path),
        extension=file_path.suffix.lower(),
        size_bytes=stat.st_size,
        modified_utc=modified,
        has_macro=has_macro,
        vba_module_count=vba_module_count,
        sheet_count=0,
        formula_count=0,
        volatile_formula_count=0,
        named_range_count=0,
        external_link_count=0,
        risk_score=min(risk_score, 100),
        notes=notes,
    )


def analyze_file(file_path: Path) -> FileReport:
    has_macro = contains_vba_project(file_path)
    vba_module_count = try_count_vba_modules(file_path)

    if file_path.suffix.lower() == ".xls":
        return analyze_legacy_xls(file_path, has_macro, vba_module_count)

    try:
        return analyze_openxml_file(file_path, has_macro, vba_module_count)
    except Exception as exc:
        stat = file_path.stat()
        modified = datetime.fromtimestamp(stat.st_mtime, tz=timezone.utc).isoformat()
        return FileReport(
            path=str(file_path),
            extension=file_path.suffix.lower(),
            size_bytes=stat.st_size,
            modified_utc=modified,
            has_macro=has_macro,
            vba_module_count=vba_module_count,
            sheet_count=0,
            formula_count=0,
            volatile_formula_count=0,
            named_range_count=0,
            external_link_count=0,
            risk_score=80,
            notes=f"Analysis failed: {exc}",
        )


def write_json_report(report: ScanReport, output_dir: Path) -> Path:
    output_path = output_dir / "report.json"
    payload = {
        "generated_utc": report.generated_utc,
        "input_root": report.input_root,
        "file_count": report.file_count,
        "files": [asdict(item) for item in report.files],
    }
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return output_path


def write_csv_report(report: ScanReport, output_dir: Path) -> Path:
    output_path = output_dir / "report.csv"
    with output_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=list(asdict(report.files[0]).keys()) if report.files else list(asdict(FileReport(
                path="",
                extension="",
                size_bytes=0,
                modified_utc="",
                has_macro=False,
                vba_module_count=0,
                sheet_count=0,
                formula_count=0,
                volatile_formula_count=0,
                named_range_count=0,
                external_link_count=0,
                risk_score=0,
                notes="",
            )).keys()),
        )
        writer.writeheader()
        for item in report.files:
            writer.writerow(asdict(item))
    return output_path


def run_scan(input_root: Path, output_dir: Path) -> ScanReport:
    files = list(find_excel_files(input_root))
    reports = [analyze_file(path) for path in files]

    scan_report = ScanReport(
        generated_utc=datetime.now(tz=timezone.utc).isoformat(),
        input_root=str(input_root),
        file_count=len(reports),
        files=reports,
    )

    output_dir.mkdir(parents=True, exist_ok=True)
    write_json_report(scan_report, output_dir)
    write_csv_report(scan_report, output_dir)
    return scan_report


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Excel migration inventory CLI")
    subparsers = parser.add_subparsers(dest="command", required=True)

    scan_parser = subparsers.add_parser("scan", help="Scan and analyze Excel files")
    scan_parser.add_argument("--input", required=True, help="Root folder to scan")
    scan_parser.add_argument("--output", required=True, help="Folder for report outputs")

    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    if args.command == "scan":
        input_root = Path(args.input).expanduser().resolve()
        output_dir = Path(args.output).expanduser().resolve()
        report = run_scan(input_root, output_dir)
        print(f"Scanned files: {report.file_count}")
        print(f"JSON report: {output_dir / 'report.json'}")
        print(f"CSV report: {output_dir / 'report.csv'}")
        return 0

    parser.print_help()
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
