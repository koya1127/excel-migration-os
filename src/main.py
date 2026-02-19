from __future__ import annotations

import argparse
import csv
import json
import zipfile
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable, List, Optional

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
DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive.file"]


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


@dataclass
class UploadResult:
    local_path: str
    file_name: str
    source_mime_type: str
    target_mime_type: str
    drive_file_id: str
    drive_web_view_link: str
    status: str
    error: str


@dataclass
class UploadReport:
    generated_utc: str
    input_root: str
    output_root: str
    file_count: int
    success_count: int
    failure_count: int
    converted_to_sheets: bool
    drive_folder_id: str
    files: List[UploadResult]


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
    except (zipfile.BadZipFile, OSError):
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
        sheet_count = len(workbook.sheetnames)
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
        sheet_count=sheet_count,
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
            fieldnames=list(asdict(report.files[0]).keys())
            if report.files
            else list(
                asdict(
                    FileReport(
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
                    )
                ).keys()
            ),
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


def load_drive_service(credentials_path: Path, token_path: Path):
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

    creds = None
    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), DRIVE_SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(credentials_path), DRIVE_SCOPES)
            creds = flow.run_local_server(port=0)
        token_path.parent.mkdir(parents=True, exist_ok=True)
        token_path.write_text(creds.to_json(), encoding="utf-8")

    return build("drive", "v3", credentials=creds, cache_discovery=False)


def get_upload_mime_type(file_path: Path) -> str:
    ext = file_path.suffix.lower()
    if ext == ".xls":
        return "application/vnd.ms-excel"
    if ext == ".xlsx":
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    if ext == ".xlsm":
        return "application/vnd.ms-excel.sheet.macroEnabled.12"
    return "application/octet-stream"


def write_upload_json(report: UploadReport, output_dir: Path) -> Path:
    output_path = output_dir / "upload_report.json"
    output_path.write_text(
        json.dumps(
            {
                "generated_utc": report.generated_utc,
                "input_root": report.input_root,
                "output_root": report.output_root,
                "file_count": report.file_count,
                "success_count": report.success_count,
                "failure_count": report.failure_count,
                "converted_to_sheets": report.converted_to_sheets,
                "drive_folder_id": report.drive_folder_id,
                "files": [asdict(f) for f in report.files],
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    return output_path


def write_upload_csv(report: UploadReport, output_dir: Path) -> Path:
    output_path = output_dir / "upload_report.csv"
    with output_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=list(asdict(report.files[0]).keys())
            if report.files
            else list(
                asdict(
                    UploadResult(
                        local_path="",
                        file_name="",
                        source_mime_type="",
                        target_mime_type="",
                        drive_file_id="",
                        drive_web_view_link="",
                        status="",
                        error="",
                    )
                ).keys()
            ),
        )
        writer.writeheader()
        for item in report.files:
            writer.writerow(asdict(item))
    return output_path


def run_upload(
    input_root: Path,
    output_dir: Path,
    credentials_path: Path,
    token_path: Path,
    drive_folder_id: Optional[str],
    convert_to_sheets: bool,
) -> UploadReport:
    from googleapiclient.http import MediaFileUpload

    service = load_drive_service(credentials_path, token_path)
    files = list(find_excel_files(input_root))
    results: List[UploadResult] = []

    for file_path in files:
        source_mime = get_upload_mime_type(file_path)
        target_mime = (
            "application/vnd.google-apps.spreadsheet"
            if convert_to_sheets
            else source_mime
        )
        target_name = file_path.stem if convert_to_sheets else file_path.name

        metadata = {"name": target_name}
        if drive_folder_id:
            metadata["parents"] = [drive_folder_id]
        if convert_to_sheets:
            metadata["mimeType"] = "application/vnd.google-apps.spreadsheet"

        try:
            media = MediaFileUpload(str(file_path), mimetype=source_mime, resumable=True)
            uploaded = (
                service.files()
                .create(
                    body=metadata,
                    media_body=media,
                    fields="id,name,mimeType,webViewLink",
                )
                .execute()
            )
            results.append(
                UploadResult(
                    local_path=str(file_path),
                    file_name=uploaded.get("name", target_name),
                    source_mime_type=source_mime,
                    target_mime_type=uploaded.get("mimeType", target_mime),
                    drive_file_id=uploaded.get("id", ""),
                    drive_web_view_link=uploaded.get("webViewLink", ""),
                    status="success",
                    error="",
                )
            )
        except Exception as exc:
            results.append(
                UploadResult(
                    local_path=str(file_path),
                    file_name=target_name,
                    source_mime_type=source_mime,
                    target_mime_type=target_mime,
                    drive_file_id="",
                    drive_web_view_link="",
                    status="failed",
                    error=str(exc),
                )
            )

    success_count = sum(1 for r in results if r.status == "success")
    failure_count = len(results) - success_count

    report = UploadReport(
        generated_utc=datetime.now(tz=timezone.utc).isoformat(),
        input_root=str(input_root),
        output_root=str(output_dir),
        file_count=len(results),
        success_count=success_count,
        failure_count=failure_count,
        converted_to_sheets=convert_to_sheets,
        drive_folder_id=drive_folder_id or "",
        files=results,
    )

    output_dir.mkdir(parents=True, exist_ok=True)
    write_upload_json(report, output_dir)
    write_upload_csv(report, output_dir)
    return report


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Excel migration inventory CLI")
    subparsers = parser.add_subparsers(dest="command", required=True)

    scan_parser = subparsers.add_parser("scan", help="Scan and analyze Excel files")
    scan_parser.add_argument("--input", required=True, help="Root folder to scan")
    scan_parser.add_argument("--output", required=True, help="Folder for report outputs")

    upload_parser = subparsers.add_parser("upload", help="Upload Excel files to Google Drive")
    upload_parser.add_argument("--input", required=True, help="Root folder to scan and upload")
    upload_parser.add_argument("--output", required=True, help="Folder for upload reports")
    upload_parser.add_argument(
        "--credentials",
        default="credentials/client_secret.json",
        help="OAuth client secret JSON from Google Cloud",
    )
    upload_parser.add_argument(
        "--token",
        default="credentials/token.json",
        help="Path to OAuth token cache file",
    )
    upload_parser.add_argument(
        "--drive-folder-id",
        default="",
        help="Drive folder ID for uploaded files (optional)",
    )
    upload_parser.add_argument(
        "--convert-to-sheets",
        action="store_true",
        help="Convert uploaded Excel files into Google Sheets format",
    )

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

    if args.command == "upload":
        input_root = Path(args.input).expanduser().resolve()
        output_dir = Path(args.output).expanduser().resolve()
        credentials_path = Path(args.credentials).expanduser().resolve()
        token_path = Path(args.token).expanduser().resolve()

        if not credentials_path.exists():
            print(f"Missing credentials file: {credentials_path}")
            return 1

        report = run_upload(
            input_root=input_root,
            output_dir=output_dir,
            credentials_path=credentials_path,
            token_path=token_path,
            drive_folder_id=args.drive_folder_id or None,
            convert_to_sheets=args.convert_to_sheets,
        )
        print(f"Upload attempted: {report.file_count}")
        print(f"Success: {report.success_count}")
        print(f"Failed: {report.failure_count}")
        print(f"JSON report: {output_dir / 'upload_report.json'}")
        print(f"CSV report: {output_dir / 'upload_report.csv'}")
        return 0 if report.failure_count == 0 else 2

    parser.print_help()
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
