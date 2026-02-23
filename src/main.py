from __future__ import annotations

import argparse
import csv
import json
import os
import re
import zipfile
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable, List, Optional

from dotenv import load_dotenv
from openpyxl import load_workbook

load_dotenv()

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
SCRIPT_SCOPES = [
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/script.projects",
]


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
class VbaModule:
    source_file: str
    module_name: str
    module_type: str
    code_lines: int
    output_path: str


@dataclass
class FormControl:
    source_file: str
    sheet_name: str
    control_name: str
    control_type: str
    label: str
    macro: str
    anchor: str


@dataclass
class ExtractReport:
    generated_utc: str
    input_root: str
    file_count: int
    module_count: int
    modules: List[VbaModule]
    controls: List[FormControl]


@dataclass
class ConvertResult:
    source_bas: str
    module_name: str
    module_type: str
    output_gs: str
    status: str
    error: str


@dataclass
class ConvertReport:
    generated_utc: str
    input_root: str
    total: int
    success: int
    failed: int
    results: List[ConvertResult]


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
class DeployReport:
    generated_utc: str
    spreadsheet_id: str
    script_id: str
    file_count: int
    files_deployed: List[str]


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
    if root.is_file():
        if root.suffix.lower() in EXCEL_EXTENSIONS:
            yield root
        return
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


def extract_form_controls_from_xlsm(file_path: Path) -> List[FormControl]:
    """Extract form control buttons from an xlsm file by parsing VML drawings."""
    controls: List[FormControl] = []
    if file_path.suffix.lower() != ".xlsm":
        return controls

    try:
        with zipfile.ZipFile(file_path, "r") as zf:
            # Map drawing relationships to sheet names via rels
            sheet_for_drawing: dict[str, str] = {}
            # Read workbook to get sheet names
            try:
                wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
            except KeyError:
                return controls

            sheet_names: list[str] = re.findall(r'<sheet[^>]+name="([^"]+)"', wb_xml)

            # Map rId -> sheet index from workbook rels
            try:
                wb_rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
            except KeyError:
                wb_rels = ""

            rid_to_target: dict[str, str] = {}
            for m in re.finditer(r'<Relationship[^>]+Id="([^"]+)"[^>]+Target="([^"]+)"', wb_rels):
                rid_to_target[m.group(1)] = m.group(2)

            # Map sheet rIds from workbook.xml
            sheet_rids = re.findall(r'<sheet[^>]+r:id="([^"]+)"', wb_xml)
            sheet_target_to_name: dict[str, str] = {}
            for i, rid in enumerate(sheet_rids):
                target = rid_to_target.get(rid, "")
                if target:
                    # Normalize: "worksheets/sheet1.xml" -> "sheet1"
                    sheet_key = target.replace("worksheets/", "").replace(".xml", "")
                    if i < len(sheet_names):
                        sheet_target_to_name[sheet_key] = sheet_names[i]

            # Find sheet -> vmlDrawing mapping from sheet rels
            for name in zf.namelist():
                m = re.match(r"xl/worksheets/_rels/(sheet\d+)\.xml\.rels$", name)
                if not m:
                    continue
                sheet_key = m.group(1)
                sheet_name = sheet_target_to_name.get(sheet_key, sheet_key)
                rels_xml = zf.read(name).decode("utf-8")
                for rel_match in re.finditer(
                    r'<Relationship[^>]+Target="([^"]*vmlDrawing[^"]*)"', rels_xml
                ):
                    vml_target = rel_match.group(1)
                    # Normalize path: ../drawings/vmlDrawing1.vml -> xl/drawings/vmlDrawing1.vml
                    if vml_target.startswith(".."):
                        vml_path = "xl/" + vml_target.lstrip("./")
                    elif not vml_target.startswith("xl/"):
                        vml_path = "xl/drawings/" + vml_target
                    else:
                        vml_path = vml_target
                    sheet_for_drawing[vml_path] = sheet_name

            # Parse each VML drawing file
            for vml_path, sheet_name in sheet_for_drawing.items():
                try:
                    raw = zf.read(vml_path)
                except KeyError:
                    continue
                # Try UTF-8 first, then Shift-JIS (common for Japanese Excel)
                for enc in ("utf-8", "cp932", "latin-1"):
                    try:
                        vml_data = raw.decode(enc)
                        break
                    except (UnicodeDecodeError, LookupError):
                        continue
                else:
                    vml_data = raw.decode("latin-1")

                # Split into shape blocks
                for shape_block in re.finditer(
                    r"<v:shape\b[^>]*>(.*?)</v:shape>", vml_data, re.DOTALL
                ):
                    block = shape_block.group(1)
                    # Check if this is a Button type
                    obj_type_m = re.search(
                        r'<x:ClientData\s+ObjectType="(\w+)"', block
                    )
                    if not obj_type_m or obj_type_m.group(1) != "Button":
                        continue

                    # Extract label (text content)
                    label = ""
                    text_m = re.search(r"<v:textbox[^>]*>(.*?)</v:textbox>", block, re.DOTALL)
                    if text_m:
                        # Strip HTML/XML tags to get plain text
                        raw_label = re.sub(r"<[^>]+>", "", text_m.group(1))
                        # Normalize whitespace (newlines, multiple spaces)
                        label = " ".join(raw_label.split())

                    # Extract macro name from FmlaMacro
                    macro = ""
                    macro_m = re.search(r"<x:FmlaMacro>(.*?)</x:FmlaMacro>", block, re.DOTALL)
                    if macro_m:
                        macro = macro_m.group(1).strip()
                        # Remove workbook index prefix like [0]!
                        macro = re.sub(r"^\[\d+\]!", "", macro)

                    # Extract anchor coordinates
                    anchor = ""
                    anchor_m = re.search(r"<x:Anchor>(.*?)</x:Anchor>", block, re.DOTALL)
                    if anchor_m:
                        coords = [c.strip() for c in anchor_m.group(1).split(",")]
                        if len(coords) >= 8:
                            # coords: col_start, offset, row_start, offset, col_end, offset, row_end, offset
                            col_start = int(coords[0])
                            row_start = int(coords[2])
                            col_end = int(coords[4])
                            row_end = int(coords[6])

                            def col_letter(n: int) -> str:
                                result = ""
                                while n >= 0:
                                    result = chr(ord("A") + n % 26) + result
                                    n = n // 26 - 1
                                return result

                            anchor = f"{col_letter(col_start)}{row_start + 1}:{col_letter(col_end)}{row_end + 1}"

                    control_name = label or "Button"
                    controls.append(
                        FormControl(
                            source_file=str(file_path),
                            sheet_name=sheet_name,
                            control_name=control_name,
                            control_type="Button",
                            label=label,
                            macro=macro,
                            anchor=anchor,
                        )
                    )
    except (zipfile.BadZipFile, OSError):
        pass

    return controls


def find_xlsm_files(root: Path) -> Iterable[Path]:
    if root.is_file():
        if root.suffix.lower() == ".xlsm":
            yield root
        return
    for file_path in root.rglob("*.xlsm"):
        if file_path.is_file():
            yield file_path


def classify_vba_module(module_name: str, code: str) -> str:
    lower_name = module_name.lower()
    if lower_name == "thisworkbook":
        return "ThisWorkbook"
    if lower_name.startswith("sheet") or "Attribute VB_Base = \"0{00020820" in code:
        return "Sheet"
    if "Attribute VB_Base = \"0{" in code:
        return "UserForm"
    if "Attribute VB_Creatable = True" in code or "VERSION 1.0 CLASS" in code:
        return "Class"
    return "Standard"


def extract_vba_from_file(
    file_path: Path, output_dir: Path
) -> List[VbaModule]:
    from oletools.olevba import VBA_Parser

    parser = VBA_Parser(str(file_path))
    modules: List[VbaModule] = []

    if not parser.detect_vba_macros():
        parser.close()
        return modules

    file_output_dir = output_dir / file_path.stem
    file_output_dir.mkdir(parents=True, exist_ok=True)

    for _, _, vba_filename, vba_code in parser.extract_macros():
        if not vba_code or not vba_code.strip():
            continue

        module_name = Path(vba_filename).stem if vba_filename else "Unknown"
        module_type = classify_vba_module(module_name, vba_code)
        code_lines = len(vba_code.splitlines())

        bas_path = file_output_dir / f"{module_name}.bas"
        bas_path.write_text(vba_code, encoding="utf-8")

        modules.append(
            VbaModule(
                source_file=str(file_path),
                module_name=module_name,
                module_type=module_type,
                code_lines=code_lines,
                output_path=str(bas_path),
            )
        )

    parser.close()
    return modules


def write_extract_json(report: ExtractReport, output_dir: Path) -> Path:
    output_path = output_dir / "extract_report.json"
    output_path.write_text(
        json.dumps(
            {
                "generated_utc": report.generated_utc,
                "input_root": report.input_root,
                "file_count": report.file_count,
                "module_count": report.module_count,
                "modules": [asdict(m) for m in report.modules],
                "controls": [asdict(c) for c in report.controls],
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    return output_path


def write_extract_csv(report: ExtractReport, output_dir: Path) -> Path:
    output_path = output_dir / "extract_report.csv"
    fieldnames = ["source_file", "module_name", "module_type", "code_lines", "output_path"]
    with output_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for m in report.modules:
            writer.writerow(asdict(m))
    return output_path


def run_extract(input_root: Path, output_dir: Path) -> ExtractReport:
    files = list(find_xlsm_files(input_root))
    all_modules: List[VbaModule] = []
    all_controls: List[FormControl] = []

    for file_path in files:
        try:
            modules = extract_vba_from_file(file_path, output_dir)
            all_modules.extend(modules)
            print(f"  {file_path.name}: {len(modules)} modules")
        except Exception as exc:
            print(f"  {file_path.name}: ERROR - {exc}")

        try:
            ctrls = extract_form_controls_from_xlsm(file_path)
            all_controls.extend(ctrls)
            if ctrls:
                print(f"  {file_path.name}: {len(ctrls)} form controls")
        except Exception as exc:
            print(f"  {file_path.name}: form control extraction error - {exc}")

    report = ExtractReport(
        generated_utc=datetime.now(tz=timezone.utc).isoformat(),
        input_root=str(input_root),
        file_count=len(files),
        module_count=len(all_modules),
        modules=all_modules,
        controls=all_controls,
    )

    output_dir.mkdir(parents=True, exist_ok=True)
    write_extract_json(report, output_dir)
    write_extract_csv(report, output_dir)
    return report


VBA_TO_GAS_SYSTEM_PROMPT = """\
You are a VBA-to-Google Apps Script converter.
Convert the given VBA module to Google Apps Script (JavaScript).

Key conversion rules:
- ThisWorkbook → SpreadsheetApp.getActiveSpreadsheet()
- ActiveSheet / Worksheets("name") → getSheetByName("name")
- Range("A1") / Cells(row, col) → getRange() / getRange(row, col)
- .Value → .getValue() / .setValue()
- MsgBox → Browser.msgBox()
- VBA Sub/Function → JavaScript function
- For...Next → for loop
- Dim → let/const
- VBA string concatenation (&) → JavaScript (+)
- VBA comments (') → JavaScript comments (//)
- On Error → try/catch
- Nothing → null

Excel form control buttons do not exist in Google Sheets.
When button information is provided, generate an onOpen() trigger that creates
a custom menu named "Macros" using SpreadsheetApp.getUi().createMenu("Macros").
Add each button as a menu item via .addItem(buttonLabel, functionName).

Output ONLY the converted code. No explanations, no markdown fences.
"""


def extract_code_block(text: str) -> str:
    """Extract code from markdown fenced block if present, else return full text."""
    m = re.search(r"```(?:javascript|js|gs|gas)?\s*\n(.*?)```", text, re.DOTALL)
    if m:
        return m.group(1).strip()
    m = re.search(r"```\s*\n(.*?)```", text, re.DOTALL)
    if m:
        return m.group(1).strip()
    return text.strip()


def convert_vba_module(
    vba_code: str, module_name: str, module_type: str, client, model: str,
    controls_context: str = "",
) -> str:
    """Call Anthropic Messages API to convert a single VBA module to Apps Script."""
    user_msg = f"Module: {module_name} (type: {module_type})\n\n{vba_code}"
    if controls_context:
        user_msg += f"\n\n{controls_context}"
    response = client.messages.create(
        model=model,
        max_tokens=4096,
        system=VBA_TO_GAS_SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_msg}],
    )
    raw = response.content[0].text
    return extract_code_block(raw)


def write_convert_json(report: ConvertReport, output_dir: Path) -> Path:
    output_path = output_dir / "convert_report.json"
    output_path.write_text(
        json.dumps(
            {
                "generated_utc": report.generated_utc,
                "input_root": report.input_root,
                "total": report.total,
                "success": report.success,
                "failed": report.failed,
                "results": [asdict(r) for r in report.results],
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    return output_path


def write_convert_csv(report: ConvertReport, output_dir: Path) -> Path:
    output_path = output_dir / "convert_report.csv"
    fieldnames = ["source_bas", "module_name", "module_type", "output_gs", "status", "error"]
    with output_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for r in report.results:
            writer.writerow(asdict(r))
    return output_path


def run_convert(
    input_dir: Path, output_dir: Path, api_key: Optional[str], model: str
) -> ConvertReport:
    from anthropic import Anthropic

    resolved_key = api_key or os.environ.get("ANTHROPIC_API_KEY", "")
    if not resolved_key:
        raise SystemExit("No API key provided. Use --api-key or set ANTHROPIC_API_KEY.")

    client = Anthropic(api_key=resolved_key)
    output_dir.mkdir(parents=True, exist_ok=True)

    bas_files = sorted(input_dir.rglob("*.bas"))
    if not bas_files:
        print(f"No .bas files found under {input_dir}")

    # Load extract report once for module types and controls
    extract_data: dict = {}
    report_json = input_dir / "extract_report.json"
    if report_json.exists():
        try:
            extract_data = json.loads(report_json.read_text(encoding="utf-8"))
        except Exception:
            pass

    # Build controls context string for ThisWorkbook modules
    controls_list = extract_data.get("controls", [])
    controls_context = ""
    if controls_list:
        lines = ["Excel form control buttons found in this workbook:"]
        for c in controls_list:
            lines.append(
                f"- Button label=\"{c.get('label', '')}\" macro=\"{c.get('macro', '')}\" "
                f"sheet=\"{c.get('sheet_name', '')}\" anchor=\"{c.get('anchor', '')}\""
            )
        lines.append(
            "\nGenerate an onOpen() function that creates a 'Macros' custom menu "
            "with menu items for each button above."
        )
        controls_context = "\n".join(lines)

    results: List[ConvertResult] = []
    for bas_path in bas_files:
        module_name = bas_path.stem
        # Preserve subdirectory structure (e.g. workbook_name/)
        rel = bas_path.parent.relative_to(input_dir)
        gs_dir = output_dir / rel
        gs_dir.mkdir(parents=True, exist_ok=True)
        gs_path = gs_dir / f"{module_name}.gs"

        # Determine module_type from extract_report.json
        module_type = "Standard"
        for m in extract_data.get("modules", []):
            if m.get("module_name") == module_name and bas_path.name in m.get("output_path", ""):
                module_type = m.get("module_type", "Standard")
                break

        # Only pass controls context for ThisWorkbook modules
        ctx = controls_context if module_type == "ThisWorkbook" else ""

        print(f"  Converting {rel / bas_path.name} ...", end=" ", flush=True)
        try:
            vba_code = bas_path.read_text(encoding="utf-8", errors="replace")
            gs_code = convert_vba_module(vba_code, module_name, module_type, client, model, controls_context=ctx)
            gs_path.write_text(gs_code, encoding="utf-8")
            results.append(
                ConvertResult(
                    source_bas=str(bas_path),
                    module_name=module_name,
                    module_type=module_type,
                    output_gs=str(gs_path),
                    status="success",
                    error="",
                )
            )
            print("OK")
        except Exception as exc:
            results.append(
                ConvertResult(
                    source_bas=str(bas_path),
                    module_name=module_name,
                    module_type=module_type,
                    output_gs="",
                    status="failed",
                    error=str(exc),
                )
            )
            print(f"FAILED - {exc}")

    success = sum(1 for r in results if r.status == "success")
    report = ConvertReport(
        generated_utc=datetime.now(tz=timezone.utc).isoformat(),
        input_root=str(input_dir),
        total=len(results),
        success=success,
        failed=len(results) - success,
        results=results,
    )

    write_convert_json(report, output_dir)
    write_convert_csv(report, output_dir)
    return report


def load_script_service(credentials_path: Path, token_path: Path):
    from google.auth.exceptions import RefreshError
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

    creds = None
    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), SCRIPT_SCOPES)

    if not creds or not creds.valid:
        refreshed = False
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                refreshed = True
            except RefreshError:
                pass
        if not refreshed:
            flow = InstalledAppFlow.from_client_secrets_file(str(credentials_path), SCRIPT_SCOPES)
            creds = flow.run_local_server(port=0)
        token_path.parent.mkdir(parents=True, exist_ok=True)
        token_path.write_text(creds.to_json(), encoding="utf-8")

    return build("script", "v1", credentials=creds, cache_discovery=False)


def write_deploy_json(report: DeployReport, output_dir: Path) -> Path:
    output_path = output_dir / "deploy_report.json"
    output_path.write_text(
        json.dumps(asdict(report), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return output_path


def run_deploy(
    input_dir: Path,
    spreadsheet_id: str,
    credentials_path: Path,
    token_path: Path,
    script_id: Optional[str] = None,
) -> DeployReport:
    service = load_script_service(credentials_path, token_path)

    gs_files = sorted(input_dir.rglob("*.gs"))
    if not gs_files:
        raise SystemExit(f"No .gs files found under {input_dir}")

    print(f"Deploying to spreadsheet: {spreadsheet_id}")

    if script_id:
        print(f"  Using existing script project: {script_id}")
    else:
        create_body = {"title": "ExcelMigration", "parentId": spreadsheet_id}
        created = service.projects().create(body=create_body).execute()
        script_id = created["scriptId"]
        print(f"  Created script project: {script_id}")

    manifest = {
        "name": "appsscript",
        "type": "JSON",
        "source": json.dumps(
            {
                "timeZone": "Asia/Tokyo",
                "dependencies": {},
                "exceptionLogging": "STACKDRIVER",
                "runtimeVersion": "V8",
                "oauthScopes": [
                    "https://www.googleapis.com/auth/spreadsheets.currentonly"
                ],
            },
            ensure_ascii=False,
        ),
    }

    files_payload = [manifest]
    deployed_names: List[str] = []
    print(f"  Pushing {len(gs_files)} files...")
    for gs_path in gs_files:
        name = gs_path.stem
        source = gs_path.read_text(encoding="utf-8")
        files_payload.append({"name": name, "type": "SERVER_JS", "source": source})
        deployed_names.append(gs_path.name)

    body = {"scriptId": script_id, "files": files_payload}
    service.projects().updateContent(scriptId=script_id, body=body).execute()

    for name in deployed_names:
        print(f"  {name} ... OK")
    print(f"Deploy complete: {len(deployed_names)} files pushed")
    print(f"Script editor: https://script.google.com/d/{script_id}/edit")

    report = DeployReport(
        generated_utc=datetime.now(tz=timezone.utc).isoformat(),
        spreadsheet_id=spreadsheet_id,
        script_id=script_id,
        file_count=len(deployed_names),
        files_deployed=deployed_names,
    )

    input_dir.parent.mkdir(parents=True, exist_ok=True)
    write_deploy_json(report, input_dir)
    return report


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Excel migration inventory CLI")
    subparsers = parser.add_subparsers(dest="command", required=True)

    scan_parser = subparsers.add_parser("scan", help="Scan and analyze Excel files")
    scan_parser.add_argument("--input", required=True, help="Root folder to scan")
    scan_parser.add_argument("--output", required=True, help="Folder for report outputs")

    extract_parser = subparsers.add_parser("extract", help="Extract VBA modules from .xlsm files")
    extract_parser.add_argument("--input", required=True, help=".xlsm file or folder to scan")
    extract_parser.add_argument("--output", required=True, help="Folder for extracted .bas files and reports")

    convert_parser = subparsers.add_parser("convert", help="Convert VBA modules to Google Apps Script via Claude API")
    convert_parser.add_argument("--input", required=True, help="Folder with extracted .bas files")
    convert_parser.add_argument("--output", required=True, help="Folder for converted .gs files and reports")
    convert_parser.add_argument("--api-key", default="", help="Anthropic API key (default: env ANTHROPIC_API_KEY)")
    convert_parser.add_argument("--model", default="claude-sonnet-4-6", help="Model to use (default: claude-sonnet-4-6)")

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

    deploy_parser = subparsers.add_parser("deploy", help="Deploy .gs files to a Google Spreadsheet via Apps Script API")
    deploy_parser.add_argument("--input", required=True, help="Folder containing .gs files")
    deploy_parser.add_argument("--spreadsheet-id", required=True, help="Target Google Spreadsheet ID")
    deploy_parser.add_argument("--script-id", default="", help="Existing Apps Script project ID (optional; creates new if omitted)")
    deploy_parser.add_argument(
        "--credentials",
        default="credentials/client_secret.json",
        help="OAuth client secret JSON from Google Cloud",
    )
    deploy_parser.add_argument(
        "--token",
        default="credentials/token.json",
        help="Path to OAuth token cache file",
    )

    migrate_parser = subparsers.add_parser(
        "migrate", help="End-to-end: upload xlsm to Sheets and deploy .gs files"
    )
    migrate_parser.add_argument("--xlsm", required=True, help="Source .xlsm file")
    migrate_parser.add_argument("--gs-dir", required=True, help="Folder with converted .gs files")
    migrate_parser.add_argument("--output", required=True, help="Folder for migration reports")
    migrate_parser.add_argument(
        "--credentials",
        default="credentials/client_secret.json",
        help="OAuth client secret JSON from Google Cloud",
    )
    migrate_parser.add_argument(
        "--token",
        default="credentials/token.json",
        help="Path to OAuth token cache file",
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

    if args.command == "extract":
        input_root = Path(args.input).expanduser().resolve()
        output_dir = Path(args.output).expanduser().resolve()
        print(f"Extracting VBA from: {input_root}")
        report = run_extract(input_root, output_dir)
        print(f"Files processed: {report.file_count}")
        print(f"Modules extracted: {report.module_count}")
        print(f"JSON report: {output_dir / 'extract_report.json'}")
        print(f"CSV report: {output_dir / 'extract_report.csv'}")
        return 0

    if args.command == "convert":
        input_dir = Path(args.input).expanduser().resolve()
        output_dir = Path(args.output).expanduser().resolve()
        print(f"Converting VBA → Apps Script from: {input_dir}")
        report = run_convert(input_dir, output_dir, args.api_key or None, args.model)
        print(f"Total: {report.total}  Success: {report.success}  Failed: {report.failed}")
        print(f"JSON report: {output_dir / 'convert_report.json'}")
        print(f"CSV report: {output_dir / 'convert_report.csv'}")
        return 0 if report.failed == 0 else 2

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

    if args.command == "deploy":
        input_dir = Path(args.input).expanduser().resolve()
        credentials_path = Path(args.credentials).expanduser().resolve()
        token_path = Path(args.token).expanduser().resolve()

        if not credentials_path.exists():
            print(f"Missing credentials file: {credentials_path}")
            return 1

        report = run_deploy(
            input_dir=input_dir,
            spreadsheet_id=args.spreadsheet_id,
            credentials_path=credentials_path,
            token_path=token_path,
            script_id=args.script_id or None,
        )
        print(f"Deploy report: {input_dir / 'deploy_report.json'}")
        return 0

    if args.command == "migrate":
        xlsm_path = Path(args.xlsm).expanduser().resolve()
        gs_dir = Path(args.gs_dir).expanduser().resolve()
        output_dir = Path(args.output).expanduser().resolve()
        credentials_path = Path(args.credentials).expanduser().resolve()
        token_path = Path(args.token).expanduser().resolve()

        if not xlsm_path.exists():
            print(f"Missing xlsm file: {xlsm_path}")
            return 1
        if not credentials_path.exists():
            print(f"Missing credentials file: {credentials_path}")
            return 1

        output_dir.mkdir(parents=True, exist_ok=True)

        # Step 1: Upload xlsm as Google Sheets
        print(f"Uploading {xlsm_path.name} to Google Sheets...")
        upload_report = run_upload(
            input_root=xlsm_path,
            output_dir=output_dir,
            credentials_path=credentials_path,
            token_path=token_path,
            drive_folder_id=None,
            convert_to_sheets=True,
        )

        if upload_report.failure_count > 0 or not upload_report.files:
            print("Upload failed.")
            return 1

        uploaded = upload_report.files[0]
        spreadsheet_id = uploaded.drive_file_id
        print(f"Spreadsheet created: {uploaded.drive_web_view_link}")

        # Step 2: Deploy .gs files to the spreadsheet
        print(f"Deploying .gs files from {gs_dir}...")
        deploy_report = run_deploy(
            input_dir=gs_dir,
            spreadsheet_id=spreadsheet_id,
            credentials_path=credentials_path,
            token_path=token_path,
        )
        print(f"Migration complete!")
        print(f"Spreadsheet: {uploaded.drive_web_view_link}")
        print(f"Script editor: https://script.google.com/d/{deploy_report.script_id}/edit")
        return 0

    parser.print_help()
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
