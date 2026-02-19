# Excel Migration MVP

Local-first CLI to inventory Excel assets before Google Sheets migration.

## Features
- Recursively scans folders for `.xls`, `.xlsx`, `.xlsm`.
- Analyzes workbook structure (`sheet_count`, `formula_count`, `named_ranges`, `external_links`).
- Detects macro-enabled files (`.xlsm` or embedded `vbaProject.bin`).
- Generates `JSON` and `CSV` reports for batch review.
- Optional VBA extraction support when `oletools` is installed.

## Quick start
```bash
cd excel_migration_mvp
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python src/main.py scan --input "C:\\path\\to\\excel-folder" --output "./out"
```

## Output
- `out/report.json`: detailed per-file analysis.
- `out/report.csv`: summary table for filtering/sorting.

## Notes
- `.xls` deep parsing is limited in this MVP (metadata + extension-based risk).
- VBA code conversion is not included yet; this stage is inventory and migration triage.
