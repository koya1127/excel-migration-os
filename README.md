# Excel Migration MVP

Local-first CLI to inventory Excel assets and upload them to Google Drive.

## Features
- Recursively scans folders for `.xls`, `.xlsx`, `.xlsm`.
- Analyzes workbook structure (`sheet_count`, `formula_count`, `named_ranges`, `external_links`).
- Detects macro-enabled files (`.xlsm` or embedded `vbaProject.bin`).
- Generates `JSON` and `CSV` reports for batch review.
- Uploads Excel files to Google Drive in batch.
- Optional conversion to Google Sheets during upload.

## Quick start
```bash
cd C:\Users\fab24\excel-migration-os
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Scan command
```bash
python src/main.py scan --input "D:\クローム　ダウンロード" --output "./out"
```

Outputs:
- `out/report.json`
- `out/report.csv`

## Drive upload command
1. Create OAuth Desktop credentials in Google Cloud Console.
2. Download the client JSON.
3. Save it to `credentials/client_secret.json`.

```bash
python src/main.py upload --input "D:\クローム　ダウンロード" --output "./out" --credentials "./credentials/client_secret.json" --token "./credentials/token.json"
```

Upload and convert to Google Sheets:
```bash
python src/main.py upload --input "D:\クローム　ダウンロード" --output "./out" --credentials "./credentials/client_secret.json" --token "./credentials/token.json" --convert-to-sheets
```

Upload to a specific Drive folder:
```bash
python src/main.py upload --input "D:\クローム　ダウンロード" --output "./out" --credentials "./credentials/client_secret.json" --token "./credentials/token.json" --drive-folder-id "YOUR_FOLDER_ID"
```

Outputs:
- `out/upload_report.json`
- `out/upload_report.csv`

## Notes
- `.xls` deep parsing is limited in this MVP (metadata + extension-based risk).
- VBA code conversion is not included yet; this stage is inventory + upload.
