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

## Extract command
Extract VBA modules from `.xlsm` files as `.bas` source files.

```bash
# Single file
python src/main.py extract --input "path/to/file.xlsm" --output "./out/extract"

# Folder (recursive)
python src/main.py extract --input "path/to/folder" --output "./out/extract"
```

Outputs:
- `out/extract/extract_report.json`
- `out/extract/extract_report.csv`
- `out/extract/<filename>/Module1.bas`, `Sheet1.bas`, etc.

## Convert command
Convert extracted VBA modules (`.bas`) to Google Apps Script (`.gs`) using Claude API.

```bash
# Convert all .bas files under the extract output
python src/main.py convert --input "./out/extract" --output "./out/convert"

# With explicit API key and model
python src/main.py convert --input "./out/extract" --output "./out/convert" --api-key "sk-..." --model "claude-sonnet-4-6"
```

Set `ANTHROPIC_API_KEY` environment variable to skip `--api-key`.

Outputs:
- `out/convert/convert_report.json`
- `out/convert/convert_report.csv`
- `out/convert/<workbook_name>/Module1.gs`, `Sheet1.gs`, etc.

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

## Deploy command
Deploy converted `.gs` files to a Google Spreadsheet as a container-bound Apps Script project.

**Prerequisites:**
- Enable **Apps Script API** in Google Cloud Console.
- OAuth client must include the `script.projects` scope (re-auth runs automatically on first use).

```bash
# Deploy to an existing spreadsheet (creates a new script project)
python src/main.py deploy \
  --input "./out/convert-textvba" \
  --spreadsheet-id "YOUR_SPREADSHEET_ID"

# Deploy to an existing script project
python src/main.py deploy \
  --input "./out/convert-textvba" \
  --spreadsheet-id "YOUR_SPREADSHEET_ID" \
  --script-id "EXISTING_SCRIPT_PROJECT_ID"
```

Options:
- `--input` (required): folder containing `.gs` files.
- `--spreadsheet-id` (required): target Google Spreadsheet ID.
- `--script-id` (optional): existing Apps Script project ID. Omit to create a new project.
- `--credentials` (default: `credentials/client_secret.json`): OAuth client secret.
- `--token` (default: `credentials/token.json`): OAuth token cache.

Outputs:
- `<input>/deploy_report.json`

## Notes
- `.xls` deep parsing is limited in this MVP (metadata + extension-based risk).
- VBA → Apps Script conversion uses Claude API (requires `ANTHROPIC_API_KEY`).
