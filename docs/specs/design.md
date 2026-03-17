# Design: 二系統移行アーキテクチャ

**作成日:** 2026-03-17
**前提:** `docs/specs/requirements.md`

---

## アーキテクチャ概要

```
                    ┌─────────────────────────┐
                    │   Excel (.xlsm) 入力     │
                    └──────────┬──────────────┘
                               │
                    ┌──────────▼──────────────┐
                    │   VBA抽出 (ExtractService)│
                    └──────────┬──────────────┘
                               │
                    ┌──────────▼──────────────┐
                    │   Track判定 (TrackRouter) │
                    │   VBAコード解析→振り分け   │
                    └────┬─────────────┬──────┘
                         │             │
              ┌──────────▼──┐   ┌──────▼──────────┐
              │  Track 1     │   │  Track 2          │
              │  GAS変換     │   │  Python変換        │
              │ (既存フロー)  │   │ (新規)             │
              └──────┬───────┘   └──────┬────────────┘
                     │                  │
              ┌──────▼───────┐   ┌──────▼────────────┐
              │ GASデプロイ   │   │ Python出力         │
              │ + 使い方シート │   │ (.py or .exe)      │
              │ (既存フロー)  │   │ + README.txt       │
              └──────────────┘   └───────────────────┘
```

---

## コンポーネント設計

### 1. TrackRouter（新規）

**場所:** `backend/ExcelMigrationApi/Services/TrackRouterService.cs`

**責務:** VBAモジュールを解析し、Track 1（GAS）かTrack 2（Python）に振り分ける。

```
入力: List<ExtractedModule>
出力: TrackResult {
    Track1Modules: List<ExtractedModule>  // GAS変換対象
    Track2Modules: List<ExtractedModule>  // Python変換対象
    Reasoning: List<TrackDecision>        // 判定理由（ログ・UI表示用）
}
```

**判定ロジック:**
1. 各モジュールのVBAコードに対してregexパターンマッチ
2. Track 2パターン（ファイルI/O、COM等）が1つでも見つかればTrack 2
3. パターンなしならTrack 1
4. モジュール単位ではなく関数単位で分割も検討（Phase 2）

### 2. PythonConvertService（新規）

**場所:** `backend/ExcelMigrationApi/Services/PythonConvertService.cs`

**責務:** VBA→Python変換。Claude APIを使用。

**システムプロンプト（概要）:**
```
VBAコードをPythonに変換する。以下のライブラリを使用:
- gspread + oauth2client: Google Sheets読み書き
- win32com.client: COM連携（Excel, Outlook等）
- os, shutil, subprocess: ファイルI/O, Shell
- openpyxl: ローカルExcelファイル操作
- csv: CSV読み書き

変換ルール:
- Range("A1").Value → worksheet.acell("A1").value
- Sheets("name") → workbook.worksheet("name")
- MsgBox → print() + input()
- Dir() → os.listdir() / glob.glob()
- FileCopy → shutil.copy2()
- CreateObject("Outlook.Application") → win32com.client.Dispatch("Outlook.Application")
- Open path For Input → open(path, "r")
- Shell → subprocess.run()
```

**モデル選択:** ConvertServiceと同じ（大: claude-sonnet-4-6、小: claude-haiku-4-5）

### 3. PythonPackager（新規）

**場所:** `backend/ExcelMigrationApi/Services/PythonPackagerService.cs`

**責務:** 変換されたPythonコードをパッケージ化して出力。

**出力物:**
```
{ファイル名}_local/
  ├── main.py              # エントリポイント（メニュー表示→関数呼び出し）
  ├── modules/
  │   ├── sheet1.py        # Sheet1のVBAから変換
  │   ├── module1.py       # Module1のVBAから変換
  │   └── ...
  ├── config.json           # Google Sheets ID、認証情報パス等
  ├── requirements.txt      # gspread, oauth2client, pywin32, openpyxl
  ├── setup.bat             # pip install -r requirements.txt
  └── README.txt            # 使い方（日本語、IT素人向け）
```

**main.pyの構造:**
```python
import sys
from modules import sheet1, module1

def main():
    print("=" * 40)
    print("  bunkai ローカル版")
    print("=" * 40)
    print()
    print("実行したい機能の番号を入力してください:")
    print("  1. 検索（部材・部品）")
    print("  2. 検索（品名）")
    print("  3. ダウンロード")
    print("  0. 終了")
    print()

    choice = input("番号: ")
    if choice == "1":
        sheet1.search_parts()
    elif choice == "2":
        sheet2.search_names()
    # ...

if __name__ == "__main__":
    main()
```

### 4. MigrateController拡張

**変更:** `backend/ExcelMigrationApi/Controllers/MigrateController.cs`

**新パラメータ:**
```json
{
  "files": [...],
  "folderId": "...",
  "trackMode": "auto" | "sheets_only" | "local_only" | "both"
}
```

**フロー（trackMode = "auto" or "both"）:**
1. Upload（既存）
2. Extract（既存）
3. **TrackRouter: モジュール振り分け**
4. Track 1モジュール → ConvertService（既存GAS変換）
5. Track 2モジュール → PythonConvertService（新規Python変換）
6. Track 1 → Deploy GAS（既存）
7. Track 2 → PythonPackager → ZIPダウンロードURL生成
8. 使い方シート更新（Track 2の機能は「ローカル版をご利用ください」と記載）

### 5. フロントエンド変更

**Migrateページ:**
- trackMode選択UI追加（ラジオボタン: 自動判定/スプシのみ/ローカルのみ/両方）
- Track 2結果表示: ZIPダウンロードリンク
- 移行プロセスにStep追加: 「Track振り分け」

---

## 技術的決定事項

| 決定 | 理由 |
|---|---|
| Python（not Electron/Tauri） | UI不要、win32comによるCOM再現率が最高（95%+） |
| gspread（not Google Sheets API直接） | Pythonから最も使いやすいSheetsライブラリ |
| PyInstaller（exe化） | IT素人でもダブルクリックで実行可能 |
| CLIメニュー（not GUI） | 機能再現率優先。print+inputで十分 |
| モジュール単位の振り分け | 関数単位は複雑すぎる。Phase 1はモジュール単位 |
| Claude API（VBA→Python変換） | VBA→GASと同じモデルを使用。プロンプトのみ変更 |

---

## Phase分け

### Phase 1（MVP）
- TrackRouter: regexパターンマッチでモジュール振り分け
- PythonConvertService: VBA→Python変換（Claude API）
- PythonPackager: .pyファイル群 + requirements.txt + README.txt出力
- MigrateController: trackModeパラメータ追加
- フロントエンド: trackMode選択UI + ZIPダウンロード

### Phase 2（改善）
- 関数単位のTrack振り分け（1モジュール内にGAS可/不可が混在する場合）
- PyInstaller exe自動ビルド（サーバー上でexe化してダウンロード提供）
- Google OAuth認証の自動セットアップ（config.json生成）
- ローカル版の使い方シート自動生成（README.txtの品質向上）

### Phase 3（将来）
- ローカル版のGUI（必要に応じてtkinter/PyQt）
- macOS対応（win32com代替としてappscript等）
- ローカル版の自動テスト生成
