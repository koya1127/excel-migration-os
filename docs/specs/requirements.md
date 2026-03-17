# Requirements: VBA移行 二系統アーキテクチャ

**作成日:** 2026-03-17
**目的:** VBA→GAS変換の機能再現率を最大化するため、移行を二系統に分ける

---

## 背景

VBA→GASの変換には原理的な限界がある（詳細: `docs/vba-gas-compatibility-matrix.md`）。
GASはクラウドサンドボックスであり、ローカルファイルI/O、COM連携、Shell実行等が不可能。
現状の移行フローではこれらを「現在お使いいただけません」として切り捨てている。

**課題:** 顧客のVBAマクロの中にファイルI/OやCOM連携が含まれる場合、移行後のスプレッドシートでは業務が回らない。

---

## 要件

### R1: 二系統の移行パスを提供する

VBAモジュールを解析し、各機能がGASで実現可能かどうかを自動判定する。
判定結果に基づき、以下の二系統に振り分ける。

| Track | 対象 | 出力 |
|---|---|---|
| Track 1: スプシ版 | GASで実現可能な機能 | Google Sheets + GAS（現行フロー） |
| Track 2: ローカル版 | GASで実現不可能な機能 | Pythonスクリプト（exe化可） |

### R2: ユーザーが移行パスを選択できる

Migrateページで以下の選択肢を提供する:
- **a) スプシ版のみ** — 現行と同じ。使えない機能は「お使いいただけません」表示
- **b) ローカル版のみ** — 全マクロをPythonに変換。スプシはデータのみ
- **c) スプシ版 + ローカル版** — GAS可能分はGAS、不可能分はPython。両方出力

### R3: ローカル版の機能再現率を最大化する

UI/UXは優先しない。機能の再現率を最大化する。
Pythonを選択する理由: win32comによるCOM連携、os/subprocessによるファイルI/O・Shell実行が可能。
VBAが使うWindows COMインフラをそのまま利用できるため、再現率95%+を目指せる。

### R4: ローカル版がGoogle Sheetsのデータを読み書きできる

ローカル版PythonスクリプトがGoogle Sheets上のデータを直接操作できること。
gspread + Google Sheets APIを使用し、データはスプシ、ロジックはローカルという分業を実現する。

### R5: 自動判定の基準

VBAコード内の以下のパターンを検出し、Track振り分けを自動判定する:

**Track 2送り（GAS不可）の判定パターン:**
- ローカルファイルI/O: `Open.*For Input`, `Open.*For Output`, `FileCopy`, `Kill`, `Dir(`, `MkDir`, `FileSystemObject`
- COM連携: `CreateObject(`, `WScript.`, `Shell(`, `SendKeys`
- フォーム部品状態: `ControlFormat.Value`, `OLEObjects`, `ActiveX`
- クリップボード: `DataObject`, `MSForms.DataObject`
- その他: `Application.Run`（外部マクロ呼び出し）、`Declare Function`（Win API）

**Track 1残留（GAS可）:**
- 上記パターンを含まないモジュール

### R6: 変換プロンプトの分岐

- Track 1: 現行のVBA→GAS変換プロンプト（ConvertService.cs）
- Track 2: 新規のVBA→Python変換プロンプト
  - win32com, os, subprocess, gspread等を使用
  - Google Sheetsとの連携コードを自動生成
  - PyInstaller対応（`__file__`パスの処理等）

---

## 非要件（スコープ外）

- ローカル版のリッチUI（tkinter/PyQt等は不要。CLIまたは最小限のprint出力で十分）
- macOS/Linux対応（win32comはWindows専用。ターゲット顧客はWindows企業）
- VBA UserFormの完全再現（機能の再現が目的であり、見た目の再現ではない）

---

## 成功指標

- bunkai.xlsmで現在「お使いいただけません」の31件の未対応機能のうち、Track 2で20件以上が動作すること
- VBA→Python変換の構文エラー率が10%以下であること
- エンドユーザーがexeダブルクリックで実行できること
