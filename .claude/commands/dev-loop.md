# 自律開発ループ

対象機能に対して、E2Eテスト＋生成物の品質チェックで問題を自分で発見し、修正→デプロイ→再テストを自律的に繰り返す。
何を直すかは自分で判断する。人間の介入なしに回し続ける。

## 引数
$ARGUMENTS — 対象機能とテストファイル（例: 「migrateページ 引取明細書.xlsm」）

## 設計思想
- **ターゲット: IT素人** — Pythonインストールもcredentials.jsonもなし
- **Python版 = ローカル完結** — gspread禁止、credentials.json不要、run.batダブルクリックで動く
- **スプシ版 = GASデプロイ** — 使い方シート付き、カスタムメニューから操作
- **テストはGUI操作** — ユーザーと同じ体験でテストする

## ループ手順

### Step 1: E2Eテスト（メインテスト — ユーザー体験そのもの）
- Playwright MCPで本番サイトにアクセス
- **実際のユーザーと同じGUI操作**でテスト（ファイルアップロード、ボタンクリック、結果確認等）
- 処理完了まで待機（3-4分）
- スクリーンショットを撮る
- 結果を「IT素人が初見で使えるか」の基準で評価
- **これが最も重要なテスト** — ユーザーが触るのはこのWebアプリのGUI

### Step 2: 生成物チェック（並列で実行）

生成物がない場合はスキップ。**サブエージェントで並列実行すること。**

#### 2A: スプシ版チェック（Agent: Playwright MCP）
1. Google Sheetsをブラウザで開く
2. マクロメニュー（カスタムメニュー）が表示されるか確認
3. メニューから実際に機能を実行（検索、データ取得等）
4. 使い方シートの内容確認:
   - 技術用語がないか（IT素人向け）
   - メニュー項目名が実際のメニューと一致しているか
   - 操作手順が具体的で再現可能か
5. 初回認証フローの説明が具体的か（「許可」ボタンの位置等）

#### 2B: Python版チェック（Agent: Bash実行）
ZIP URLをcurlでダウンロード→展開→**embedded pythonで実際に起動テスト**:

1. ZIPの構造確認: `python-embed/` `main.py` `modules/` `run.bat` `README.txt` があるか
2. **python-embed/python.exe で main.py を起動**: メニューが表示されるか
3. **各メニュー番号を入力して実行**: 機能が動くか、エラーが出ないか
4. 構文チェック: `python-embed/python.exe -m py_compile *.py`
5. 禁止チェック: `import gspread` `import google.oauth2` `credentials.json` が含まれていないこと
6. README.txt: 「run.batダブルクリック」だけで説明が完結しているか
7. メニュー項目名が日本語で意味がわかるか

### Step 3: 問題発見
- テスト結果と生成物チェックから問題を洗い出す
- **問題なし → 完了報告して終了**
- 問題あり → 優先度の高いものから修正へ

### Step 4: 修正実装
- 問題に対してコードを修正。修正先の判断基準:
  - UI/表示の問題 → フロントエンド（frontend/src/）
  - 使い方シートの問題 → DeployService.cs
  - Python変換がgspread使ってる → PythonConvertService.cs（プロンプト修正）
  - Python構文エラー → PythonConvertService.cs（プロンプト or max_tokens）
  - Python実行エラー → PythonConvertService.cs（変換マッピング）
  - main.pyの問題 → PythonPackagerService.cs
  - Track判定の問題 → TrackRouterService.cs
  - メニュー表示/CLI UXの問題 → PythonPackagerService.cs（テンプレート）
  - ZIP構造の問題 → PythonPackagerService.cs + Dockerfile
- `dotnet build` でビルド確認

### Step 5: デプロイ
- `git add` → `git commit` → `git push origin main`
- `gh run list` でデプロイ完了確認（3-4分）
- バックエンドウォームアップ（curl でAPI叩く）

### Step 6: Step 1 に戻る

## 並列化ルール
- Step 2A（スプシ）と Step 2B（Python）は **サブエージェントで同時実行**
- Step 1 の移行待ち中に修正方針の調査を並列で走らせてよい
- デプロイ待ち中にスキル改善やメモリ更新を並列で走らせてよい

## ルール
- 最大10イテレーション
- 何を直すかは自分で判断。ユーザーに聞かない
- 「問題なし」と判断したら即終了
- 各イテレーションで「何を見つけて何を直したか」を1行で記録
- 生成されたPythonコードを直接編集しない（バックエンドのプロンプト/ロジックを修正）
- **gspreadが生成コードに含まれていたら即修正** — PythonConvertService.csのプロンプトを見直す
- テスト環境が整わない場合は理由を記録してスキップ（ループをブロックしない）

## テスト観点
- 機能が正常に動くか
- エラー表示が適切か（local_onlyでデプロイエラーにならないか等）
- IT素人が見て意味がわかるか
- 日本語が自然か
- 生成されたスプシの使い方シート品質
- 生成されたPythonが**ローカル完結か（gspread/credentials.json不要か）**
- Python CLIのメニュー表示・操作性
- run.batダブルクリックで動くか

## スキル改善ルール（メタルール）
- **1周回ったらこのスキル自体を改善する** — 発見したパターンや知見をスキルに追記
- 新しい修正先パターン → Step 4 に追加
- 新しいテスト観点 → テスト観点に追加
- 新しい既知問題 → 既知パターンに追加

## 既知パターン
- 構文エラー除外されたGASがonOpenを壊す → BuildUsageSheetをDeploy後に除外ファイルを除いて呼ぶ（修正済み）
- ボタンなしExcelでカスタムメニューが出ない → ConvertServiceがButtonContext空でonOpen生成をスキップ → ThisWorkbookモジュールには常にonOpen生成（修正済み）
- 使い方シートのTODOコメントに技術用語が混入 → DeployServiceの使い方シート生成ロジックでフィルタリング（未対応）
- Python版にgspreadが混入 → PythonConvertServiceのプロンプトでFORBIDDEN指定（修正済み）
- Python Embeddable の ._pth に `..` が必要 → 親ディレクトリのmodules/を見つけるため（修正済み）
- pip download で --platform win_amd64 必要 → Linux上でWindows用wheelを取得するため（修正済み）

## テストファイル
- `C:\Users\fab24\excel-migration-os\bunkai.xlsm`
- `D:\クローム　ダウンロード\Excelマクロツール・ACCESS\引取明細書作成補助ツール(改)全受注区分.xlsm`

## 本番URL
- https://frontend-self-ten-98.vercel.app
- https://excel-migration-api.salmonbay-4f8a43a0.japaneast.azurecontainerapps.io
