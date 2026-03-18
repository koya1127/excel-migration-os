# 自律開発ループ

対象機能に対して、E2Eテスト＋生成物の品質チェックで問題を自分で発見し、修正→デプロイ→再テストを自律的に繰り返す。
何を直すかは自分で判断する。人間の介入なしに回し続ける。

## 引数
$ARGUMENTS — 対象機能とテストファイル（例: 「migrateページ 引取明細書.xlsm」）

## 設計思想
- **ターゲット: IT素人** — Pythonインストールもcredentials.jsonもなし
- **Python版 = ローカル完結** — gspread禁止、credentials.json不要、run.batダブルクリックで動く
- **スプシ版 = GASデプロイ** — 使い方シート付き、カスタムメニューから操作
- **GUIテストは初回のみ** — 2回目以降はcurl API直接でブロッキングを回避

## ブロッキング対策（重要）
- **migrate APIは3-5分かかる** — Playwright wait_forでブロック中は何もできない
- **初回のみGUIテスト** — UX確認用。以降はcurl + run_in_backgroundで非ブロッキング
- **JWTトークン取得** — 初回Playwrightログイン時に `page.evaluate(() => window.Clerk.session.getToken())` で取得→Bash変数に保存
- **Python版テストはtest.py一発** — ZIPに同梱済み、手動チェック不要

## ループ手順

### Step 1: 初回セットアップ（1回だけ）
1. Playwright MCPで本番サイトにログイン
2. JWTトークンを取得:
   ```
   mcp__claude-in-chrome__javascript_tool: window.Clerk.session.getToken()
   ```
3. 取得したトークンをBash変数に保存:
   ```bash
   JWT="eyJ..."
   API_URL="https://excel-migration-api.salmonbay-4f8a43a0.japaneast.azurecontainerapps.io"
   ```

### Step 2: E2Eテスト（初回: GUI / 2回目以降: curl API）

#### 初回: GUIテスト（UX確認）
- Playwright MCPで本番サイトにアクセス
- **移行モード: 「両方出力」を必ず選択**
- ファイルアップロード→移行開始→完了まで wait_for(300)
- スクリーンショット→「IT素人が初見で使えるか」で評価
- 以降のイテレーションはcurlに切り替え

#### 2回目以降: curl API直接（非ブロッキング）
```bash
# run_in_backgroundで実行 → ブロックなし
curl -s -X POST "$API_URL/api/migrate" \
  -H "Authorization: Bearer $JWT" \
  -F "files=@test.xlsm" \
  -F "trackMode=both" \
  -o /tmp/migrate_result.json
```
- run_in_backgroundで実行 → 待機中にStep 3の調査を並列実行
- 完了通知後に結果を確認

### Step 3: 生成物チェック（並列で実行）

生成物がない場合はスキップ。**サブエージェントで並列実行すること。**

#### 3A: スプシ版チェック（Agent: Playwright MCP）
1. Google Sheetsをブラウザで開く
2. マクロメニュー（カスタムメニュー）が表示されるか確認
3. メニューから実際に機能を実行
4. 使い方シートの内容確認（技術用語なし、メニュー名一致、手順が具体的）

#### 3B: Python版チェック（Agent: Bash実行）— test.py一発
```bash
# ZIPをダウンロード→展開→test.py実行
curl -s "$ZIP_URL" -o /tmp/python_local.zip
cd /tmp && unzip -o python_local.zip
cd *_local && python-embed/python.exe test.py
```
test.pyが自動チェックする内容:
- 起動→終了テスト
- 無効番号エラー
- 各メニュー番号のクラッシュチェック
- 全.pyファイルの構文チェック
- gspread/google.oauth2禁止ワードチェック
- 必須ファイル存在チェック

### Step 4: 問題発見
- テスト結果と生成物チェックから問題を洗い出す
- **問題なし → 完了報告して終了**
- 問題あり → 優先度の高いものから修正へ

### Step 5: 修正実装
- 問題に対してコードを修正。修正先の判断基準:
  - UI/表示の問題 → フロントエンド（frontend/src/）
  - 使い方シートの問題 → DeployService.cs
  - Python変換がgspread使ってる → PythonConvertService.cs（プロンプト修正）
  - Python構文エラー → PythonConvertService.cs（プロンプト or max_tokens）
  - Python実行エラー → PythonConvertService.cs（変換マッピング）
  - main.pyの問題 → PythonPackagerService.cs
  - test.pyの問題 → PythonPackagerService.cs（BuildTestPyメソッド）
  - Track判定の問題 → TrackRouterService.cs
  - メニュー表示/CLI UXの問題 → PythonPackagerService.cs（テンプレート）
  - ZIP構造の問題 → PythonPackagerService.cs + Dockerfile
- `dotnet build` でビルド確認

### Step 6: デプロイ
- `git add` → `git commit` → `git push origin main`
- `gh run list` でデプロイ完了確認（3-4分）
- バックエンドウォームアップ（curl でAPI叩く）

### Step 7: Step 2 に戻る（curlモード）

## 並列化ルール
- Step 3A（スプシ）と Step 3B（Python）は **サブエージェントで同時実行**
- **curl API実行中（run_in_background）に修正方針の調査を並列で走らせる** — これがループが回る鍵
- デプロイ待ち中にスキル改善やメモリ更新を並列で走らせてよい
- JWTトークンは1時間で期限切れ → 長時間ループ時は再取得

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
- **test.pyが全パスすること**

## スキル改善ルール（メタルール）
- **1周回ったらこのスキル自体を改善する** — 発見したパターンや知見をスキルに追記
- 新しい修正先パターン → Step 5 に追加
- 新しいテスト観点 → テスト観点に追加
- 新しい既知問題 → 既知パターンに追加

## 既知パターン
- 構文エラー除外されたGASがonOpenを壊す → BuildUsageSheetをDeploy後に除外ファイルを除いて呼ぶ（修正済み）
- ボタンなしExcelでカスタムメニューが出ない → ConvertServiceがButtonContext空でonOpen生成をスキップ → ThisWorkbookモジュールには常にonOpen生成（修正済み）
- 使い方シートのTODOコメントに技術用語が混入 → DeployServiceの使い方シート生成ロジックでフィルタリング（未対応）
- Python版にgspreadが混入 → PythonConvertServiceのプロンプトでFORBIDDEN指定（修正済み）
- Python Embeddable の ._pth に `..` が必要 → 親ディレクトリのmodules/を見つけるため（修正済み）
- pip download で --platform win_amd64 必要 → Linux上でWindows用wheelを取得するため（修正済み）
- 「両方出力」モードでPython ZIPが出ない → MigrateControllerがTrack2Modulesのみ使用→allModulesに修正（修正済み）
- 自動判定モードではPython版が出ないファイルがある → テスト時は「両方出力」固定
- **Playwright wait_forがブロックしてループが回らない** → 初回のみGUI、以降curl API（修正済み）

## テストファイル
- `C:\Users\fab24\excel-migration-os\bunkai.xlsm`
- `D:\クローム　ダウンロード\Excelマクロツール・ACCESS\引取明細書作成補助ツール(改)全受注区分.xlsm`

## 本番URL
- https://frontend-self-ten-98.vercel.app
- https://excel-migration-api.salmonbay-4f8a43a0.japaneast.azurecontainerapps.io
