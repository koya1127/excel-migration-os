# 自律開発ループ

対象機能に対して、E2Eテスト＋生成物の品質チェックで問題を自分で発見し、修正→デプロイ→再テストを自律的に繰り返す。
何を直すかは自分で判断する。人間の介入なしに回し続ける。

## 引数
$ARGUMENTS — 対象機能（例: 「migrateページ全体」「自動判定モードの二系統移行」「Scanページ」）

## ループ手順

### Step 1: E2Eテスト
- Playwright MCPで本番サイトにアクセス
- 対象機能を実際に操作（ファイルアップロード、ボタンクリック等）
- 処理完了まで待機
- スクリーンショットを撮る
- 結果を「IT素人が初見で使えるか」の基準で評価

### Step 2: 生成物チェック
- **スプシ版**: Google Sheetsを開いて使い方シートの内容確認。技術用語がないか、メニュー項目は正しいか
- **Python版**: ZIP URLがあればダウンロード→展開→以下を実行:
  - `python -m py_compile *.py` で全ファイル構文チェック
  - `echo "0" | python main.py` で起動テスト
  - main.pyの関数バインドが正しいか確認
  - requirements.txtに必要なパッケージが揃っているか確認
- 生成物がない場合はスキップ

### Step 3: 問題発見
- テスト結果と生成物チェックから問題を洗い出す
- **問題なし → 完了報告して終了**
- 問題あり → 優先度の高いものから修正へ

### Step 4: 修正実装
- 問題に対してコードを修正。修正先の判断基準:
  - UI/表示の問題 → フロントエンド（frontend/src/）
  - 使い方シートの問題 → DeployService.cs
  - Python構文エラー → PythonConvertService.cs（プロンプト or max_tokens）
  - Python実行エラー → PythonConvertService.cs（変換マッピング）
  - main.pyの問題 → PythonPackagerService.cs
  - Track判定の問題 → TrackRouterService.cs
- `dotnet build` / `npx next build` でビルド確認

### Step 5: デプロイ
- `git add` → `git commit` → `git push origin main`
- `gh run list` でデプロイ完了確認（3-4分）
- バックエンドウォームアップ（curl でAPI叩く）

### Step 6: Step 1 に戻る

## ルール
- 最大10イテレーション
- 何を直すかは自分で判断。ユーザーに聞かない
- 「問題なし」と判断したら即終了
- 各イテレーションで「何を見つけて何を直したか」を1行で記録
- 生成されたPythonコードを直接編集しない（バックエンドのプロンプト/ロジックを修正）
- gspread/win32com連携テストはスキップ（認証・環境依存）
- Python構文チェック + import + main.py起動がパスすればPython側はOK

## テスト観点
- 機能が正常に動くか
- エラー表示が適切か（local_onlyでデプロイエラーにならないか等）
- IT素人が見て意味がわかるか
- 日本語が自然か
- 生成されたスプシの使い方シート品質
- 生成されたPythonの構文エラー・実行エラー

## テストファイル
- `C:\Users\fab24\excel-migration-os\bunkai.xlsm`

## 本番URL
- https://frontend-self-ten-98.vercel.app
- https://excel-migration-api.salmonbay-4f8a43a0.japaneast.azurecontainerapps.io
