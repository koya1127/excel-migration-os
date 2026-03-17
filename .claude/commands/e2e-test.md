# E2Eテスト実行

Playwright MCPで本番環境のE2Eテストを実行する。

## 引数
$ARGUMENTS — テスト対象（例: 「migrateページでbunkai.xlsmをローカルのみモードで移行」「出来上がったスプシの使い方シートを確認」）

## 手順

### 1. テスト準備
- テスト対象に応じてPlaywright MCPで本番URLにアクセス
- 必要ならファイルアップロード（`C:\Users\fab24\excel-migration-os\bunkai.xlsm` をテストファイルとして使用）

### 2. テスト実行
- ボタンクリック、フォーム入力等の操作をPlaywright MCPで実行
- 処理完了まで待機（バックグラウンドタスクで待つ）
- 結果画面のスナップショットとスクリーンショットを取得

### 3. 結果検証
- 期待される結果と実際の結果を比較
- 問題があれば具体的に記録（行番号、テキスト内容、スクリーンショット）
- IT素人目線で「初見で使えるか」を評価

### 4. 報告
- テスト結果をまとめて報告
- 問題があれば改善提案を含める

## 本番URL
- フロントエンド: https://frontend-self-ten-98.vercel.app
- バックエンドAPI: https://excel-migration-api.salmonbay-4f8a43a0.japaneast.azurecontainerapps.io

## テストファイル
- `C:\Users\fab24\excel-migration-os\bunkai.xlsm`
- テスト用Excelフォルダ: `D:\クローム　ダウンロード\Excelマクロツール・ACCESS\`
