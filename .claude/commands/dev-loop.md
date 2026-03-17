# 自律開発ループ

対象機能に対して、E2Eテストで問題を自分で発見し、修正→デプロイ→再テストを自律的に繰り返す。
何を直すかは自分で判断する。人間の介入なしに回し続ける。

## 引数
$ARGUMENTS — 対象機能（例: 「migrateページのローカルのみモード」「使い方シート」「Scanページ」）

## ループ手順

### Step 1: E2Eテスト
- Playwright MCPで本番サイト（https://frontend-self-ten-98.vercel.app）にアクセス
- 対象機能を実際に操作する（ファイルアップロード、ボタンクリック等）
- 処理完了まで待機
- スクリーンショットを撮る
- 結果を「IT素人が初見で使えるか」の基準で評価

### Step 2: 問題発見
- テスト結果から問題点を自分で洗い出す
- 問題がなければ → 完了報告して終了
- 問題があれば → 優先度の高いものから修正

### Step 3: 修正実装
- 問題に対してコードを修正
- `dotnet build` / `npx next build` でビルド確認

### Step 4: デプロイ
- `git add` → `git commit` → `git push origin main`
- `gh run list` でデプロイ完了を確認（3-4分待つ）
- バックエンドのウォームアップ（curl でAPIを叩く）

### Step 5: Step 1 に戻る

## ルール
- 最大10イテレーションで打ち切り
- 何を直すかは自分で判断する。ユーザーに聞かない
- 「問題なし」と判断したら即終了（無駄に回さない）
- 各イテレーションで「何を見つけて何を直したか」を1行で記録
- デプロイ待ちはバックグラウンドタスク（sleep + gh run list）
- コミットメッセージに修正内容を明記

## テスト観点
- 機能が正常に動くか
- エラー表示が適切か
- IT素人が見て意味がわかるか（専門用語がないか）
- 日本語が自然か
- 表示が崩れていないか
- 生成物（スプシ、Pythonコード）の品質

## テストファイル
- `C:\Users\fab24\excel-migration-os\bunkai.xlsm`

## 本番URL
- https://frontend-self-ten-98.vercel.app
- https://excel-migration-api.salmonbay-4f8a43a0.japaneast.azurecontainerapps.io
