# Technology Stack

## Architecture

フロントエンド（Next.js on Vercel）+ バックエンド（C# ASP.NET Core on Azure Container Apps）の分離構成。
VBA→GAS/Python変換はClaude APIを使用。Google操作はDrive/Sheets/Apps Script APIを直接呼び出し。

## Core Technologies

- **Frontend**: TypeScript, Next.js 14 (App Router), Tailwind CSS
- **Backend**: C# (.NET 9), ASP.NET Core
- **AI**: Claude API (Sonnet 4.6 / Haiku 4.5) for VBA conversion
- **Auth**: Clerk (JWT)
- **Billing**: Stripe (metered)
- **Hosting**: Vercel (frontend), Azure Container Apps (backend)

## Key Libraries

- **Backend**: NPOI + OpenMcdf (Excel parsing), System.Text.Json
- **Frontend**: @clerk/nextjs, stripe (server-side)
- **Python output**: gspread, pywin32, openpyxl (generated code dependencies)

## Development Standards

### Type Safety
- C#: nullable enabled, strict typing
- TypeScript: strict mode, explicit types on API boundaries

### Code Quality
- Backend: Japanese user-facing messages, English code/comments
- Frontend: All UI text in Japanese
- 使い方シート: 技術用語禁止、平易な日本語のみ

### Testing
- E2Eテスト: Playwright MCP (ブラウザで実運用と同じ手順)
- 単体テスト: 未導入（今後検討）

## Development Environment

### Required Tools
- .NET 9 SDK, Node.js 18+, Python 3.10+ (ローカル版テスト用)

### Common Commands
```bash
# Backend dev: cd backend/ExcelMigrationApi && dotnet run
# Frontend dev: cd frontend && npm run dev
# Backend build: cd backend/ExcelMigrationApi && dotnet build
# Frontend build: cd frontend && npx next build
```

## Key Technical Decisions

- **二系統変換**: GASで不可能な機能をPythonで救済（win32comでCOM再現率95%+）
- **EPPlus→NPOI+OpenMcdf**: 商用ライセンス回避
- **使い方シート**: SimplifyForEndUserでサブジェクトマップ方式（regex連鎖は崩壊するため廃止）
- **Track判定**: regexパターンマッチでモジュール単位振り分け（関数単位はPhase 2）

---
_Document standards and patterns, not every dependency_
