# Project Structure

## Organization Philosophy

フロントエンド/バックエンド分離。バックエンドはMVC風（Controllers + Services + Models）。
仕様は `docs/specs/` にMarkdownで管理（SDD方式）。

## Directory Patterns

### Backend Services
**Location**: `backend/ExcelMigrationApi/Services/`
**Purpose**: ビジネスロジック。各機能を独立したServiceクラスに分離
**Pattern**: `{Feature}Service.cs` — ScanService, ConvertService, PythonConvertService, TrackRouterService, DeployService等

### Backend Models
**Location**: `backend/ExcelMigrationApi/Models/`
**Purpose**: リクエスト/レスポンスのDTO
**Pattern**: `{Feature}Models.cs` — 1ファイルに関連モデルをまとめる

### Backend Controllers
**Location**: `backend/ExcelMigrationApi/Controllers/`
**Purpose**: APIエンドポイント定義。薄いコントローラ（ロジックはServiceに委譲）
**Pattern**: `{Feature}Controller.cs`

### Frontend Pages
**Location**: `frontend/src/app/{route}/page.tsx`
**Purpose**: Next.js App Router のページコンポーネント
**Pattern**: 1ページ1ファイル、"use client"

### Frontend API
**Location**: `frontend/src/lib/api.ts`
**Purpose**: 全APIの型定義と呼び出し関数を集約
**Pattern**: interface定義 + export async function

### Specifications
**Location**: `docs/specs/`
**Purpose**: 機能仕様（Requirements, Design）。実装前に書き、コードと一緒にgit管理
**Pattern**: `requirements.md`, `design.md`, `{feature}.md`

## Naming Conventions

- **C# files**: PascalCase (`TrackRouterService.cs`)
- **TypeScript files**: kebab-case for routes, camelCase for lib
- **API endpoints**: lowercase (`/api/migrate`, `/api/scan`)
- **Database/config**: snake_case in JSON, PascalCase in C#

## Import Organization

```typescript
// Frontend: absolute imports with @/
import { migrateFiles, type MigrateReport } from "@/lib/api";
```

```csharp
// Backend: namespace-based
using ExcelMigrationApi.Models;
using ExcelMigrationApi.Services;
```

## Code Organization Principles

- Service層にビジネスロジック集中、Controllerは薄く
- フロントエンドの型定義はapi.tsに集約
- 日本語のユーザー向けメッセージはバックエンドで生成（フロントは表示のみ）
- 仕様変更時は `docs/specs/` を先に更新

---
_Document patterns, not file trees. New files following patterns shouldn't require updates_
