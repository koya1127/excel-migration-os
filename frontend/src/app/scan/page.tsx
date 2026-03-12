"use client";

import { useState, useCallback, useRef, Fragment } from "react";
import { scanFiles, type ScanReport, type FileReport, type GroupSummary } from "@/lib/api";

function formatBytes(bytes: number): string {
  if (bytes === 0) return "0 B";
  const k = 1024;
  const sizes = ["B", "KB", "MB", "GB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + " " + sizes[i];
}

function DifficultyBadge({ difficulty }: { difficulty: string }) {
  const colors: Record<string, string> = {
    Easy: "bg-green-100 text-green-800",
    Medium: "bg-yellow-100 text-yellow-800",
    Hard: "bg-red-100 text-red-800",
  };
  return (
    <span
      className={`inline-flex items-center rounded-full px-2.5 py-0.5 text-xs font-medium ${colors[difficulty] || "bg-gray-100 text-gray-800"}`}
    >
      {difficulty}
    </span>
  );
}

function RiskBar({ score }: { score: number }) {
  const color = score < 40 ? "bg-green-500" : score < 70 ? "bg-yellow-500" : "bg-red-500";
  return (
    <div className="flex items-center gap-2">
      <div className="h-2 w-20 rounded-full bg-gray-200">
        <div
          className={`h-2 rounded-full ${color}`}
          style={{ width: `${Math.min(score, 100)}%` }}
        />
      </div>
      <span className="text-xs text-gray-600">{score}</span>
    </div>
  );
}

type SortKey = "name" | "sizeBytes" | "hasMacro" | "formulaCount" | "incompatibleFunctionCount" | "riskScore";
type SortDir = "asc" | "desc";

function SortHeader({ label, sortKey, currentKey, currentDir, onSort, title }: {
  label: string; sortKey: SortKey; currentKey: SortKey; currentDir: SortDir;
  onSort: (key: SortKey) => void; title?: string;
}) {
  const active = currentKey === sortKey;
  return (
    <th
      className="px-4 py-3 cursor-pointer select-none hover:text-gray-900 transition-colors"
      onClick={() => onSort(sortKey)}
      title={title}
    >
      <span className="inline-flex items-center gap-1">
        {label}
        <span className="text-[10px]">{active ? (currentDir === "asc" ? "\u25B2" : "\u25BC") : "\u25B4\u25BE"}</span>
      </span>
    </th>
  );
}

function noteTag(text: string) {
  const isWarning = text.includes("非互換") || text.includes("外部リンク") || text.includes("揮発性");
  const isInfo = text.includes("数式が多い");
  const cls = isWarning
    ? "bg-red-50 text-red-700 border border-red-200"
    : isInfo
    ? "bg-yellow-50 text-yellow-700 border border-yellow-200"
    : "bg-gray-100 text-gray-600";
  return cls;
}

function FileTable({ files }: { files: FileReport[]; nested?: boolean }) {
  const [sortKey, setSortKey] = useState<SortKey>("riskScore");
  const [sortDir, setSortDir] = useState<SortDir>("desc");

  const handleSort = (key: SortKey) => {
    if (sortKey === key) {
      setSortDir(d => d === "asc" ? "desc" : "asc");
    } else {
      setSortKey(key);
      setSortDir("desc");
    }
  };

  const sorted = [...files].sort((a, b) => {
    let va: string | number, vb: string | number;
    switch (sortKey) {
      case "name": va = a.path.toLowerCase(); vb = b.path.toLowerCase(); break;
      case "sizeBytes": va = a.sizeBytes; vb = b.sizeBytes; break;
      case "hasMacro": va = a.vbaModuleCount; vb = b.vbaModuleCount; break;
      case "formulaCount": va = a.formulaCount; vb = b.formulaCount; break;
      case "incompatibleFunctionCount": va = a.incompatibleFunctionCount; vb = b.incompatibleFunctionCount; break;
      case "riskScore": va = a.riskScore; vb = b.riskScore; break;
      default: va = 0; vb = 0;
    }
    const cmp = va < vb ? -1 : va > vb ? 1 : 0;
    return sortDir === "asc" ? cmp : -cmp;
  });

  const sortProps = { currentKey: sortKey, currentDir: sortDir, onSort: handleSort };

  return (
    <div className="overflow-x-auto">
      <table className="min-w-full text-sm">
        <thead>
          <tr className="border-b border-gray-200 text-left text-xs font-medium uppercase text-gray-500">
            <SortHeader label="ファイル名" sortKey="name" {...sortProps} />
            <th className="px-4 py-3">形式</th>
            <SortHeader label="サイズ" sortKey="sizeBytes" {...sortProps} />
            <SortHeader label="マクロ" sortKey="hasMacro" title="VBAマクロの有無とモジュール数" {...sortProps} />
            <SortHeader label="数式" sortKey="formulaCount" title="セル内の数式（=で始まるセル）の数" {...sortProps} />
            <SortHeader label="Sheets非対応" sortKey="incompatibleFunctionCount" title="Google Sheetsに存在しない、または動作が異なる関数の数" {...sortProps} />
            <SortHeader label="移行リスク" sortKey="riskScore" title="移行の難しさ（マクロ・外部リンク・非対応関数などから算出）" {...sortProps} />
            <th className="px-4 py-3">詳細</th>
          </tr>
        </thead>
        <tbody className="divide-y divide-gray-100">
          {sorted.map((f, i) => {
            const fileName = f.path.split(/[/\\]/).pop() || f.path;
            return (
              <tr key={i} className="hover:bg-gray-50">
                <td className="px-4 py-3 font-medium text-gray-900 max-w-[280px]" title={fileName}>
                  <span className="block truncate">{fileName}</span>
                </td>
                <td className="px-4 py-3 text-gray-500 text-xs">{f.extension}</td>
                <td className="px-4 py-3 text-gray-600 whitespace-nowrap">{formatBytes(f.sizeBytes)}</td>
                <td className="px-4 py-3">
                  {f.hasMacro ? (
                    <span className="cursor-help" title={`VBAモジュール ${f.vbaModuleCount} 個を含む`}>
                      <span className="inline-flex items-center rounded-full bg-amber-50 border border-amber-200 px-2 py-0.5 text-xs font-medium text-amber-700">
                        {f.vbaModuleCount}個
                      </span>
                    </span>
                  ) : (
                    <span className="text-gray-300 text-xs">-</span>
                  )}
                </td>
                <td className="px-4 py-3 text-gray-600">
                  {f.formulaCount > 0 ? f.formulaCount.toLocaleString() : <span className="text-gray-300">-</span>}
                </td>
                <td className="px-4 py-3">
                  {f.incompatibleFunctionCount > 0 ? (
                    <span className="inline-flex items-center rounded-full bg-red-50 border border-red-200 px-2 py-0.5 text-xs font-medium text-red-700">
                      {f.incompatibleFunctionCount}個
                    </span>
                  ) : (
                    <span className="text-gray-300 text-xs">-</span>
                  )}
                </td>
                <td className="px-4 py-3">
                  <RiskBar score={f.riskScore} />
                </td>
                <td className="px-4 py-3 max-w-[300px]">
                  <div className="flex flex-wrap gap-1">
                    {Array.isArray(f.notes)
                      ? f.notes.map((n, j) => (
                          <span key={j} className={`inline-block rounded-full px-2 py-0.5 text-[11px] ${noteTag(n)}`}>
                            {n}
                          </span>
                        ))
                      : f.notes && f.notes !== "リスク低"
                        ? <span className={`inline-block rounded-full px-2 py-0.5 text-[11px] ${noteTag(String(f.notes))}`}>{f.notes}</span>
                        : <span className="text-gray-300 text-xs">-</span>
                    }
                  </div>
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

function GroupTable({ groups, files }: { groups: GroupSummary[]; files: FileReport[] }) {
  const [expanded, setExpanded] = useState<Set<number>>(new Set());

  const toggle = (idx: number) => {
    setExpanded((prev) => {
      const next = new Set(prev);
      if (next.has(idx)) next.delete(idx);
      else next.add(idx);
      return next;
    });
  };

  return (
    <div className="space-y-2">
      <table className="min-w-full text-sm">
        <thead>
          <tr className="border-b border-gray-200 text-left text-xs font-medium uppercase text-gray-500">
            <th className="px-4 py-3 w-8"></th>
            <th className="px-4 py-3">グループ名</th>
            <th className="px-4 py-3">ファイル数</th>
            <th className="px-4 py-3" title="グループ内でマクロを含むファイルの数">マクロあり</th>
            <th className="px-4 py-3" title="グループ内のVBAモジュール合計数">VBA合計</th>
            <th className="px-4 py-3" title="グループ内ファイルの移行リスク平均">平均リスク</th>
            <th className="px-4 py-3" title="グループ内で最もリスクが高いファイルのスコア">最大リスク</th>
            <th className="px-4 py-3" title="Google Sheetsに非対応の関数の合計数">Sheets非対応</th>
            <th className="px-4 py-3" title="Easy=そのまま移行可 / Medium=マクロ変換が必要 / Hard=非対応関数あり・要手動対応">移行難易度</th>
          </tr>
        </thead>
        <tbody className="divide-y divide-gray-100">
          {groups.map((g, i) => (
            <Fragment key={i}>
              <tr
                className="hover:bg-gray-50 cursor-pointer"
                onClick={() => toggle(i)}
              >
                <td className="px-4 py-3 text-gray-400">
                  {expanded.has(i) ? "\u25BC" : "\u25B6"}
                </td>
                <td className="px-4 py-3 font-medium text-gray-900">{g.groupName}</td>
                <td className="px-4 py-3 text-gray-600">{g.fileCount}</td>
                <td className="px-4 py-3 text-gray-600">{g.macroFileCount}</td>
                <td className="px-4 py-3 text-gray-600">{g.totalVbaModules}</td>
                <td className="px-4 py-3">
                  <RiskBar score={Math.round(g.avgRiskScore)} />
                </td>
                <td className="px-4 py-3">
                  <RiskBar score={g.maxRiskScore} />
                </td>
                <td className="px-4 py-3">
                  {g.totalIncompatibleFunctions > 0 ? (
                    <span className="text-red-600 font-medium">{g.totalIncompatibleFunctions}</span>
                  ) : (
                    <span className="text-gray-400">0</span>
                  )}
                </td>
                <td className="px-4 py-3">
                  <DifficultyBadge difficulty={g.migrationDifficulty} />
                </td>
              </tr>
              {expanded.has(i) && (
                <tr key={`detail-${i}`}>
                  <td colSpan={9} className="bg-gray-50 px-4 py-4">
                    <FileTable files={g.fileIndices.map((idx) => files[idx])} />
                  </td>
                </tr>
              )}
            </Fragment>
          ))}
        </tbody>
      </table>
    </div>
  );
}

export default function ScanPage() {
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [groupBy, setGroupBy] = useState("subfolder");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [report, setReport] = useState<ScanReport | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const folderInputRef = useRef<HTMLInputElement>(null);

  const addExcelFiles = useCallback((files: File[]) => {
    const accepted = files.filter((f) =>
      /\.(xls|xlsx|xlsm)$/i.test(f.name)
    );
    if (accepted.length > 0) {
      setSelectedFiles((prev) => {
        const combined = [...prev, ...accepted];
        if (combined.length > 1000) {
          setError(`ファイル数が${combined.length}件です。1回のスキャンは最大1,000件までです。フォルダを分けてお試しください。`);
          return prev;
        }
        return combined;
      });
    }
  }, []);

  const handleFiles = useCallback((files: FileList | null) => {
    if (!files) return;
    addExcelFiles(Array.from(files));
  }, [addExcelFiles]);

  const handleDrop = useCallback(
    async (e: React.DragEvent) => {
      e.preventDefault();
      setIsDragging(false);

      // Recursively read directory entries from drag & drop
      // Preserves folder structure by encoding fullPath into the File name
      const readEntries = async (entry: FileSystemEntry): Promise<File[]> => {
        if (entry.isFile) {
          return new Promise((resolve) => {
            (entry as FileSystemFileEntry).file((f) => {
              // Encode relative path into file name so backend can reconstruct folder structure
              const relativePath = entry.fullPath.replace(/^\//, "");
              const fileWithPath = new File([f], relativePath, { type: f.type, lastModified: f.lastModified });
              resolve([fileWithPath]);
            });
          });
        }
        if (entry.isDirectory) {
          const reader = (entry as FileSystemDirectoryEntry).createReader();
          const entries = await new Promise<FileSystemEntry[]>((resolve) => {
            const all: FileSystemEntry[] = [];
            const readBatch = () => {
              reader.readEntries((batch) => {
                if (batch.length === 0) { resolve(all); return; }
                all.push(...batch);
                readBatch();
              });
            };
            readBatch();
          });
          const nested = await Promise.all(entries.map(readEntries));
          return nested.flat();
        }
        return [];
      };

      const items = e.dataTransfer.items;
      if (items) {
        const entries = Array.from(items)
          .map((item) => item.webkitGetAsEntry())
          .filter((e): e is FileSystemEntry => e !== null);
        if (entries.some((e) => e.isDirectory)) {
          const allFiles = await Promise.all(entries.map(readEntries));
          addExcelFiles(allFiles.flat());
          return;
        }
      }
      handleFiles(e.dataTransfer.files);
    },
    [handleFiles, addExcelFiles]
  );

  const handleScan = async () => {
    if (selectedFiles.length === 0) return;
    setLoading(true);
    setError(null);
    try {
      const result = await scanFiles(selectedFiles, groupBy);
      setReport(result);
    } catch (err) {
      setError(err instanceof Error ? err.message : "スキャンに失敗しました");
    } finally {
      setLoading(false);
    }
  };

  // Per-file difficulty for summary cards
  const fileDifficulty = (f: FileReport) => {
    if (f.riskScore >= 70 || f.incompatibleFunctionCount > 0) return "Hard";
    if (f.riskScore >= 40 || f.hasMacro) return "Medium";
    return "Easy";
  };

  const difficultyCounts = report
    ? {
        Easy: report.files.filter((f) => fileDifficulty(f) === "Easy").length,
        Medium: report.files.filter((f) => fileDifficulty(f) === "Medium").length,
        Hard: report.files.filter((f) => fileDifficulty(f) === "Hard").length,
      }
    : null;

  const showGroups = report && report.groups.length > 1;

  return (
    <div className="mx-auto max-w-7xl px-6 py-8">
      <h1 className="text-2xl font-bold text-gray-900">スキャン</h1>
      <p className="mt-1 text-sm text-gray-500">
        Excel ファイルをアップロードして移行リスクを分析します
      </p>

      {/* Upload area */}
      <div className="mt-6 rounded-xl border border-gray-200 bg-white p-6 shadow-sm">
        <div
          className={`flex flex-col items-center justify-center rounded-lg border-2 border-dashed p-10 transition-colors ${
            isDragging
              ? "border-blue-400 bg-blue-50"
              : "border-gray-300 hover:border-gray-400"
          }`}
          onDragOver={(e) => {
            e.preventDefault();
            setIsDragging(true);
          }}
          onDragLeave={() => setIsDragging(false)}
          onDrop={handleDrop}
        >
          <input
            ref={fileInputRef}
            type="file"
            className="hidden"
            accept=".xls,.xlsx,.xlsm"
            multiple
            onChange={(e) => handleFiles(e.target.files)}
          />
          <input
            ref={folderInputRef}
            type="file"
            className="hidden"
            /* @ts-expect-error webkitdirectory is not in types */
            webkitdirectory=""
            multiple
            onChange={(e) => handleFiles(e.target.files)}
          />
          <svg className="h-10 w-10 text-gray-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M12 16V4m0 0l-4 4m4-4l4 4M4 20h16" />
          </svg>
          <p className="mt-3 text-sm text-gray-600">
            ここにファイル・フォルダをドロップ
          </p>
          <div className="mt-3 flex gap-3">
            <button
              type="button"
              className="rounded-md bg-white px-3 py-1.5 text-xs font-medium text-gray-700 border border-gray-300 hover:bg-gray-50"
              onClick={(e) => { e.stopPropagation(); fileInputRef.current?.click(); }}
            >
              ファイルを選択
            </button>
            <button
              type="button"
              className="rounded-md bg-white px-3 py-1.5 text-xs font-medium text-gray-700 border border-gray-300 hover:bg-gray-50"
              onClick={(e) => { e.stopPropagation(); folderInputRef.current?.click(); }}
            >
              フォルダを選択
            </button>
          </div>
          <p className="mt-2 text-xs text-gray-400">.xls, .xlsx, .xlsm に対応（フォルダ内も再帰的にスキャン）</p>
        </div>

        {selectedFiles.length > 0 && (
          <div className="mt-4">
            <div className="flex items-center justify-between">
              <p className="text-sm font-medium text-gray-700">
                {selectedFiles.length} 件のファイルを選択中
              </p>
              <button
                className="text-xs text-red-500 hover:text-red-700"
                onClick={() => setSelectedFiles([])}
              >
                すべてクリア
              </button>
            </div>
            <div className="mt-2 flex flex-wrap gap-2">
              {selectedFiles.map((f, i) => (
                <span
                  key={i}
                  className="inline-flex items-center gap-1 rounded-md bg-gray-100 px-2.5 py-1 text-xs text-gray-700"
                >
                  {f.name}
                  <button
                    className="ml-1 text-gray-400 hover:text-gray-600"
                    onClick={(e) => {
                      e.stopPropagation();
                      setSelectedFiles((prev) => prev.filter((_, j) => j !== i));
                    }}
                  >
                    x
                  </button>
                </span>
              ))}
            </div>
          </div>
        )}

        {/* Group by + scan button */}
        <div className="mt-6 flex flex-wrap items-center gap-6">
          <div className="flex items-center gap-3">
            <span className="text-sm font-medium text-gray-700">グループ分け:</span>
            {([
              { value: "subfolder", label: "フォルダ別", desc: "フォルダごとにまとめる" },
              { value: "prefix", label: "ファイル名別", desc: "「_」より前の名前でまとめる" },
              { value: "none", label: "グループ分けしない", desc: "" },
            ] as const).map((opt) => (
              <label key={opt.value} className="flex items-center gap-1.5 text-sm text-gray-600 cursor-pointer" title={opt.desc}>
                <input
                  type="radio"
                  name="groupBy"
                  value={opt.value}
                  checked={groupBy === opt.value}
                  onChange={() => setGroupBy(opt.value)}
                  className="accent-blue-600"
                />
                {opt.label}
              </label>
            ))}
          </div>
          <button
            onClick={handleScan}
            disabled={selectedFiles.length === 0 || loading}
            className="rounded-lg bg-blue-600 px-6 py-2.5 text-sm font-semibold text-white shadow-sm hover:bg-blue-500 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
          >
            {loading ? (
              <span className="flex items-center gap-2">
                <svg className="h-4 w-4 animate-spin" viewBox="0 0 24 24" fill="none">
                  <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                  <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" />
                </svg>
                スキャン中...
              </span>
            ) : (
              "スキャン開始"
            )}
          </button>
        </div>
      </div>

      {/* Error */}
      {error && (
        <div className="mt-6 rounded-lg border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">
          {error}
        </div>
      )}

      {/* Results */}
      {report && (
        <div className="mt-8 space-y-6">
          {/* Summary metrics */}
          <div className="grid grid-cols-2 gap-4 sm:grid-cols-4">
            <div className="rounded-xl border border-gray-200 bg-white p-5 shadow-sm">
              <p className="text-sm text-gray-500">ファイル数</p>
              <p className="mt-1 text-3xl font-bold text-gray-900">{report.fileCount}</p>
            </div>
            {difficultyCounts && (
              <>
                <div className="rounded-xl border border-green-200 bg-green-50 p-5 shadow-sm">
                  <p className="text-sm text-green-700">Easy</p>
                  <p className="mt-1 text-3xl font-bold text-green-800">{difficultyCounts.Easy}</p>
                  <p className="text-xs text-green-600 mt-1">そのまま移行可</p>
                </div>
                <div className="rounded-xl border border-yellow-200 bg-yellow-50 p-5 shadow-sm">
                  <p className="text-sm text-yellow-700">Medium</p>
                  <p className="mt-1 text-3xl font-bold text-yellow-800">{difficultyCounts.Medium}</p>
                  <p className="text-xs text-yellow-600 mt-1">マクロ変換が必要</p>
                </div>
                <div className="rounded-xl border border-red-200 bg-red-50 p-5 shadow-sm">
                  <p className="text-sm text-red-700">Hard</p>
                  <p className="mt-1 text-3xl font-bold text-red-800">{difficultyCounts.Hard}</p>
                  <p className="text-xs text-red-600 mt-1">手動対応が必要</p>
                </div>
              </>
            )}
          </div>

          {/* Group summary */}
          {showGroups && (
            <div className="rounded-xl border border-gray-200 bg-white shadow-sm overflow-hidden">
              <div className="border-b border-gray-200 px-6 py-4">
                <h2 className="text-lg font-semibold text-gray-900">グループサマリー</h2>
                <p className="text-xs text-gray-500 mt-0.5">
                  グループ化: {report.groupBy} / クリックで詳細を展開
                </p>
              </div>
              <GroupTable groups={report.groups} files={report.files} />
            </div>
          )}

          {/* All files */}
          <div className="rounded-xl border border-gray-200 bg-white shadow-sm overflow-hidden">
            <div className="border-b border-gray-200 px-6 py-4">
              <h2 className="text-lg font-semibold text-gray-900">全ファイル詳細</h2>
            </div>
            <FileTable files={report.files} />
          </div>
        </div>
      )}
    </div>
  );
}
