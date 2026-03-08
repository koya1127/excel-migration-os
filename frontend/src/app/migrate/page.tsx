"use client";

import { useState, useCallback, useRef } from "react";
import {
  migrateFiles,
  type MigrateReport,
  type UploadReport,
  type ExtractReport,
  type ConvertReport,
  type DeployReport,
} from "@/lib/api";

function Spinner() {
  return (
    <svg className="h-4 w-4 animate-spin" viewBox="0 0 24 24" fill="none">
      <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
      <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" />
    </svg>
  );
}

function StepIndicator({
  step,
  label,
  status,
}: {
  step: number;
  label: string;
  status: "pending" | "active" | "success" | "error";
}) {
  const colors = {
    pending: "border-gray-300 bg-gray-100 text-gray-400",
    active: "border-blue-500 bg-blue-50 text-blue-600 animate-pulse",
    success: "border-green-500 bg-green-50 text-green-600",
    error: "border-red-500 bg-red-50 text-red-600",
  };
  const icons = {
    pending: <span className="text-sm font-bold">{step}</span>,
    active: <Spinner />,
    success: (
      <svg className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7" />
      </svg>
    ),
    error: (
      <svg className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12" />
      </svg>
    ),
  };

  return (
    <div className="flex items-center gap-3">
      <div
        className={`flex h-10 w-10 items-center justify-center rounded-full border-2 ${colors[status]}`}
      >
        {icons[status]}
      </div>
      <span
        className={`text-sm font-medium ${
          status === "pending" ? "text-gray-400" : status === "error" ? "text-red-700" : "text-gray-900"
        }`}
      >
        {label}
      </span>
    </div>
  );
}

function UploadStepDetail({ data }: { data: UploadReport }) {
  return (
    <div className="space-y-2">
      <div className="flex gap-4 text-sm">
        <span className="text-gray-500">ファイル数: <span className="font-medium text-gray-900">{data.fileCount}</span></span>
        <span className="text-green-600">成功: {data.successCount}</span>
        {data.failureCount > 0 && <span className="text-red-600">失敗: {data.failureCount}</span>}
      </div>
      {data.files.map((f, i) => (
        <div key={i} className="flex items-center gap-3 text-xs">
          <span className={`inline-flex items-center rounded-full px-2 py-0.5 font-medium ${
            f.status.toLowerCase() === "success" ? "bg-green-100 text-green-800" : "bg-red-100 text-red-800"
          }`}>
            {f.status}
          </span>
          <span className="text-gray-700">{f.fileName}</span>
          {f.webViewLink && (
            <a href={f.webViewLink} target="_blank" rel="noopener noreferrer" className="text-blue-600 hover:underline">
              Drive で開く
            </a>
          )}
          {f.error && <span className="text-red-500">{f.error}</span>}
        </div>
      ))}
    </div>
  );
}

function ExtractStepDetail({ data }: { data: ExtractReport }) {
  return (
    <div className="flex gap-6 text-sm">
      <span className="text-gray-500">ファイル: <span className="font-medium text-gray-900">{data.fileCount}</span></span>
      <span className="text-gray-500">VBA モジュール: <span className="font-medium text-blue-700">{data.moduleCount}</span></span>
      <span className="text-gray-500">コントロール: <span className="font-medium text-purple-700">{data.controls.length}</span></span>
    </div>
  );
}

function ConvertStepDetail({ data }: { data: ConvertReport }) {
  return (
    <div className="space-y-2">
      <div className="flex gap-4 text-sm">
        <span className="text-gray-500">合計: <span className="font-medium text-gray-900">{data.total}</span></span>
        <span className="text-green-600">成功: {data.success}</span>
        {data.failed > 0 && <span className="text-red-600">失敗: {data.failed}</span>}
      </div>
      {data.results.map((r, i) => (
        <div key={i} className="flex items-center gap-3 text-xs">
          <span className={`inline-flex items-center rounded-full px-2 py-0.5 font-medium ${
            r.status.toLowerCase() === "success" ? "bg-green-100 text-green-800" : "bg-red-100 text-red-800"
          }`}>
            {r.status}
          </span>
          <span className="text-gray-700">{r.moduleName}</span>
          {r.error && <span className="text-red-500">{r.error}</span>}
        </div>
      ))}
    </div>
  );
}

function DeployStepDetail({ data }: { data: DeployReport }) {
  return (
    <div className="space-y-2">
      <div className="flex flex-col gap-1 text-sm">
        {data.scriptId && (
          <span className="text-gray-500">Script ID: <span className="font-mono text-xs text-gray-700">{data.scriptId}</span></span>
        )}
        {data.spreadsheetId && (
          <span className="text-gray-500">Spreadsheet ID: <span className="font-mono text-xs text-gray-700">{data.spreadsheetId}</span></span>
        )}
      </div>
      {data.filesDeployed && data.filesDeployed.length > 0 && (
        <div>
          <p className="text-xs font-medium text-gray-500 mb-1">デプロイ済みファイル:</p>
          <div className="flex flex-wrap gap-2">
            {data.filesDeployed.map((f, i) => (
              <span key={i} className="inline-flex items-center rounded-md bg-blue-50 border border-blue-200 px-2 py-0.5 text-xs text-blue-700">
                {f}
              </span>
            ))}
          </div>
        </div>
      )}
      {data.error && (
        <p className="text-xs text-red-600">{data.error}</p>
      )}
    </div>
  );
}

export default function MigratePage() {
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [convertToSheets, setConvertToSheets] = useState(true);
  const [folderId, setFolderId] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [report, setReport] = useState<MigrateReport | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [expandedSteps, setExpandedSteps] = useState<Set<number>>(new Set());
  const fileInputRef = useRef<HTMLInputElement>(null);
  const folderInputRef = useRef<HTMLInputElement>(null);

  const addFiles = useCallback((files: File[]) => {
    const accepted = files.filter((f) => /\.xlsm$/i.test(f.name));
    if (accepted.length > 0) {
      setSelectedFiles((prev) => [...prev, ...accepted]);
    }
  }, []);

  const handleFiles = useCallback((files: FileList | null) => {
    if (!files) return;
    addFiles(Array.from(files));
  }, [addFiles]);

  const readEntries = async (entry: FileSystemEntry): Promise<File[]> => {
    if (entry.isFile) {
      return new Promise((resolve) => {
        (entry as FileSystemFileEntry).file((f) => resolve([f]));
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

  const handleDrop = useCallback(
    async (e: React.DragEvent) => {
      e.preventDefault();
      setIsDragging(false);
      const items = e.dataTransfer.items;
      if (items) {
        const entries = Array.from(items)
          .map((item) => item.webkitGetAsEntry())
          .filter((e): e is FileSystemEntry => e !== null);
        if (entries.some((e) => e.isDirectory)) {
          const allFiles = await Promise.all(entries.map(readEntries));
          addFiles(allFiles.flat());
          return;
        }
      }
      handleFiles(e.dataTransfer.files);
    },
    [handleFiles, addFiles]
  );

  const toggleStep = (step: number) => {
    setExpandedSteps((prev) => {
      const next = new Set(prev);
      if (next.has(step)) next.delete(step);
      else next.add(step);
      return next;
    });
  };

  const handleMigrate = async () => {
    if (selectedFiles.length === 0) return;
    setLoading(true);
    setError(null);
    setReport(null);
    try {
      const result = await migrateFiles(selectedFiles, convertToSheets, folderId || undefined);
      setReport(result);
      // Auto-expand all steps
      setExpandedSteps(new Set([1, 2, 3, 4]));
    } catch (err) {
      setError(err instanceof Error ? err.message : "移行に失敗しました");
    } finally {
      setLoading(false);
    }
  };

  const getStepStatus = (stepData: UploadReport | ExtractReport | ConvertReport | DeployReport | null, stepKey: string): "pending" | "active" | "success" | "error" => {
    if (loading && !report) return "active";
    if (!report) return "pending";
    if (!stepData) return "error";
    // Check for errors based on step type
    if (stepKey === "upload") {
      const d = stepData as UploadReport;
      return d.failureCount > 0 ? "error" : "success";
    }
    if (stepKey === "convert") {
      const d = stepData as ConvertReport;
      return d.failed > 0 ? "error" : "success";
    }
    if (stepKey === "deploy") {
      const d = stepData as DeployReport;
      return d.error ? "error" : "success";
    }
    return "success";
  };

  return (
    <div className="mx-auto max-w-7xl px-6 py-8">
      <div className="text-center mb-8">
        <h1 className="text-3xl font-bold text-gray-900">エンドツーエンド移行</h1>
        <p className="mt-2 text-sm text-gray-500">
          Excel/VBA ファイルを Google Sheets + Apps Script に一括移行します
        </p>
      </div>

      {/* Upload area */}
      <div className="rounded-xl border border-gray-200 bg-white p-6 shadow-sm">
        <div
          className={`flex flex-col items-center justify-center rounded-lg border-2 border-dashed p-10 transition-colors ${
            isDragging ? "border-blue-400 bg-blue-50" : "border-gray-300 hover:border-gray-400"
          }`}
          onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
          onDragLeave={() => setIsDragging(false)}
          onDrop={handleDrop}
        >
          <input
            ref={fileInputRef}
            type="file"
            className="hidden"
            accept=".xlsm"
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
          <p className="mt-3 text-sm text-gray-600">ここにファイル・フォルダをドロップ</p>
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
          <p className="mt-2 text-xs text-gray-400">.xlsm のみ対応（VBA マクロを含むファイル）</p>
        </div>

        {selectedFiles.length > 0 && (
          <div className="mt-4">
            <div className="flex items-center justify-between">
              <p className="text-sm font-medium text-gray-700">
                {selectedFiles.length} 件のファイルを選択中
              </p>
              <button className="text-xs text-red-500 hover:text-red-700" onClick={() => setSelectedFiles([])}>
                すべてクリア
              </button>
            </div>
            <div className="mt-2 flex flex-wrap gap-2">
              {selectedFiles.map((f, i) => (
                <span key={i} className="inline-flex items-center gap-1 rounded-md bg-gray-100 px-2.5 py-1 text-xs text-gray-700">
                  {f.name}
                  <button
                    className="ml-1 text-gray-400 hover:text-gray-600"
                    onClick={() => setSelectedFiles((prev) => prev.filter((_, j) => j !== i))}
                  >
                    x
                  </button>
                </span>
              ))}
            </div>
          </div>
        )}

        {/* Options */}
        <div className="mt-6 space-y-4">
          <div className="flex items-center gap-3">
            <label className="relative inline-flex cursor-pointer items-center">
              <input
                type="checkbox"
                checked={convertToSheets}
                onChange={(e) => setConvertToSheets(e.target.checked)}
                className="peer sr-only"
              />
              <div className="h-6 w-11 rounded-full bg-gray-200 after:absolute after:left-[2px] after:top-[2px] after:h-5 after:w-5 after:rounded-full after:border after:border-gray-300 after:bg-white after:transition-all after:content-[''] peer-checked:bg-blue-600 peer-checked:after:translate-x-full peer-checked:after:border-white"></div>
            </label>
            <span className="text-sm text-gray-700">Google Sheets 形式に変換</span>
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">
              Google Drive フォルダ ID（オプション）
            </label>
            <input
              type="text"
              value={folderId}
              onChange={(e) => setFolderId(e.target.value)}
              className="w-full max-w-md rounded-lg border border-gray-300 px-3 py-2 text-sm focus:border-blue-500 focus:ring-1 focus:ring-blue-500 outline-none"
              placeholder="空欄の場合はマイドライブのルートにアップロード"
            />
          </div>
        </div>

        <div className="mt-6">
          <button
            onClick={handleMigrate}
            disabled={selectedFiles.length === 0 || loading}
            className="rounded-lg bg-gradient-to-r from-blue-600 to-indigo-600 px-8 py-3 text-sm font-semibold text-white shadow-md hover:from-blue-500 hover:to-indigo-500 disabled:opacity-50 disabled:cursor-not-allowed transition-all"
          >
            {loading ? (
              <span className="flex items-center gap-2"><Spinner />移行中...</span>
            ) : (
              "移行開始"
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

      {/* Progress / Results */}
      {(loading || report) && (
        <div className="mt-8">
          <div className="rounded-xl border border-gray-200 bg-white shadow-sm overflow-hidden">
            <div className="border-b border-gray-200 px-6 py-4">
              <h2 className="text-lg font-semibold text-gray-900">移行プロセス</h2>
            </div>
            <div className="p-6 space-y-0">
              {/* Step 1: Upload */}
              <div>
                <div
                  className={`flex items-center justify-between py-4 ${report ? "cursor-pointer hover:bg-gray-50 -mx-6 px-6" : ""}`}
                  onClick={() => report && toggleStep(1)}
                >
                  <StepIndicator
                    step={1}
                    label="アップロード"
                    status={report ? getStepStatus(report.upload, "upload") : (loading ? "active" : "pending")}
                  />
                  {report && (
                    <span className="text-gray-400 text-sm">
                      {expandedSteps.has(1) ? "\u25BC" : "\u25B6"}
                    </span>
                  )}
                </div>
                {report && expandedSteps.has(1) && report.upload && (
                  <div className="ml-[52px] pb-4">
                    <UploadStepDetail data={report.upload} />
                  </div>
                )}
                <div className="ml-5 h-6 border-l-2 border-gray-200"></div>
              </div>

              {/* Step 2: Extract */}
              <div>
                <div
                  className={`flex items-center justify-between py-4 ${report ? "cursor-pointer hover:bg-gray-50 -mx-6 px-6" : ""}`}
                  onClick={() => report && toggleStep(2)}
                >
                  <StepIndicator
                    step={2}
                    label="VBA 抽出"
                    status={report ? getStepStatus(report.extract, "extract") : "pending"}
                  />
                  {report && (
                    <span className="text-gray-400 text-sm">
                      {expandedSteps.has(2) ? "\u25BC" : "\u25B6"}
                    </span>
                  )}
                </div>
                {report && expandedSteps.has(2) && report.extract && (
                  <div className="ml-[52px] pb-4">
                    <ExtractStepDetail data={report.extract} />
                  </div>
                )}
                <div className="ml-5 h-6 border-l-2 border-gray-200"></div>
              </div>

              {/* Step 3: Convert */}
              <div>
                <div
                  className={`flex items-center justify-between py-4 ${report ? "cursor-pointer hover:bg-gray-50 -mx-6 px-6" : ""}`}
                  onClick={() => report && toggleStep(3)}
                >
                  <StepIndicator
                    step={3}
                    label="VBA → GAS 変換"
                    status={report ? getStepStatus(report.convert, "convert") : "pending"}
                  />
                  {report && (
                    <span className="text-gray-400 text-sm">
                      {expandedSteps.has(3) ? "\u25BC" : "\u25B6"}
                    </span>
                  )}
                </div>
                {report && expandedSteps.has(3) && report.convert && (
                  <div className="ml-[52px] pb-4">
                    <ConvertStepDetail data={report.convert} />
                  </div>
                )}
                <div className="ml-5 h-6 border-l-2 border-gray-200"></div>
              </div>

              {/* Step 4: Deploy */}
              <div>
                <div
                  className={`flex items-center justify-between py-4 ${report ? "cursor-pointer hover:bg-gray-50 -mx-6 px-6" : ""}`}
                  onClick={() => report && toggleStep(4)}
                >
                  <StepIndicator
                    step={4}
                    label="Apps Script デプロイ"
                    status={report ? getStepStatus(report.deploy, "deploy") : "pending"}
                  />
                  {report && (
                    <span className="text-gray-400 text-sm">
                      {expandedSteps.has(4) ? "\u25BC" : "\u25B6"}
                    </span>
                  )}
                </div>
                {report && expandedSteps.has(4) && report.deploy && (
                  <div className="ml-[52px] pb-4">
                    <DeployStepDetail data={report.deploy} />
                  </div>
                )}
              </div>
            </div>
          </div>

          {/* Summary card after completion */}
          {report && !loading && (
            <div className="mt-6 rounded-xl border border-green-200 bg-green-50 p-6 shadow-sm">
              <h3 className="font-semibold text-green-800">移行完了</h3>
              <p className="mt-1 text-sm text-green-700">
                {report.upload?.successCount || 0} 件のファイルをアップロードし、
                {report.convert?.success || 0} 件の VBA モジュールを変換しました。
              </p>
              {report.deploy?.scriptId && (
                <p className="mt-1 text-xs text-green-600 font-mono">
                  Script ID: {report.deploy.scriptId}
                </p>
              )}
            </div>
          )}
        </div>
      )}
    </div>
  );
}
