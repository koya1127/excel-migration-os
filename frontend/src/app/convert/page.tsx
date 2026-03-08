"use client";

import { useState, useCallback, useRef } from "react";
import {
  extractFiles,
  convertBatch,
  convertSingle,
  checkSubscription,
  type ExtractReport,
  type ConvertReport,
  type ConvertResult,
} from "@/lib/api";

function Spinner() {
  return (
    <svg className="h-4 w-4 animate-spin" viewBox="0 0 24 24" fill="none">
      <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
      <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" />
    </svg>
  );
}

function StatusBadge({ status }: { status: string }) {
  const isSuccess = status.toLowerCase() === "success";
  return (
    <span
      className={`inline-flex items-center rounded-full px-2.5 py-0.5 text-xs font-medium ${
        isSuccess ? "bg-green-100 text-green-800" : "bg-red-100 text-red-800"
      }`}
    >
      {status}
    </span>
  );
}

function CopyButton({ text }: { text: string }) {
  const [copied, setCopied] = useState(false);
  const handleCopy = async () => {
    await navigator.clipboard.writeText(text);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };
  return (
    <button
      onClick={handleCopy}
      className="rounded-md bg-gray-700 px-2 py-1 text-xs text-gray-200 hover:bg-gray-600 transition-colors"
    >
      {copied ? "Copied!" : "Copy"}
    </button>
  );
}

export default function ConvertPage() {
  const [mode, setMode] = useState<"file" | "code">("file");

  // File mode state
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const folderInputRef = useRef<HTMLInputElement>(null);

  // Code mode state
  const [vbaCode, setVbaCode] = useState("");
  const [moduleName, setModuleName] = useState("Module1");

  // Shared state
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [report, setReport] = useState<ConvertReport | null>(null);
  const [extractReport, setExtractReport] = useState<ExtractReport | null>(null);
  const [expandedResult, setExpandedResult] = useState<number | null>(null);

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

  const ensureSubscription = async (): Promise<boolean> => {
    const { hasSubscription } = await checkSubscription();
    if (!hasSubscription) {
      setError("変換機能を使用するには従量課金プランへの登録が必要です。設定ページから登録してください。");
      return false;
    }
    return true;
  };

  const handleFileConvert = async () => {
    if (selectedFiles.length === 0) return;
    setLoading(true);
    setError(null);
    try {
      if (!(await ensureSubscription())) return;
      const ext = await extractFiles(selectedFiles);
      setExtractReport(ext);
      const result = await convertBatch(ext);
      setReport(result);
    } catch (err) {
      setError(err instanceof Error ? err.message : "変換に失敗しました");
    } finally {
      setLoading(false);
    }
  };

  const handleCodeConvert = async () => {
    if (!vbaCode.trim()) return;
    setLoading(true);
    setError(null);
    setExtractReport(null);
    try {
      if (!(await ensureSubscription())) return;
      const result = await convertSingle({
        vbaCode,
        moduleName: moduleName || "Module1",
        moduleType: "Module",
      });
      setReport(result);
    } catch (err) {
      setError(err instanceof Error ? err.message : "変換に失敗しました");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="mx-auto max-w-7xl px-6 py-8">
      <h1 className="text-2xl font-bold text-gray-900">VBA → GAS 変換</h1>
      <p className="mt-1 text-sm text-gray-500">
        VBA コードを Google Apps Script に変換します
      </p>

      {/* Mode selector */}
      <div className="mt-6 flex gap-2">
        <button
          onClick={() => setMode("file")}
          className={`rounded-lg px-4 py-2 text-sm font-medium transition-colors ${
            mode === "file"
              ? "bg-blue-600 text-white"
              : "bg-white text-gray-700 border border-gray-300 hover:bg-gray-50"
          }`}
        >
          ファイルから変換
        </button>
        <button
          onClick={() => setMode("code")}
          className={`rounded-lg px-4 py-2 text-sm font-medium transition-colors ${
            mode === "code"
              ? "bg-blue-600 text-white"
              : "bg-white text-gray-700 border border-gray-300 hover:bg-gray-50"
          }`}
        >
          コードを直接入力
        </button>
      </div>

      {/* File mode */}
      {mode === "file" && (
        <div className="mt-6 rounded-xl border border-gray-200 bg-white p-6 shadow-sm">
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
            <p className="mt-2 text-xs text-gray-400">.xlsm のみ対応</p>
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

          <div className="mt-6">
            <button
              onClick={handleFileConvert}
              disabled={selectedFiles.length === 0 || loading}
              className="rounded-lg bg-blue-600 px-6 py-2.5 text-sm font-semibold text-white shadow-sm hover:bg-blue-500 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
            >
              {loading ? (
                <span className="flex items-center gap-2"><Spinner />変換中...</span>
              ) : (
                "変換開始"
              )}
            </button>
          </div>
        </div>
      )}

      {/* Code mode */}
      {mode === "code" && (
        <div className="mt-6 rounded-xl border border-gray-200 bg-white p-6 shadow-sm">
          <div className="space-y-4">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">モジュール名</label>
              <input
                type="text"
                value={moduleName}
                onChange={(e) => setModuleName(e.target.value)}
                className="w-full max-w-xs rounded-lg border border-gray-300 px-3 py-2 text-sm focus:border-blue-500 focus:ring-1 focus:ring-blue-500 outline-none"
                placeholder="Module1"
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">VBA コード</label>
              <textarea
                value={vbaCode}
                onChange={(e) => setVbaCode(e.target.value)}
                className="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm font-mono focus:border-blue-500 focus:ring-1 focus:ring-blue-500 outline-none"
                rows={15}
                placeholder={"Sub HelloWorld()\n    MsgBox \"Hello, World!\"\nEnd Sub"}
              />
            </div>
            <button
              onClick={handleCodeConvert}
              disabled={!vbaCode.trim() || loading}
              className="rounded-lg bg-blue-600 px-6 py-2.5 text-sm font-semibold text-white shadow-sm hover:bg-blue-500 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
            >
              {loading ? (
                <span className="flex items-center gap-2"><Spinner />変換中...</span>
              ) : (
                "変換開始"
              )}
            </button>
          </div>
        </div>
      )}

      {/* Error */}
      {error && (
        <div className="mt-6 rounded-lg border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">
          {error}
        </div>
      )}

      {/* Results */}
      {report && (
        <div className="mt-8 space-y-6">
          {/* Summary cards */}
          <div className="grid grid-cols-3 gap-4">
            <div className="rounded-xl border border-gray-200 bg-white p-5 shadow-sm">
              <p className="text-sm text-gray-500">合計</p>
              <p className="mt-1 text-3xl font-bold text-gray-900">{report.total}</p>
            </div>
            <div className="rounded-xl border border-green-200 bg-green-50 p-5 shadow-sm">
              <p className="text-sm text-green-700">成功</p>
              <p className="mt-1 text-3xl font-bold text-green-800">{report.success}</p>
            </div>
            <div className="rounded-xl border border-red-200 bg-red-50 p-5 shadow-sm">
              <p className="text-sm text-red-700">失敗</p>
              <p className="mt-1 text-3xl font-bold text-red-800">{report.failed}</p>
            </div>
          </div>

          {/* Conversion results */}
          <div className="space-y-4">
            {report.results.map((r: ConvertResult, i: number) => (
              <div key={i} className="rounded-xl border border-gray-200 bg-white shadow-sm overflow-hidden">
                <div
                  className="flex items-center justify-between px-6 py-4 cursor-pointer hover:bg-gray-50"
                  onClick={() => setExpandedResult(expandedResult === i ? null : i)}
                >
                  <div className="flex items-center gap-3">
                    <span className="text-gray-400">
                      {expandedResult === i ? "\u25BC" : "\u25B6"}
                    </span>
                    <span className="font-medium text-gray-900">{r.moduleName}</span>
                    <StatusBadge status={r.status} />
                  </div>
                  {r.error && (
                    <span className="text-xs text-red-600">{r.error}</span>
                  )}
                </div>
                {expandedResult === i && (
                  <div className="border-t border-gray-200 px-6 py-4">
                    <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                      {/* Original VBA */}
                      {extractReport && extractReport.modules[i] && (
                        <div>
                          <p className="text-xs font-medium text-gray-500 mb-2">VBA (元のコード)</p>
                          <pre className="rounded-lg bg-gray-900 text-gray-100 p-4 text-xs font-mono overflow-x-auto max-h-96 overflow-y-auto">
                            {extractReport.modules[i].code}
                          </pre>
                        </div>
                      )}
                      {mode === "code" && (
                        <div>
                          <p className="text-xs font-medium text-gray-500 mb-2">VBA (元のコード)</p>
                          <pre className="rounded-lg bg-gray-900 text-gray-100 p-4 text-xs font-mono overflow-x-auto max-h-96 overflow-y-auto">
                            {vbaCode}
                          </pre>
                        </div>
                      )}
                      {/* Converted GAS */}
                      <div>
                        <div className="flex items-center justify-between mb-2">
                          <p className="text-xs font-medium text-gray-500">Google Apps Script (変換後)</p>
                          {r.gasCode && <CopyButton text={r.gasCode} />}
                        </div>
                        <pre className="rounded-lg bg-gray-900 text-gray-100 p-4 text-xs font-mono overflow-x-auto max-h-96 overflow-y-auto">
                          {r.gasCode || "変換結果なし"}
                        </pre>
                      </div>
                    </div>
                  </div>
                )}
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}
