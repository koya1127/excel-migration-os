"use client";

import { useState, useCallback, useRef } from "react";
import { extractFiles, type ExtractReport, type VbaModule } from "@/lib/api";

function Spinner() {
  return (
    <svg className="h-4 w-4 animate-spin" viewBox="0 0 24 24" fill="none">
      <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
      <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" />
    </svg>
  );
}

export default function ExtractPage() {
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [report, setReport] = useState<ExtractReport | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [expandedModule, setExpandedModule] = useState<number | null>(null);
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

  const handleDrop = useCallback(
    async (e: React.DragEvent) => {
      e.preventDefault();
      setIsDragging(false);

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

  const handleExtract = async () => {
    if (selectedFiles.length === 0) return;
    setLoading(true);
    setError(null);
    try {
      const result = await extractFiles(selectedFiles);
      setReport(result);
    } catch (err) {
      setError(err instanceof Error ? err.message : "抽出に失敗しました");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="mx-auto max-w-7xl px-6 py-8">
      <h1 className="text-2xl font-bold text-gray-900">VBA 抽出</h1>
      <p className="mt-1 text-sm text-gray-500">
        .xlsm ファイルから VBA モジュールとフォームコントロールを抽出します
      </p>

      {/* Upload area */}
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
            onClick={handleExtract}
            disabled={selectedFiles.length === 0 || loading}
            className="rounded-lg bg-blue-600 px-6 py-2.5 text-sm font-semibold text-white shadow-sm hover:bg-blue-500 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
          >
            {loading ? (
              <span className="flex items-center gap-2"><Spinner />抽出中...</span>
            ) : (
              "抽出開始"
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
          {/* Summary cards */}
          <div className="grid grid-cols-3 gap-4">
            <div className="rounded-xl border border-gray-200 bg-white p-5 shadow-sm">
              <p className="text-sm text-gray-500">ファイル数</p>
              <p className="mt-1 text-3xl font-bold text-gray-900">{report.fileCount}</p>
            </div>
            <div className="rounded-xl border border-blue-200 bg-blue-50 p-5 shadow-sm">
              <p className="text-sm text-blue-700">VBA モジュール</p>
              <p className="mt-1 text-3xl font-bold text-blue-800">{report.moduleCount}</p>
            </div>
            <div className="rounded-xl border border-purple-200 bg-purple-50 p-5 shadow-sm">
              <p className="text-sm text-purple-700">フォームコントロール</p>
              <p className="mt-1 text-3xl font-bold text-purple-800">{report.controls.length}</p>
            </div>
          </div>

          {/* VBA Modules table */}
          {report.modules.length > 0 && (
            <div className="rounded-xl border border-gray-200 bg-white shadow-sm overflow-hidden">
              <div className="border-b border-gray-200 px-6 py-4">
                <h2 className="text-lg font-semibold text-gray-900">VBA モジュール</h2>
              </div>
              <div className="overflow-x-auto">
                <table className="min-w-full text-sm">
                  <thead>
                    <tr className="border-b border-gray-200 text-left text-xs font-medium uppercase text-gray-500">
                      <th className="px-4 py-3 w-8"></th>
                      <th className="px-4 py-3">ソースファイル</th>
                      <th className="px-4 py-3">モジュール名</th>
                      <th className="px-4 py-3">種類</th>
                      <th className="px-4 py-3">行数</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100">
                    {report.modules.map((m: VbaModule, i: number) => (
                      <>
                        <tr
                          key={`mod-${i}`}
                          className="hover:bg-gray-50 cursor-pointer"
                          onClick={() => setExpandedModule(expandedModule === i ? null : i)}
                        >
                          <td className="px-4 py-3 text-gray-400">
                            {expandedModule === i ? "\u25BC" : "\u25B6"}
                          </td>
                          <td className="px-4 py-3 text-gray-600">{m.sourceFile}</td>
                          <td className="px-4 py-3 font-medium text-gray-900">{m.moduleName}</td>
                          <td className="px-4 py-3">
                            <span className="inline-flex items-center rounded-full bg-gray-100 px-2.5 py-0.5 text-xs font-medium text-gray-700">
                              {m.moduleType}
                            </span>
                          </td>
                          <td className="px-4 py-3 text-gray-600">{m.codeLines}</td>
                        </tr>
                        {expandedModule === i && (
                          <tr key={`code-${i}`}>
                            <td colSpan={5} className="px-4 py-4 bg-gray-50">
                              <pre className="rounded-lg bg-gray-900 text-gray-100 p-4 text-xs font-mono overflow-x-auto max-h-96 overflow-y-auto">
                                {m.code}
                              </pre>
                            </td>
                          </tr>
                        )}
                      </>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* Form Controls table */}
          {report.controls.length > 0 && (
            <div className="rounded-xl border border-gray-200 bg-white shadow-sm overflow-hidden">
              <div className="border-b border-gray-200 px-6 py-4">
                <h2 className="text-lg font-semibold text-gray-900">フォームコントロール</h2>
              </div>
              <div className="overflow-x-auto">
                <table className="min-w-full text-sm">
                  <thead>
                    <tr className="border-b border-gray-200 text-left text-xs font-medium uppercase text-gray-500">
                      <th className="px-4 py-3">ソースファイル</th>
                      <th className="px-4 py-3">シート</th>
                      <th className="px-4 py-3">コントロール名</th>
                      <th className="px-4 py-3">種類</th>
                      <th className="px-4 py-3">ラベル</th>
                      <th className="px-4 py-3">マクロ</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100">
                    {report.controls.map((c, i) => (
                      <tr key={i} className="hover:bg-gray-50">
                        <td className="px-4 py-3 text-gray-600">{c.sourceFile}</td>
                        <td className="px-4 py-3 text-gray-600">{c.sheetName}</td>
                        <td className="px-4 py-3 font-medium text-gray-900">{c.controlName}</td>
                        <td className="px-4 py-3">
                          <span className="inline-flex items-center rounded-full bg-purple-50 border border-purple-200 px-2.5 py-0.5 text-xs font-medium text-purple-700">
                            {c.controlType}
                          </span>
                        </td>
                        <td className="px-4 py-3 text-gray-600">{c.label}</td>
                        <td className="px-4 py-3">
                          {c.macro ? (
                            <span className="inline-flex items-center rounded-full bg-amber-50 border border-amber-200 px-2.5 py-0.5 text-xs font-medium text-amber-700">
                              {c.macro}
                            </span>
                          ) : (
                            <span className="text-gray-300">-</span>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
