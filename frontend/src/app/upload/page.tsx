"use client";

import { useState, useCallback, useRef } from "react";
import { uploadFiles, type UploadReport } from "@/lib/api";

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

export default function UploadPage() {
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [convertToSheets, setConvertToSheets] = useState(true);
  const [folderId, setFolderId] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [report, setReport] = useState<UploadReport | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const folderInputRef = useRef<HTMLInputElement>(null);

  const addFiles = useCallback((files: File[]) => {
    const accepted = files.filter((f) => /\.(xls|xlsx|xlsm)$/i.test(f.name));
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

  const handleUpload = async () => {
    if (selectedFiles.length === 0) return;
    setLoading(true);
    setError(null);
    try {
      const result = await uploadFiles(selectedFiles, convertToSheets, folderId || undefined);
      setReport(result);
    } catch (err) {
      setError(err instanceof Error ? err.message : "アップロードに失敗しました");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="mx-auto max-w-7xl px-6 py-8">
      <h1 className="text-2xl font-bold text-gray-900">Google Drive アップロード</h1>
      <p className="mt-1 text-sm text-gray-500">
        Excel ファイルを Google Drive にアップロードします
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
          <p className="mt-2 text-xs text-gray-400">.xls, .xlsx, .xlsm に対応</p>
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
            onClick={handleUpload}
            disabled={selectedFiles.length === 0 || loading}
            className="rounded-lg bg-blue-600 px-6 py-2.5 text-sm font-semibold text-white shadow-sm hover:bg-blue-500 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
          >
            {loading ? (
              <span className="flex items-center gap-2"><Spinner />アップロード中...</span>
            ) : (
              "アップロード開始"
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
            <div className="rounded-xl border border-green-200 bg-green-50 p-5 shadow-sm">
              <p className="text-sm text-green-700">成功</p>
              <p className="mt-1 text-3xl font-bold text-green-800">{report.successCount}</p>
            </div>
            <div className="rounded-xl border border-red-200 bg-red-50 p-5 shadow-sm">
              <p className="text-sm text-red-700">失敗</p>
              <p className="mt-1 text-3xl font-bold text-red-800">{report.failureCount}</p>
            </div>
          </div>

          {/* File list */}
          <div className="rounded-xl border border-gray-200 bg-white shadow-sm overflow-hidden">
            <div className="border-b border-gray-200 px-6 py-4">
              <h2 className="text-lg font-semibold text-gray-900">アップロード結果</h2>
            </div>
            <div className="overflow-x-auto">
              <table className="min-w-full text-sm">
                <thead>
                  <tr className="border-b border-gray-200 text-left text-xs font-medium uppercase text-gray-500">
                    <th className="px-4 py-3">ファイル名</th>
                    <th className="px-4 py-3">ステータス</th>
                    <th className="px-4 py-3">リンク</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {report.files.map((f, i) => (
                    <tr key={i} className="hover:bg-gray-50">
                      <td className="px-4 py-3 font-medium text-gray-900">{f.fileName}</td>
                      <td className="px-4 py-3">
                        <StatusBadge status={f.status} />
                        {f.error && (
                          <span className="ml-2 text-xs text-red-600">{f.error}</span>
                        )}
                      </td>
                      <td className="px-4 py-3">
                        {f.webViewLink ? (
                          <a
                            href={f.webViewLink}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="text-blue-600 hover:text-blue-800 underline text-xs"
                          >
                            Google Drive で開く
                          </a>
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
        </div>
      )}
    </div>
  );
}
