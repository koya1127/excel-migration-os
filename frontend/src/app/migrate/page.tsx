"use client";

import { useState, useCallback, useRef } from "react";
import {
  uploadFiles,
  extractFiles,
  convertBatch,
  deployGas,
  checkSubscription,
  type UploadReport,
  type ExtractReport,
  type ConvertReport,
  type DeployReport,
  type GasFile,
} from "@/lib/api";

function Spinner() {
  return (
    <svg className="h-4 w-4 animate-spin" viewBox="0 0 24 24" fill="none">
      <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
      <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" />
    </svg>
  );
}

type StepStatus = "pending" | "active" | "success" | "error";

function StepIndicator({
  step,
  label,
  status,
  detail,
}: {
  step: number;
  label: string;
  status: StepStatus;
  detail?: string;
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
      <div className="flex flex-col">
        <span
          className={`text-sm font-medium ${
            status === "pending" ? "text-gray-400" : status === "error" ? "text-red-700" : "text-gray-900"
          }`}
        >
          {label}
        </span>
        {detail && (
          <span className="text-xs text-gray-500">{detail}</span>
        )}
      </div>
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
        <span className="text-gray-400">トークン: {(data.totalInputTokens + data.totalOutputTokens).toLocaleString()}</span>
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

interface StepState {
  upload: { status: StepStatus; detail: string; data: UploadReport | null };
  extract: { status: StepStatus; detail: string; data: ExtractReport | null };
  convert: { status: StepStatus; detail: string; data: ConvertReport | null };
  deploy: { status: StepStatus; detail: string; data: DeployReport | null };
}

const initialSteps: StepState = {
  upload: { status: "pending", detail: "", data: null },
  extract: { status: "pending", detail: "", data: null },
  convert: { status: "pending", detail: "", data: null },
  deploy: { status: "pending", detail: "", data: null },
};

export default function MigratePage() {
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [convertToSheets, setConvertToSheets] = useState(true);
  const [folderId, setFolderId] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [steps, setSteps] = useState<StepState>(initialSteps);
  const [isDragging, setIsDragging] = useState(false);
  const [expandedSteps, setExpandedSteps] = useState<Set<number>>(new Set());
  const [elapsed, setElapsed] = useState(0);
  const timerRef = useRef<ReturnType<typeof setInterval> | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const folderInputRef = useRef<HTMLInputElement>(null);

  const started = steps.upload.status !== "pending";

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

  const updateStep = (key: keyof StepState, update: Partial<StepState[keyof StepState]>) => {
    setSteps((prev) => ({ ...prev, [key]: { ...prev[key], ...update } }));
  };

  const handleMigrate = async () => {
    if (selectedFiles.length === 0) return;
    setLoading(true);
    setError(null);
    setSteps(initialSteps);
    setExpandedSteps(new Set());
    setElapsed(0);

    // Start timer
    const start = Date.now();
    timerRef.current = setInterval(() => {
      setElapsed(Math.floor((Date.now() - start) / 1000));
    }, 1000);

    try {
      // Subscription check
      const { hasSubscription } = await checkSubscription();
      if (!hasSubscription) {
        setError("移行機能を使用するには従量課金プランへの登録が必要です。設定ページから登録してください。");
        return;
      }

      // Step 1: Upload
      updateStep("upload", { status: "active", detail: `${selectedFiles.length} 件アップロード中...` });
      const uploadResult = await uploadFiles(selectedFiles, convertToSheets, folderId || undefined);
      updateStep("upload", {
        status: uploadResult.failureCount > 0 ? "error" : "success",
        detail: `${uploadResult.successCount}/${uploadResult.fileCount} 件成功`,
        data: uploadResult,
      });
      setExpandedSteps((prev) => new Set([...prev, 1]));

      // Step 2: Extract
      updateStep("extract", { status: "active", detail: "VBA コード抽出中..." });
      const extractResult = await extractFiles(selectedFiles);
      updateStep("extract", {
        status: "success",
        detail: `${extractResult.moduleCount} モジュール、${extractResult.controls.length} コントロール`,
        data: extractResult,
      });
      setExpandedSteps((prev) => new Set([...prev, 2]));

      if (extractResult.moduleCount === 0) {
        updateStep("convert", { status: "success", detail: "変換対象の VBA モジュールなし" });
        updateStep("deploy", { status: "success", detail: "デプロイ対象なし" });
        return;
      }

      // Step 3: Convert
      updateStep("convert", { status: "active", detail: `${extractResult.moduleCount} モジュールを AI 変換中...` });
      const convertResult = await convertBatch(extractResult);
      updateStep("convert", {
        status: convertResult.failed > 0 ? "error" : "success",
        detail: `${convertResult.success}/${convertResult.total} 成功（${(convertResult.totalInputTokens + convertResult.totalOutputTokens).toLocaleString()} トークン）`,
        data: convertResult,
      });
      setExpandedSteps((prev) => new Set([...prev, 3]));

      // Step 4: Deploy — per-file: each spreadsheet gets its own GAS
      const successUploads = uploadResult.files.filter((f) => f.status.toLowerCase() === "success" && f.driveFileId);
      const successConverts = convertResult.results.filter((r) => r.status.toLowerCase() === "success" && r.gasCode);

      if (successConverts.length === 0 || successUploads.length === 0) {
        updateStep("deploy", { status: "error", detail: "デプロイ可能なファイルがありません" });
        return;
      }

      // Group converted GAS by source file
      const gasByFile = new Map<string, GasFile[]>();
      for (const r of successConverts) {
        const key = r.sourceFile || "";
        if (!gasByFile.has(key)) gasByFile.set(key, []);
        gasByFile.get(key)!.push({ name: r.moduleName + ".gs", source: r.gasCode, type: "SERVER_JS" });
      }

      // Match each upload to its GAS files and deploy
      const deployResults: { fileName: string; scriptId?: string; error?: string }[] = [];
      let deployedCount = 0;
      let deployErrorCount = 0;

      updateStep("deploy", { status: "active", detail: `${successUploads.length} スプレッドシートにデプロイ中...` });

      for (const upload of successUploads) {
        const gasFiles = gasByFile.get(upload.fileName);
        if (!gasFiles || gasFiles.length === 0) {
          deployResults.push({ fileName: upload.fileName, error: "VBAモジュールなし（スキップ）" });
          continue;
        }

        try {
          const result = await deployGas(upload.driveFileId, gasFiles);
          if (result.error) {
            deployResults.push({ fileName: upload.fileName, error: result.error });
            deployErrorCount++;
          } else {
            deployResults.push({ fileName: upload.fileName, scriptId: result.scriptId });
            deployedCount++;
          }
        } catch (e) {
          deployResults.push({ fileName: upload.fileName, error: e instanceof Error ? e.message : "デプロイ失敗" });
          deployErrorCount++;
        }
      }

      const deployDetail = deployErrorCount > 0
        ? `${deployedCount}/${deployedCount + deployErrorCount} 成功`
        : `${deployedCount} スプレッドシートにデプロイ完了`;

      updateStep("deploy", {
        status: deployErrorCount > 0 && deployedCount === 0 ? "error" : deployErrorCount > 0 ? "error" : "success",
        detail: deployDetail,
        data: {
          generatedUtc: new Date().toISOString(),
          spreadsheetId: "",
          scriptId: "",
          fileCount: deployedCount,
          filesDeployed: deployResults.filter(d => d.scriptId).map(d => d.fileName),
          status: deployErrorCount > 0 ? "partial" : "success",
          error: deployResults.filter(d => d.error).map(d => `${d.fileName}: ${d.error}`).join("; "),
        } as DeployReport,
      });
      setExpandedSteps((prev) => new Set([...prev, 4]));

    } catch (err) {
      setError(err instanceof Error ? err.message : "移行に失敗しました");
      // Mark current active step as error
      const errMsg = err instanceof Error ? err.message : "エラー";
      setSteps((prev) => ({
        ...prev,
        upload: prev.upload.status === "active" ? { ...prev.upload, status: "error" as const, detail: errMsg } : prev.upload,
        extract: prev.extract.status === "active" ? { ...prev.extract, status: "error" as const, detail: errMsg } : prev.extract,
        convert: prev.convert.status === "active" ? { ...prev.convert, status: "error" as const, detail: errMsg } : prev.convert,
        deploy: prev.deploy.status === "active" ? { ...prev.deploy, status: "error" as const, detail: errMsg } : prev.deploy,
      }));
    } finally {
      setLoading(false);
      if (timerRef.current) {
        clearInterval(timerRef.current);
        timerRef.current = null;
      }
    }
  };

  const formatTime = (s: number) => {
    const m = Math.floor(s / 60);
    const sec = s % 60;
    return m > 0 ? `${m}分${sec}秒` : `${sec}秒`;
  };

  const allDone = !loading && started && steps.upload.status !== "pending";

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

        <div className="mt-6 flex items-center gap-4">
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
          {loading && (
            <span className="text-sm text-gray-500">経過時間: {formatTime(elapsed)}</span>
          )}
        </div>
      </div>

      {/* Error */}
      {error && (
        <div className="mt-6 rounded-lg border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">
          {error}
        </div>
      )}

      {/* Progress / Results */}
      {started && (
        <div className="mt-8">
          <div className="rounded-xl border border-gray-200 bg-white shadow-sm overflow-hidden">
            <div className="border-b border-gray-200 px-6 py-4 flex items-center justify-between">
              <h2 className="text-lg font-semibold text-gray-900">移行プロセス</h2>
              {!loading && allDone && (
                <span className="text-sm text-gray-500">完了 — {formatTime(elapsed)}</span>
              )}
            </div>
            <div className="p-6 space-y-0">
              {/* Step 1: Upload */}
              <div>
                <div
                  className={`flex items-center justify-between py-4 ${steps.upload.data ? "cursor-pointer hover:bg-gray-50 -mx-6 px-6" : ""}`}
                  onClick={() => steps.upload.data && toggleStep(1)}
                >
                  <StepIndicator step={1} label="アップロード" status={steps.upload.status} detail={steps.upload.detail} />
                  {steps.upload.data && (
                    <span className="text-gray-400 text-sm">
                      {expandedSteps.has(1) ? "\u25BC" : "\u25B6"}
                    </span>
                  )}
                </div>
                {expandedSteps.has(1) && steps.upload.data && (
                  <div className="ml-[52px] pb-4">
                    <UploadStepDetail data={steps.upload.data} />
                  </div>
                )}
                <div className="ml-5 h-6 border-l-2 border-gray-200"></div>
              </div>

              {/* Step 2: Extract */}
              <div>
                <div
                  className={`flex items-center justify-between py-4 ${steps.extract.data ? "cursor-pointer hover:bg-gray-50 -mx-6 px-6" : ""}`}
                  onClick={() => steps.extract.data && toggleStep(2)}
                >
                  <StepIndicator step={2} label="VBA 抽出" status={steps.extract.status} detail={steps.extract.detail} />
                  {steps.extract.data && (
                    <span className="text-gray-400 text-sm">
                      {expandedSteps.has(2) ? "\u25BC" : "\u25B6"}
                    </span>
                  )}
                </div>
                {expandedSteps.has(2) && steps.extract.data && (
                  <div className="ml-[52px] pb-4">
                    <ExtractStepDetail data={steps.extract.data} />
                  </div>
                )}
                <div className="ml-5 h-6 border-l-2 border-gray-200"></div>
              </div>

              {/* Step 3: Convert */}
              <div>
                <div
                  className={`flex items-center justify-between py-4 ${steps.convert.data ? "cursor-pointer hover:bg-gray-50 -mx-6 px-6" : ""}`}
                  onClick={() => steps.convert.data && toggleStep(3)}
                >
                  <StepIndicator step={3} label="VBA → GAS 変換" status={steps.convert.status} detail={steps.convert.detail} />
                  {steps.convert.data && (
                    <span className="text-gray-400 text-sm">
                      {expandedSteps.has(3) ? "\u25BC" : "\u25B6"}
                    </span>
                  )}
                </div>
                {expandedSteps.has(3) && steps.convert.data && (
                  <div className="ml-[52px] pb-4">
                    <ConvertStepDetail data={steps.convert.data} />
                  </div>
                )}
                <div className="ml-5 h-6 border-l-2 border-gray-200"></div>
              </div>

              {/* Step 4: Deploy */}
              <div>
                <div
                  className={`flex items-center justify-between py-4 ${steps.deploy.data ? "cursor-pointer hover:bg-gray-50 -mx-6 px-6" : ""}`}
                  onClick={() => steps.deploy.data && toggleStep(4)}
                >
                  <StepIndicator step={4} label="Apps Script デプロイ" status={steps.deploy.status} detail={steps.deploy.detail} />
                  {steps.deploy.data && (
                    <span className="text-gray-400 text-sm">
                      {expandedSteps.has(4) ? "\u25BC" : "\u25B6"}
                    </span>
                  )}
                </div>
                {expandedSteps.has(4) && steps.deploy.data && (
                  <div className="ml-[52px] pb-4">
                    <DeployStepDetail data={steps.deploy.data} />
                  </div>
                )}
              </div>
            </div>
          </div>

          {/* Summary card after completion */}
          {allDone && !error && (
            <div className={`mt-6 rounded-xl border p-6 shadow-sm ${
              steps.deploy.status === "error" || steps.convert.status === "error"
                ? "border-yellow-200 bg-yellow-50"
                : "border-green-200 bg-green-50"
            }`}>
              <h3 className={`font-semibold ${
                steps.deploy.status === "error" || steps.convert.status === "error"
                  ? "text-yellow-800"
                  : "text-green-800"
              }`}>
                {steps.deploy.status === "error" || steps.convert.status === "error"
                  ? "移行完了（一部エラーあり）"
                  : "移行完了"
                }
              </h3>
              <p className="mt-1 text-sm text-gray-700">
                {steps.upload.data?.successCount || 0} 件アップロード、
                {steps.convert.data?.success || 0} 件変換、
                経過時間 {formatTime(elapsed)}
              </p>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
