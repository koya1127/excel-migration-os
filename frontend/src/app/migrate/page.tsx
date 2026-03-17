"use client";

import { useState, useCallback, useRef } from "react";
import {
  migrateFiles,
  checkSubscription,
  type MigrateReport,
  type UploadReport,
  type ExtractReport,
  type ConvertReport,
  type DeployReport,
  type TrackMode,
  type PythonConvertReport,
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
      <div className={`flex h-10 w-10 items-center justify-center rounded-full border-2 ${colors[status]}`}>
        {icons[status]}
      </div>
      <div className="flex flex-col">
        <span className={`text-sm font-medium ${status === "pending" ? "text-gray-400" : status === "error" ? "text-red-700" : "text-gray-900"}`}>
          {label}
        </span>
        {detail && <span className="text-xs text-gray-500">{detail}</span>}
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
          }`}>{f.status}</span>
          <span className="text-gray-700">{f.fileName}</span>
          {f.webViewLink && <a href={f.webViewLink} target="_blank" rel="noopener noreferrer" className="text-blue-600 hover:underline">Drive で開く</a>}
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

function DeployStepDetail({ deploys }: { deploys: DeployReport[] }) {
  return (
    <div className="space-y-3">
      {deploys.map((d, i) => (
        <div key={i} className="space-y-1">
          {d.webViewLink && (
            <a href={d.webViewLink} target="_blank" rel="noopener noreferrer" className="text-sm text-blue-600 hover:underline font-medium">
              スプレッドシートを開く
            </a>
          )}
          {d.filesDeployed && d.filesDeployed.length > 0 && (
            <div className="flex flex-wrap gap-2">
              {d.filesDeployed.map((f, j) => (
                <span key={j} className="inline-flex items-center rounded-md bg-blue-50 border border-blue-200 px-2 py-0.5 text-xs text-blue-700">
                  {f}
                </span>
              ))}
            </div>
          )}
          {d.error && <p className="text-xs text-red-600">{d.error}</p>}
        </div>
      ))}
    </div>
  );
}

interface StepState {
  upload: { status: StepStatus; detail: string; data: UploadReport | null };
  extract: { status: StepStatus; detail: string; data: ExtractReport | null };
  convert: { status: StepStatus; detail: string; data: ConvertReport | null };
  deploy: { status: StepStatus; detail: string; deploys: DeployReport[] };
}

const initialSteps: StepState = {
  upload: { status: "pending", detail: "", data: null },
  extract: { status: "pending", detail: "", data: null },
  convert: { status: "pending", detail: "", data: null },
  deploy: { status: "pending", detail: "", deploys: [] },
};

export default function MigratePage() {
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [folderId, setFolderId] = useState("");
  const [trackMode, setTrackMode] = useState<TrackMode>("auto");
  const [pythonDownloadUrl, setPythonDownloadUrl] = useState<string | null>(null);
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
    setSteps(initialSteps);
    setExpandedSteps(new Set());
    setElapsed(0);

    const start = Date.now();
    timerRef.current = setInterval(() => {
      setElapsed(Math.floor((Date.now() - start) / 1000));
    }, 1000);

    try {
      const { hasSubscription } = await checkSubscription();
      if (!hasSubscription) {
        setError("移行機能を使用するには従量課金プランへの登録が必要です。設定ページから登録してください。");
        return;
      }

      // Show all steps as active/pending
      setSteps({
        upload: { status: "active", detail: "処理中...", data: null },
        extract: { status: "pending", detail: "", data: null },
        convert: { status: "pending", detail: "", data: null },
        deploy: { status: "pending", detail: "", deploys: [] },
      });

      // Call backend migrate API (does upload → extract → convert → deploy in one request)
      const report: MigrateReport = await migrateFiles(selectedFiles, true, folderId || undefined, trackMode);

      // Store Python download URL if available
      if (report.pythonPackageUrl) {
        const apiBase = process.env.NEXT_PUBLIC_API_URL || "";
        setPythonDownloadUrl(`${apiBase}${report.pythonPackageUrl}`);
      }

      // Update upload step
      const upload = report.upload;
      setSteps((prev) => ({
        ...prev,
        upload: {
          status: upload && upload.failureCount > 0 ? "error" : "success",
          detail: upload ? `${upload.successCount}/${upload.fileCount} 件成功` : "スキップ",
          data: upload || null,
        },
      }));

      // Update extract step
      const extract = report.extract;
      setSteps((prev) => ({
        ...prev,
        extract: {
          status: "success",
          detail: `${extract?.moduleCount || 0} モジュール、${extract?.controls?.length || 0} コントロール`,
          data: extract || null,
        },
      }));

      // Update convert step
      const convert = report.convert;
      if (convert) {
        setSteps((prev) => ({
          ...prev,
          convert: {
            status: convert.failed > 0 ? "error" : "success",
            detail: `${convert.success}/${convert.total} 成功（${(convert.totalInputTokens + convert.totalOutputTokens).toLocaleString()} トークン）`,
            data: convert,
          },
        }));
      }

      // Update deploy step
      const deploys = report.deploys || [];
      const deploySuccess = deploys.filter((d) => d.status === "success").length;
      const deployTotal = deploys.length;
      const hasDeployError = deploys.some((d) => d.status !== "success");

      setSteps((prev) => ({
        ...prev,
        deploy: {
          status: deployTotal === 0 ? "error" : hasDeployError ? "error" : "success",
          detail: deployTotal === 0
            ? "デプロイ可能なファイルがありません"
            : `${deploySuccess} スプレッドシートにデプロイ完了`,
          deploys,
        },
      }));

      setExpandedSteps(new Set([1, 2, 3, 4]));

    } catch (err) {
      setError(err instanceof Error ? err.message : "移行に失敗しました");
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

  const allDone = !loading && started;

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
          <input ref={fileInputRef} type="file" className="hidden" accept=".xlsm" multiple onChange={(e) => handleFiles(e.target.files)} />
          {/* @ts-expect-error webkitdirectory is not in types */}
          <input ref={folderInputRef} type="file" className="hidden" webkitdirectory="" multiple onChange={(e) => handleFiles(e.target.files)} />
          <svg className="h-10 w-10 text-gray-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M12 16V4m0 0l-4 4m4-4l4 4M4 20h16" />
          </svg>
          <p className="mt-3 text-sm text-gray-600">ここにファイル・フォルダをドロップ</p>
          <div className="mt-3 flex gap-3">
            <button type="button" className="rounded-md bg-white px-3 py-1.5 text-xs font-medium text-gray-700 border border-gray-300 hover:bg-gray-50" onClick={(e) => { e.stopPropagation(); fileInputRef.current?.click(); }}>ファイルを選択</button>
            <button type="button" className="rounded-md bg-white px-3 py-1.5 text-xs font-medium text-gray-700 border border-gray-300 hover:bg-gray-50" onClick={(e) => { e.stopPropagation(); folderInputRef.current?.click(); }}>フォルダを選択</button>
          </div>
          <p className="mt-2 text-xs text-gray-400">.xlsm のみ対応（VBA マクロを含むファイル）</p>
        </div>

        {selectedFiles.length > 0 && (
          <div className="mt-4">
            <div className="flex items-center justify-between">
              <p className="text-sm font-medium text-gray-700">{selectedFiles.length} 件のファイルを選択中</p>
              <button className="text-xs text-red-500 hover:text-red-700" onClick={() => setSelectedFiles([])}>すべてクリア</button>
            </div>
            <div className="mt-2 flex flex-wrap gap-2">
              {selectedFiles.map((f, i) => (
                <span key={i} className="inline-flex items-center gap-1 rounded-md bg-gray-100 px-2.5 py-1 text-xs text-gray-700">
                  {f.name}
                  <button className="ml-1 text-gray-400 hover:text-gray-600" onClick={() => setSelectedFiles((prev) => prev.filter((_, j) => j !== i))}>x</button>
                </span>
              ))}
            </div>
          </div>
        )}

        {/* Options */}
        <div className="mt-6">
          <label className="block text-sm font-medium text-gray-700 mb-1">Google Drive フォルダ ID（オプション）</label>
          <input type="text" value={folderId} onChange={(e) => setFolderId(e.target.value)} className="w-full max-w-md rounded-lg border border-gray-300 px-3 py-2 text-sm focus:border-blue-500 focus:ring-1 focus:ring-blue-500 outline-none" placeholder="空欄の場合はマイドライブのルートに作成" />
        </div>

        {/* Track Mode */}
        <div className="mt-4">
          <label className="block text-sm font-medium text-gray-700 mb-2">移行モード</label>
          <div className="flex flex-wrap gap-3">
            {([
              { value: "auto", label: "自動判定", desc: "マクロの内容に応じて自動で振り分け" },
              { value: "sheets_only", label: "スプシのみ", desc: "すべてGASに変換" },
              { value: "local_only", label: "ローカルのみ", desc: "すべてPythonに変換" },
              { value: "both", label: "両方出力", desc: "GAS + Python両方を生成" },
            ] as { value: TrackMode; label: string; desc: string }[]).map(({ value, label, desc }) => (
              <label key={value} className={`flex items-start gap-2 rounded-lg border px-3 py-2 cursor-pointer transition-colors ${trackMode === value ? "border-blue-500 bg-blue-50" : "border-gray-200 hover:border-gray-300"}`}>
                <input type="radio" name="trackMode" value={value} checked={trackMode === value} onChange={() => setTrackMode(value)} className="mt-0.5" />
                <div>
                  <span className="text-sm font-medium text-gray-900">{label}</span>
                  <p className="text-xs text-gray-500">{desc}</p>
                </div>
              </label>
            ))}
          </div>
        </div>

        <div className="mt-6 flex items-center gap-4">
          <button onClick={handleMigrate} disabled={selectedFiles.length === 0 || loading} className="rounded-lg bg-gradient-to-r from-blue-600 to-indigo-600 px-8 py-3 text-sm font-semibold text-white shadow-md hover:from-blue-500 hover:to-indigo-500 disabled:opacity-50 disabled:cursor-not-allowed transition-all">
            {loading ? <span className="flex items-center gap-2"><Spinner />移行中...</span> : "移行開始"}
          </button>
          {loading && <span className="text-sm text-gray-500">経過時間: {formatTime(elapsed)}</span>}
        </div>
      </div>

      {/* Error */}
      {error && (
        <div className="mt-6 rounded-lg border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">{error}</div>
      )}

      {/* Progress / Results */}
      {started && (
        <div className="mt-8">
          <div className="rounded-xl border border-gray-200 bg-white shadow-sm overflow-hidden">
            <div className="border-b border-gray-200 px-6 py-4 flex items-center justify-between">
              <h2 className="text-lg font-semibold text-gray-900">移行プロセス</h2>
              {allDone && <span className="text-sm text-gray-500">完了 — {formatTime(elapsed)}</span>}
            </div>
            <div className="p-6 space-y-0">
              {/* Step 1: Upload */}
              <div>
                <div className={`flex items-center justify-between py-4 ${steps.upload.data ? "cursor-pointer hover:bg-gray-50 -mx-6 px-6" : ""}`} onClick={() => steps.upload.data && toggleStep(1)}>
                  <StepIndicator step={1} label="アップロード" status={steps.upload.status} detail={steps.upload.detail} />
                  {steps.upload.data && <span className="text-gray-400 text-sm">{expandedSteps.has(1) ? "\u25BC" : "\u25B6"}</span>}
                </div>
                {expandedSteps.has(1) && steps.upload.data && (
                  <div className="ml-[52px] pb-4"><UploadStepDetail data={steps.upload.data} /></div>
                )}
                <div className="ml-5 h-6 border-l-2 border-gray-200"></div>
              </div>

              {/* Step 2: Extract */}
              <div>
                <div className={`flex items-center justify-between py-4 ${steps.extract.data ? "cursor-pointer hover:bg-gray-50 -mx-6 px-6" : ""}`} onClick={() => steps.extract.data && toggleStep(2)}>
                  <StepIndicator step={2} label="VBA 抽出" status={steps.extract.status} detail={steps.extract.detail} />
                  {steps.extract.data && <span className="text-gray-400 text-sm">{expandedSteps.has(2) ? "\u25BC" : "\u25B6"}</span>}
                </div>
                {expandedSteps.has(2) && steps.extract.data && (
                  <div className="ml-[52px] pb-4"><ExtractStepDetail data={steps.extract.data} /></div>
                )}
                <div className="ml-5 h-6 border-l-2 border-gray-200"></div>
              </div>

              {/* Step 3: Convert */}
              <div>
                <div className={`flex items-center justify-between py-4 ${steps.convert.data ? "cursor-pointer hover:bg-gray-50 -mx-6 px-6" : ""}`} onClick={() => steps.convert.data && toggleStep(3)}>
                  <StepIndicator step={3} label="VBA → GAS 変換" status={steps.convert.status} detail={steps.convert.detail} />
                  {steps.convert.data && <span className="text-gray-400 text-sm">{expandedSteps.has(3) ? "\u25BC" : "\u25B6"}</span>}
                </div>
                {expandedSteps.has(3) && steps.convert.data && (
                  <div className="ml-[52px] pb-4"><ConvertStepDetail data={steps.convert.data} /></div>
                )}
                <div className="ml-5 h-6 border-l-2 border-gray-200"></div>
              </div>

              {/* Step 4: Deploy */}
              <div>
                <div className={`flex items-center justify-between py-4 ${steps.deploy.deploys.length > 0 ? "cursor-pointer hover:bg-gray-50 -mx-6 px-6" : ""}`} onClick={() => steps.deploy.deploys.length > 0 && toggleStep(4)}>
                  <StepIndicator step={4} label="Apps Script デプロイ" status={steps.deploy.status} detail={steps.deploy.detail} />
                  {steps.deploy.deploys.length > 0 && <span className="text-gray-400 text-sm">{expandedSteps.has(4) ? "\u25BC" : "\u25B6"}</span>}
                </div>
                {expandedSteps.has(4) && steps.deploy.deploys.length > 0 && (
                  <div className="ml-[52px] pb-4"><DeployStepDetail deploys={steps.deploy.deploys} /></div>
                )}
              </div>
            </div>
          </div>

          {/* Summary */}
          {allDone && !error && (
            <div className={`mt-6 rounded-xl border p-6 shadow-sm ${
              steps.deploy.status === "error" || steps.convert.status === "error"
                ? "border-yellow-200 bg-yellow-50" : "border-green-200 bg-green-50"
            }`}>
              <h3 className={`font-semibold ${
                steps.deploy.status === "error" || steps.convert.status === "error" ? "text-yellow-800" : "text-green-800"
              }`}>
                {steps.deploy.status === "error" || steps.convert.status === "error" ? "移行完了（一部エラーあり）" : "移行完了"}
              </h3>
              <p className="mt-1 text-sm text-gray-700">
                {steps.upload.data?.successCount || 0} 件アップロード、{steps.convert.data?.success || 0} 件変換、経過時間 {formatTime(elapsed)}
              </p>
              {steps.deploy.deploys.filter((d) => d.webViewLink).map((d, i) => (
                <a key={i} href={d.webViewLink} target="_blank" rel="noopener noreferrer" className="mt-2 inline-block text-sm text-blue-600 hover:underline font-medium">
                  移行先スプレッドシートを開く
                </a>
              ))}
              {pythonDownloadUrl && (
                <div className="mt-3">
                  <a href={pythonDownloadUrl} download className="inline-flex items-center gap-2 rounded-lg bg-gray-800 px-4 py-2 text-sm font-medium text-white hover:bg-gray-700 transition-colors">
                    <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 10v6m0 0l-3-3m3 3l3-3M3 17v3a2 2 0 002 2h14a2 2 0 002-2v-3" /></svg>
                    ローカル版をダウンロード（Python）
                  </a>
                  <p className="mt-1 text-xs text-gray-500">スプレッドシートでは動かせなかった機能が含まれています</p>
                </div>
              )}
            </div>
          )}
        </div>
      )}
    </div>
  );
}
