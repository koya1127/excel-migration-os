// All endpoints call the backend directly to avoid Vercel's 4.5MB body limit
// and 30-second serverless function timeout (convert/batch can take minutes).
const API_BASE = process.env.NEXT_PUBLIC_API_URL || "";

function requireApiBase(): string {
  if (!API_BASE) {
    if (typeof window !== "undefined" && window.location.hostname !== "localhost") {
      throw new Error("NEXT_PUBLIC_API_URL が設定されていません。バックエンドAPIに接続できません。");
    }
  }
  return API_BASE;
}

async function getAuthHeaders(): Promise<Record<string, string>> {
  if (typeof window === "undefined") return {};
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const clerk = (window as any).Clerk;
  if (!clerk?.session) return {};
  const token = await clerk.session.getToken();
  return token ? { Authorization: `Bearer ${token}` } : {};
}

// Google token is now fetched server-side by the backend from Clerk privateMetadata.
// No need to pass X-Google-Token from the browser (eliminates XSS token theft risk).

export interface FileReport {
  path: string;
  extension: string;
  sizeBytes: number;
  modifiedUtc: string;
  hasMacro: boolean;
  vbaModuleCount: number | null;
  vbaTotalCodeLength: number | null;
  analysisFailed: boolean;
  sheetCount: number;
  formulaCount: number;
  volatileFormulaCount: number;
  namedRangeCount: number;
  externalLinkCount: number;
  incompatibleFunctionCount: number;
  riskScore: number;
  notes: string | string[];
}

export interface GroupSummary {
  groupName: string;
  fileCount: number;
  totalSizeBytes: number;
  macroFileCount: number;
  totalVbaModules: number;
  avgRiskScore: number;
  maxRiskScore: number;
  totalFormulas: number;
  totalIncompatibleFunctions: number;
  migrationDifficulty: string;
  fileIndices: number[];
}

export interface ScanReport {
  generatedUtc: string;
  inputRoot: string;
  fileCount: number;
  files: FileReport[];
  groupBy: string;
  groups: GroupSummary[];
}

export async function scanFiles(files: File[], groupBy: string = "subfolder"): Promise<ScanReport> {
  const authHeaders = await getAuthHeaders();
  const formData = new FormData();
  // Preserve folder structure: use webkitRelativePath (folder picker) or name (D&D with encoded path)
  files.forEach(f => formData.append("files", f, f.webkitRelativePath || f.name));
  formData.append("groupBy", groupBy);

  const res = await fetch(`${requireApiBase()}/api/scan`, {
    method: "POST",
    headers: { ...authHeaders },
    body: formData,
  });
  if (!res.ok) {
    const body = await res.json().catch(() => null);
    throw new Error(body?.error || `スキャンに失敗しました（${res.status}）`);
  }
  return res.json();
}

// Extract types
export interface VbaEvent {
  vbaEventName: string;
  sheetName: string;
  gasTriggerType: string;
  gasNotes: string;
}

export interface VbaModule {
  sourceFile: string;
  moduleName: string;
  moduleType: string;
  codeLines: number;
  code: string;
  sheetName: string;
  detectedEvents: VbaEvent[];
}

export interface FormControl {
  sourceFile: string;
  sheetName: string;
  controlName: string;
  controlType: string;
  label: string;
  macro: string;
}

export interface ExtractReport {
  generatedUtc: string;
  fileCount: number;
  moduleCount: number;
  modules: VbaModule[];
  controls: FormControl[];
}

// Convert types
export interface ConvertRequest {
  vbaCode: string;
  moduleName: string;
  moduleType: string;
  sheetName?: string;
  buttonContext?: FormControl[];
  detectedEvents?: VbaEvent[];
}

export interface ConvertResult {
  moduleName: string;
  gasCode: string;
  status: string;
  error: string;
  sourceFile: string;
  inputTokens: number;
  outputTokens: number;
}

export interface ConvertReport {
  generatedUtc: string;
  total: number;
  success: number;
  failed: number;
  totalInputTokens: number;
  totalOutputTokens: number;
  results: ConvertResult[];
}

// Upload types
export interface UploadResult {
  fileName: string;
  driveFileId: string;
  webViewLink: string;
  status: string;
  error: string;
}

export interface UploadReport {
  generatedUtc: string;
  fileCount: number;
  successCount: number;
  failureCount: number;
  convertedToSheets: boolean;
  files: UploadResult[];
}

// Deploy types
export interface GasFile {
  name: string;
  source: string;
  type: string;
}

export interface DeployReport {
  generatedUtc: string;
  spreadsheetId: string;
  scriptId: string;
  fileCount: number;
  filesDeployed: string[];
  status: string;
  error: string;
}

// Migrate types
export interface MigrateReport {
  generatedUtc: string;
  upload: UploadReport;
  extract: ExtractReport;
  convert: ConvertReport;
  deploys: DeployReport[];
}

// API functions
export async function extractFiles(files: File[]): Promise<ExtractReport> {
  const authHeaders = await getAuthHeaders();
  const formData = new FormData();
  files.forEach(f => formData.append("files", f));
  const res = await fetch(`${requireApiBase()}/api/extract`, { method: "POST", headers: { ...authHeaders }, body: formData });
  if (!res.ok) throw new Error(`Extract failed: ${res.statusText}`);
  return res.json();
}

export async function convertBatch(extractReport: ExtractReport): Promise<ConvertReport> {
  const authHeaders = await getAuthHeaders();
  const res = await fetch(`${requireApiBase()}/api/convert/batch`, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...authHeaders },
    body: JSON.stringify(extractReport),
  });
  if (!res.ok) {
    const err = await res.json().catch(() => null);
    throw new Error(err?.error || `Convert failed: ${res.statusText}`);
  }
  return res.json();
}

export async function convertSingle(request: ConvertRequest): Promise<ConvertReport> {
  const authHeaders = await getAuthHeaders();
  const res = await fetch(`${requireApiBase()}/api/convert`, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...authHeaders },
    body: JSON.stringify([request]),
  });
  if (!res.ok) {
    const err = await res.json().catch(() => null);
    throw new Error(err?.error || `Convert failed: ${res.statusText}`);
  }
  return res.json();
}

export async function checkSubscription(): Promise<{ hasSubscription: boolean; subscriptionStatus?: string }> {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const clerk = (window as any).Clerk;
    if (!clerk?.user) return { hasSubscription: false };
    const meta = clerk.user.publicMetadata as Record<string, unknown> | undefined;
    const status = meta?.subscriptionStatus as string | undefined;
    return {
      hasSubscription: status === "active",
      subscriptionStatus: status,
    };
  } catch {
    return { hasSubscription: false };
  }
}

export async function uploadFiles(files: File[], convertToSheets: boolean = true, folderId?: string): Promise<UploadReport> {
  const authHeaders = await getAuthHeaders();
  const formData = new FormData();
  files.forEach(f => formData.append("files", f));
  formData.append("convertToSheets", String(convertToSheets));
  if (folderId) formData.append("folderId", folderId);
  const res = await fetch(`${requireApiBase()}/api/upload`, { method: "POST", headers: { ...authHeaders }, body: formData });
  if (!res.ok) {
    const err = await res.json().catch(() => null);
    throw new Error(err?.error || `Upload failed: ${res.statusText}`);
  }
  return res.json();
}

export async function deployGas(spreadsheetId: string, gasFiles: GasFile[]): Promise<DeployReport> {
  const authHeaders = await getAuthHeaders();
  const res = await fetch(`${requireApiBase()}/api/deploy`, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...authHeaders },
    body: JSON.stringify({ spreadsheetId, gasFiles }),
  });
  if (!res.ok) {
    const err = await res.json().catch(() => null);
    throw new Error(err?.error || `Deploy failed: ${res.statusText}`);
  }
  return res.json();
}

export async function migrateFiles(files: File[], convertToSheets: boolean = true, folderId?: string): Promise<MigrateReport> {
  const authHeaders = await getAuthHeaders();
  const formData = new FormData();
  files.forEach(f => formData.append("files", f));
  formData.append("convertToSheets", String(convertToSheets));
  if (folderId) formData.append("folderId", folderId);
  const res = await fetch(`${requireApiBase()}/api/migrate`, { method: "POST", headers: { ...authHeaders }, body: formData });
  if (!res.ok) {
    const err = await res.json().catch(() => null);
    throw new Error(err?.error || `Migrate failed: ${res.statusText}`);
  }
  // Usage is now reported server-side in MigrateController
  return res.json();
}
