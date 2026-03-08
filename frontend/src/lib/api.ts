const API_BASE = process.env.NEXT_PUBLIC_API_URL || "http://localhost:5269";

async function getAuthHeaders(): Promise<Record<string, string>> {
  if (typeof window === "undefined") return {};
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const clerk = (window as any).Clerk;
  if (!clerk?.session) return {};
  const token = await clerk.session.getToken();
  return token ? { Authorization: `Bearer ${token}` } : {};
}

async function getGoogleToken(): Promise<string | null> {
  try {
    const res = await fetch("/api/auth/google/token");
    if (!res.ok) {
      const err = await res.json().catch(() => null);
      console.error("Google token取得失敗:", res.status, err);
      return null;
    }
    const data = await res.json();
    return data.accessToken || null;
  } catch (e) {
    console.error("Google token取得エラー:", e);
    return null;
  }
}

async function getAuthHeadersWithGoogle(): Promise<Record<string, string>> {
  const headers = await getAuthHeaders();
  const googleToken = await getGoogleToken();
  if (googleToken) {
    headers["X-Google-Token"] = googleToken;
  }
  return headers;
}

export interface FileReport {
  path: string;
  extension: string;
  sizeBytes: number;
  modifiedUtc: string;
  hasMacro: boolean;
  vbaModuleCount: number;
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
  files.forEach(f => formData.append("files", f));
  formData.append("groupBy", groupBy);

  const res = await fetch(`${API_BASE}/api/scan`, {
    method: "POST",
    headers: { ...authHeaders },
    body: formData,
  });
  if (!res.ok) throw new Error(`Scan failed: ${res.statusText}`);
  return res.json();
}

// Extract types
export interface VbaModule {
  sourceFile: string;
  moduleName: string;
  moduleType: string;
  codeLines: number;
  code: string;
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
  buttonContext?: FormControl[];
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
  deploy: DeployReport;
}

// API functions
export async function extractFiles(files: File[]): Promise<ExtractReport> {
  const authHeaders = await getAuthHeaders();
  const formData = new FormData();
  files.forEach(f => formData.append("files", f));
  const res = await fetch(`${API_BASE}/api/extract`, { method: "POST", headers: { ...authHeaders }, body: formData });
  if (!res.ok) throw new Error(`Extract failed: ${res.statusText}`);
  return res.json();
}

export async function convertBatch(extractReport: ExtractReport): Promise<ConvertReport> {
  const authHeaders = await getAuthHeaders();
  const res = await fetch(`${API_BASE}/api/convert/batch`, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...authHeaders },
    body: JSON.stringify(extractReport),
  });
  if (!res.ok) throw new Error(`Convert failed: ${res.statusText}`);
  const report: ConvertReport = await res.json();
  await reportUsage(report.totalInputTokens, report.totalOutputTokens);
  return report;
}

export async function convertSingle(request: ConvertRequest): Promise<ConvertReport> {
  const authHeaders = await getAuthHeaders();
  const res = await fetch(`${API_BASE}/api/convert`, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...authHeaders },
    body: JSON.stringify([request]),
  });
  if (!res.ok) throw new Error(`Convert failed: ${res.statusText}`);
  const report: ConvertReport = await res.json();
  await reportUsage(report.totalInputTokens, report.totalOutputTokens);
  return report;
}

export async function checkSubscription(): Promise<{ hasSubscription: boolean; subscriptionStatus?: string }> {
  try {
    const res = await fetch("/api/stripe/subscription");
    if (!res.ok) return { hasSubscription: false };
    return await res.json();
  } catch {
    return { hasSubscription: false };
  }
}

async function reportUsage(inputTokens: number, outputTokens: number): Promise<void> {
  if (inputTokens + outputTokens <= 0) return;
  const res = await fetch("/api/stripe/report-usage", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ inputTokens, outputTokens }),
  });
  if (!res.ok) {
    const data = await res.json().catch(() => null);
    console.error("課金報告に失敗しました:", data?.error || res.statusText);
  }
}

export async function uploadFiles(files: File[], convertToSheets: boolean = true, folderId?: string): Promise<UploadReport> {
  const authHeaders = await getAuthHeadersWithGoogle();
  const formData = new FormData();
  files.forEach(f => formData.append("files", f));
  formData.append("convertToSheets", String(convertToSheets));
  if (folderId) formData.append("folderId", folderId);
  const res = await fetch(`${API_BASE}/api/upload`, { method: "POST", headers: { ...authHeaders }, body: formData });
  if (!res.ok) {
    const err = await res.json().catch(() => null);
    throw new Error(err?.error || `Upload failed: ${res.statusText}`);
  }
  return res.json();
}

export async function deployGas(spreadsheetId: string, gasFiles: GasFile[]): Promise<DeployReport> {
  const authHeaders = await getAuthHeadersWithGoogle();
  const res = await fetch(`${API_BASE}/api/deploy`, {
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
  const authHeaders = await getAuthHeadersWithGoogle();
  const formData = new FormData();
  files.forEach(f => formData.append("files", f));
  formData.append("convertToSheets", String(convertToSheets));
  if (folderId) formData.append("folderId", folderId);
  const res = await fetch(`${API_BASE}/api/migrate`, { method: "POST", headers: { ...authHeaders }, body: formData });
  if (!res.ok) {
    const err = await res.json().catch(() => null);
    throw new Error(err?.error || `Migrate failed: ${res.statusText}`);
  }
  const report: MigrateReport = await res.json();
  // Report usage for the conversion step
  if (report.convert) {
    await reportUsage(report.convert.totalInputTokens, report.convert.totalOutputTokens);
  }
  return report;
}
