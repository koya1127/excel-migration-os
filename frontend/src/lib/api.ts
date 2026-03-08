const API_BASE = process.env.NEXT_PUBLIC_API_URL || "http://localhost:5269";

async function getAuthHeaders(): Promise<Record<string, string>> {
  if (typeof window === "undefined") return {};
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const clerk = (window as any).Clerk;
  if (!clerk?.session) return {};
  const token = await clerk.session.getToken();
  return token ? { Authorization: `Bearer ${token}` } : {};
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
}

export interface ConvertReport {
  generatedUtc: string;
  total: number;
  success: number;
  failed: number;
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
  return res.json();
}

export async function convertSingle(request: ConvertRequest): Promise<ConvertReport> {
  const authHeaders = await getAuthHeaders();
  const res = await fetch(`${API_BASE}/api/convert`, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...authHeaders },
    body: JSON.stringify([request]),
  });
  if (!res.ok) throw new Error(`Convert failed: ${res.statusText}`);
  return res.json();
}

export async function uploadFiles(files: File[], convertToSheets: boolean = true, folderId?: string): Promise<UploadReport> {
  const authHeaders = await getAuthHeaders();
  const formData = new FormData();
  files.forEach(f => formData.append("files", f));
  formData.append("convertToSheets", String(convertToSheets));
  if (folderId) formData.append("folderId", folderId);
  const res = await fetch(`${API_BASE}/api/upload`, { method: "POST", headers: { ...authHeaders }, body: formData });
  if (!res.ok) throw new Error(`Upload failed: ${res.statusText}`);
  return res.json();
}

export async function deployGas(spreadsheetId: string, gasFiles: GasFile[]): Promise<DeployReport> {
  const authHeaders = await getAuthHeaders();
  const res = await fetch(`${API_BASE}/api/deploy`, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...authHeaders },
    body: JSON.stringify({ spreadsheetId, gasFiles }),
  });
  if (!res.ok) throw new Error(`Deploy failed: ${res.statusText}`);
  return res.json();
}

export async function migrateFiles(files: File[], convertToSheets: boolean = true, folderId?: string): Promise<MigrateReport> {
  const authHeaders = await getAuthHeaders();
  const formData = new FormData();
  files.forEach(f => formData.append("files", f));
  formData.append("convertToSheets", String(convertToSheets));
  if (folderId) formData.append("folderId", folderId);
  const res = await fetch(`${API_BASE}/api/migrate`, { method: "POST", headers: { ...authHeaders }, body: formData });
  if (!res.ok) throw new Error(`Migrate failed: ${res.statusText}`);
  return res.json();
}
