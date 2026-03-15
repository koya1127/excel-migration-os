namespace ExcelMigrationApi.Models;

public class UploadRequest
{
    public bool ConvertToSheets { get; set; } = true;
    public string? FolderId { get; set; }
}

public class UploadResult
{
    public string FileName { get; set; } = string.Empty;
    public string DriveFileId { get; set; } = string.Empty;
    public string WebViewLink { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
    public string Error { get; set; } = string.Empty;
}

public class UploadReport
{
    public string GeneratedUtc { get; set; } = string.Empty;
    public int FileCount { get; set; }
    public int SuccessCount { get; set; }
    public int FailureCount { get; set; }
    public bool ConvertedToSheets { get; set; }
    public List<UploadResult> Files { get; set; } = new();
}

public class DeployRequest
{
    public string SpreadsheetId { get; set; } = string.Empty;
    public List<GasFile> GasFiles { get; set; } = new();
}

public class GasFile
{
    public string Name { get; set; } = string.Empty;
    public string Source { get; set; } = string.Empty;
    public string Type { get; set; } = "SERVER_JS"; // SERVER_JS or HTML
}

public class DeployReport
{
    public string GeneratedUtc { get; set; } = string.Empty;
    public string SpreadsheetId { get; set; } = string.Empty;
    public string WebViewLink { get; set; } = string.Empty;
    public string ScriptId { get; set; } = string.Empty;
    public int FileCount { get; set; }
    public List<string> FilesDeployed { get; set; } = new();
    public string Status { get; set; } = string.Empty;
    public string Error { get; set; } = string.Empty;
}

public class UsageSheetInfo
{
    public string OriginalFileName { get; set; } = string.Empty;
    public List<string> MenuItems { get; set; } = new();
    public List<string> Limitations { get; set; } = new();
}

public class MigrateRequest
{
    public bool ConvertToSheets { get; set; } = true;
    public string? FolderId { get; set; }
}

public class MigrateReport
{
    public string GeneratedUtc { get; set; } = string.Empty;
    public UploadReport? Upload { get; set; }
    public ExtractReport? Extract { get; set; }
    public ConvertReport? Convert { get; set; }
    public List<DeployReport> Deploys { get; set; } = new();
}
