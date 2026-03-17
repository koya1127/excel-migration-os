namespace ExcelMigrationApi.Models;

public class TrackDecision
{
    public string ModuleName { get; set; } = string.Empty;
    public string SourceFile { get; set; } = string.Empty;
    public int Track { get; set; } // 1 = GAS, 2 = Python
    public string Reason { get; set; } = string.Empty;
    public List<string> DetectedPatterns { get; set; } = new();
}

public class TrackResult
{
    public List<VbaModule> Track1Modules { get; set; } = new(); // GAS-compatible
    public List<VbaModule> Track2Modules { get; set; } = new(); // Requires local (Python)
    public List<TrackDecision> Decisions { get; set; } = new();
}

public class PythonConvertRequest
{
    public string VbaCode { get; set; } = string.Empty;
    public string ModuleName { get; set; } = string.Empty;
    public string ModuleType { get; set; } = string.Empty;
    public string SourceFile { get; set; } = string.Empty;
    public string SheetName { get; set; } = string.Empty;
    public string? SpreadsheetId { get; set; }
}

public class PythonConvertResult
{
    public string ModuleName { get; set; } = string.Empty;
    public string PythonCode { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
    public string Error { get; set; } = string.Empty;
    public string SourceFile { get; set; } = string.Empty;
    public int InputTokens { get; set; }
    public int OutputTokens { get; set; }
}

public class PythonConvertReport
{
    public string GeneratedUtc { get; set; } = string.Empty;
    public int Total { get; set; }
    public int Success { get; set; }
    public int Failed { get; set; }
    public int TotalInputTokens { get; set; }
    public int TotalOutputTokens { get; set; }
    public List<PythonConvertResult> Results { get; set; } = new();
}

public class PythonPackage
{
    public string FileName { get; set; } = string.Empty;
    public byte[] ZipData { get; set; } = Array.Empty<byte>();
    public List<string> ModuleNames { get; set; } = new();
    public string ReadmeContent { get; set; } = string.Empty;
}
