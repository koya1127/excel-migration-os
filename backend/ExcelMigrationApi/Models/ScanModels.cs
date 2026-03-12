namespace ExcelMigrationApi.Models;

public class FileReport
{
    public string Path { get; set; } = string.Empty;
    public string Extension { get; set; } = string.Empty;
    public long SizeBytes { get; set; }
    public string ModifiedUtc { get; set; } = string.Empty;
    public bool HasMacro { get; set; }
    public int? VbaModuleCount { get; set; }
    public int? VbaTotalCodeLength { get; set; }
    public bool AnalysisFailed { get; set; }
    public int SheetCount { get; set; }
    public int FormulaCount { get; set; }
    public int VolatileFormulaCount { get; set; }
    public int NamedRangeCount { get; set; }
    public int ExternalLinkCount { get; set; }
    public int IncompatibleFunctionCount { get; set; }
    public int RiskScore { get; set; }
    public List<string> Notes { get; set; } = new();
}

public class GroupSummary
{
    public string GroupName { get; set; } = string.Empty;
    public int FileCount { get; set; }
    public long TotalSizeBytes { get; set; }
    public int MacroFileCount { get; set; }
    public int TotalVbaModules { get; set; }
    public double AvgRiskScore { get; set; }
    public int MaxRiskScore { get; set; }
    public int TotalFormulas { get; set; }
    public int TotalIncompatibleFunctions { get; set; }
    public string MigrationDifficulty { get; set; } = string.Empty;
    public List<int> FileIndices { get; set; } = new();
}

public class ScanReport
{
    public string GeneratedUtc { get; set; } = string.Empty;
    public string InputRoot { get; set; } = string.Empty;
    public int FileCount { get; set; }
    public List<FileReport> Files { get; set; } = new();
    public string GroupBy { get; set; } = "none";
    public List<GroupSummary> Groups { get; set; } = new();
}

public class ScanRequest
{
    public string GroupBy { get; set; } = "none";
}
