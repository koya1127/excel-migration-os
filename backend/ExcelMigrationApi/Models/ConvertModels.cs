namespace ExcelMigrationApi.Models;

public class ConvertRequest
{
    public string VbaCode { get; set; } = string.Empty;
    public string ModuleName { get; set; } = string.Empty;
    public string ModuleType { get; set; } = string.Empty;
    public List<FormControl>? ButtonContext { get; set; }
}

public class ConvertResult
{
    public string ModuleName { get; set; } = string.Empty;
    public string GasCode { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty; // "success" or "error"
    public string Error { get; set; } = string.Empty;
    public int InputTokens { get; set; }
    public int OutputTokens { get; set; }
}

public class ConvertReport
{
    public string GeneratedUtc { get; set; } = string.Empty;
    public int Total { get; set; }
    public int Success { get; set; }
    public int Failed { get; set; }
    public int TotalInputTokens { get; set; }
    public int TotalOutputTokens { get; set; }
    public List<ConvertResult> Results { get; set; } = new();
}
