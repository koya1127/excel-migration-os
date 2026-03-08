namespace ExcelMigrationApi.Models;

public class VbaModule
{
    public string SourceFile { get; set; } = string.Empty;
    public string ModuleName { get; set; } = string.Empty;
    public string ModuleType { get; set; } = string.Empty; // "Standard", "Class", "Form", "Document"
    public int CodeLines { get; set; }
    public string Code { get; set; } = string.Empty;
}

public class FormControl
{
    public string SourceFile { get; set; } = string.Empty;
    public string SheetName { get; set; } = string.Empty;
    public string ControlName { get; set; } = string.Empty;
    public string ControlType { get; set; } = string.Empty;
    public string Label { get; set; } = string.Empty;
    public string Macro { get; set; } = string.Empty;
}

public class ExtractReport
{
    public string GeneratedUtc { get; set; } = string.Empty;
    public int FileCount { get; set; }
    public int ModuleCount { get; set; }
    public List<VbaModule> Modules { get; set; } = new();
    public List<FormControl> Controls { get; set; } = new();
}
