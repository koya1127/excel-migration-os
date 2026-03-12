using System.IO.Compression;
using ExcelMigrationApi.Models;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelMigrationApi.Services;

public class ScanService
{
    private readonly ILogger<ScanService> _logger;

    public ScanService(ILogger<ScanService> logger)
    {
        _logger = logger;
    }

    private static readonly HashSet<string> ExcelExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".xls", ".xlsx", ".xlsm"
    };

    private static readonly HashSet<string> VolatileFunctions = new(StringComparer.OrdinalIgnoreCase)
    {
        "NOW(", "TODAY(", "RAND(", "RANDBETWEEN(", "OFFSET(", "INDIRECT(", "CELL(", "INFO("
    };

    private static readonly Dictionary<string, List<string>> IncompatibleFunctions = new()
    {
        ["missing"] = new()
        {
            "AGGREGATE", "CUBEVALUE", "CUBEMEMBER", "CUBESET", "CUBERANKEDMEMBER",
            "CUBEKPIMEMBER", "CUBESETCOUNT", "CUBEMEMBERPROPERTY", "CALL", "REGISTER.ID",
            "RTD", "SQL.REQUEST", "EUROCONVERT", "WEBSERVICE", "FILTERXML",
            "PHONETIC", "JIS", "ASC", "BAHTTEXT"
        },
        ["partial"] = new()
        {
            "XLOOKUP", "XMATCH", "LET", "LAMBDA", "GETPIVOTDATA", "INFO", "CELL", "ERROR.TYPE"
        },
        ["check"] = new()
        {
            "SUBTOTAL", "INDIRECT", "OFFSET"
        }
    };

    private static readonly HashSet<string> AllIncompatibleNames;

    static ScanService()
    {
        AllIncompatibleNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var list in IncompatibleFunctions.Values)
        {
            foreach (var name in list)
            {
                AllIncompatibleNames.Add(name);
            }
        }
    }

    public List<string> FindExcelFiles(string rootPath)
    {
        var files = new List<string>();
        foreach (var file in Directory.EnumerateFiles(rootPath, "*.*", SearchOption.AllDirectories))
        {
            var ext = System.IO.Path.GetExtension(file);
            if (ExcelExtensions.Contains(ext))
            {
                files.Add(file);
            }
        }
        files.Sort(StringComparer.OrdinalIgnoreCase);
        return files;
    }

    public FileReport AnalyzeFile(string filePath)
    {
        var fileInfo = new FileInfo(filePath);
        var ext = fileInfo.Extension.ToLowerInvariant();
        bool hasMacro = ext == ".xlsm";

        if (ext == ".xls")
        {
            return AnalyzeLegacyXls(filePath, fileInfo);
        }

        // For .xlsx and .xlsm, check for VBA project
        if (ext == ".xlsm")
        {
            hasMacro = ContainsVbaProject(filePath);
        }

        return AnalyzeOpenXmlFile(filePath, fileInfo, hasMacro);
    }

    private FileReport AnalyzeOpenXmlFile(string filePath, FileInfo fileInfo, bool hasMacro)
    {
        var report = new FileReport
        {
            Path = filePath,
            Extension = fileInfo.Extension.ToLowerInvariant(),
            SizeBytes = fileInfo.Length,
            ModifiedUtc = fileInfo.LastWriteTimeUtc.ToString("o"),
            HasMacro = hasMacro
        };

        try
        {
            using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            var workbook = new XSSFWorkbook(fs);

            report.SheetCount = workbook.NumberOfSheets;

            // Count VBA modules via VbaExtractor
            if (hasMacro)
            {
                try
                {
                    var vbaModules = VbaExtractor.ExtractModules(filePath);
                    report.VbaModuleCount = vbaModules.Count;
                    report.VbaTotalCodeLength = vbaModules.Sum(m => m.Code?.Length ?? 0);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "VBA extraction failed for {FilePath}", filePath);
                    // Leave VbaModuleCount/VbaTotalCodeLength as null to indicate failure
                }
            }

            // Named ranges
            try { report.NamedRangeCount = workbook.NumberOfNames; } catch { }

            // Walk all cells to count formulas
            int formulaCount = 0;
            int volatileCount = 0;
            int externalLinkCount = 0;
            var incompatibleFound = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int s = 0; s < workbook.NumberOfSheets; s++)
            {
                var sheet = workbook.GetSheetAt(s);
                if (sheet == null) continue;

                for (int r = sheet.FirstRowNum; r <= sheet.LastRowNum; r++)
                {
                    var row = sheet.GetRow(r);
                    if (row == null) continue;

                    for (int c = row.FirstCellNum; c < row.LastCellNum; c++)
                    {
                        var cell = row.GetCell(c);
                        if (cell == null || cell.CellType != CellType.Formula) continue;

                        string formula;
                        try { formula = cell.CellFormula; } catch { continue; }

                        formulaCount++;
                        var upper = formula.ToUpperInvariant();

                        // Volatile functions
                        foreach (var vf in VolatileFunctions)
                        {
                            if (upper.Contains(vf, StringComparison.OrdinalIgnoreCase))
                            {
                                volatileCount++;
                                break;
                            }
                        }

                        // External links (references containing [)
                        if (upper.Contains('['))
                        {
                            externalLinkCount++;
                        }

                        // Incompatible functions
                        foreach (var funcName in AllIncompatibleNames)
                        {
                            if (upper.Contains(funcName + "(", StringComparison.OrdinalIgnoreCase))
                            {
                                incompatibleFound.Add(funcName);
                            }
                        }
                    }
                }
            }

            report.FormulaCount = formulaCount;
            report.VolatileFormulaCount = volatileCount;
            report.ExternalLinkCount = externalLinkCount;
            report.IncompatibleFunctionCount = incompatibleFound.Count;

            // Calculate risk score
            CalculateRiskScore(report, incompatibleFound);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Scan analysis failed for {FilePath}", filePath);
            report.AnalysisFailed = true;
            report.Notes.Add("ファイルが複雑なため解析できませんでした");
            report.RiskScore = 50;
        }

        return report;
    }

    private FileReport AnalyzeLegacyXls(string filePath, FileInfo fileInfo)
    {
        var report = new FileReport
        {
            Path = filePath,
            Extension = ".xls",
            SizeBytes = fileInfo.Length,
            ModifiedUtc = fileInfo.LastWriteTimeUtc.ToString("o"),
        };

        try
        {
            using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            var workbook = new HSSFWorkbook(fs);

            report.SheetCount = workbook.NumberOfSheets;

            // Check for VBA macros
            try
            {
                var vba = workbook.GetType().GetProperty("VBAProject")?.GetValue(workbook);
                report.HasMacro = vba != null;
            }
            catch
            {
                // Fallback: check OLE2 stream for _VBA_PROJECT_CUR
                report.HasMacro = ContainsVbaProjectOle2(filePath);
            }

            // Named ranges
            try { report.NamedRangeCount = workbook.NumberOfNames; } catch { }

            int formulaCount = 0;
            int volatileCount = 0;
            int externalLinkCount = 0;
            var incompatibleFound = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int s = 0; s < workbook.NumberOfSheets; s++)
            {
                var sheet = workbook.GetSheetAt(s);
                if (sheet == null) continue;

                for (int r = sheet.FirstRowNum; r <= sheet.LastRowNum; r++)
                {
                    var row = sheet.GetRow(r);
                    if (row == null) continue;

                    for (int c = row.FirstCellNum; c < row.LastCellNum; c++)
                    {
                        var cell = row.GetCell(c);
                        if (cell == null || cell.CellType != CellType.Formula) continue;

                        string formula;
                        try { formula = cell.CellFormula; } catch { continue; }

                        formulaCount++;
                        var upper = formula.ToUpperInvariant();

                        foreach (var vf in VolatileFunctions)
                        {
                            if (upper.Contains(vf, StringComparison.OrdinalIgnoreCase))
                            {
                                volatileCount++;
                                break;
                            }
                        }

                        if (upper.Contains('['))
                            externalLinkCount++;

                        foreach (var funcName in AllIncompatibleNames)
                        {
                            if (upper.Contains(funcName + "(", StringComparison.OrdinalIgnoreCase))
                                incompatibleFound.Add(funcName);
                        }
                    }
                }
            }

            report.FormulaCount = formulaCount;
            report.VolatileFormulaCount = volatileCount;
            report.ExternalLinkCount = externalLinkCount;
            report.IncompatibleFunctionCount = incompatibleFound.Count;

            CalculateRiskScore(report, incompatibleFound);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Legacy .xls scan failed for {FilePath}", filePath);
            report.RiskScore = 30;
            report.Notes = new List<string> { "旧形式(.xls)", "ファイル分析中にエラーが発生しました" };
        }

        return report;
    }

    private bool ContainsVbaProjectOle2(string filePath)
    {
        try
        {
            using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            var ole2 = new NPOI.POIFS.FileSystem.POIFSFileSystem(fs);
            return ole2.Root.HasEntry("_VBA_PROJECT_CUR") || ole2.Root.HasEntry("Macros");
        }
        catch
        {
            return false;
        }
    }

    private bool ContainsVbaProject(string filePath)
    {
        try
        {
            using var zip = ZipFile.OpenRead(filePath);
            return zip.Entries.Any(e =>
                e.FullName.Equals("xl/vbaProject.bin", StringComparison.OrdinalIgnoreCase));
        }
        catch
        {
            return false;
        }
    }

    private void CalculateRiskScore(FileReport report, HashSet<string> incompatibleFound)
    {
        int score = 0;
        var notes = new List<string>();

        if (report.HasMacro)
        {
            score += 40;
            notes.Add("マクロ有り");
        }

        if (report.VolatileFormulaCount > 0)
        {
            score += 15;
            notes.Add("揮発性の数式あり");
        }

        if (report.ExternalLinkCount > 0)
        {
            score += 20;
            notes.Add("外部リンクあり");
        }

        if (report.FormulaCount > 10000)
        {
            score += 10;
            notes.Add("数式が多い");
        }

        if (incompatibleFound.Count > 0)
        {
            score += 25;
            notes.Add($"非互換関数: {string.Join(", ", incompatibleFound.OrderBy(x => x))}");
        }

        report.RiskScore = Math.Min(score, 100);
        report.Notes = notes.Count > 0 ? notes : new List<string> { "リスク低" };
    }

    public List<GroupSummary> BuildGroupSummaries(List<FileReport> files, string groupBy, string inputRoot)
    {
        if (string.IsNullOrEmpty(groupBy) || groupBy == "none")
        {
            return new List<GroupSummary>();
        }

        var groups = new Dictionary<string, List<int>>();

        for (int i = 0; i < files.Count; i++)
        {
            string groupName = groupBy switch
            {
                "prefix" => GetPrefixGroup(files[i].Path),
                "subfolder" => GetSubfolderGroup(files[i].Path, inputRoot),
                _ => "all"
            };

            if (!groups.ContainsKey(groupName))
            {
                groups[groupName] = new List<int>();
            }
            groups[groupName].Add(i);
        }

        var summaries = new List<GroupSummary>();
        foreach (var (groupName, indices) in groups.OrderBy(g => g.Key))
        {
            var groupFiles = indices.Select(i => files[i]).ToList();

            int macroCount = groupFiles.Count(f => f.HasMacro);
            int totalVba = groupFiles.Sum(f => f.VbaModuleCount ?? 0);
            double avgRisk = groupFiles.Average(f => f.RiskScore);
            int maxRisk = groupFiles.Max(f => f.RiskScore);
            int totalFormulas = groupFiles.Sum(f => f.FormulaCount);
            int totalIncompat = groupFiles.Sum(f => f.IncompatibleFunctionCount);

            string difficulty = DetermineDifficulty(maxRisk, macroCount, totalIncompat);

            summaries.Add(new GroupSummary
            {
                GroupName = groupName,
                FileCount = groupFiles.Count,
                TotalSizeBytes = groupFiles.Sum(f => f.SizeBytes),
                MacroFileCount = macroCount,
                TotalVbaModules = totalVba,
                AvgRiskScore = Math.Round(avgRisk, 1),
                MaxRiskScore = maxRisk,
                TotalFormulas = totalFormulas,
                TotalIncompatibleFunctions = totalIncompat,
                MigrationDifficulty = difficulty,
                FileIndices = indices
            });
        }

        return summaries;
    }

    private string GetPrefixGroup(string filePath)
    {
        var fileName = System.IO.Path.GetFileNameWithoutExtension(filePath);
        var underscoreIndex = fileName.IndexOf('_');
        return underscoreIndex > 0 ? fileName[..underscoreIndex] : "(未分類)";
    }

    private string GetSubfolderGroup(string filePath, string inputRoot)
    {
        var relativePath = System.IO.Path.GetRelativePath(inputRoot, filePath);
        var parts = relativePath.Split(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar);

        if (parts.Length <= 1)
        {
            // File is directly in root
            return new DirectoryInfo(inputRoot).Name;
        }

        return parts[0];
    }

    private string DetermineDifficulty(int maxRisk, int macroCount, int totalIncompat)
    {
        if (maxRisk >= 70 || totalIncompat > 0)
            return "Hard";
        if (maxRisk >= 40 || macroCount > 0)
            return "Medium";
        return "Easy";
    }
}
