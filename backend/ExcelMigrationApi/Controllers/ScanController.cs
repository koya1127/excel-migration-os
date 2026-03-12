using ExcelMigrationApi.Models;
using ExcelMigrationApi.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.RateLimiting;

namespace ExcelMigrationApi.Controllers;

[AllowAnonymous]
[ApiController]
[Route("api/[controller]")]
[EnableRateLimiting("per-user")]
public class ScanController : ControllerBase
{
    private readonly ScanService _scanService;

    public ScanController(ScanService scanService)
    {
        _scanService = scanService;
    }

    private static readonly HashSet<string> AllowedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".xlsx", ".xlsm", ".xls", ".csv"
    };

    [HttpPost]
    [RequestSizeLimit(500 * 1024 * 1024)] // 500MB ASP.NET limit (actual limit enforced in code below)
    public async Task<ActionResult<ScanReport>> Scan(
        [FromForm] List<IFormFile> files,
        [FromForm] string groupBy = "none")
    {
        if (files == null || files.Count == 0)
        {
            return BadRequest(new { error = "ファイルがアップロードされていません。Excelファイル（.xlsx, .xlsm, .xls）を選択してください。" });
        }

        // Validate file extensions
        var unsupported = files
            .Select(f => Path.GetExtension(f.FileName))
            .Where(ext => !AllowedExtensions.Contains(ext))
            .Distinct()
            .ToList();
        if (unsupported.Count > 0)
        {
            return BadRequest(new { error = $"サポートされていないファイル形式が含まれています: {string.Join(", ", unsupported)}。対応形式: .xlsx, .xlsm, .xls, .csv" });
        }

        if (files.Count > 1000)
        {
            return BadRequest(new { error = $"ファイル数が{files.Count}件です。1回のスキャンは最大1,000件までです。フォルダを分けて再度お試しください。" });
        }

        // Reject individual files over 30MB
        var tooLarge = files.Where(f => f.Length > 30 * 1024 * 1024).Select(f => f.FileName).ToList();
        if (tooLarge.Count > 0)
        {
            return BadRequest(new { error = $"30MBを超えるファイルがあります: {string.Join(", ", tooLarge.Take(5))}。各ファイルは30MB以下にしてください。" });
        }

        // Check total size
        var totalBytes = files.Sum(f => f.Length);
        if (totalBytes > 30 * 1024 * 1024) // TODO: raise back to 300MB after testing
        {
            return BadRequest(new { error = $"合計サイズが{totalBytes / (1024 * 1024)}MBです。1回のスキャンは合計300MBまでです。ファイル数を減らして再度お試しください。" });
        }

        // Validate groupBy parameter
        var validGroupBy = new HashSet<string> { "none", "prefix", "subfolder" };
        if (!validGroupBy.Contains(groupBy))
        {
            return BadRequest(new { error = "groupByはnone, prefix, subfolderのいずれかを指定してください" });
        }

        var tempDir = Path.Combine(Path.GetTempPath(), "excel-migration-scan-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        try
        {
            // Save uploaded files preserving folder structure for subfolder grouping.
            // The browser sends webkitRelativePath or D&D fullPath as the filename.
            var usedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var file in files)
            {
                if (file.Length == 0) continue;

                // Sanitize: normalize separators, remove path traversal components
                var relativePath = file.FileName.Replace('\\', '/');
                relativePath = string.Join('/',
                    relativePath.Split('/').Where(p => p != ".." && p != "." && !string.IsNullOrEmpty(p)));
                if (string.IsNullOrEmpty(relativePath))
                    relativePath = "unknown.xlsx";

                // Deduplicate full relative paths
                if (!usedPaths.Add(relativePath))
                {
                    var dir = Path.GetDirectoryName(relativePath)?.Replace('\\', '/') ?? "";
                    var nameWithoutExt = Path.GetFileNameWithoutExtension(relativePath);
                    var ext = Path.GetExtension(relativePath);
                    var counter = 2;
                    do
                    {
                        relativePath = string.IsNullOrEmpty(dir)
                            ? $"{nameWithoutExt}_{counter}{ext}"
                            : $"{dir}/{nameWithoutExt}_{counter}{ext}";
                        counter++;
                    } while (!usedPaths.Add(relativePath));
                }

                var filePath = Path.Combine(tempDir, relativePath.Replace('/', Path.DirectorySeparatorChar));
                Directory.CreateDirectory(Path.GetDirectoryName(filePath)!);

                await using var stream = new FileStream(filePath, FileMode.Create);
                await file.CopyToAsync(stream);
            }

            // Scan all files
            var excelFiles = _scanService.FindExcelFiles(tempDir);
            var fileReports = new List<FileReport>();

            foreach (var filePath in excelFiles)
            {
                var report = _scanService.AnalyzeFile(filePath);
                // Use relative path from tempDir for subfolder grouping and display
                report.Path = Path.GetRelativePath(tempDir, report.Path).Replace('\\', '/');
                fileReports.Add(report);
            }

            // Build group summaries
            var groups = _scanService.BuildGroupSummaries(fileReports, groupBy, tempDir);

            var scanReport = new ScanReport
            {
                GeneratedUtc = DateTime.UtcNow.ToString("o"),
                InputRoot = "upload",
                FileCount = fileReports.Count,
                Files = fileReports,
                GroupBy = groupBy,
                Groups = groups
            };

            return Ok(scanReport);
        }
        finally
        {
            // Clean up temp files
            try
            {
                Directory.Delete(tempDir, recursive: true);
            }
            catch
            {
                // Best effort cleanup
            }
        }
    }
}
