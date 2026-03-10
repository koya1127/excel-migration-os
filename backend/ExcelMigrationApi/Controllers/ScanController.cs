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
    [RequestSizeLimit(50 * 1024 * 1024)] // 50MB limit
    public async Task<ActionResult<ScanReport>> Scan(
        [FromForm] List<IFormFile> files,
        [FromForm] string groupBy = "none")
    {
        if (files == null || files.Count == 0)
        {
            return BadRequest(new { error = "ファイルがアップロードされていません" });
        }

        // Validate file extensions
        foreach (var file in files)
        {
            var ext = Path.GetExtension(file.FileName);
            if (!AllowedExtensions.Contains(ext))
            {
                return BadRequest(new { error = $"サポートされていないファイル形式です: {ext}（.xlsx, .xlsm, .xls, .csv のみ対応）" });
            }
        }

        if (files.Count > 100)
        {
            return BadRequest(new { error = "1回のリクエストでアップロードできるファイル数は最大100です" });
        }

        // Reject individual files over 30MB
        foreach (var file in files)
        {
            if (file.Length > 30 * 1024 * 1024)
            {
                return BadRequest(new { error = $"ファイル「{file.FileName}」が30MBを超えています" });
            }
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
