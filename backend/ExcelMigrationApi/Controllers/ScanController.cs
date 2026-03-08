using ExcelMigrationApi.Models;
using ExcelMigrationApi.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace ExcelMigrationApi.Controllers;

[Authorize]
[ApiController]
[Route("api/[controller]")]
public class ScanController : ControllerBase
{
    private readonly ScanService _scanService;

    public ScanController(ScanService scanService)
    {
        _scanService = scanService;
    }

    [HttpPost]
    [RequestSizeLimit(200 * 1024 * 1024)] // 200MB limit
    public async Task<ActionResult<ScanReport>> Scan(
        [FromForm] List<IFormFile> files,
        [FromForm] string groupBy = "none")
    {
        if (files == null || files.Count == 0)
        {
            return BadRequest(new { error = "No files uploaded" });
        }

        // Validate groupBy parameter
        var validGroupBy = new HashSet<string> { "none", "prefix", "subfolder" };
        if (!validGroupBy.Contains(groupBy))
        {
            return BadRequest(new { error = "groupBy must be one of: none, prefix, subfolder" });
        }

        var tempDir = Path.Combine(Path.GetTempPath(), "excel-migration-scan-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        try
        {
            // Save uploaded files to temp directory
            foreach (var file in files)
            {
                if (file.Length == 0) continue;

                // Preserve original filename (sanitize path separators)
                var fileName = Path.GetFileName(file.FileName);
                var filePath = Path.Combine(tempDir, fileName);

                await using var stream = new FileStream(filePath, FileMode.Create);
                await file.CopyToAsync(stream);
            }

            // Scan all files
            var excelFiles = _scanService.FindExcelFiles(tempDir);
            var fileReports = new List<FileReport>();

            foreach (var filePath in excelFiles)
            {
                var report = _scanService.AnalyzeFile(filePath);
                // Replace temp path with original filename for cleaner output
                report.Path = Path.GetFileName(report.Path);
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
