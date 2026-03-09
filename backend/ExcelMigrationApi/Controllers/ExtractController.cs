using ExcelMigrationApi.Models;
using ExcelMigrationApi.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.RateLimiting;

namespace ExcelMigrationApi.Controllers;

[Authorize]
[ApiController]
[Route("api/[controller]")]
[EnableRateLimiting("per-user")]
public class ExtractController : ControllerBase
{
    private readonly ExtractService _extractService;

    public ExtractController(ExtractService extractService)
    {
        _extractService = extractService;
    }

    private static readonly HashSet<string> AllowedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".xlsm", ".xls"
    };

    [HttpPost]
    [RequestSizeLimit(50 * 1024 * 1024)] // 50MB limit
    public async Task<ActionResult<ExtractReport>> Extract([FromForm] List<IFormFile> files)
    {
        if (files == null || files.Count == 0)
        {
            return BadRequest(new { error = "ファイルがアップロードされていません" });
        }

        // Validate file extensions (extract only supports macro-enabled files)
        foreach (var file in files)
        {
            var ext = Path.GetExtension(file.FileName);
            if (!AllowedExtensions.Contains(ext))
            {
                return BadRequest(new { error = $"VBA抽出は .xlsm, .xls のみ対応しています: {file.FileName}" });
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

        var tempDir = Path.Combine(Path.GetTempPath(), "excel-migration-extract-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        try
        {
            var filePaths = new List<string>();
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var file in files)
            {
                if (file.Length == 0) continue;

                var fileName = Path.GetFileName(file.FileName);
                if (!usedNames.Add(fileName))
                {
                    var nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                    var ext = Path.GetExtension(fileName);
                    var counter = 2;
                    do { fileName = $"{nameWithoutExt}_{counter}{ext}"; counter++; } while (!usedNames.Add(fileName));
                }
                var filePath = Path.Combine(tempDir, fileName);

                await using var stream = new FileStream(filePath, FileMode.Create);
                await file.CopyToAsync(stream);

                filePaths.Add(filePath);
            }

            var report = _extractService.Extract(filePaths);
            return Ok(report);
        }
        finally
        {
            try { Directory.Delete(tempDir, recursive: true); } catch { }
        }
    }
}
