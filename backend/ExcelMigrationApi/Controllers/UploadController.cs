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
public class UploadController : ControllerBase
{
    private readonly UploadService _uploadService;
    private readonly ClerkService _clerkService;

    public UploadController(UploadService uploadService, ClerkService clerkService)
    {
        _uploadService = uploadService;
        _clerkService = clerkService;
    }

    private static readonly HashSet<string> AllowedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".xlsx", ".xlsm", ".xls", ".csv"
    };

    [HttpPost]
    [RequestSizeLimit(50 * 1024 * 1024)] // 50MB total
    public async Task<ActionResult<UploadReport>> Upload(
        [FromForm] List<IFormFile> files,
        [FromForm] bool convertToSheets = true,
        [FromForm] string? folderId = null)
    {
        var userId = User.FindFirst("sub")?.Value ?? User.FindFirst(System.Security.Claims.ClaimTypes.NameIdentifier)?.Value;
        var googleToken = !string.IsNullOrEmpty(userId) ? await _clerkService.GetGoogleToken(userId) : null;
        if (string.IsNullOrEmpty(googleToken))
        {
            return BadRequest(new { error = "Googleアカウントが未連携です。設定画面からGoogleアカウントを連携してください。" });
        }

        if (files == null || files.Count == 0)
        {
            return BadRequest(new { error = "ファイルがアップロードされていません" });
        }

        if (files.Count > 100)
        {
            return BadRequest(new { error = "1回のリクエストでアップロードできるファイル数は最大100です" });
        }

        // Validate file extensions and sizes
        foreach (var file in files)
        {
            var ext = Path.GetExtension(file.FileName);
            if (!AllowedExtensions.Contains(ext))
            {
                return BadRequest(new { error = $"サポートされていないファイル形式です: {ext}（.xlsx, .xlsm, .xls, .csv のみ対応）" });
            }
            if (file.Length > 30 * 1024 * 1024)
            {
                return BadRequest(new { error = $"ファイル「{file.FileName}」が30MBを超えています" });
            }
        }

        var tempDir = Path.Combine(Path.GetTempPath(), "excel-migration-upload-" + Guid.NewGuid().ToString("N"));
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

            var report = await _uploadService.UploadFiles(filePaths, convertToSheets, folderId, googleToken);
            return Ok(report);
        }
        finally
        {
            try { Directory.Delete(tempDir, recursive: true); } catch { }
        }
    }
}
