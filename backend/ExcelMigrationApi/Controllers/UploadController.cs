using ExcelMigrationApi.Models;
using ExcelMigrationApi.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace ExcelMigrationApi.Controllers;

[Authorize]
[ApiController]
[Route("api/[controller]")]
public class UploadController : ControllerBase
{
    private readonly UploadService _uploadService;

    public UploadController(UploadService uploadService)
    {
        _uploadService = uploadService;
    }

    [HttpPost]
    [RequestSizeLimit(200 * 1024 * 1024)]
    public async Task<ActionResult<UploadReport>> Upload(
        [FromForm] List<IFormFile> files,
        [FromForm] bool convertToSheets = true,
        [FromForm] string? folderId = null)
    {
        var googleToken = Request.Headers["X-Google-Token"].FirstOrDefault();
        if (string.IsNullOrEmpty(googleToken))
        {
            return BadRequest(new { error = "Google account not connected. Please link your Google account in Settings." });
        }

        if (files == null || files.Count == 0)
        {
            return BadRequest(new { error = "No files uploaded" });
        }

        var tempDir = Path.Combine(Path.GetTempPath(), "excel-migration-upload-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        try
        {
            var filePaths = new List<string>();

            foreach (var file in files)
            {
                if (file.Length == 0) continue;

                var fileName = Path.GetFileName(file.FileName);
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
