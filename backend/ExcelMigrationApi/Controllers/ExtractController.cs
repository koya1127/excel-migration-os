using ExcelMigrationApi.Models;
using ExcelMigrationApi.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace ExcelMigrationApi.Controllers;

[Authorize]
[ApiController]
[Route("api/[controller]")]
public class ExtractController : ControllerBase
{
    private readonly ExtractService _extractService;

    public ExtractController(ExtractService extractService)
    {
        _extractService = extractService;
    }

    [HttpPost]
    [RequestSizeLimit(200 * 1024 * 1024)]
    public async Task<ActionResult<ExtractReport>> Extract([FromForm] List<IFormFile> files)
    {
        if (files == null || files.Count == 0)
        {
            return BadRequest(new { error = "No files uploaded" });
        }

        var tempDir = Path.Combine(Path.GetTempPath(), "excel-migration-extract-" + Guid.NewGuid().ToString("N"));
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

            var report = _extractService.Extract(filePaths);
            return Ok(report);
        }
        finally
        {
            try { Directory.Delete(tempDir, recursive: true); } catch { }
        }
    }
}
