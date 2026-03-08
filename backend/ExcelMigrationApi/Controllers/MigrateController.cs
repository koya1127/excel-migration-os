using ExcelMigrationApi.Models;
using ExcelMigrationApi.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace ExcelMigrationApi.Controllers;

[Authorize]
[ApiController]
[Route("api/[controller]")]
public class MigrateController : ControllerBase
{
    private readonly ExtractService _extractService;
    private readonly ConvertService _convertService;
    private readonly UploadService _uploadService;
    private readonly DeployService _deployService;

    public MigrateController(
        ExtractService extractService,
        ConvertService convertService,
        UploadService uploadService,
        DeployService deployService)
    {
        _extractService = extractService;
        _convertService = convertService;
        _uploadService = uploadService;
        _deployService = deployService;
    }

    /// <summary>
    /// End-to-end migration: upload -> extract -> convert -> deploy
    /// </summary>
    [HttpPost]
    [RequestSizeLimit(200 * 1024 * 1024)]
    public async Task<ActionResult<MigrateReport>> Migrate(
        [FromForm] List<IFormFile> files,
        [FromForm] bool convertToSheets = true,
        [FromForm] string? folderId = null)
    {
        var googleToken = Request.Headers["X-Google-Token"].FirstOrDefault();
        if (string.IsNullOrEmpty(googleToken))
        {
            return BadRequest(new { error = "Googleアカウントが未連携です。設定画面からGoogleアカウントを連携してください。" });
        }

        if (files == null || files.Count == 0)
        {
            return BadRequest(new { error = "ファイルがアップロードされていません" });
        }

        var migrateReport = new MigrateReport
        {
            GeneratedUtc = DateTime.UtcNow.ToString("o")
        };

        var tempDir = Path.Combine(Path.GetTempPath(), "excel-migration-migrate-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        try
        {
            // Save uploaded files to temp directory
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

            // Step 1: Upload to Google Drive (convert to Sheets)
            var uploadReport = await _uploadService.UploadFiles(filePaths, convertToSheets, folderId, googleToken);
            migrateReport.Upload = uploadReport;

            // Step 2: Extract VBA modules from .xlsm files
            var xlsmPaths = filePaths.Where(f =>
                Path.GetExtension(f).Equals(".xlsm", StringComparison.OrdinalIgnoreCase)).ToList();

            var extractReport = _extractService.Extract(xlsmPaths);
            migrateReport.Extract = extractReport;

            // If no VBA modules found, skip convert and deploy
            if (extractReport.Modules.Count == 0 || extractReport.Modules.All(m => m.CodeLines <= 1))
            {
                migrateReport.Convert = new ConvertReport
                {
                    GeneratedUtc = DateTime.UtcNow.ToString("o"),
                    Total = 0
                };
                migrateReport.Deploy = new DeployReport
                {
                    GeneratedUtc = DateTime.UtcNow.ToString("o"),
                    Status = "skipped",
                    Error = "変換するVBAモジュールがありません"
                };
                return Ok(migrateReport);
            }

            // Step 3: Convert VBA to GAS
            var convertRequests = new List<ConvertRequest>();
            foreach (var module in extractReport.Modules)
            {
                if (string.IsNullOrWhiteSpace(module.Code) || module.CodeLines <= 1)
                    continue;

                var buttonContext = extractReport.Controls
                    .Where(c => c.SourceFile == module.SourceFile && !string.IsNullOrEmpty(c.Macro))
                    .ToList();

                convertRequests.Add(new ConvertRequest
                {
                    VbaCode = module.Code,
                    ModuleName = module.ModuleName,
                    ModuleType = module.ModuleType,
                    ButtonContext = buttonContext.Count > 0 ? buttonContext : null
                });
            }

            var convertReport = await _convertService.ConvertBatch(convertRequests);
            migrateReport.Convert = convertReport;

            // Step 4: Deploy GAS to the first successfully uploaded spreadsheet
            var firstSuccess = uploadReport.Files.FirstOrDefault(f => f.Status == "success");
            if (firstSuccess == null || string.IsNullOrEmpty(firstSuccess.DriveFileId))
            {
                migrateReport.Deploy = new DeployReport
                {
                    GeneratedUtc = DateTime.UtcNow.ToString("o"),
                    Status = "skipped",
                    Error = "デプロイ先のスプレッドシートがありません"
                };
                return Ok(migrateReport);
            }

            var successfulConversions = convertReport.Results
                .Where(r => r.Status == "success")
                .ToList();

            if (successfulConversions.Count == 0)
            {
                migrateReport.Deploy = new DeployReport
                {
                    GeneratedUtc = DateTime.UtcNow.ToString("o"),
                    SpreadsheetId = firstSuccess.DriveFileId,
                    Status = "skipped",
                    Error = "デプロイする変換結果がありません"
                };
                return Ok(migrateReport);
            }

            var deployRequest = new DeployRequest
            {
                SpreadsheetId = firstSuccess.DriveFileId,
                GasFiles = successfulConversions.Select(r => new GasFile
                {
                    Name = r.ModuleName,
                    Source = r.GasCode,
                    Type = "SERVER_JS"
                }).ToList()
            };

            var deployReport = await _deployService.Deploy(deployRequest, googleToken);
            migrateReport.Deploy = deployReport;

            return Ok(migrateReport);
        }
        finally
        {
            try { Directory.Delete(tempDir, recursive: true); } catch { }
        }
    }
}
