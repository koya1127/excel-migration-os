using System.Security.Cryptography;
using System.Text;
using ExcelMigrationApi.Filters;
using ExcelMigrationApi.Models;
using ExcelMigrationApi.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.RateLimiting;

namespace ExcelMigrationApi.Controllers;

[Authorize]
[ApiController]
[Route("api/[controller]")]
[EnableRateLimiting("convert")]
public class MigrateController : ControllerBase
{
    private const int MaxModulesPerRequest = 50;
    private static readonly HashSet<string> AllowedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".xlsx", ".xlsm", ".csv"
        // .xls excluded: VBA extraction not supported after EPPlus removal
    };

    private readonly ExtractService _extractService;
    private readonly ConvertService _convertService;
    private readonly UploadService _uploadService;
    private readonly DeployService _deployService;
    private readonly StripeUsageService _stripeUsageService;
    private readonly ClerkService _clerkService;
    private readonly ILogger<MigrateController> _logger;

    public MigrateController(
        ExtractService extractService,
        ConvertService convertService,
        UploadService uploadService,
        DeployService deployService,
        StripeUsageService stripeUsageService,
        ClerkService clerkService,
        ILogger<MigrateController> logger)
    {
        _extractService = extractService;
        _convertService = convertService;
        _uploadService = uploadService;
        _deployService = deployService;
        _stripeUsageService = stripeUsageService;
        _clerkService = clerkService;
        _logger = logger;
    }

    /// <summary>
    /// End-to-end migration: upload -> extract -> convert -> deploy GAS + usage sheet
    /// </summary>
    [HttpPost]
    [RequireSubscription]
    [RequestSizeLimit(50 * 1024 * 1024)] // 50MB total (down from 200MB)
    public async Task<ActionResult<MigrateReport>> Migrate(
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

        // Validate folderId format (Google Drive folder IDs: alphanumeric, hyphens, underscores, 10-128 chars)
        if (!string.IsNullOrEmpty(folderId) && !System.Text.RegularExpressions.Regex.IsMatch(folderId, @"^[a-zA-Z0-9_-]{10,128}$"))
        {
            return BadRequest(new { error = "フォルダIDの形式が不正です（英数字・ハイフン・アンダースコア、10〜128文字）" });
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
                return BadRequest(new { error = $"サポートされていないファイル形式です: {ext}（.xlsx, .xlsm, .csv のみ対応）" });
            }
            if (file.Length > 30 * 1024 * 1024)
            {
                return BadRequest(new { error = $"ファイル「{file.FileName}」が30MBを超えています" });
            }
        }

        var migrateReport = new MigrateReport
        {
            GeneratedUtc = DateTime.UtcNow.ToString("o")
        };

        var tempDir = Path.Combine(Path.GetTempPath(), "excel-migration-migrate-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        try
        {
            // Save uploaded files to temp directory (deduplicate names to prevent overwrite)
            var filePaths = new List<string>();
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var file in files)
            {
                if (file.Length == 0) continue;

                var fileName = Path.GetFileName(file.FileName);
                // If same basename already exists, prefix with a counter to avoid overwrite
                if (!usedNames.Add(fileName))
                {
                    var nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                    var ext = Path.GetExtension(fileName);
                    var counter = 2;
                    do
                    {
                        fileName = $"{nameWithoutExt}_{counter}{ext}";
                        counter++;
                    } while (!usedNames.Add(fileName));
                }

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
                    SourceFile = module.SourceFile,
                    SheetName = module.SheetName,
                    ButtonContext = buttonContext.Count > 0 ? buttonContext : null,
                    DetectedEvents = module.DetectedEvents.Count > 0 ? module.DetectedEvents : null
                });
            }

            if (convertRequests.Count > MaxModulesPerRequest)
            {
                return BadRequest(new { error = $"1回のリクエストで変換できるモジュール数は最大{MaxModulesPerRequest}です（検出: {convertRequests.Count}）" });
            }

            var convertReport = await _convertService.ConvertBatch(convertRequests);
            migrateReport.Convert = convertReport;

            // Report usage — block response if billing fails
            var totalTokens = convertReport.TotalInputTokens + convertReport.TotalOutputTokens;
            if (totalTokens > 0)
            {
                if (HttpContext.Items.TryGetValue("ClerkUserMeta", out var metaObj)
                    && metaObj is ClerkService.ClerkUserMeta meta
                    && !string.IsNullOrEmpty(meta.StripeCustomerId))
                {
                    try
                    {
                        var idempotencyKey = BuildIdempotencyKey("migrate", convertRequests);
                        await _stripeUsageService.ReportUsage(meta.StripeCustomerId, totalTokens, idempotencyKey);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "Stripe usage report failed during migrate");
                        return StatusCode(502, new { error = "課金処理に失敗しました。変換は完了していますが、再度お試しください。" });
                    }
                }
                else
                {
                    return StatusCode(502, new { error = "課金情報が取得できません" });
                }
            }

            // Step 4: Deploy GAS to uploaded spreadsheets + add usage sheet
            var successfulConversions = convertReport.Results
                .Where(r => r.Status == "success")
                .ToList();

            if (successfulConversions.Count == 0)
            {
                return Ok(migrateReport);
            }

            // Build mapping from source file to uploaded spreadsheet
            var uploadMap = uploadReport.Files
                .Where(f => f.Status == "success" && !string.IsNullOrEmpty(f.DriveFileId))
                .ToDictionary(f => f.FileName, f => f, StringComparer.OrdinalIgnoreCase);

            // Group converted modules by source file
            var modulesByFile = successfulConversions
                .GroupBy(r => r.SourceFile ?? "")
                .ToList();

            foreach (var group in modulesByFile)
            {
                if (!uploadMap.TryGetValue(group.Key, out var uploadResult))
                {
                    _logger.LogWarning("Deploy skipped: no matching upload for source file '{SourceFile}'", group.Key);
                    migrateReport.Deploys.Add(new DeployReport
                    {
                        Status = "skipped",
                        Error = $"アップロード先が見つかりません: {group.Key}"
                    });
                    continue;
                }

                var gasFiles = group.Select(r => new GasFile
                {
                    Name = r.ModuleName,
                    Source = r.GasCode,
                    Type = "SERVER_JS"
                }).ToList();

                // Deploy GAS
                var deployRequest = new DeployRequest
                {
                    SpreadsheetId = uploadResult.DriveFileId,
                    GasFiles = gasFiles
                };
                var deployReport = await _deployService.Deploy(deployRequest, googleToken);
                deployReport.WebViewLink = uploadResult.WebViewLink;
                migrateReport.Deploys.Add(deployReport);

                // Add usage sheet to the spreadsheet
                var usageSheet = DeployService.BuildUsageSheet(group.Key, group);
                await _deployService.AddUsageSheet(uploadResult.DriveFileId, usageSheet, googleToken);
            }

            return Ok(migrateReport);
        }
        finally
        {
            try { Directory.Delete(tempDir, recursive: true); } catch { }
        }
    }

    private static string BuildIdempotencyKey(string prefix, List<ConvertRequest> requests)
    {
        using var sha = SHA256.Create();
        var sb = new StringBuilder();
        foreach (var r in requests)
        {
            sb.Append(r.ModuleName).Append(':').Append(r.VbaCode?.Length ?? 0).Append(':');
            sb.Append(r.VbaCode ?? "");
        }
        var hash = Convert.ToHexString(sha.ComputeHash(Encoding.UTF8.GetBytes(sb.ToString()))).ToLowerInvariant();
        return $"{prefix}-{hash}";
    }
}
