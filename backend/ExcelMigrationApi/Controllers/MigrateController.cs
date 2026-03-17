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
    private readonly TrackRouterService _trackRouterService;
    private readonly PythonConvertService _pythonConvertService;
    private readonly PythonPackagerService _pythonPackagerService;
    private readonly StripeUsageService _stripeUsageService;
    private readonly ClerkService _clerkService;
    private readonly ILogger<MigrateController> _logger;

    public MigrateController(
        ExtractService extractService,
        ConvertService convertService,
        UploadService uploadService,
        DeployService deployService,
        TrackRouterService trackRouterService,
        PythonConvertService pythonConvertService,
        PythonPackagerService pythonPackagerService,
        StripeUsageService stripeUsageService,
        ClerkService clerkService,
        ILogger<MigrateController> logger)
    {
        _extractService = extractService;
        _convertService = convertService;
        _uploadService = uploadService;
        _deployService = deployService;
        _trackRouterService = trackRouterService;
        _pythonConvertService = pythonConvertService;
        _pythonPackagerService = pythonPackagerService;
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
        [FromForm] string? folderId = null,
        [FromForm] string trackMode = "auto")
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

            // Step 3a: Track routing — determine which modules go to GAS vs Python
            var validModules = extractReport.Modules
                .Where(m => !string.IsNullOrWhiteSpace(m.Code) && m.CodeLines > 1)
                .ToList();
            var trackResult = _trackRouterService.Route(validModules);
            migrateReport.TrackRouting = trackResult;

            // Determine which tracks to execute based on trackMode
            var runGas = trackMode is "auto" or "sheets_only" or "both";
            var runPython = trackMode is "auto" or "local_only" or "both";

            // For "auto": only run each track if there are modules for it
            if (trackMode == "auto")
            {
                runGas = trackResult.Track1Modules.Count > 0;
                runPython = trackResult.Track2Modules.Count > 0;
            }

            // Filter convertRequests to Track 1 modules only (for GAS)
            var track1Names = new HashSet<string>(trackResult.Track1Modules.Select(m => m.ModuleName));
            var gasConvertRequests = trackMode == "local_only"
                ? new List<ConvertRequest>()
                : trackMode == "sheets_only" || trackMode == "both"
                    ? convertRequests // all modules go to GAS
                    : convertRequests.Where(r => track1Names.Contains(r.ModuleName)).ToList(); // auto: only Track 1

            // GAS conversion
            ConvertReport convertReport;
            if (gasConvertRequests.Count > 0 && runGas)
            {
                convertReport = await _convertService.ConvertBatch(gasConvertRequests);
            }
            else
            {
                convertReport = new ConvertReport { GeneratedUtc = DateTime.UtcNow.ToString("o") };
            }
            migrateReport.Convert = convertReport;

            // Python conversion (Track 2)
            PythonConvertReport? pythonReport = null;
            if (runPython && trackResult.Track2Modules.Count > 0)
            {
                var spreadsheetId = uploadReport.Files
                    .FirstOrDefault(f => f.Status == "success")?.DriveFileId;

                var pythonRequests = trackResult.Track2Modules.Select(m => new PythonConvertRequest
                {
                    VbaCode = m.Code,
                    ModuleName = m.ModuleName,
                    ModuleType = m.ModuleType,
                    SourceFile = m.SourceFile,
                    SheetName = m.SheetName,
                    SpreadsheetId = spreadsheetId
                }).ToList();

                pythonReport = await _pythonConvertService.ConvertBatch(pythonRequests);
                migrateReport.PythonConvert = pythonReport;

                // Package Python files into ZIP
                if (pythonReport.Success > 0)
                {
                    var sourceFileName = files.FirstOrDefault()?.FileName ?? "output";
                    var package = _pythonPackagerService.Package(sourceFileName, pythonReport.Results, spreadsheetId);

                    // Save ZIP to temp and create download URL
                    var zipPath = Path.Combine(tempDir, package.FileName);
                    await System.IO.File.WriteAllBytesAsync(zipPath, package.ZipData);

                    // Store ZIP in memory cache for download endpoint
                    var zipId = Guid.NewGuid().ToString("N");
                    PythonZipCache.Store(zipId, package);
                    migrateReport.PythonPackageUrl = $"/api/migrate/download/{zipId}";
                }
            }

            // Report usage — block response if billing fails (include both GAS + Python tokens)
            var totalTokens = convertReport.TotalInputTokens + convertReport.TotalOutputTokens
                + (pythonReport?.TotalInputTokens ?? 0) + (pythonReport?.TotalOutputTokens ?? 0);
            if (totalTokens > 0)
            {
                if (HttpContext.Items.TryGetValue("ClerkUserMeta", out var metaObj)
                    && metaObj is ClerkService.ClerkUserMeta meta
                    && !string.IsNullOrEmpty(meta.StripeCustomerId))
                {
                    try
                    {
                        var idempotencyKey = BuildIdempotencyKey("migrate", convertRequests, totalTokens);
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

    /// <summary>
    /// Download endpoint for Python ZIP packages
    /// </summary>
    [HttpGet("download/{zipId}")]
    [AllowAnonymous] // ZIP download doesn't need auth (short-lived URL)
    public ActionResult DownloadPythonPackage(string zipId)
    {
        var package = PythonZipCache.Get(zipId);
        if (package == null)
            return NotFound(new { error = "ダウンロードリンクの有効期限が切れました" });

        return File(package.ZipData, "application/zip", package.FileName);
    }

    private static string BuildIdempotencyKey(string prefix, List<ConvertRequest> requests, int totalTokens = 0)
    {
        using var sha = SHA256.Create();
        var sb = new StringBuilder();
        foreach (var r in requests)
        {
            sb.Append(r.ModuleName).Append(':').Append(r.VbaCode?.Length ?? 0).Append(':');
            sb.Append(r.VbaCode ?? "");
        }
        // Include token count + timestamp (hour granularity) to avoid idempotency conflicts
        sb.Append(':').Append(totalTokens);
        sb.Append(':').Append(DateTimeOffset.UtcNow.ToString("yyyyMMddHH"));
        var hash = Convert.ToHexString(sha.ComputeHash(Encoding.UTF8.GetBytes(sb.ToString()))).ToLowerInvariant();
        return $"{prefix}-{hash}";
    }
}

/// <summary>
/// Simple in-memory cache for Python ZIP packages (auto-expires after 30 minutes)
/// </summary>
public static class PythonZipCache
{
    private static readonly Dictionary<string, (PythonPackage Package, DateTime ExpiresAt)> _cache = new();
    private static readonly object _lock = new();

    public static void Store(string id, PythonPackage package)
    {
        lock (_lock)
        {
            // Clean expired entries
            var expired = _cache.Where(kv => kv.Value.ExpiresAt < DateTime.UtcNow).Select(kv => kv.Key).ToList();
            foreach (var key in expired) _cache.Remove(key);

            _cache[id] = (package, DateTime.UtcNow.AddMinutes(30));
        }
    }

    public static PythonPackage? Get(string id)
    {
        lock (_lock)
        {
            if (_cache.TryGetValue(id, out var entry) && entry.ExpiresAt > DateTime.UtcNow)
            {
                _cache.Remove(id); // One-time download
                return entry.Package;
            }
            return null;
        }
    }
}
