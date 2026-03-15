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
public class ConvertController : ControllerBase
{
    private readonly ConvertService _convertService;
    private readonly StripeUsageService _stripeUsageService;
    private readonly ILogger<ConvertController> _logger;
    private const int MaxModulesPerRequest = 50;

    public ConvertController(ConvertService convertService, StripeUsageService stripeUsageService, ILogger<ConvertController> logger)
    {
        _convertService = convertService;
        _stripeUsageService = stripeUsageService;
        _logger = logger;
    }

    /// <summary>
    /// Convert a single VBA module to Google Apps Script.
    /// </summary>
    [HttpPost]
    [RequireSubscription]
    [RequestSizeLimit(10 * 1024 * 1024)] // 10MB for JSON body
    public async Task<ActionResult<ConvertReport>> Convert([FromBody] List<ConvertRequest> requests)
    {
        if (requests == null || requests.Count == 0)
            return BadRequest(new { error = "No convert requests provided" });

        if (requests.Count > MaxModulesPerRequest)
            return BadRequest(new { error = $"1回のリクエストで変換できるモジュール数は最大{MaxModulesPerRequest}です" });

        var report = await _convertService.ConvertBatch(requests);

        // Report usage — fail the response if billing fails (prevent free usage)
        var idempotencyKey = BuildIdempotencyKey("convert", requests);
        if (!await TryReportUsage(report, idempotencyKey))
        {
            return StatusCode(502, new { error = "課金処理に失敗しました。しばらくしてから再度お試しください。" });
        }

        return Ok(report);
    }

    /// <summary>
    /// Convert all modules from an ExtractReport (batch conversion).
    /// </summary>
    [HttpPost("batch")]
    [RequireSubscription]
    [RequestSizeLimit(10 * 1024 * 1024)] // 10MB for JSON body
    public async Task<ActionResult<ConvertReport>> ConvertBatch([FromBody] ExtractReport extractReport)
    {
        if (extractReport == null || extractReport.Modules.Count == 0)
            return BadRequest(new { error = "No modules in extract report" });

        var requests = new List<ConvertRequest>();

        foreach (var module in extractReport.Modules)
        {
            if (string.IsNullOrWhiteSpace(module.Code) || module.CodeLines <= 1)
                continue;

            var buttonContext = extractReport.Controls
                .Where(c => c.SourceFile == module.SourceFile && !string.IsNullOrEmpty(c.Macro))
                .ToList();

            requests.Add(new ConvertRequest
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

        if (requests.Count > MaxModulesPerRequest)
            return BadRequest(new { error = $"1回のリクエストで変換できるモジュール数は最大{MaxModulesPerRequest}です" });

        if (requests.Count == 0)
        {
            return Ok(new ConvertReport
            {
                GeneratedUtc = DateTime.UtcNow.ToString("o"),
                Total = 0,
                Success = 0,
                Failed = 0
            });
        }

        var report = await _convertService.ConvertBatch(requests);

        var idempotencyKey = BuildIdempotencyKey("batch", requests);
        if (!await TryReportUsage(report, idempotencyKey))
        {
            return StatusCode(502, new { error = "課金処理に失敗しました。しばらくしてから再度お試しください。" });
        }

        return Ok(report);
    }

    private static string BuildIdempotencyKey(string prefix, List<ConvertRequest> requests)
    {
        // Deterministic key from content: same VBA input = same idempotency key.
        // Prevents double-billing on client retries while allowing re-conversions
        // of truly different content.
        using var sha = SHA256.Create();
        var sb = new StringBuilder();
        foreach (var r in requests)
        {
            sb.Append(r.ModuleName).Append(':').Append(r.VbaCode?.Length ?? 0).Append(':');
            sb.Append(r.VbaCode ?? "");
        }
        var hash = System.Convert.ToHexString(sha.ComputeHash(Encoding.UTF8.GetBytes(sb.ToString()))).ToLowerInvariant();
        return $"{prefix}-{hash}";
    }

    private async Task<bool> TryReportUsage(ConvertReport report, string idempotencyKey)
    {
        var totalTokens = report.TotalInputTokens + report.TotalOutputTokens;
        if (totalTokens <= 0) return true;

        if (HttpContext.Items.TryGetValue("ClerkUserMeta", out var metaObj)
            && metaObj is ClerkService.ClerkUserMeta meta
            && !string.IsNullOrEmpty(meta.StripeCustomerId))
        {
            try
            {
                await _stripeUsageService.ReportUsage(meta.StripeCustomerId, totalTokens, idempotencyKey);
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Stripe usage report failed");
                return false;
            }
        }

        // No Stripe customer ID — should not happen with RequireSubscription, but fail safe
        return false;
    }
}
