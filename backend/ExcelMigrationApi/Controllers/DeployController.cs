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
[EnableRateLimiting("upload-deploy")]
public class DeployController : ControllerBase
{
    private readonly DeployService _deployService;
    private readonly ClerkService _clerkService;

    public DeployController(DeployService deployService, ClerkService clerkService)
    {
        _deployService = deployService;
        _clerkService = clerkService;
    }

    [HttpPost]
    [RequireSubscription]
    [RequestSizeLimit(10 * 1024 * 1024)] // 10MB for JSON body
    public async Task<ActionResult<DeployReport>> Deploy([FromBody] DeployRequest request)
    {
        var userId = User.FindFirst("sub")?.Value ?? User.FindFirst(System.Security.Claims.ClaimTypes.NameIdentifier)?.Value;
        var googleToken = !string.IsNullOrEmpty(userId) ? await _clerkService.GetGoogleToken(userId) : null;
        if (string.IsNullOrEmpty(googleToken))
        {
            return BadRequest(new { error = "Googleアカウントが未連携です。設定画面からGoogleアカウントを連携してください。" });
        }

        if (request == null || string.IsNullOrEmpty(request.SpreadsheetId))
        {
            return BadRequest(new { error = "SpreadsheetIdは必須です" });
        }

        if (request.GasFiles == null || request.GasFiles.Count == 0)
        {
            return BadRequest(new { error = "GASファイルが1つ以上必要です" });
        }

        var report = await _deployService.Deploy(request, googleToken);
        return Ok(report);
    }
}
