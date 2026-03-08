using ExcelMigrationApi.Models;
using ExcelMigrationApi.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace ExcelMigrationApi.Controllers;

[Authorize]
[ApiController]
[Route("api/[controller]")]
public class DeployController : ControllerBase
{
    private readonly DeployService _deployService;

    public DeployController(DeployService deployService)
    {
        _deployService = deployService;
    }

    [HttpPost]
    public async Task<ActionResult<DeployReport>> Deploy([FromBody] DeployRequest request)
    {
        var googleToken = Request.Headers["X-Google-Token"].FirstOrDefault();
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
