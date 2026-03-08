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
        if (request == null || string.IsNullOrEmpty(request.SpreadsheetId))
        {
            return BadRequest(new { error = "SpreadsheetId is required" });
        }

        if (request.GasFiles == null || request.GasFiles.Count == 0)
        {
            return BadRequest(new { error = "At least one GAS file is required" });
        }

        var report = await _deployService.Deploy(request);
        return Ok(report);
    }
}
