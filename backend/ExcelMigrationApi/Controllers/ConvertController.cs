using ExcelMigrationApi.Models;
using ExcelMigrationApi.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace ExcelMigrationApi.Controllers;

[Authorize]
[ApiController]
[Route("api/[controller]")]
public class ConvertController : ControllerBase
{
    private readonly ConvertService _convertService;

    public ConvertController(ConvertService convertService)
    {
        _convertService = convertService;
    }

    /// <summary>
    /// Convert a single VBA module to Google Apps Script.
    /// </summary>
    [HttpPost]
    public async Task<ActionResult<ConvertReport>> Convert([FromBody] List<ConvertRequest> requests)
    {
        if (requests == null || requests.Count == 0)
        {
            return BadRequest(new { error = "No convert requests provided" });
        }

        var report = await _convertService.ConvertBatch(requests);
        return Ok(report);
    }

    /// <summary>
    /// Convert all modules from an ExtractReport (batch conversion).
    /// </summary>
    [HttpPost("batch")]
    public async Task<ActionResult<ConvertReport>> ConvertBatch([FromBody] ExtractReport extractReport)
    {
        if (extractReport == null || extractReport.Modules.Count == 0)
        {
            return BadRequest(new { error = "No modules in extract report" });
        }

        var requests = new List<ConvertRequest>();

        foreach (var module in extractReport.Modules)
        {
            // Skip empty/document modules with no real code
            if (string.IsNullOrWhiteSpace(module.Code) || module.CodeLines <= 1)
                continue;

            // Find button context for this module's source file
            var buttonContext = extractReport.Controls
                .Where(c => c.SourceFile == module.SourceFile && !string.IsNullOrEmpty(c.Macro))
                .ToList();

            requests.Add(new ConvertRequest
            {
                VbaCode = module.Code,
                ModuleName = module.ModuleName,
                ModuleType = module.ModuleType,
                ButtonContext = buttonContext.Count > 0 ? buttonContext : null
            });
        }

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
        return Ok(report);
    }
}
