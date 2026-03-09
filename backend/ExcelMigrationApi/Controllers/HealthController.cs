using Microsoft.AspNetCore.Mvc;

namespace ExcelMigrationApi.Controllers;

[ApiController]
[Route("[controller]")]
public class HealthController : ControllerBase
{
    [HttpGet("/health")]
    public IActionResult Health() => Ok(new { status = "healthy", timestamp = DateTime.UtcNow.ToString("o") });
}
