using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using ExcelMigrationApi.Models;

namespace ExcelMigrationApi.Services;

public class DeployService
{
    private readonly HttpClient _httpClient = new();
    private const string ScriptApiBase = "https://script.googleapis.com/v1/projects";

    public async Task<DeployReport> Deploy(DeployRequest request, string googleToken)
    {
        var report = new DeployReport
        {
            GeneratedUtc = DateTime.UtcNow.ToString("o"),
            SpreadsheetId = request.SpreadsheetId,
            FileCount = request.GasFiles.Count
        };

        try
        {
            // Step 1: Create a new Apps Script project bound to the spreadsheet
            var createBody = JsonSerializer.Serialize(new
            {
                title = "Migration Script",
                parentId = request.SpreadsheetId
            });

            var createRes = await SendRequest(HttpMethod.Post, ScriptApiBase, createBody, googleToken);
            if (!createRes.IsSuccess)
            {
                report.Status = "error";
                report.Error = $"Create project failed: {createRes.Body}";
                return report;
            }

            using var createDoc = JsonDocument.Parse(createRes.Body);
            report.ScriptId = createDoc.RootElement.GetProperty("scriptId").GetString() ?? string.Empty;

            // Step 2: Build file list for updateContent
            var files = new List<object>
            {
                // appsscript.json manifest
                new
                {
                    name = "appsscript",
                    type = "JSON",
                    source = JsonSerializer.Serialize(new
                    {
                        timeZone = "Asia/Tokyo",
                        dependencies = new { },
                        runtimeVersion = "V8"
                    }, new JsonSerializerOptions { WriteIndented = true })
                }
            };

            foreach (var gasFile in request.GasFiles)
            {
                var name = gasFile.Name;
                // Remove extension for API (it uses type to determine)
                if (name.EndsWith(".gs", StringComparison.OrdinalIgnoreCase))
                    name = name[..^3];
                else if (name.EndsWith(".html", StringComparison.OrdinalIgnoreCase))
                    name = name[..^5];

                files.Add(new
                {
                    name,
                    type = gasFile.Type?.ToUpperInvariant() == "HTML" ? "HTML" : "SERVER_JS",
                    source = gasFile.Source
                });

                report.FilesDeployed.Add(gasFile.Name);
            }

            // Step 3: Update project content
            var updateBody = JsonSerializer.Serialize(new { files });
            var updateUrl = $"{ScriptApiBase}/{report.ScriptId}/content";
            var updateRes = await SendRequest(HttpMethod.Put, updateUrl, updateBody, googleToken);

            if (!updateRes.IsSuccess)
            {
                report.Status = "error";
                report.Error = $"Update content failed: {updateRes.Body}";
                return report;
            }

            // Step 4: Create a new version
            var versionBody = JsonSerializer.Serialize(new
            {
                description = "Auto-deployed by Excel Migration OS"
            });
            var versionUrl = $"{ScriptApiBase}/{report.ScriptId}/versions";
            var versionRes = await SendRequest(HttpMethod.Post, versionUrl, versionBody, googleToken);

            if (!versionRes.IsSuccess)
            {
                // Content updated but version creation failed - partial success
                report.Status = "partial";
                report.Error = $"Content pushed but version creation failed: {versionRes.Body}";
                return report;
            }

            report.Status = "success";
        }
        catch (Exception ex)
        {
            report.Status = "error";
            report.Error = ex.Message;
        }

        return report;
    }

    private async Task<(bool IsSuccess, string Body)> SendRequest(
        HttpMethod method, string url, string jsonBody, string googleToken)
    {
        var request = new HttpRequestMessage(method, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", googleToken);
        request.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

        var response = await _httpClient.SendAsync(request);
        var body = await response.Content.ReadAsStringAsync();
        return (response.IsSuccessStatusCode, body);
    }
}
