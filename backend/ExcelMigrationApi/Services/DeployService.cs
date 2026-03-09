using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using ExcelMigrationApi.Models;

namespace ExcelMigrationApi.Services;

public class DeployService
{
    private readonly IHttpClientFactory _httpClientFactory;
    private readonly ILogger<DeployService> _logger;
    private const string ScriptApiBase = "https://script.googleapis.com/v1/projects";

    public DeployService(ILogger<DeployService> logger, IHttpClientFactory httpClientFactory)
    {
        _logger = logger;
        _httpClientFactory = httpClientFactory;
    }

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
                _logger.LogError("Create Apps Script project failed for spreadsheet {SpreadsheetId}", request.SpreadsheetId);
                report.Status = "error";
                report.Error = "Apps Scriptプロジェクトの作成に失敗しました";
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

            // Step 3: Update project content (retry with problematic files excluded)
            var updateUrl = $"{ScriptApiBase}/{report.ScriptId}/content";
            var filesToDeploy = new List<object>(files);
            var excludedFiles = new List<string>();
            var maxRetries = 5;

            for (var attempt = 0; attempt <= maxRetries; attempt++)
            {
                var updateBody = JsonSerializer.Serialize(new { files = filesToDeploy });
                var updateRes = await SendRequest(HttpMethod.Put, updateUrl, updateBody, googleToken);

                if (updateRes.IsSuccess)
                    break;

                // Try to extract the problematic file name from error
                var errorFileName = ExtractErrorFileName(updateRes.Body);

                if (errorFileName == null || attempt == maxRetries)
                {
                    // Can't identify the file or max retries reached
                    if (excludedFiles.Count > 0)
                    {
                        // Partial success - some files were excluded
                        _logger.LogWarning("Partial deploy: {Count} files excluded for spreadsheet {SpreadsheetId}", excludedFiles.Count, request.SpreadsheetId);
                        report.Status = "partial";
                        report.Error = $"{excludedFiles.Count} ファイルを構文エラーのため除外: {string.Join(", ", excludedFiles)}";
                        return report;
                    }
                    _logger.LogError("Deploy update content failed for spreadsheet {SpreadsheetId}", request.SpreadsheetId);
                    report.Status = "error";
                    report.Error = "GASコードのデプロイに失敗しました";
                    return report;
                }

                // Remove the problematic file and retry
                var nameWithoutExt = errorFileName.EndsWith(".gs") ? errorFileName[..^3] : errorFileName;
                filesToDeploy.RemoveAll(f =>
                {
                    var json = JsonSerializer.Serialize(f);
                    using var doc = JsonDocument.Parse(json);
                    var name = doc.RootElement.GetProperty("name").GetString();
                    return name == nameWithoutExt;
                });
                excludedFiles.Add(errorFileName);
                report.FilesDeployed.Remove(errorFileName);
            }

            if (excludedFiles.Count > 0)
            {
                report.Error = $"{excludedFiles.Count} ファイルを構文エラーのため除外: {string.Join(", ", excludedFiles)}";
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
                _logger.LogWarning("Version creation failed for script {ScriptId}", report.ScriptId);
                report.Status = "partial";
                report.Error = "コードは更新されましたが、バージョンの作成に失敗しました";
                return report;
            }

            report.Status = "success";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Deploy failed for spreadsheet {SpreadsheetId}", request.SpreadsheetId);
            report.Status = "error";
            report.Error = "GASデプロイ中にエラーが発生しました";
        }

        return report;
    }

    private static string? ExtractErrorFileName(string errorBody)
    {
        // Error format: "... file: Module1.gs"
        try
        {
            var match = System.Text.RegularExpressions.Regex.Match(errorBody, @"file:\s*(\S+\.gs)");
            if (match.Success)
                return match.Groups[1].Value;
        }
        catch { }
        return null;
    }

    private async Task<(bool IsSuccess, string Body)> SendRequest(
        HttpMethod method, string url, string jsonBody, string googleToken)
    {
        var request = new HttpRequestMessage(method, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", googleToken);
        request.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

        var httpClient = _httpClientFactory.CreateClient("AppsScript");
        var response = await httpClient.SendAsync(request);
        var body = await response.Content.ReadAsStringAsync();
        return (response.IsSuccessStatusCode, body);
    }
}
