using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using ExcelMigrationApi.Models;

namespace ExcelMigrationApi.Services;

public class DeployService
{
    private readonly IHttpClientFactory _httpClientFactory;
    private readonly ILogger<DeployService> _logger;
    private const string ScriptApiBase = "https://script.googleapis.com/v1/projects";
    private const string SheetsApiBase = "https://sheets.googleapis.com/v4/spreadsheets";
    private const string DriveApiBase = "https://www.googleapis.com/drive/v3/files";

    public DeployService(ILogger<DeployService> logger, IHttpClientFactory httpClientFactory)
    {
        _logger = logger;
        _httpClientFactory = httpClientFactory;
    }

    /// <summary>
    /// Create a new empty spreadsheet, write a "使い方" sheet, then deploy GAS.
    /// </summary>
    public async Task<DeployReport> CreateAndDeploy(
        string originalFileName, List<GasFile> gasFiles, UsageSheetInfo usageSheet,
        string? folderId, string googleToken)
    {
        var title = Path.GetFileNameWithoutExtension(originalFileName) + "（移行済み）";

        // Step 1: Create empty spreadsheet
        var createBody = JsonSerializer.Serialize(new { properties = new { title } });
        var createRes = await SendRequest(HttpMethod.Post, SheetsApiBase, createBody, googleToken);
        if (!createRes.IsSuccess)
        {
            _logger.LogError("Failed to create spreadsheet: {Body}", createRes.Body);
            return new DeployReport
            {
                GeneratedUtc = DateTime.UtcNow.ToString("o"),
                Status = "error",
                Error = "スプレッドシートの作成に失敗しました"
            };
        }

        using var ssDoc = JsonDocument.Parse(createRes.Body);
        var spreadsheetId = ssDoc.RootElement.GetProperty("spreadsheetId").GetString() ?? "";
        var webViewLink = ssDoc.RootElement.GetProperty("spreadsheetUrl").GetString() ?? "";

        // Step 2: Move to folder if specified
        if (!string.IsNullOrEmpty(folderId))
        {
            var moveUrl = $"{DriveApiBase}/{spreadsheetId}?addParents={folderId}&removeParents=root";
            await SendRequest(HttpMethod.Patch, moveUrl, "{}", googleToken);
        }

        // Step 3: Write "使い方" sheet (delete default Sheet1 since this is a new empty spreadsheet)
        await WriteUsageSheet(spreadsheetId, usageSheet, googleToken, deleteDefault: true);

        // Step 4: Deploy GAS
        var deployRequest = new DeployRequest
        {
            SpreadsheetId = spreadsheetId,
            GasFiles = gasFiles
        };
        var report = await Deploy(deployRequest, googleToken);
        report.WebViewLink = webViewLink;
        return report;
    }

    /// <summary>
    /// Add a "使い方" sheet to an existing spreadsheet (does not delete other sheets).
    /// </summary>
    public async Task AddUsageSheet(string spreadsheetId, UsageSheetInfo info, string googleToken)
    {
        await WriteUsageSheet(spreadsheetId, info, googleToken);
    }

    private async Task WriteUsageSheet(string spreadsheetId, UsageSheetInfo info, string googleToken, bool deleteDefault = false)
    {
        try
        {
            // Add "使い方" sheet
            var requests = new List<object>
            {
                new { addSheet = new { properties = new { title = "使い方", index = 0 } } }
            };

            if (deleteDefault)
            {
                // Get the default sheet ID to delete (only for empty spreadsheets)
                var getRes = await SendRequest(HttpMethod.Get, $"{SheetsApiBase}/{spreadsheetId}?fields=sheets.properties", "", googleToken);
                if (getRes.IsSuccess)
                {
                    using var doc = JsonDocument.Parse(getRes.Body);
                    var defaultSheetId = doc.RootElement.GetProperty("sheets")[0]
                        .GetProperty("properties").GetProperty("sheetId").GetInt32();
                    requests.Add(new { deleteSheet = new { sheetId = defaultSheetId } });
                }
            }

            var batchBody = JsonSerializer.Serialize(new { requests });
            await SendRequest(HttpMethod.Post, $"{SheetsApiBase}/{spreadsheetId}:batchUpdate", batchBody, googleToken);

            // Build content rows
            var rows = new List<IList<object>>();
            rows.Add(new object[] { "📋 このスプレッドシートについて" });
            rows.Add(new object[] { "" });
            rows.Add(new object[] { $"元ファイル: {info.OriginalFileName}" });
            rows.Add(new object[] { $"移行日: {DateTime.Now:yyyy/MM/dd}" });
            rows.Add(new object[] { $"移行ツール: Excel Migration OS" });
            rows.Add(new object[] { "" });
            rows.Add(new object[] { "📖 マクロの使い方" });
            rows.Add(new object[] { "" });
            rows.Add(new object[] { "1. スプレッドシートを開くと、メニューバーに「カスタムメニュー」が自動で追加されます" });
            rows.Add(new object[] { "2. メニューから実行したい機能を選んでクリックしてください" });
            rows.Add(new object[] { "3. 初回実行時に「承認が必要です」と表示されたら「許可」を選んでください" });
            rows.Add(new object[] { "" });

            if (info.MenuItems.Count > 0)
            {
                rows.Add(new object[] { "📌 メニュー項目一覧" });
                rows.Add(new object[] { "" });
                foreach (var item in info.MenuItems)
                {
                    rows.Add(new object[] { $"・{item}" });
                }
                rows.Add(new object[] { "" });
            }

            if (info.Limitations.Count > 0)
            {
                rows.Add(new object[] { "⚠️ 注意事項（自動変換で対応しきれなかった箇所）" });
                rows.Add(new object[] { "" });
                foreach (var limitation in info.Limitations)
                {
                    rows.Add(new object[] { $"・{limitation}" });
                }
                rows.Add(new object[] { "" });
            }

            rows.Add(new object[] { "❓ 困ったときは" });
            rows.Add(new object[] { "" });
            rows.Add(new object[] { "・メニューが表示されない → ページを再読み込み（F5）してください" });
            rows.Add(new object[] { "・「承認が必要」が何度も出る → メニューの「拡張機能」→「Apps Script」を開き、関数を手動で一度実行してください" });
            rows.Add(new object[] { "・エラーが出る → 上の「注意事項」に記載の未対応箇所が原因の可能性があります" });

            // Write values
            var valueBody = JsonSerializer.Serialize(new { values = rows });
            var url = $"{SheetsApiBase}/{spreadsheetId}/values/使い方!A1?valueInputOption=RAW";
            await SendRequest(HttpMethod.Put, url, valueBody, googleToken);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to write usage sheet for {SpreadsheetId}", spreadsheetId);
            // Non-fatal: GAS deploy can still proceed
        }
    }

    /// <summary>
    /// Extract menu items and TODO/limitations from GAS code.
    /// </summary>
    public static UsageSheetInfo BuildUsageSheet(string originalFileName, IEnumerable<ConvertResult> results)
    {
        var info = new UsageSheetInfo { OriginalFileName = originalFileName };
        var menuRegex = new Regex(@"\.addItem\(['""](.+?)['""]", RegexOptions.Compiled);
        var todoRegex = new Regex(@"//\s*TODO:?\s*(.*)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        foreach (var result in results)
        {
            if (string.IsNullOrEmpty(result.GasCode)) continue;

            // Extract menu items (deduplicate)
            foreach (Match m in menuRegex.Matches(result.GasCode))
            {
                var item = m.Groups[1].Value;
                if (!info.MenuItems.Contains(item))
                {
                    info.MenuItems.Add(item);
                }
            }

            // Extract TODO/limitations
            foreach (Match m in todoRegex.Matches(result.GasCode))
            {
                var text = m.Groups[1].Value.Trim();
                if (!string.IsNullOrEmpty(text) && !info.Limitations.Contains(text))
                {
                    info.Limitations.Add(text);
                }
            }
        }

        return info;
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
        if (method != HttpMethod.Get && !string.IsNullOrEmpty(jsonBody))
        {
            request.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");
        }

        var httpClient = _httpClientFactory.CreateClient("AppsScript");
        var response = await httpClient.SendAsync(request);
        var body = await response.Content.ReadAsStringAsync();
        return (response.IsSuccessStatusCode, body);
    }
}
