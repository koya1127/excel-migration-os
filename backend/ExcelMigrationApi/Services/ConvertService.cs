using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using ExcelMigrationApi.Models;

namespace ExcelMigrationApi.Services;

public class ConvertService
{
    private readonly HttpClient _httpClient;
    private readonly string _apiKey;
    private const string AnthropicApiUrl = "https://api.anthropic.com/v1/messages";
    private const string Model = "claude-sonnet-4-6";

    private const string SystemPrompt = @"You are a specialist in converting Excel VBA macros to Google Apps Script.
Convert the following VBA module to equivalent Google Apps Script (.gs) code.

Rules:
- Use V8 runtime syntax (let, const, arrow functions OK)
- Replace Excel object model with Google Sheets equivalents
- Replace MsgBox with Browser.msgBox or SpreadsheetApp.getUi().alert()
- Replace InputBox with Browser.inputBox
- If button context is provided, generate an onOpen() function that adds a custom menu
- Output ONLY the .gs code, no explanations";

    public ConvertService(IConfiguration config)
    {
        _httpClient = new HttpClient();
        _apiKey = config["ANTHROPIC_API_KEY"]
            ?? Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY")
            ?? string.Empty;
    }

    public async Task<ConvertResult> ConvertModule(ConvertRequest request)
    {
        if (string.IsNullOrEmpty(_apiKey))
        {
            return new ConvertResult
            {
                ModuleName = request.ModuleName,
                Status = "error",
                Error = "ANTHROPIC_API_KEY is not configured"
            };
        }

        try
        {
            var userMessage = new StringBuilder();
            userMessage.AppendLine($"Module Name: {request.ModuleName}");
            userMessage.AppendLine($"Module Type: {request.ModuleType}");
            userMessage.AppendLine();
            userMessage.AppendLine("VBA Code:");
            userMessage.AppendLine("```vba");
            userMessage.AppendLine(request.VbaCode);
            userMessage.AppendLine("```");

            if (request.ButtonContext != null && request.ButtonContext.Count > 0)
            {
                userMessage.AppendLine();
                userMessage.AppendLine("Button Context (form controls that trigger macros):");
                foreach (var btn in request.ButtonContext)
                {
                    userMessage.AppendLine($"- Button \"{btn.Label}\" calls macro \"{btn.Macro}\"");
                }
                userMessage.AppendLine();
                userMessage.AppendLine("Generate an onOpen() function that creates a custom menu with these buttons.");
            }

            var requestBody = new
            {
                model = Model,
                max_tokens = 4096,
                system = SystemPrompt,
                messages = new[]
                {
                    new { role = "user", content = userMessage.ToString() }
                }
            };

            var json = JsonSerializer.Serialize(requestBody);
            var httpRequest = new HttpRequestMessage(HttpMethod.Post, AnthropicApiUrl);
            httpRequest.Content = new StringContent(json, Encoding.UTF8, "application/json");
            httpRequest.Headers.Add("x-api-key", _apiKey);
            httpRequest.Headers.Add("anthropic-version", "2023-06-01");

            var response = await _httpClient.SendAsync(httpRequest);
            var responseBody = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                return new ConvertResult
                {
                    ModuleName = request.ModuleName,
                    Status = "error",
                    Error = $"API error {(int)response.StatusCode}: {responseBody}"
                };
            }

            // Parse Claude API response
            var apiResponse = JsonSerializer.Deserialize<ClaudeApiResponse>(responseBody);
            var gasCode = string.Empty;
            var inputTokens = apiResponse?.Usage?.InputTokens ?? 0;
            var outputTokens = apiResponse?.Usage?.OutputTokens ?? 0;

            if (apiResponse?.Content != null)
            {
                foreach (var block in apiResponse.Content)
                {
                    if (block.Type == "text")
                    {
                        gasCode = block.Text ?? string.Empty;
                    }
                }
            }

            // Strip markdown code fences if present
            gasCode = StripCodeFences(gasCode);

            return new ConvertResult
            {
                ModuleName = request.ModuleName,
                GasCode = gasCode,
                Status = "success",
                InputTokens = inputTokens,
                OutputTokens = outputTokens
            };
        }
        catch (Exception ex)
        {
            return new ConvertResult
            {
                ModuleName = request.ModuleName,
                Status = "error",
                Error = ex.Message
            };
        }
    }

    public async Task<ConvertReport> ConvertBatch(List<ConvertRequest> requests)
    {
        var report = new ConvertReport
        {
            GeneratedUtc = DateTime.UtcNow.ToString("o"),
            Total = requests.Count
        };

        foreach (var request in requests)
        {
            var result = await ConvertModule(request);
            report.Results.Add(result);

            if (result.Status == "success")
                report.Success++;
            else
                report.Failed++;

            report.TotalInputTokens += result.InputTokens;
            report.TotalOutputTokens += result.OutputTokens;
        }

        return report;
    }

    private static string StripCodeFences(string code)
    {
        if (string.IsNullOrEmpty(code)) return code;

        var lines = code.Split('\n').ToList();

        // Remove leading ```javascript or ```gs
        if (lines.Count > 0 && lines[0].TrimStart().StartsWith("```"))
        {
            lines.RemoveAt(0);
        }

        // Remove trailing ```
        if (lines.Count > 0 && lines[^1].Trim() == "```")
        {
            lines.RemoveAt(lines.Count - 1);
        }

        return string.Join('\n', lines).Trim();
    }

    // Internal classes for Claude API response deserialization
    private class ClaudeApiResponse
    {
        [JsonPropertyName("content")]
        public List<ContentBlock>? Content { get; set; }

        [JsonPropertyName("usage")]
        public UsageInfo? Usage { get; set; }
    }

    private class ContentBlock
    {
        [JsonPropertyName("type")]
        public string Type { get; set; } = string.Empty;

        [JsonPropertyName("text")]
        public string? Text { get; set; }
    }

    private class UsageInfo
    {
        [JsonPropertyName("input_tokens")]
        public int InputTokens { get; set; }

        [JsonPropertyName("output_tokens")]
        public int OutputTokens { get; set; }
    }
}
