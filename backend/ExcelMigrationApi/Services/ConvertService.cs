using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using ExcelMigrationApi.Models;

namespace ExcelMigrationApi.Services;

public class ConvertService
{
    private readonly IHttpClientFactory _httpClientFactory;
    private readonly ILogger<ConvertService> _logger;
    private readonly string _apiKey;
    private const string AnthropicApiUrl = "https://api.anthropic.com/v1/messages";
    private const string ModelLarge = "claude-sonnet-4-6";
    private const string ModelSmall = "claude-haiku-4-5-20251001";
    private const int SmallModuleThreshold = 50; // lines
    private const int MaxConcurrency = 5;
    private static readonly SemaphoreSlim _concurrencySemaphore = new(MaxConcurrency);

    private const string SystemPrompt = @"You are a specialist in converting Excel VBA macros to Google Apps Script.
Convert the following VBA module to equivalent Google Apps Script (.gs) code.

## VBA → GAS Event Handler Mapping
When converting Document modules (ThisWorkbook, Sheet modules), apply these mappings:

| VBA Event | GAS Equivalent | Notes |
|---|---|---|
| Workbook_Open / Auto_Open | function onOpen(e) | Simple trigger — runs automatically when spreadsheet opens. Use e.source for SpreadsheetApp reference. |
| Worksheet_Change | function onEdit(e) | Simple trigger. Use e.range, e.value, e.oldValue, e.source. Filter by e.range.getSheet().getName() if the handler is sheet-specific. |
| Worksheet_SelectionChange | installable trigger | Requires: ScriptApp.newTrigger('functionName').forSpreadsheet(SpreadsheetApp.getActive()).onSelectionChange().create() — add a setupTriggers() function. |
| Workbook_BeforeClose | — | No GAS equivalent. Add TODO comment. |
| Workbook_BeforeSave | installable onChange | Approximate with onChange trigger (e.changeType). |
| Worksheet_Activate/Deactivate | — | No GAS equivalent. Add TODO comment suggesting onOpen or custom menu. |
| Workbook_NewSheet | installable onChange | Use onChange with e.changeType == 'INSERT_SHEET'. |
| Worksheet_Calculate | installable onChange | Approximate with onChange trigger. |

## VBA Patterns to Remove or Replace
- Application.ScreenUpdating = True/False → DELETE (GAS has no equivalent, not needed)
- Application.EnableEvents = True/False → DELETE
- Application.DisplayAlerts = True/False → DELETE
- Application.Calculation = xlManual/xlAutomatic → DELETE (use SpreadsheetApp.flush() if needed)
- On Error Resume Next → Use try/catch blocks
- On Error GoTo → Use try/catch blocks
- DoEvents → DELETE (GAS is single-threaded)
- ActiveWorkbook → SpreadsheetApp.getActiveSpreadsheet()
- ActiveSheet → SpreadsheetApp.getActiveSheet()
- ThisWorkbook → SpreadsheetApp.getActiveSpreadsheet()
- Cells(row, col) → sheet.getRange(row, col)
- Range(""A1"") → sheet.getRange(""A1"")
- Sheets(""name"") → ss.getSheetByName(""name"")
- .Value / .Value2 → .getValue() / .getValues()
- MsgBox → SpreadsheetApp.getUi().alert()
- InputBox → SpreadsheetApp.getUi().prompt()

## Installable Triggers
When an event requires an installable trigger, generate a setupTriggers() function:
```javascript
function setupTriggers() {
  // Remove existing triggers to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  // Create new triggers
  ScriptApp.newTrigger('onSelectionChange').forSpreadsheet(SpreadsheetApp.getActive()).onSelectionChange().create();
}
```
And add a comment at the top: // Run setupTriggers() once to install event triggers

## Button/Menu Handling
- If button context is provided with sheet grouping, create a custom menu in onOpen()
- Group menu items by sheet when buttons come from multiple sheets
- If the module already has a Workbook_Open/Auto_Open handler, merge menu creation into the converted onOpen()
- Each button's macro becomes a menu item calling the converted function

## Output Rules
- Use V8 runtime syntax (let, const, arrow functions OK)
- Output ONLY the .gs code, no explanations
- CRITICAL: Ensure all braces {}, parentheses (), and brackets [] are properly closed
- CRITICAL: The output must be syntactically valid JavaScript. Never truncate code.
- If the VBA is too complex to fully convert, output a simplified working stub with TODO comments";

    public ConvertService(IConfiguration config, ILogger<ConvertService> logger, IHttpClientFactory httpClientFactory)
    {
        _logger = logger;
        _httpClientFactory = httpClientFactory;
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
                SourceFile = request.SourceFile,
                Status = "error",
                Error = "サーバー設定エラーが発生しました"
            };
        }

        try
        {
            var userMessage = new StringBuilder();
            userMessage.AppendLine($"Module Name: {request.ModuleName}");
            userMessage.AppendLine($"Module Type: {request.ModuleType}");
            if (!string.IsNullOrEmpty(request.SheetName))
            {
                userMessage.AppendLine($"Sheet Name: {request.SheetName}");
            }
            userMessage.AppendLine();

            // Include detected events context
            if (request.DetectedEvents != null && request.DetectedEvents.Count > 0)
            {
                userMessage.AppendLine("Detected VBA Events in this module:");
                foreach (var evt in request.DetectedEvents)
                {
                    userMessage.AppendLine($"- {evt.VbaEventName} → GAS: {evt.GasTriggerType} | {evt.GasNotes}");
                }
                userMessage.AppendLine();
            }

            userMessage.AppendLine("VBA Code:");
            userMessage.AppendLine("```vba");
            userMessage.AppendLine(request.VbaCode);
            userMessage.AppendLine("```");

            // Button context grouped by sheet
            if (request.ButtonContext != null && request.ButtonContext.Count > 0)
            {
                userMessage.AppendLine();
                userMessage.AppendLine("Button Context (form controls that trigger macros):");

                var buttonsBySheet = request.ButtonContext
                    .GroupBy(b => string.IsNullOrEmpty(b.SheetName) ? "(unknown sheet)" : b.SheetName);

                foreach (var group in buttonsBySheet)
                {
                    userMessage.AppendLine($"  Sheet: {group.Key}");
                    foreach (var btn in group)
                    {
                        userMessage.AppendLine($"    - {btn.ControlType} \"{btn.Label}\" → calls \"{btn.Macro}\"");
                    }
                }

                userMessage.AppendLine();
                userMessage.AppendLine("Generate an onOpen() function that creates a custom menu with these buttons, grouped by sheet if they come from multiple sheets.");

                // Check if this module already has Workbook_Open
                var hasOpenEvent = request.DetectedEvents?.Any(e =>
                    e.VbaEventName.Equals("Workbook_Open", StringComparison.OrdinalIgnoreCase) ||
                    e.VbaEventName.Equals("Auto_Open", StringComparison.OrdinalIgnoreCase)) ?? false;

                if (hasOpenEvent)
                {
                    userMessage.AppendLine("IMPORTANT: This module already contains Workbook_Open/Auto_Open. Merge the menu creation INTO the converted onOpen() function — do NOT create a separate onOpen().");
                }
            }

            var model = request.VbaCode.Split('\n').Length <= SmallModuleThreshold ? ModelSmall : ModelLarge;

            var requestBody = new
            {
                model,
                max_tokens = 8192,
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

            var httpClient = _httpClientFactory.CreateClient("Anthropic");
            var response = await httpClient.SendAsync(httpRequest);
            var responseBody = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                _logger.LogError("Anthropic API error {StatusCode} for module {Module}", (int)response.StatusCode, request.ModuleName);
                return new ConvertResult
                {
                    ModuleName = request.ModuleName,
                    SourceFile = request.SourceFile,
                    Status = "error",
                    Error = $"AI変換APIでエラーが発生しました（ステータス: {(int)response.StatusCode}）"
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
                SourceFile = request.SourceFile,
                GasCode = gasCode,
                Status = "success",
                InputTokens = inputTokens,
                OutputTokens = outputTokens
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "ConvertModule failed for {ModuleName}", request.ModuleName);
            return new ConvertResult
            {
                ModuleName = request.ModuleName,
                SourceFile = request.SourceFile,
                Status = "error",
                Error = "VBA変換中にエラーが発生しました"
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

        // Process in parallel with class-level concurrency limit
        var tasks = requests.Select(async (request, index) =>
        {
            await _concurrencySemaphore.WaitAsync();
            try
            {
                return (Index: index, Result: await ConvertModule(request));
            }
            finally
            {
                _concurrencySemaphore.Release();
            }
        }).ToList();

        var results = await Task.WhenAll(tasks);

        // Add results in original order
        foreach (var item in results.OrderBy(r => r.Index))
        {
            report.Results.Add(item.Result);

            if (item.Result.Status == "success")
                report.Success++;
            else
                report.Failed++;

            report.TotalInputTokens += item.Result.InputTokens;
            report.TotalOutputTokens += item.Result.OutputTokens;
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
