using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using ExcelMigrationApi.Models;

namespace ExcelMigrationApi.Services;

public class PythonConvertService
{
    private readonly IHttpClientFactory _httpClientFactory;
    private readonly ILogger<PythonConvertService> _logger;
    private readonly string _apiKey;
    private const string AnthropicApiUrl = "https://api.anthropic.com/v1/messages";
    private const string ModelLarge = "claude-sonnet-4-6";
    private const string ModelSmall = "claude-haiku-4-5-20251001";
    private const int SmallModuleThreshold = 50;
    private const int MaxConcurrency = 5;
    private static readonly SemaphoreSlim _concurrencySemaphore = new(MaxConcurrency);

    private const string SystemPrompt = @"You are a specialist in converting Excel VBA macros to Python.
Convert the following VBA module to equivalent Python code.

## Libraries to use
- gspread + google.oauth2.service_account: Google Sheets read/write
- win32com.client: COM automation (Excel, Outlook, Word, etc.)
- os, shutil, subprocess: File I/O, shell execution
- openpyxl: Local Excel file manipulation
- csv: CSV read/write
- glob: File pattern matching
- re: Regular expressions

## VBA → Python Mapping

| VBA | Python |
|---|---|
| Range(""A1"").Value | worksheet.acell(""A1"").value |
| Cells(row, col).Value | worksheet.cell(row, col).value |
| Range(""A1:B10"").Value | worksheet.get(""A1:B10"") |
| Sheets(""name"") | workbook.worksheet(""name"") |
| ActiveSheet | workbook.sheet1 (or specify) |
| ThisWorkbook.Path | os.path.dirname(os.path.abspath(__file__)) |
| Dir(path) | glob.glob(path) or os.listdir(folder) |
| FileCopy src, dst | shutil.copy2(src, dst) |
| Kill path | os.remove(path) |
| MkDir path | os.makedirs(path, exist_ok=True) |
| Open path For Input As #1 | with open(path, 'r', encoding='cp932') as f: |
| Open path For Output As #1 | with open(path, 'w', encoding='cp932') as f: |
| Print #1, text | f.write(text + '\n') |
| Line Input #1, var | var = f.readline().strip() |
| Close #1 | (automatic with 'with' statement) |
| Shell cmd | subprocess.run(cmd, shell=True) |
| CreateObject(""Outlook.Application"") | win32com.client.Dispatch(""Outlook.Application"") |
| CreateObject(""Scripting.FileSystemObject"") | (use os/shutil directly) |
| MsgBox text | print(text) |
| InputBox prompt | input(prompt) |
| Application.ScreenUpdating = False | (delete — not needed) |
| Application.EnableEvents = False | (delete — not needed) |
| On Error Resume Next | try/except: pass |
| On Error GoTo label | try/except |
| DoEvents | (delete — not needed) |
| StrConv(s, vbNarrow) | unicodedata.normalize('NFKC', s) |
| ASC(char) | (use unicodedata.normalize) |

## Google Sheets Integration
When the VBA code reads/writes to worksheets, use gspread:
```python
import gspread
from google.oauth2.service_account import Credentials

def get_spreadsheet():
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_file('credentials.json', scopes=scopes)
    gc = gspread.authorize(creds)
    return gc.open_by_key(SPREADSHEET_ID)
```

## Output Rules
- Use Python 3.10+ syntax
- Use type hints where appropriate
- Output ONLY the .py code, no explanations
- Use encoding='cp932' for Japanese file I/O (Shift-JIS)
- All print output should be in Japanese
- If a VBA function can't be fully converted, add a # TODO comment in Japanese
- CRITICAL: All TODO comments MUST be in Japanese
- Each module should be a standalone .py file with functions
- Do NOT include if __name__ == '__main__' block (main.py handles entry)";

    public PythonConvertService(IConfiguration config, ILogger<PythonConvertService> logger, IHttpClientFactory httpClientFactory)
    {
        _logger = logger;
        _httpClientFactory = httpClientFactory;
        _apiKey = config["ANTHROPIC_API_KEY"]
            ?? Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY")
            ?? string.Empty;
    }

    public async Task<PythonConvertResult> ConvertModule(PythonConvertRequest request)
    {
        if (string.IsNullOrEmpty(_apiKey))
        {
            return new PythonConvertResult
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
                userMessage.AppendLine($"Sheet Name: {request.SheetName}");
            if (!string.IsNullOrEmpty(request.SpreadsheetId))
                userMessage.AppendLine($"Google Spreadsheet ID: {request.SpreadsheetId}");
            userMessage.AppendLine();
            userMessage.AppendLine("VBA Code:");
            userMessage.AppendLine("```vba");
            userMessage.AppendLine(request.VbaCode);
            userMessage.AppendLine("```");

            var model = request.VbaCode.Split('\n').Length <= SmallModuleThreshold ? ModelSmall : ModelLarge;

            var requestBody = new
            {
                model,
                max_tokens = 16384,
                system = SystemPrompt,
                messages = new[]
                {
                    new { role = "user", content = userMessage.ToString() }
                }
            };

            var requestJson = JsonSerializer.Serialize(requestBody);
            var httpRequest = new HttpRequestMessage(HttpMethod.Post, AnthropicApiUrl);
            httpRequest.Content = new StringContent(requestJson, Encoding.UTF8, "application/json");
            httpRequest.Headers.Add("x-api-key", _apiKey);
            httpRequest.Headers.Add("anthropic-version", "2023-06-01");

            var httpClient = _httpClientFactory.CreateClient("Anthropic");
            var response = await httpClient.SendAsync(httpRequest);
            var responseBody = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                _logger.LogError("Anthropic API error {StatusCode} for Python convert {Module}: {Body}",
                    (int)response.StatusCode, request.ModuleName, responseBody);
                return new PythonConvertResult
                {
                    ModuleName = request.ModuleName,
                    SourceFile = request.SourceFile,
                    Status = "error",
                    Error = $"AI変換APIでエラーが発生しました（ステータス: {(int)response.StatusCode}）"
                };
            }

            var apiResponse = JsonSerializer.Deserialize<ClaudeApiResponse>(responseBody);
            var pythonCode = string.Empty;
            var inputTokens = apiResponse?.Usage?.InputTokens ?? 0;
            var outputTokens = apiResponse?.Usage?.OutputTokens ?? 0;

            if (apiResponse?.Content != null)
            {
                foreach (var block in apiResponse.Content)
                {
                    if (block.Type == "text")
                        pythonCode = block.Text ?? string.Empty;
                }
            }

            // Strip markdown code fences if present
            pythonCode = StripCodeFences(pythonCode);

            // Check for truncated output (unclosed brackets/strings)
            if (IsTruncated(pythonCode))
            {
                _logger.LogWarning("Python code for {Module} appears truncated, attempting to fix", request.ModuleName);
                pythonCode = FixTruncatedCode(pythonCode);
            }

            return new PythonConvertResult
            {
                ModuleName = request.ModuleName,
                SourceFile = request.SourceFile,
                PythonCode = pythonCode,
                Status = "success",
                InputTokens = inputTokens,
                OutputTokens = outputTokens
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Python conversion failed for module {Module}", request.ModuleName);
            return new PythonConvertResult
            {
                ModuleName = request.ModuleName,
                SourceFile = request.SourceFile,
                Status = "error",
                Error = $"変換中にエラーが発生しました: {ex.Message}"
            };
        }
    }

    public async Task<PythonConvertReport> ConvertBatch(List<PythonConvertRequest> requests)
    {
        var report = new PythonConvertReport
        {
            GeneratedUtc = DateTime.UtcNow.ToString("o"),
            Total = requests.Count
        };

        var tasks = requests.Select(async request =>
        {
            await _concurrencySemaphore.WaitAsync();
            try
            {
                return await ConvertModule(request);
            }
            finally
            {
                _concurrencySemaphore.Release();
            }
        });

        var results = await Task.WhenAll(tasks);
        report.Results = results.ToList();
        report.Success = results.Count(r => r.Status == "success");
        report.Failed = results.Count(r => r.Status == "error");
        report.TotalInputTokens = results.Sum(r => r.InputTokens);
        report.TotalOutputTokens = results.Sum(r => r.OutputTokens);

        return report;
    }

    private static string StripCodeFences(string code)
    {
        code = code.Trim();
        if (code.StartsWith("```python"))
            code = code["```python".Length..];
        else if (code.StartsWith("```py"))
            code = code["```py".Length..];
        else if (code.StartsWith("```"))
            code = code[3..];

        if (code.EndsWith("```"))
            code = code[..^3];

        return code.Trim();
    }

    /// <summary>
    /// Check if Python code was truncated (incomplete last function).
    /// Simple heuristic: last non-empty line is inside a function body (indented)
    /// and doesn't look like a natural end.
    /// </summary>
    private static bool IsTruncated(string code)
    {
        if (string.IsNullOrEmpty(code)) return false;
        var lines = code.TrimEnd().Split('\n');
        if (lines.Length == 0) return false;

        var lastLine = lines[^1].TrimEnd();
        // If last line is indented and doesn't end with a complete statement, likely truncated
        if (lastLine.Length > 0 && char.IsWhiteSpace(lastLine[0]))
        {
            // Truncated if ends mid-string, mid-dict, mid-call
            if (lastLine.EndsWith(",") || lastLine.EndsWith("(") || lastLine.EndsWith("{") || lastLine.EndsWith("["))
                return true;
        }
        // Check for unclosed triple-quoted strings
        var tripleDouble = System.Text.RegularExpressions.Regex.Matches(code, "\"\"\"").Count;
        var tripleSingle = System.Text.RegularExpressions.Regex.Matches(code, "'''").Count;
        if (tripleDouble % 2 != 0 || tripleSingle % 2 != 0)
            return true;

        return false;
    }

    /// <summary>
    /// Fix truncated Python code by cutting at the last complete top-level function.
    /// </summary>
    private static string FixTruncatedCode(string code)
    {
        var lines = code.Split('\n');

        // Find all top-level function start positions
        var funcStarts = new List<int>();
        for (var i = 0; i < lines.Length; i++)
        {
            if (lines[i].StartsWith("def "))
                funcStarts.Add(i);
        }

        if (funcStarts.Count < 2)
            return code; // Can't fix if only one function

        // Remove the last (incomplete) function — keep up to the start of the last function
        var cutAt = funcStarts[^1];
        var result = string.Join('\n', lines.Take(cutAt)).TrimEnd();
        result += "\n\n# TODO: 最後の関数はAI変換時に途中で切れたため省略されました。元のVBAを確認してください。\n";

        return result;
    }

    // Reuse the same Claude API response models from ConvertService
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
