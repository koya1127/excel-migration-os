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
            rows.Add(new object[] { "このスプレッドシートについて" });
            rows.Add(new object[] { "" });
            rows.Add(new object[] { $"元ファイル: {info.OriginalFileName}" });
            rows.Add(new object[] { $"移行日: {DateTime.Now:yyyy/MM/dd}" });
            rows.Add(new object[] { "" });
            rows.Add(new object[] { "マクロの使い方" });
            rows.Add(new object[] { "" });
            rows.Add(new object[] { "手順1. このスプレッドシートを開くと、画面上部のメニューバー（「ファイル」「編集」...の並び）の右端に" });
            rows.Add(new object[] { "　　　 マクロ用のメニューが自動で追加されます（数秒かかる場合があります）" });
            rows.Add(new object[] { "" });
            rows.Add(new object[] { "手順2. 追加されたメニューをクリックし、実行したい機能を選んでください" });
            rows.Add(new object[] { "" });
            rows.Add(new object[] { "手順3.【初回のみ】「承認が必要です」という画面が出ます。以下の手順で許可してください：" });
            rows.Add(new object[] { "　　　 ① 「権限を確認」をクリック" });
            rows.Add(new object[] { "　　　 ② Googleアカウントを選択" });
            rows.Add(new object[] { "　　　 ③ 「詳細」→「○○（安全ではないページ）に移動」をクリック" });
            rows.Add(new object[] { "　　　 ④ 「許可」をクリック" });
            rows.Add(new object[] { "　　　 ※ これは移行したマクロにスプレッドシートへのアクセスを許可するためです" });
            rows.Add(new object[] { "　　　 ※ 一度許可すれば、次回以降は表示されません" });
            rows.Add(new object[] { "" });

            if (info.MenuItems.Count > 0)
            {
                rows.Add(new object[] { "使える機能の一覧" });
                rows.Add(new object[] { "" });
                foreach (var item in info.MenuItems)
                {
                    rows.Add(new object[] { $"　・{item}" });
                }
                rows.Add(new object[] { "" });
            }

            if (info.Limitations.Count > 0)
            {
                rows.Add(new object[] { "ご注意（一部、自動変換できなかった機能があります）" });
                rows.Add(new object[] { "" });
                foreach (var limitation in info.Limitations)
                {
                    rows.Add(new object[] { $"　・{limitation}" });
                }
                rows.Add(new object[] { "" });
            }

            rows.Add(new object[] { "困ったときは" });
            rows.Add(new object[] { "" });
            rows.Add(new object[] { "　・メニューが出てこない" });
            rows.Add(new object[] { "　　→ ページを再読み込み（キーボードの F5 キー）してください。数秒待つと表示されます" });
            rows.Add(new object[] { "" });
            rows.Add(new object[] { "　・「承認が必要です」が何度も出る" });
            rows.Add(new object[] { "　　→ 上部メニュー「拡張機能」→「Apps Script」を開いてください" });
            rows.Add(new object[] { "　　→ 表示されたコードの上にある「▶ 実行」ボタンを押して、画面の指示に従って許可してください" });
            rows.Add(new object[] { "" });
            rows.Add(new object[] { "　・機能を使ったらエラーが出た" });
            rows.Add(new object[] { "　　→ 上の「ご注意」に書かれている機能は、完全には移行できていない可能性があります" });
            rows.Add(new object[] { "　　→ 元のExcelファイルで同じ操作を試し、結果を比較してみてください" });

            // Write values
            var valueBody = JsonSerializer.Serialize(new { values = rows });
            var url = $"{SheetsApiBase}/{spreadsheetId}/values/使い方!A1?valueInputOption=RAW";
            await SendRequest(HttpMethod.Put, url, valueBody, googleToken);

            // Format the sheet: widen column A, bold headers, background colors
            await FormatUsageSheet(spreadsheetId, rows.Count, googleToken);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to write usage sheet for {SpreadsheetId}", spreadsheetId);
            // Non-fatal: GAS deploy can still proceed
        }
    }

    /// <summary>
    /// Format the usage sheet: widen column A, bold section headers, light background colors.
    /// </summary>
    private async Task FormatUsageSheet(string spreadsheetId, int totalRows, string googleToken)
    {
        try
        {
            // First, get the sheet ID for "使い方"
            var getRes = await SendRequest(HttpMethod.Get,
                $"{SheetsApiBase}/{spreadsheetId}?fields=sheets.properties", "", googleToken);
            if (!getRes.IsSuccess) return;

            int sheetId = 0;
            using (var doc = JsonDocument.Parse(getRes.Body))
            {
                foreach (var sheet in doc.RootElement.GetProperty("sheets").EnumerateArray())
                {
                    var title = sheet.GetProperty("properties").GetProperty("title").GetString();
                    if (title == "使い方")
                    {
                        sheetId = sheet.GetProperty("properties").GetProperty("sheetId").GetInt32();
                        break;
                    }
                }
            }

            var formatRequests = new List<object>
            {
                // Widen column A to 800px so text doesn't overflow
                new
                {
                    updateDimensionProperties = new
                    {
                        range = new { sheetId, dimension = "COLUMNS", startIndex = 0, endIndex = 1 },
                        properties = new { pixelSize = 800 },
                        fields = "pixelSize"
                    }
                },
                // Bold + larger font for title row (row 0)
                new
                {
                    repeatCell = new
                    {
                        range = new { sheetId, startRowIndex = 0, endRowIndex = 1, startColumnIndex = 0, endColumnIndex = 1 },
                        cell = new
                        {
                            userEnteredFormat = new
                            {
                                textFormat = new { bold = true, fontSize = 14 },
                                backgroundColor = new { red = 0.9, green = 0.94, blue = 1.0 }
                            }
                        },
                        fields = "userEnteredFormat(textFormat,backgroundColor)"
                    }
                },
                // Freeze top row
                new
                {
                    updateSheetProperties = new
                    {
                        properties = new { sheetId, gridProperties = new { frozenRowCount = 1 } },
                        fields = "gridProperties.frozenRowCount"
                    }
                }
            };

            var batchBody = JsonSerializer.Serialize(new { requests = formatRequests });
            await SendRequest(HttpMethod.Post, $"{SheetsApiBase}/{spreadsheetId}:batchUpdate", batchBody, googleToken);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to format usage sheet (non-fatal)");
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
        var rawMenuItems = new List<string>();
        var rawLimitations = new List<string>();

        foreach (var result in results)
        {
            if (string.IsNullOrEmpty(result.GasCode)) continue;

            foreach (Match m in menuRegex.Matches(result.GasCode))
                rawMenuItems.Add(m.Groups[1].Value.Trim());

            foreach (Match m in todoRegex.Matches(result.GasCode))
            {
                var text = m.Groups[1].Value.Trim();
                if (!string.IsNullOrEmpty(text))
                    rawLimitations.Add(text);
            }
        }

        // Deduplicate menu items: normalize by removing brackets/parens/prefixes
        info.MenuItems = DeduplicateMenuItems(rawMenuItems);

        // Deduplicate limitations: group similar items, translate common English patterns to Japanese
        info.Limitations = DeduplicateLimitations(rawLimitations);

        return info;
    }

    /// <summary>
    /// Aggressively deduplicate menu items that represent the same function.
    /// Strategy: extract "core action" (Japanese text before any brackets/parens),
    /// group by core, keep only the most descriptive unique variant per core.
    /// </summary>
    private static List<string> DeduplicateMenuItems(List<string> items)
    {
        var result = new List<string>();
        var seenCores = new HashSet<string>();
        // Track full normalized forms to catch exact dupes
        var seenFull = new HashSet<string>();

        foreach (var item in items)
        {
            // Skip items that look like internal function references (e.g. "search(部材・部品) - 検索")
            if (Regex.IsMatch(item, @"^[a-zA-Z]+\(.+\)\s*-\s*"))
                continue;

            // Extract core action: Japanese text before first bracket/paren/dash, trimmed
            var core = Regex.Replace(item, @"[\s　]*[\[\(（\-].*$", "").Trim();
            if (string.IsNullOrEmpty(core)) core = item;

            // Full normalized: strip everything non-essential for exact dupe check
            var fullNorm = Regex.Replace(item, @"[\[\]\(\)（）\s　\""\'\-]", "").ToLowerInvariant();
            // Remove "search" function name prefix that leaks from code
            fullNorm = Regex.Replace(fullNorm, @"search", "");

            if (seenFull.Contains(fullNorm))
                continue;
            seenFull.Add(fullNorm);

            // If this core action was already seen, skip (e.g. "検索" already added, skip "検索 (部材・部品)")
            if (seenCores.Contains(core))
                continue;

            seenCores.Add(core);
            result.Add(item);
        }

        return result;
    }

    /// <summary>
    /// Deduplicate and translate TODO/limitation items to Japanese.
    /// </summary>
    private static List<string> DeduplicateLimitations(List<string> items)
    {
        // Common English→Japanese translations for TODO comments
        var translations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "implement search logic", "検索機能の実装が必要です" },
            { "implement search functionality", "検索機能の実装が必要です" },
            { "implement the search functionality", "検索機能の実装が必要です" },
            { "implement go back logic", "「戻る」機能の実装が必要です" },
            { "implement go back functionality", "「戻る」機能の実装が必要です" },
            { "implement the go back functionality", "「戻る」機能の実装が必要です" },
            { "implement back navigation", "「戻る」機能の実装が必要です" },
            { "implement csv download logic", "CSVダウンロード機能の実装が必要です" },
            { "implement csv download functionality", "CSVダウンロード機能の実装が必要です" },
            { "implement the csv download functionality", "CSVダウンロード機能の実装が必要です" },
            { "implement hsearch function", "検索機能の実装が必要です" },
            { "implement goback function", "「戻る」機能の実装が必要です" },
            { "implement csvdownload function", "CSVダウンロード機能の実装が必要です" },
        };

        // Regex-based translations for patterns
        var patternTranslations = new (Regex Pattern, string Japanese)[]
        {
            (new Regex(@"is not (available|supported) in (Google Apps Script|GAS)", RegexOptions.IgnoreCase),
                "はGoogle Apps Scriptでは利用できません"),
            (new Regex(@"No GAS equivalent", RegexOptions.IgnoreCase),
                "Google Apps Scriptに対応する機能がありません"),
            (new Regex(@"Implement actual (.+)", RegexOptions.IgnoreCase),
                "の実装が必要です"),
            (new Regex(@"Replace with actual (.+)", RegexOptions.IgnoreCase),
                "の実装が必要です"),
            (new Regex(@"Update this (.+)", RegexOptions.IgnoreCase),
                "の更新が必要です"),
            (new Regex(@"Verify that (.+)", RegexOptions.IgnoreCase),
                "の確認が必要です"),
            (new Regex(@"Consider alternative", RegexOptions.IgnoreCase),
                "代替手段の検討が必要です"),
        };

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var result = new List<string>();

        foreach (var item in items)
        {
            var translated = item;

            // Try exact translation first
            if (translations.TryGetValue(item.TrimEnd('.'), out var exact))
            {
                translated = exact;
            }
            else
            {
                // Try pattern-based translation
                foreach (var (pattern, japanese) in patternTranslations)
                {
                    if (pattern.IsMatch(item))
                    {
                        // Extract the subject (e.g. "Worksheet_FollowHyperlink")
                        var subject = Regex.Replace(item, @"\s*(is not available|is not supported|is unsupported).*$", "", RegexOptions.IgnoreCase).Trim();
                        subject = Regex.Replace(subject, @"^(Implement actual|Replace with actual|Consider alternative:?\s*)", "", RegexOptions.IgnoreCase).Trim();

                        if (!string.IsNullOrEmpty(subject) && subject != item)
                            translated = $"{subject} {japanese}";
                        else
                            translated = japanese;
                        break;
                    }
                }
            }

            // Final pass: strip technical jargon for end users
            translated = SimplifyForEndUser(translated);

            // Normalize for dedup: lowercase, remove punctuation
            var normalized = Regex.Replace(translated.ToLowerInvariant(), @"[。、．\.\s]", "");

            if (seen.Contains(normalized))
                continue;

            seen.Add(normalized);
            result.Add(translated);
        }

        // Cap at 10 items to avoid overwhelming non-technical users
        if (result.Count > 10)
        {
            var truncated = result.Take(10).ToList();
            truncated.Add($"（他 {result.Count - 10} 件）");
            return truncated;
        }

        return result;
    }

    /// <summary>
    /// Strip technical jargon from limitation text so non-technical users can understand.
    /// </summary>
    private static string SimplifyForEndUser(string text)
    {
        // Remove VBA/GAS technical terms that mean nothing to end users
        text = Regex.Replace(text, @"Workbook_BeforeClose\S*", "「閉じる前の処理」");
        text = Regex.Replace(text, @"Worksheet_FollowHyperlink\S*", "「リンクをクリックした時の処理」");
        text = Regex.Replace(text, @"Worksheet_SelectionChange\S*", "「セル選択時の処理」");
        text = Regex.Replace(text, @"Workbook_BeforeSave\S*", "「保存前の処理」");
        text = Regex.Replace(text, @"Worksheet_Activate\S*", "「シート切替時の処理」");
        text = Regex.Replace(text, @"Worksheet_Calculate\S*", "「再計算時の処理」");

        // Replace function names with descriptive Japanese
        text = Regex.Replace(text, @"\bhSearch\b", "検索", RegexOptions.IgnoreCase);
        text = Regex.Replace(text, @"\bgoBack\b", "「戻る」", RegexOptions.IgnoreCase);
        text = Regex.Replace(text, @"\bcsvDownload\b", "CSVダウンロード", RegexOptions.IgnoreCase);
        text = Regex.Replace(text, @"\bsearchRng\b", "検索範囲", RegexOptions.IgnoreCase);

        // Replace technical GAS/VBA terms
        text = Regex.Replace(text, @"Google Apps Script", "スプレッドシートのマクロ");
        text = Regex.Replace(text, @"\bGAS\b", "スプレッドシートのマクロ");
        text = Regex.Replace(text, @"\bVBA\b", "Excelマクロ");
        text = Regex.Replace(text, @"ControlFormat", "フォームの部品");
        text = Regex.Replace(text, @"Shapes", "図形・ボタン");
        text = Regex.Replace(text, @"Radio button", "ラジオボタン（選択肢）");
        text = Regex.Replace(text, @"WScript\.Network", "Windowsのネットワーク機能");
        text = Regex.Replace(text, @"Local file I/O.*$", "パソコン上のファイル読み書き機能");
        text = Regex.Replace(text, @"Dir\(.*?\)", "フォルダ内のファイル一覧取得");
        text = Regex.Replace(text, @"FileCopy", "ファイルコピー");
        text = Regex.Replace(text, @"StrConv\(.*?\)", "文字の全角半角変換");
        text = Regex.Replace(text, @"ThisWorkbook\.Path", "ファイルの保存場所");
        text = Regex.Replace(text, @"DriveApp", "Googleドライブ");
        text = Regex.Replace(text, @"Module1\.", "");
        text = Regex.Replace(text, @"cell-based selection", "セル選択方式");
        text = Regex.Replace(text, @"onOpen", "スプレッドシート起動時の処理");
        text = Regex.Replace(text, @"onChange trigger", "変更検知の仕組み");
        text = Regex.Replace(text, @"installable trigger", "自動実行の仕組み");

        // Clean up double spaces / leading dots
        text = Regex.Replace(text, @"\s{2,}", " ").Trim();

        return text;
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
