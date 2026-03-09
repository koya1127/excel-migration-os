using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using ExcelMigrationApi.Models;

namespace ExcelMigrationApi.Services;

public class UploadService
{
    private readonly HttpClient _httpClient = new() { Timeout = TimeSpan.FromMinutes(2) };
    private readonly ILogger<UploadService> _logger;

    public UploadService(ILogger<UploadService> logger)
    {
        _logger = logger;
    }

    public async Task<UploadResult> UploadFile(string filePath, bool convertToSheets, string? folderId, string googleToken)
    {
        var fileName = Path.GetFileName(filePath);

        try
        {
            // Build Drive API metadata
            var metadata = new Dictionary<string, object>
            {
                // When converting to Sheets, use name without extension (Google adds its own).
                // When keeping original format, preserve the full filename with extension.
                ["name"] = convertToSheets ? Path.GetFileNameWithoutExtension(fileName) : fileName
            };

            if (convertToSheets)
            {
                metadata["mimeType"] = "application/vnd.google-apps.spreadsheet";
            }

            if (!string.IsNullOrEmpty(folderId))
            {
                metadata["parents"] = new[] { folderId };
            }

            var metadataJson = JsonSerializer.Serialize(metadata);

            // Determine upload MIME type
            var ext = Path.GetExtension(fileName).ToLowerInvariant();
            var mimeType = ext switch
            {
                ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ".xlsm" => "application/vnd.ms-excel.sheet.macroEnabled.12",
                ".xls" => "application/vnd.ms-excel",
                ".csv" => "text/csv",
                _ => "application/octet-stream"
            };

            // Use multipart upload to Google Drive API (streaming to avoid loading entire file into memory)
            using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 81920);

            using var content = new MultipartContent("related");

            // Part 1: metadata
            var metadataPart = new StringContent(metadataJson, Encoding.UTF8, "application/json");
            content.Add(metadataPart);

            // Part 2: file content (streamed)
            var filePart = new StreamContent(fileStream);
            filePart.Headers.ContentType = new MediaTypeHeaderValue(mimeType);
            content.Add(filePart);

            var request = new HttpRequestMessage(HttpMethod.Post,
                "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,webViewLink");
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", googleToken);
            request.Content = content;

            var response = await _httpClient.SendAsync(request);
            var responseBody = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                _logger.LogError("Drive API error {StatusCode} for file {FileName}", (int)response.StatusCode, fileName);
                return new UploadResult
                {
                    FileName = fileName,
                    Status = "error",
                    Error = $"Google Driveへのアップロードに失敗しました（ステータス: {(int)response.StatusCode}）"
                };
            }

            var driveFileId = string.Empty;
            var webViewLink = string.Empty;

            using var doc = JsonDocument.Parse(responseBody);
            var root = doc.RootElement;
            if (root.TryGetProperty("id", out var idProp))
                driveFileId = idProp.GetString() ?? string.Empty;
            if (root.TryGetProperty("webViewLink", out var linkProp))
                webViewLink = linkProp.GetString() ?? string.Empty;

            return new UploadResult
            {
                FileName = fileName,
                DriveFileId = driveFileId,
                WebViewLink = webViewLink,
                Status = "success"
            };
        }
        catch (Exception ex)
        {
            return new UploadResult
            {
                FileName = fileName,
                Status = "error",
                Error = ex.Message
            };
        }
    }

    private const int MaxUploadConcurrency = 3;
    private readonly SemaphoreSlim _uploadSemaphore = new(MaxUploadConcurrency);

    public async Task<UploadReport> UploadFiles(List<string> filePaths, bool convertToSheets, string? folderId, string googleToken)
    {
        var report = new UploadReport
        {
            GeneratedUtc = DateTime.UtcNow.ToString("o"),
            FileCount = filePaths.Count,
            ConvertedToSheets = convertToSheets
        };

        var tasks = filePaths.Select(async (filePath, index) =>
        {
            await _uploadSemaphore.WaitAsync();
            try
            {
                return (Index: index, Result: await UploadFile(filePath, convertToSheets, folderId, googleToken));
            }
            finally
            {
                _uploadSemaphore.Release();
            }
        }).ToList();

        var results = await Task.WhenAll(tasks);

        foreach (var item in results.OrderBy(r => r.Index))
        {
            report.Files.Add(item.Result);

            if (item.Result.Status == "success")
                report.SuccessCount++;
            else
                report.FailureCount++;
        }

        return report;
    }
}
