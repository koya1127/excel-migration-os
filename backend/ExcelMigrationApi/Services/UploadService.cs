using System.Text.Json;
using ExcelMigrationApi.Models;

namespace ExcelMigrationApi.Services;

public class UploadService
{
    public async Task<UploadResult> UploadFile(string filePath, bool convertToSheets, string? folderId)
    {
        var fileName = Path.GetFileName(filePath);

        try
        {
            // Build gws metadata JSON
            var metadata = new Dictionary<string, object>
            {
                ["name"] = Path.GetFileNameWithoutExtension(fileName)
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
            // Escape for command line - use single quotes on the outside
            var escapedJson = metadataJson.Replace("\"", "\\\"");

            var args = $"drive files create --upload \"{filePath.Replace("\\", "/")}\" --json \"{escapedJson}\"";

            var (exitCode, stdout, stderr) = await ProcessHelper.RunProcessAsync("gws", args, timeoutMs: 120000);

            if (exitCode != 0)
            {
                return new UploadResult
                {
                    FileName = fileName,
                    Status = "error",
                    Error = $"gws exited with code {exitCode}: {stderr}"
                };
            }

            // Parse JSON response from gws
            var driveFileId = string.Empty;
            var webViewLink = string.Empty;

            try
            {
                using var doc = JsonDocument.Parse(stdout);
                var root = doc.RootElement;

                if (root.TryGetProperty("id", out var idProp))
                    driveFileId = idProp.GetString() ?? string.Empty;

                if (root.TryGetProperty("webViewLink", out var linkProp))
                    webViewLink = linkProp.GetString() ?? string.Empty;
            }
            catch
            {
                // If we can't parse JSON but exit code was 0, treat as partial success
                driveFileId = "(unknown - could not parse gws output)";
            }

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

    public async Task<UploadReport> UploadFiles(List<string> filePaths, bool convertToSheets, string? folderId)
    {
        var report = new UploadReport
        {
            GeneratedUtc = DateTime.UtcNow.ToString("o"),
            FileCount = filePaths.Count,
            ConvertedToSheets = convertToSheets
        };

        foreach (var filePath in filePaths)
        {
            var result = await UploadFile(filePath, convertToSheets, folderId);
            report.Files.Add(result);

            if (result.Status == "success")
                report.SuccessCount++;
            else
                report.FailureCount++;
        }

        return report;
    }
}
