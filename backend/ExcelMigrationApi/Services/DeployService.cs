using System.Text.Json;
using System.Text.RegularExpressions;
using ExcelMigrationApi.Models;

namespace ExcelMigrationApi.Services;

public class DeployService
{
    public async Task<DeployReport> Deploy(DeployRequest request)
    {
        var report = new DeployReport
        {
            GeneratedUtc = DateTime.UtcNow.ToString("o"),
            SpreadsheetId = request.SpreadsheetId,
            FileCount = request.GasFiles.Count
        };

        // Create a temp directory for the GAS project
        var tempDir = Path.Combine(Path.GetTempPath(), "excel-migration-deploy-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        try
        {
            // Step 1: Create a new Apps Script project bound to the spreadsheet
            var (createExit, createOut, createErr) = await ProcessHelper.RunProcessAsync(
                "clasp", $"create --type sheets --parentId {request.SpreadsheetId} --title \"Migration Script\"",
                workingDir: tempDir, timeoutMs: 60000);

            if (createExit != 0)
            {
                report.Status = "error";
                report.Error = $"clasp create failed (exit {createExit}): {createErr}";
                return report;
            }

            // Extract scriptId from .clasp.json that clasp created
            var claspJsonPath = Path.Combine(tempDir, ".clasp.json");
            if (File.Exists(claspJsonPath))
            {
                try
                {
                    var claspJson = File.ReadAllText(claspJsonPath);
                    using var doc = JsonDocument.Parse(claspJson);
                    if (doc.RootElement.TryGetProperty("scriptId", out var scriptIdProp))
                    {
                        report.ScriptId = scriptIdProp.GetString() ?? string.Empty;
                    }
                }
                catch { }
            }

            // Step 2: Write appsscript.json manifest
            var manifest = new
            {
                timeZone = "Asia/Tokyo",
                dependencies = new { },
                runtimeVersion = "V8"
            };
            var manifestJson = JsonSerializer.Serialize(manifest, new JsonSerializerOptions { WriteIndented = true });
            await File.WriteAllTextAsync(Path.Combine(tempDir, "appsscript.json"), manifestJson);

            // Step 3: Write each .gs file
            foreach (var gasFile in request.GasFiles)
            {
                var ext = gasFile.Type == "HTML" ? ".html" : ".gs";
                var safeName = gasFile.Name;
                if (!safeName.EndsWith(ext, StringComparison.OrdinalIgnoreCase))
                {
                    safeName += ext;
                }
                await File.WriteAllTextAsync(Path.Combine(tempDir, safeName), gasFile.Source);
                report.FilesDeployed.Add(safeName);
            }

            // Step 4: Push to Apps Script
            var (pushExit, pushOut, pushErr) = await ProcessHelper.RunProcessAsync(
                "clasp", "push --force",
                workingDir: tempDir, timeoutMs: 60000);

            if (pushExit != 0)
            {
                report.Status = "error";
                report.Error = $"clasp push failed (exit {pushExit}): {pushErr}";
                return report;
            }

            // Step 5: Deploy
            var (deployExit, deployOut, deployErr) = await ProcessHelper.RunProcessAsync(
                "clasp", "deploy --description \"Auto-deployed by Excel Migration OS\"",
                workingDir: tempDir, timeoutMs: 60000);

            if (deployExit != 0)
            {
                // Push succeeded but deploy failed - partial success
                report.Status = "partial";
                report.Error = $"clasp push succeeded but deploy failed (exit {deployExit}): {deployErr}";
                return report;
            }

            report.Status = "success";
        }
        catch (Exception ex)
        {
            report.Status = "error";
            report.Error = ex.Message;
        }
        finally
        {
            // Clean up temp dir
            try
            {
                Directory.Delete(tempDir, recursive: true);
            }
            catch { }
        }

        return report;
    }
}
