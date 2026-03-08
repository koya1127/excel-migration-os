using System.Diagnostics;

namespace ExcelMigrationApi.Services;

public static class ProcessHelper
{
    public static async Task<(int ExitCode, string StdOut, string StdErr)> RunProcessAsync(
        string command, string args, string? workingDir = null, int timeoutMs = 60000)
    {
        using var process = new Process();
        process.StartInfo = new ProcessStartInfo
        {
            FileName = command,
            Arguments = args,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
            WorkingDirectory = workingDir ?? Directory.GetCurrentDirectory()
        };

        try
        {
            process.Start();
        }
        catch (Exception ex)
        {
            return (-1, string.Empty, $"Failed to start '{command}': {ex.Message}");
        }

        using var cts = new CancellationTokenSource(timeoutMs);

        try
        {
            var stdOutTask = process.StandardOutput.ReadToEndAsync(cts.Token);
            var stdErrTask = process.StandardError.ReadToEndAsync(cts.Token);

            await process.WaitForExitAsync(cts.Token);

            var stdout = await stdOutTask;
            var stderr = await stdErrTask;

            return (process.ExitCode, stdout, stderr);
        }
        catch (OperationCanceledException)
        {
            try { process.Kill(entireProcessTree: true); } catch { }
            return (-1, string.Empty, $"Process '{command}' timed out after {timeoutMs}ms");
        }
    }
}
