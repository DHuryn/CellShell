using System.Diagnostics;
using System.IO;

namespace CellShell.Core;

public static class CommandExecutor
{
    private static string _workingDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

    public static string WorkingDirectory => _workingDirectory;

    public static async Task<string> ExecuteAsync(string command, int timeoutMs = 30000)
    {
        // Handle cd commands to update working directory
        var trimmed = command.Trim();
        if (trimmed.Equals("cd", StringComparison.OrdinalIgnoreCase))
        {
            _workingDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            return _workingDirectory;
        }
        if (trimmed.StartsWith("cd ", StringComparison.OrdinalIgnoreCase))
        {
            var target = trimmed[3..].Trim();
            // Strip /d flag (cmd.exe drive-change switch)
            if (target.StartsWith("/d ", StringComparison.OrdinalIgnoreCase))
                target = target[3..].Trim();
            // Only strip matching surrounding quotes
            if (target.Length >= 2 && target.StartsWith('"') && target.EndsWith('"'))
                target = target[1..^1];
            // Handle ~ as home directory
            if (target == "~")
                target = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var newDir = Path.GetFullPath(Path.Combine(_workingDirectory, target));
            if (Directory.Exists(newDir))
            {
                _workingDirectory = newDir;
                return _workingDirectory;
            }
            return $"The system cannot find the path specified: {target}";
        }

        var psi = new ProcessStartInfo
        {
            FileName = "cmd.exe",
            Arguments = $"/c {command}",
            WorkingDirectory = _workingDirectory,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        using var process = new Process { StartInfo = psi };
        using var cts = new CancellationTokenSource(timeoutMs);

        process.Start();

        // Read both streams concurrently to avoid pipe deadlock
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();

        try
        {
            await process.WaitForExitAsync(cts.Token);
        }
        catch (OperationCanceledException)
        {
            try { process.Kill(entireProcessTree: true); } catch { }
            // Wait briefly for stream tasks to complete after kill
            try { await Task.WhenAll(stdoutTask, stderrTask).WaitAsync(TimeSpan.FromSeconds(2)); } catch { }
            var partialOut = stdoutTask.IsCompletedSuccessfully ? stdoutTask.Result : "";
            var partialErr = stderrTask.IsCompletedSuccessfully ? stderrTask.Result : "";
            return partialOut + partialErr + "\n[Timed out after 30s]";
        }

        var stdout = await stdoutTask;
        var stderr = await stderrTask;

        var result = stdout;
        if (!string.IsNullOrEmpty(stderr))
            result += (string.IsNullOrEmpty(result) ? "" : "\n") + stderr;
        return result.TrimEnd('\r', '\n');
    }
}
