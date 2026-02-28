using System.Diagnostics;
using System.IO;

namespace CellShell.Core;

public enum ShellType { Cmd, PowerShell }

public static class CommandExecutor
{
    private static string _workingDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

    public static ShellType CurrentShell { get; set; } = ShellType.Cmd;

    public static string WorkingDirectory => _workingDirectory;

    public static async Task<string> ExecuteAsync(string command, int timeoutMs = 0)
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

        var (fileName, arguments) = GetShellCommand(command);
        var psi = new ProcessStartInfo
        {
            FileName = fileName,
            Arguments = arguments,
            WorkingDirectory = _workingDirectory,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        using var process = new Process { StartInfo = psi };
        process.Start();

        // Read both streams concurrently to avoid pipe deadlock
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();

        if (timeoutMs > 0)
        {
            using var cts = new CancellationTokenSource(timeoutMs);
            try
            {
                await process.WaitForExitAsync(cts.Token);
            }
            catch (OperationCanceledException)
            {
                try { process.Kill(entireProcessTree: true); } catch { }
                try { await Task.WhenAll(stdoutTask, stderrTask).WaitAsync(TimeSpan.FromSeconds(2)); } catch { }
                var partialOut = stdoutTask.IsCompletedSuccessfully ? stdoutTask.Result : "";
                var partialErr = stderrTask.IsCompletedSuccessfully ? stderrTask.Result : "";
                return partialOut + partialErr + $"\n[Timed out after {timeoutMs / 1000}s]";
            }
        }
        else
        {
            await process.WaitForExitAsync();
        }

        var stdout = await stdoutTask;
        var stderr = await stderrTask;

        var result = stdout;
        if (!string.IsNullOrEmpty(stderr))
            result += (string.IsNullOrEmpty(result) ? "" : "\n") + stderr;
        return result.TrimEnd('\r', '\n');
    }

    /// <summary>
    /// Streams command output line-by-line via onOutput callback.
    /// The caller's Process reference is set via processCallback so it can be killed externally.
    /// </summary>
    public static async Task ExecuteStreamingAsync(
        string command,
        Action<string> onOutput,
        Action<Process?>? processCallback = null,
        int timeoutMs = 0)
    {
        // Handle cd commands the same as the sync version
        var trimmed = command.Trim();
        if (trimmed.Equals("cd", StringComparison.OrdinalIgnoreCase))
        {
            _workingDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            onOutput(_workingDirectory);
            return;
        }
        if (trimmed.StartsWith("cd ", StringComparison.OrdinalIgnoreCase))
        {
            var target = trimmed[3..].Trim();
            if (target.StartsWith("/d ", StringComparison.OrdinalIgnoreCase))
                target = target[3..].Trim();
            if (target.Length >= 2 && target.StartsWith('"') && target.EndsWith('"'))
                target = target[1..^1];
            if (target == "~")
                target = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var newDir = Path.GetFullPath(Path.Combine(_workingDirectory, target));
            if (Directory.Exists(newDir))
            {
                _workingDirectory = newDir;
                onOutput(_workingDirectory);
            }
            else
            {
                onOutput($"The system cannot find the path specified: {target}");
            }
            return;
        }

        var (fileName, arguments) = GetShellCommand(command);
        var psi = new ProcessStartInfo
        {
            FileName = fileName,
            Arguments = arguments,
            WorkingDirectory = _workingDirectory,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        var process = new Process { StartInfo = psi, EnableRaisingEvents = true };
        var tcs = new TaskCompletionSource<bool>();

        process.OutputDataReceived += (_, e) =>
        {
            if (e.Data != null) onOutput(e.Data);
        };
        process.ErrorDataReceived += (_, e) =>
        {
            if (e.Data != null) onOutput(e.Data);
        };
        process.Exited += (_, _) => tcs.TrySetResult(true);

        process.Start();
        processCallback?.Invoke(process);
        process.BeginOutputReadLine();
        process.BeginErrorReadLine();

        if (timeoutMs > 0)
        {
            using var cts = new CancellationTokenSource(timeoutMs);
            using var reg = cts.Token.Register(() =>
            {
                try { process.Kill(entireProcessTree: true); } catch { }
                onOutput($"[Timed out after {timeoutMs / 1000}s]");
                tcs.TrySetResult(false);
            });
            await tcs.Task;
        }
        else
        {
            await tcs.Task;
        }
        processCallback?.Invoke(null);
        process.Dispose();
    }

    private static (string FileName, string Arguments) GetShellCommand(string command)
    {
        if (CurrentShell == ShellType.Cmd)
            return ("cmd.exe", $"/c {command}");

        // Prefer pwsh (PowerShell 7+), fall back to powershell (Windows PowerShell 5.1)
        var pwsh = FindExecutable("pwsh");
        var exe = pwsh ?? "powershell";
        return (exe, $"-NoProfile -Command {command}");
    }

    private static string? FindExecutable(string name)
    {
        var pathVar = Environment.GetEnvironmentVariable("PATH") ?? "";
        foreach (var dir in pathVar.Split(Path.PathSeparator))
        {
            var full = Path.Combine(dir, name + ".exe");
            if (File.Exists(full)) return full;
        }
        return null;
    }
}
