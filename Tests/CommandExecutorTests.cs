using CellShell.Core;

namespace CellShell.Tests;

public class CommandExecutorTests
{
    [Fact]
    public async Task Execute_EchoCommand_ReturnsOutput()
    {
        var result = await CommandExecutor.ExecuteAsync("echo hello");
        Assert.Contains("hello", result);
    }

    [Fact]
    public async Task Execute_CdHome_ReturnsHomeDir()
    {
        var result = await CommandExecutor.ExecuteAsync("cd");
        var expected = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        Assert.Equal(expected, result);
    }

    [Fact]
    public async Task Execute_CdWithSlashD_Works()
    {
        var result = await CommandExecutor.ExecuteAsync("cd /d C:\\");
        Assert.Equal("C:\\", result);
    }

    [Fact]
    public async Task Execute_CdTilde_GoesHome()
    {
        var result = await CommandExecutor.ExecuteAsync("cd ~");
        var expected = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        Assert.Equal(expected, result);
    }

    [Fact]
    public async Task Execute_CdNonexistent_ReturnsError()
    {
        var result = await CommandExecutor.ExecuteAsync("cd nonexistent_dir_12345");
        Assert.Contains("cannot find the path", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Execute_InvalidCommand_ReturnsStderr()
    {
        var result = await CommandExecutor.ExecuteAsync("this_command_does_not_exist_xyz");
        Assert.False(string.IsNullOrEmpty(result));
    }

    [Fact]
    public async Task ExecuteStreaming_EchoCommand_CallsOnOutput()
    {
        var lines = new List<string>();
        await CommandExecutor.ExecuteStreamingAsync("echo hello", line => lines.Add(line));
        Assert.Contains(lines, l => l.Contains("hello"));
    }

    [Fact]
    public async Task ExecuteStreaming_CdHome_CallsOnOutput()
    {
        var lines = new List<string>();
        await CommandExecutor.ExecuteStreamingAsync("cd", line => lines.Add(line));
        var expected = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        Assert.Contains(expected, lines);
    }

    [Fact]
    public async Task ExecuteStreaming_MultipleLines_CallsOnOutputPerLine()
    {
        var lines = new List<string>();
        await CommandExecutor.ExecuteStreamingAsync("echo line1 & echo line2", line => lines.Add(line));
        Assert.True(lines.Count >= 2);
        Assert.Contains(lines, l => l.Contains("line1"));
        Assert.Contains(lines, l => l.Contains("line2"));
    }
}
