using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using CellShell;
using CellShell.Core;

namespace CellShell.PerfTest;

static class PerfHarness
{
    private static readonly List<double> _frameTimes = new();
    private static TimeSpan _lastRenderTime;
    private static bool _recording;

    [STAThread]
    static int Main()
    {
        try
        {
            var app = new Application();
            var window = new MainWindow();
            window.Show();

            // Let the window fully render and settle
            PumpDispatcher(1000);

            Console.WriteLine("=== CellShell Performance Report ===");
            Console.WriteLine();

            RunColdRedraw(window);
            RunStreamingOutput(window);
            RunColumnDragResize(window);
            RunScroll(window);

            app.Shutdown();
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"FATAL: {ex}");
            return 1;
        }
    }

    // ─── Scenarios ────────────────────────────────────────────────

    static void RunColdRedraw(MainWindow window)
    {
        Settle();
        var redrawTimes = new List<double>();

        StartRecording();
        for (int i = 0; i < 10; i++)
        {
            var sw = Stopwatch.StartNew();
            window.ForceRedraw();
            PumpDispatcher(50); // let WPF render the frame
            sw.Stop();
            redrawTimes.Add(sw.Elapsed.TotalMilliseconds);
        }
        StopRecording();

        var stats = ComputeStats();
        Console.WriteLine("Cold Redraw (10 iterations):");
        Console.WriteLine($"  Avg redraw:    {redrawTimes.Average():F1}ms");
        PrintFrameStats(stats);
        Console.WriteLine();
    }

    static void RunStreamingOutput(MainWindow window)
    {
        Settle();

        // Submit a command that produces many lines of output quickly
        var model = window.Model;
        model.SelectRow(model.Rows.Count);
        window.ForceRedraw();
        PumpDispatcher(100);

        StartRecording();
        var overallSw = Stopwatch.StartNew();

        // Simulate streaming: use a fast for /L loop that prints 50 lines
        var command = "for /L %i in (1,1,50) do @echo Line %i";
        var data = model.SubmitCommand(command);
        var executingRow = model.ActiveRowIndex;
        model.MoveToNextEmptyRow();
        window.ForceRedraw();

        // Fire the streaming execution
        _ = CommandExecutor.ExecuteStreamingAsync(
            command,
            line => window.Dispatcher.InvokeAsync(() =>
            {
                data.Output = string.IsNullOrEmpty(data.Output) ? line : data.Output + "\n" + line;
            }),
            process => window.Dispatcher.InvokeAsync(() => data.RunningProcess = process));

        // Wait for command to finish (poll status), redrawing so the UI reflects streaming output
        var timeout = Stopwatch.StartNew();
        while (data.Status == CellStatus.Running && timeout.ElapsedMilliseconds < 30000)
        {
            window.ForceRedraw();
            PumpDispatcher(50);
        }

        if (data.Status == CellStatus.Running)
        {
            data.Status = CellStatus.Error;
            data.Output = "[Timed out]";
        }
        else if (data.Status != CellStatus.Error)
        {
            data.Status = CellStatus.Complete;
        }
        data.RunningProcess = null;

        overallSw.Stop();
        StopRecording();

        var stats = ComputeStats();
        Console.WriteLine("Streaming Output (for /L 50 lines):");
        Console.WriteLine($"  Duration:      {overallSw.Elapsed.TotalSeconds:F1}s");
        if (stats.FrameCount > 1)
            Console.WriteLine($"  Avg FPS:       {stats.AvgFps:F1}");
        PrintFrameStats(stats);
        Console.WriteLine();
    }

    static void RunColumnDragResize(MainWindow window)
    {
        Settle();
        var model = window.Model;
        var originalWidth = model.ColWidths[0];
        const int steps = 100;

        StartRecording();
        for (int i = 0; i < steps; i++)
        {
            // Oscillate width between original and original+200
            var t = (double)i / steps;
            model.ColWidths[0] = originalWidth + Math.Sin(t * Math.PI * 2) * 100;
            window.ForceRedraw();
            PumpDispatcher(16); // ~60fps pacing
        }
        // Restore
        model.ColWidths[0] = originalWidth;
        window.ForceRedraw();
        StopRecording();

        var stats = ComputeStats();
        Console.WriteLine("Column Drag Resize (100 steps):");
        if (stats.FrameCount > 1)
            Console.WriteLine($"  Avg FPS:       {stats.AvgFps:F1}");
        PrintFrameStats(stats);
        Console.WriteLine();
    }

    static void RunScroll(MainWindow window)
    {
        Settle();
        var scroller = window.GridScrollerAccess;

        // First, ensure we have enough rows to scroll through
        window.ForceRedraw();
        PumpDispatcher(100);

        const int scrollSteps = 200;
        var model = window.Model;
        var rowHeight = model.DefaultRowHeight;

        StartRecording();
        for (int i = 0; i < scrollSteps; i++)
        {
            scroller.ScrollToVerticalOffset(i * rowHeight);
            PumpDispatcher(16);
        }
        // Scroll back to top
        scroller.ScrollToVerticalOffset(0);
        PumpDispatcher(100);
        StopRecording();

        var stats = ComputeStats();
        Console.WriteLine("Scroll (200 rows):");
        if (stats.FrameCount > 1)
            Console.WriteLine($"  Avg FPS:       {stats.AvgFps:F1}");
        PrintFrameStats(stats);
        Console.WriteLine();
    }

    // ─── Frame timing infrastructure ──────────────────────────────

    static void StartRecording()
    {
        _frameTimes.Clear();
        _lastRenderTime = TimeSpan.Zero;
        _recording = true;
        CompositionTarget.Rendering += OnRendering;
    }

    static void StopRecording()
    {
        _recording = false;
        CompositionTarget.Rendering -= OnRendering;
    }

    static void OnRendering(object? sender, EventArgs e)
    {
        if (!_recording) return;
        if (e is not RenderingEventArgs args) return;

        var now = args.RenderingTime;
        if (_lastRenderTime != TimeSpan.Zero)
        {
            var delta = (now - _lastRenderTime).TotalMilliseconds;
            if (delta > 0)
                _frameTimes.Add(delta);
        }
        _lastRenderTime = now;
    }

    static FrameStats ComputeStats()
    {
        if (_frameTimes.Count == 0)
            return new FrameStats();

        var sorted = _frameTimes.OrderBy(t => t).ToList();
        return new FrameStats
        {
            FrameCount = sorted.Count,
            P50 = Percentile(sorted, 0.50),
            P95 = Percentile(sorted, 0.95),
            P99 = Percentile(sorted, 0.99),
            JankCount = sorted.Count(t => t > 33),
            AvgFps = sorted.Count > 0 ? 1000.0 / sorted.Average() : 0
        };
    }

    static double Percentile(List<double> sorted, double p)
    {
        if (sorted.Count == 0) return 0;
        var index = (int)Math.Floor(p * (sorted.Count - 1));
        return sorted[Math.Min(index, sorted.Count - 1)];
    }

    static void PrintFrameStats(FrameStats stats)
    {
        if (stats.FrameCount == 0)
        {
            Console.WriteLine("  (no frame data captured)");
            return;
        }
        Console.WriteLine($"  Frame P50:     {stats.P50:F1}ms");
        Console.WriteLine($"  Frame P95:     {stats.P95:F1}ms");
        Console.WriteLine($"  Frame P99:     {stats.P99:F1}ms");
        Console.WriteLine($"  Jank (>33ms):  {stats.JankCount} frames");
    }

    // ─── Helpers ──────────────────────────────────────────────────

    static void Settle()
    {
        _frameTimes.Clear();
        _lastRenderTime = TimeSpan.Zero;
        PumpDispatcher(500);
    }

    static void PumpDispatcher(int ms)
    {
        var sw = Stopwatch.StartNew();
        while (sw.ElapsedMilliseconds < ms)
        {
            Dispatcher.CurrentDispatcher.Invoke(
                DispatcherPriority.Background,
                new Action(() => { }));
            Thread.Sleep(1);
        }
    }

    record struct FrameStats
    {
        public int FrameCount;
        public double P50;
        public double P95;
        public double P99;
        public int JankCount;
        public double AvgFps;
    }
}
