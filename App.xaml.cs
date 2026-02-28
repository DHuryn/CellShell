using System.Runtime.InteropServices;
using System.Windows;

namespace CellShell;

public partial class App : Application
{
    [DllImport("kernel32.dll")]
    private static extern bool AttachConsole(int dwProcessId);
    private const int ATTACH_PARENT_PROCESS = -1;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // WinExe apps aren't attached to a console, so Console.CancelKeyPress
        // never fires. Attach to the parent console (terminal / dotnet run) first.
        if (AttachConsole(ATTACH_PARENT_PROCESS))
        {
            Console.CancelKeyPress += (_, args) =>
            {
                args.Cancel = true;
                Dispatcher.Invoke(() =>
                {
                    KillAllRunningProcesses();
                    Shutdown();
                });
            };
        }

        // Fallback: kill child processes during any process exit
        AppDomain.CurrentDomain.ProcessExit += (_, _) =>
        {
            try { Dispatcher.Invoke(KillAllRunningProcesses); } catch { }
        };
    }

    private void KillAllRunningProcesses()
    {
        if (MainWindow is not MainWindow mw) return;
        foreach (var row in mw.Model.Rows)
        {
            if (row.RunningProcess != null)
            {
                try { row.RunningProcess.Kill(entireProcessTree: true); } catch { }
                row.RunningProcess = null;
            }
        }
    }
}
