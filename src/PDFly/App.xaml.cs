using System.IO;
using System.Threading.Tasks;
using Microsoft.UI.Xaml;

namespace PDFly;

public partial class App : Application
{
    /// <summary>The single application window. Available app-wide for pickers / interop.</summary>
    public static Window Window { get; private set; } = null!;

    /// <summary>The UI thread's dispatcher. Use to marshal background-thread work back
    /// to the UI. Fully qualified to avoid CS0104 with <see cref="Windows.System.DispatcherQueue"/>
    /// and to avoid shadowing the static GetForCurrentThread call below.</summary>
    public static Microsoft.UI.Dispatching.DispatcherQueue DispatcherQueue { get; private set; } = null!;

    /// <summary>The native HWND of the main window — needed by file pickers and any
    /// WinRT interop that requires <c>InitializeWithWindow</c>.</summary>
    public static nint WindowHandle => WinRT.Interop.WindowNative.GetWindowHandle(Window);

    public App()
    {
        InitializeComponent();

        // Best-effort crash logging. Anything we can catch goes to %AppData%\PDFly\crash.log.
        UnhandledException += (_, ev) =>
        {
            LogCrash(ev.Exception, "UnhandledException");
            ev.Handled = true;
        };
        AppDomain.CurrentDomain.UnhandledException += (_, ev) =>
            LogCrash(ev.ExceptionObject as Exception, "AppDomain");
        TaskScheduler.UnobservedTaskException += (_, ev) =>
        {
            LogCrash(ev.Exception, "Task");
            ev.SetObserved();
        };
    }

    protected override void OnLaunched(LaunchActivatedEventArgs args)
    {
        Window = new MainWindow();
        DispatcherQueue = Microsoft.UI.Dispatching.DispatcherQueue.GetForCurrentThread();
        Window.Activate();
    }

    private static readonly object CrashLogLock = new();
    private static void LogCrash(Exception? ex, string source)
    {
        if (ex is null) return;
        try
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PDFly");
            Directory.CreateDirectory(dir);
            var path = Path.Combine(dir, "crash.log");
            var entry = $"[{DateTime.Now:O}] [{source}]\n{ex}\n\n";
            lock (CrashLogLock) File.AppendAllText(path, entry);
        }
        catch { /* logging best-effort */ }
    }
}
