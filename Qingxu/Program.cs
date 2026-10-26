using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Diagnostics;
using System.Text.Json;
using System.Windows;
using System.Windows.Threading;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace Qingxu;

public static class Program
{
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int SetCurrentProcessExplicitAppUserModelID(string applicationId);

    [STAThread]
    public static int Main(string[] arguments)
    {
        try
        {
            if (arguments.Contains("--check"))
            {
                var checkedBook = new XlsxBook(BookArgument(arguments));
                Console.WriteLine(checkedBook.Version); return 0;
            }
            var app = new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };
            app.DispatcherUnhandledException += (_, eventArgs) => { MessageBox.Show(eventArgs.Exception.Message, "青序", MessageBoxButton.OK, MessageBoxImage.Warning); eventArgs.Handled = true; };
            Marshal.ThrowExceptionForHR(SetCurrentProcessExplicitAppUserModelID("Qingxu.Native.Agenda"));
            // 与 PowerShell 版共用同一组内核对象名，避免两个实现同时持有同一个工作簿
            using var mutex = new Mutex(true, "Local\\CodexTodayAgenda_27107", out var created);
            using var restore = new EventWaitHandle(false, EventResetMode.AutoReset, "Local\\QingxuRestore_27107");
            if (!created) { restore.Set(); return 0; }
            try
            {
                var pageOption = Array.IndexOf(arguments, "--page");
                var startPage = pageOption >= 0 && pageOption + 1 < arguments.Length && int.TryParse(arguments[pageOption + 1], out var parsedPage) ? parsedPage : 0;
                var window = new MainWindow(BookArgument(arguments), restore, startPage);
                app.MainWindow = window;
                var smoke = Array.IndexOf(arguments, "--smoke-test");
                if (smoke >= 0 && smoke + 1 < arguments.Length) InstallSmokeTest(window, Path.GetFullPath(arguments[smoke + 1]));
                return app.Run(window);
            }
            finally { mutex.ReleaseMutex(); }
        }
        catch (Exception error)
        {
            var log = Path.Combine(AppContext.BaseDirectory, "启动错误.txt");
            try { File.WriteAllText(log, error.ToString()); } catch { }
            MessageBox.Show("青序未能启动：" + error.Message + "\n旧版程序和个人数据没有改动。", "青序", MessageBoxButton.OK, MessageBoxImage.Warning);
            return 1;
        }
    }
    private static void InstallSmokeTest(MainWindow window, string reportPath)
    {
        window.ContentRendered += (_, _) =>
        {
            window.UpdateLayout();
            var image = new RenderTargetBitmap((int)window.ActualWidth, (int)window.ActualHeight, 96, 96, PixelFormats.Pbgra32); image.Render(window);
            var encoder = new PngBitmapEncoder(); encoder.Frames.Add(BitmapFrame.Create(image));
            using (var output = File.Create(reportPath + ".png")) encoder.Save(output);
            var process = Process.GetCurrentProcess();
            var cpu = process.TotalProcessorTime.TotalSeconds; var clock = Stopwatch.StartNew();
            var warmed = false;
            var timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(3) };
            timer.Tick += (_, _) =>
            {
                if (!warmed) { warmed = true; process.Refresh(); cpu = process.TotalProcessorTime.TotalSeconds; clock.Restart(); timer.Interval = TimeSpan.FromSeconds(10); return; }
                timer.Stop(); process.Refresh();
                var report = new { WorkingSetMiB = process.WorkingSet64 / 1048576.0, PrivateMiB = process.PrivateMemorySize64 / 1048576.0,
                    CpuPercentOfMachine = 100 * (process.TotalProcessorTime.TotalSeconds - cpu) / clock.Elapsed.TotalSeconds / Environment.ProcessorCount,
                    Seconds = clock.Elapsed.TotalSeconds, LogicalProcessors = Environment.ProcessorCount, window.IsVisible, window.ShowInTaskbar };
                File.WriteAllText(reportPath, JsonSerializer.Serialize(report, AgendaRules.JsonOptions)); window.Exit();
            };
            timer.Start();
        };
    }
    public static string BookArgument(string[] arguments)
    {
        var option = Array.IndexOf(arguments, "--book");
        if (option >= 0 && option + 1 < arguments.Length) return Path.GetFullPath(arguments[option + 1]);
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            var path = Path.Combine(directory.FullName, "outputs", "study-plan", "计划.xlsx");
            if (File.Exists(path)) return path;
            directory = directory.Parent;
        }
        throw new FileNotFoundException("没有找到计划.xlsx。请使用 --book 指定工作簿路径。");
    }
}
