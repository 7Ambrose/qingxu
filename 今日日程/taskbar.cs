using System;
using System.IO;
using System.Threading;
using System.Runtime.InteropServices;

public sealed class QingxuWorkbookWatcher : IDisposable
{
    private readonly FileSystemWatcher watcher;
    private int changed;

    public QingxuWorkbookWatcher(string path)
    {
        watcher = new FileSystemWatcher(Path.GetDirectoryName(Path.GetFullPath(path)), Path.GetFileName(path));
        watcher.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.Size;
        watcher.Changed += MarkChanged;
        watcher.Created += MarkChanged;
        watcher.Deleted += MarkChanged;
        watcher.Renamed += MarkChanged;
        watcher.EnableRaisingEvents = true;
    }

    private void MarkChanged(object sender, FileSystemEventArgs args)
    {
        Interlocked.Exchange(ref changed, 1);
    }

    public bool ConsumeChange()
    {
        return Interlocked.Exchange(ref changed, 0) != 0;
    }

    public void Dispose()
    {
        watcher.Dispose();
    }
}

public static class QingxuTaskbar
{
    public const string ApplicationId = "Qingxu.Desktop.Agenda";

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int SetCurrentProcessExplicitAppUserModelID(string applicationId);

    [DllImport("shell32.dll")]
    private static extern int GetCurrentProcessExplicitAppUserModelID(out IntPtr applicationId);

    [DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr window, uint message, IntPtr parameter, IntPtr value);

    public static void Initialize()
    {
        Marshal.ThrowExceptionForHR(SetCurrentProcessExplicitAppUserModelID(ApplicationId));
    }

    public static string GetApplicationId()
    {
        IntPtr applicationId;
        Marshal.ThrowExceptionForHR(GetCurrentProcessExplicitAppUserModelID(out applicationId));
        try { return Marshal.PtrToStringUni(applicationId); }
        finally { Marshal.FreeCoTaskMem(applicationId); }
    }

    public static void SetWindowIcon(IntPtr window, IntPtr icon)
    {
        SendMessage(window, 0x0080, IntPtr.Zero, icon);
        SendMessage(window, 0x0080, new IntPtr(1), icon);
    }

    public static IntPtr GetWindowIcon(IntPtr window, bool large)
    {
        return SendMessage(window, 0x007F, new IntPtr(large ? 1 : 0), IntPtr.Zero);
    }
}
