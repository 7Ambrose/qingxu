using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Reflection;

[assembly: AssemblyTitle("Qingxu PowerShell Launcher")]
[assembly: AssemblyDescription("青序 PowerShell 版启动器，非原生版")]
[assembly: AssemblyProduct("Qingxu PowerShell")]
[assembly: AssemblyVersion("0.1.0.0")]
[assembly: AssemblyFileVersion("0.1.0.0")]

internal static class QingxuLauncher
{
    [STAThread]
    private static int Main(string[] arguments)
    {
        bool diagnose = Array.IndexOf(arguments, "--diagnose") >= 0;
        string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
        string scriptPath = Path.Combine(baseDirectory, "今日日程.ps1");
        if (!File.Exists(scriptPath))
        {
            return ReportFailure(baseDirectory, "找不到今日日程.ps1，请确认青序文件夹完整。", diagnose, 1);
        }

        string powershellPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.System),
            "WindowsPowerShell\\v1.0\\powershell.exe");
        ProcessStartInfo startInfo = new ProcessStartInfo
        {
            FileName = powershellPath,
            Arguments = "-NoProfile -STA -WindowStyle Hidden -File \"" + scriptPath + "\"",
            WorkingDirectory = baseDirectory,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            WindowStyle = ProcessWindowStyle.Hidden
        };

        try
        {
            using (Process process = new Process())
            {
                StringBuilder output = new StringBuilder();
                process.StartInfo = startInfo;
                DataReceivedEventHandler capture = delegate(object sender, DataReceivedEventArgs received)
                {
                    if (received.Data != null)
                    {
                        lock (output)
                        {
                            if (output.Length < 65536) output.AppendLine(received.Data);
                        }
                    }
                };
                process.OutputDataReceived += capture;
                process.ErrorDataReceived += capture;
                process.Start();
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();
                process.WaitForExit();
                if (process.ExitCode != 0)
                {
                    return ReportFailure(baseDirectory, output.ToString(), diagnose, process.ExitCode);
                }
            }
            return 0;
        }
        catch (Exception error)
        {
            return ReportFailure(baseDirectory, error.ToString(), diagnose, 1);
        }
    }

    private static int ReportFailure(string baseDirectory, string details, bool diagnose, int exitCode)
    {
        string logPath = Path.Combine(baseDirectory, "启动错误.txt");
        try
        {
            File.WriteAllText(logPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + Environment.NewLine + details, Encoding.UTF8);
        }
        catch
        {
            logPath = "无法写入启动日志";
        }
        bool policyBlocked = details.IndexOf("PSSecurityException", StringComparison.OrdinalIgnoreCase) >= 0
            || details.IndexOf("UnauthorizedAccess", StringComparison.OrdinalIgnoreCase) >= 0
            || details.IndexOf("ExecutionPolicy", StringComparison.OrdinalIgnoreCase) >= 0;
        string message = policyBlocked
            ? "PowerShell 的执行策略阻止了青序脚本。请确认文件来源，并由你或管理员按可信脚本规范配置。程序未修改系统执行策略。"
            : "青序启动或运行失败，请检查文件是否完整、Python 是否可用，并查看启动日志。";
        if (!diagnose) MessageBox.Show(message + Environment.NewLine + logPath, "青序启动失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
        return exitCode;
    }
}
