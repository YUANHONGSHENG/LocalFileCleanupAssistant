using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

internal static class Program
{
    private const string Version = "20261029-1";
    private const string ResourcePrefix = "LocalFileCleanupAssistant.";

    [STAThread]
    private static void Main(string[] args)
    {
        try
        {
            string root = FindWritableRoot();

            string appDir = Path.Combine(root, "app-" + Version);
            try
            {
                ExtractAll(appDir);
            }
            catch
            {
                appDir = Path.Combine(root, "app-" + Version + "-" + Guid.NewGuid().ToString("N").Substring(0, 8));
                ExtractAll(appDir);
            }

            string bootstrap = Path.Combine(appDir, "Bootstrap.ps1");
            string ps = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.System),
                @"WindowsPowerShell\v1.0\powershell.exe");
            if (!File.Exists(ps))
                ps = "powershell.exe";

            var sb = new StringBuilder();
            sb.Append("-STA -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"");
            sb.Append(bootstrap);
            sb.Append("\"");
            foreach (string a in args)
            {
                sb.Append(" ");
                sb.Append(Quote(a));
            }

            var psi = new ProcessStartInfo();
            psi.FileName = ps;
            psi.Arguments = sb.ToString();
            psi.WorkingDirectory = root;
            psi.UseShellExecute = false;
            psi.CreateNoWindow = true;
            Process.Start(psi);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "启动失败：\r\n" + ex.Message,
                "本地文件清理助手",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private static string FindWritableRoot()
    {
        string exeDir = AppDomain.CurrentDomain.BaseDirectory;
        string[] candidates = new string[]
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "LocalFileCleanupAssistant"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "LocalFileCleanupAssistant"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "LocalFileCleanupAssistant"),
            Path.Combine(Path.GetTempPath(), "LocalFileCleanupAssistant"),
            Path.Combine(exeDir, "LocalFileCleanupAssistantData")
        };

        StringBuilder errors = new StringBuilder();
        foreach (string root in candidates)
        {
            if (string.IsNullOrEmpty(root)) continue;
            try
            {
                Directory.CreateDirectory(root);
                string probe = Path.Combine(root, ".write-test-" + Guid.NewGuid().ToString("N"));
                File.WriteAllText(probe, "ok");
                File.Delete(probe);
                return root;
            }
            catch (Exception ex)
            {
                errors.AppendLine(root + " -> " + ex.Message);
            }
        }

        throw new IOException("找不到可写的数据目录。已尝试：\r\n" + errors.ToString());
    }

    private static void ExtractAll(string appDir)
    {
        Directory.CreateDirectory(appDir);
        Extract("Bootstrap.ps1", Path.Combine(appDir, "Bootstrap.ps1"));
        Extract("App.ps1", Path.Combine(appDir, "App.ps1"));
        Extract("Scanner.ps1", Path.Combine(appDir, "Scanner.ps1"));
        Extract("MainWindow.xaml", Path.Combine(appDir, "MainWindow.xaml"));
    }

    private static void Extract(string name, string destination)
    {
        Assembly asm = Assembly.GetExecutingAssembly();
        using (Stream input = asm.GetManifestResourceStream(ResourcePrefix + name))
        {
            if (input == null)
                throw new InvalidOperationException("缺少内置资源：" + name);
            using (FileStream output = new FileStream(destination, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                input.CopyTo(output);
            }
        }
    }

    private static string Quote(string value)
    {
        if (value == null) return "\"\"";
        if (value.Length > 0 && value.IndexOfAny(new char[] { ' ', '\t', '"' }) < 0)
            return value;
        return "\"" + value.Replace("\"", "\\\"") + "\"";
    }
}
