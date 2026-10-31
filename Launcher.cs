using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        try
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string script = Path.Combine(baseDir, "app", "App.ps1");
            if (!File.Exists(script))
            {
                MessageBox.Show("找不到 app\\App.ps1。请保持整个文件夹完整。", "本地文件清理助手",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string ps = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.System),
                @"WindowsPowerShell\v1.0\powershell.exe");
            if (!File.Exists(ps))
                ps = "powershell.exe";

            var psi = new ProcessStartInfo();
            psi.FileName = ps;
            psi.Arguments = "-STA -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"" +
                            script + "\"";
            psi.WorkingDirectory = baseDir;
            psi.UseShellExecute = false;
            psi.CreateNoWindow = true;
            Process.Start(psi);
        }
        catch (Exception ex)
        {
            MessageBox.Show("启动失败：\r\n" + ex.Message, "本地文件清理助手",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
