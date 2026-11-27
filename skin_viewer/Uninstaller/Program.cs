using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Win32;

internal static partial class Program
{
    const string UninstallRegSubKey = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\HonorOfKingsSkinViewer";
    const string AppTitle = "王者荣耀皮肤格子查看器 - 卸载程序";
    const string ViewerExeName = "王者荣耀皮肤格子查看器.exe";
    const string ExcelFileName = "hero_skin_scan.xlsx";

    const uint MB_YESNO = 0x00000004;
    const uint MB_ICONQUESTION = 0x00000020;
    const uint MB_ICONINFORMATION = 0x00000040;
    const int IDYES = 6;

    [LibraryImport("user32.dll", EntryPoint = "MessageBoxW", StringMarshalling = StringMarshalling.Utf16)]
    private static partial int MessageBoxW(IntPtr hWnd, string lpText, string lpCaption, uint uType);

    [LibraryImport("user32.dll", EntryPoint = "SetProcessDpiAwarenessContext")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool SetProcessDpiAwarenessContext(IntPtr value);

    [LibraryImport("user32.dll", EntryPoint = "SetProcessDPIAware")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool SetProcessDPIAware();

    static int Main(string[] args)
    {
        try
        {
            // DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4
            if (!SetProcessDpiAwarenessContext(new IntPtr(-4)))
            {
                SetProcessDPIAware();
            }
        }
        catch { }

        bool silent = args.Any(a => string.Equals(a, "/S", StringComparison.OrdinalIgnoreCase)
                                 || string.Equals(a, "--silent", StringComparison.OrdinalIgnoreCase));

        string installDir = AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        string selfPath = Environment.ProcessPath ?? Path.Combine(installDir, "Uninstall.exe");

        if (!silent)
        {
            int ans = MessageBoxW(
                IntPtr.Zero,
                $"确定要卸载「王者荣耀皮肤格子查看器」吗？\n\n安装位置：{installDir}\n\n这会移除程序文件、本地 Excel 表格以及桌面/开始菜单快捷方式。",
                AppTitle,
                MB_YESNO | MB_ICONQUESTION);

            if (ans != IDYES)
            {
                return 0;
            }
        }

        // 1) 关闭正在运行的查看器进程（按路径匹配）
        string viewerPath = Path.Combine(installDir, ViewerExeName);
        try
        {
            foreach (var proc in Process.GetProcesses())
            {
                try
                {
                    if (proc.Id == Environment.ProcessId) continue;
                    string? p = proc.MainModule?.FileName;
                    if (p is not null && string.Equals(p, viewerPath, StringComparison.OrdinalIgnoreCase))
                    {
                        proc.Kill();
                        proc.WaitForExit(3000);
                    }
                }
                catch { }
                finally { proc.Dispose(); }
            }
        }
        catch { }

        // 2) 删除桌面与开始菜单快捷方式
        TryDeleteFile(Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
            "王者荣耀皮肤格子查看器.lnk"));

        string startMenuDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Programs),
            "王者荣耀皮肤格子查看器");
        TryDeleteDirectory(startMenuDir);
        TryDeleteFile(Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Programs),
            "王者荣耀皮肤格子查看器.lnk"));

        // 3) 删除注册表卸载项
        try
        {
            Registry.CurrentUser.DeleteSubKeyTree(UninstallRegSubKey, throwOnMissingSubKey: false);
        }
        catch { }

        // 4) 删除程序与数据文件（不动用户自己放在目录里的其他无关文件）
        foreach (string target in new[]
        {
            viewerPath,
            viewerPath + ".old",
            viewerPath + ".new",
            viewerPath + ".tmp",
            Path.Combine(installDir, ExcelFileName),
            Path.Combine(installDir, ExcelFileName + ".tmp"),
            selfPath + ".new",
        })
        {
            TryDeleteFile(target);
        }

        if (!silent)
        {
            MessageBoxW(
                IntPtr.Zero,
                "「王者荣耀皮肤格子查看器」已卸载完成。",
                AppTitle,
                MB_ICONINFORMATION);
        }

        // 5) 退出后静默自删 Uninstall.exe，并在目录已空时移除空目录（不用 rmdir /s，绝不误删用户无关文件）
        try
        {
            string cmdArgs = $"/c timeout /t 1 /nobreak >nul & del /f /q \"{selfPath}\" >nul 2>&1 & rmdir \"{installDir}\" >nul 2>&1";
            Process.Start(new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = cmdArgs,
                WorkingDirectory = Path.GetTempPath(),
                CreateNoWindow = true,
                UseShellExecute = false,
            });
        }
        catch { }

        return 0;
    }

    static void TryDeleteFile(string path)
    {
        for (int i = 0; i < 5; i++)
        {
            try
            {
                if (!File.Exists(path)) return;
                File.SetAttributes(path, FileAttributes.Normal);
                File.Delete(path);
                return;
            }
            catch
            {
                Thread.Sleep(150);
            }
        }
    }

    static void TryDeleteDirectory(string dir)
    {
        try
        {
            if (Directory.Exists(dir))
            {
                Directory.Delete(dir, recursive: true);
            }
        }
        catch { }
    }
}
