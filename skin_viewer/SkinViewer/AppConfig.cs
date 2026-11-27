using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace SkinViewer;

internal static class AppConfig
{
    public const string AppName = "王者荣耀皮肤格子查看器";
    public const string ExeFileName = "王者荣耀皮肤格子查看器.exe";
    public const string UninstallFileName = "Uninstall.exe";
    public const string ExcelFileName = "hero_skin_scan.xlsx";

    public const string RepoOwner = "Guichen12138";
    public const string RepoName = "HonorOfKingsCNServer-Skin-Slot-Scanner";
    public const string RepoBranch = "main";
    public const string RepoWebUrl = $"https://github.com/{RepoOwner}/{RepoName}";

    public const string UninstallRegSubKey = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\HonorOfKingsSkinViewer";

    public static string CurrentVersion
    {
        get
        {
            var v = Assembly.GetExecutingAssembly().GetName().Version;
            return v is null ? "1.0.0" : $"{v.Major}.{v.Minor}.{v.Build}";
        }
    }

    public static string DefaultInstallDir()
    {
        // 优先选非系统固定盘根目录下的应用文件夹，没有第二块盘则放 LocalAppData，避开 Program Files 权限限制
        try
        {
            string sysRoot = Path.GetPathRoot(Environment.SystemDirectory) ?? string.Empty;
            foreach (var drive in DriveInfo.GetDrives())
            {
                if (drive.DriveType == DriveType.Fixed
                    && drive.IsReady
                    && !string.Equals(drive.Name, sysRoot, StringComparison.OrdinalIgnoreCase)
                    && drive.AvailableFreeSpace > 200L * 1024 * 1024)
                {
                    return Path.Combine(drive.Name, AppName);
                }
            }
        }
        catch { }

        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            AppName);
    }

    public static void WriteUninstallRegistry(string installDir)
    {
        string exePath = Path.Combine(installDir, ExeFileName);
        string uninstallPath = Path.Combine(installDir, UninstallFileName);

        using var key = Registry.CurrentUser.CreateSubKey(UninstallRegSubKey);
        if (key is null) return;

        long totalBytes = 0;
        try
        {
            foreach (var f in Directory.GetFiles(installDir))
            {
                totalBytes += new FileInfo(f).Length;
            }
        }
        catch { }

        key.SetValue("DisplayName", AppName);
        key.SetValue("DisplayVersion", CurrentVersion);
        key.SetValue("Publisher", RepoOwner);
        key.SetValue("InstallLocation", installDir);
        key.SetValue("DisplayIcon", $"\"{exePath}\",0");
        key.SetValue("UninstallString", $"\"{uninstallPath}\"");
        key.SetValue("QuietUninstallString", $"\"{uninstallPath}\" /S");
        key.SetValue("URLInfoAbout", RepoWebUrl);
        key.SetValue("NoModify", 1, RegistryValueKind.DWord);
        key.SetValue("NoRepair", 1, RegistryValueKind.DWord);
        key.SetValue("EstimatedSize", (int)Math.Max(1024, totalBytes / 1024), RegistryValueKind.DWord);
    }

    public static bool CreateShortcut(string shortcutPath, string targetPath, string workingDir, string description)
    {
        object? shell = null;
        object? shortcut = null;
        try
        {
            string? dir = Path.GetDirectoryName(shortcutPath);
            if (!string.IsNullOrEmpty(dir))
            {
                Directory.CreateDirectory(dir);
            }

            Type? shellType = Type.GetTypeFromProgID("WScript.Shell");
            if (shellType is null) return false;

            shell = Activator.CreateInstance(shellType);
            if (shell is null) return false;

            shortcut = shellType.InvokeMember(
                "CreateShortcut",
                BindingFlags.InvokeMethod,
                binder: null,
                target: shell,
                args: new object[] { shortcutPath });
            if (shortcut is null) return false;

            Type scType = shortcut.GetType();
            scType.InvokeMember("TargetPath", BindingFlags.SetProperty, null, shortcut, new object[] { targetPath });
            scType.InvokeMember("WorkingDirectory", BindingFlags.SetProperty, null, shortcut, new object[] { workingDir });
            scType.InvokeMember("Description", BindingFlags.SetProperty, null, shortcut, new object[] { description });
            scType.InvokeMember("IconLocation", BindingFlags.SetProperty, null, shortcut, new object[] { $"{targetPath},0" });
            scType.InvokeMember("Save", BindingFlags.InvokeMethod, null, shortcut, null);

            return File.Exists(shortcutPath);
        }
        catch
        {
            return false;
        }
        finally
        {
            if (shortcut is not null && Marshal.IsComObject(shortcut))
            {
                try { Marshal.FinalReleaseComObject(shortcut); } catch { }
            }
            if (shell is not null && Marshal.IsComObject(shell))
            {
                try { Marshal.FinalReleaseComObject(shell); } catch { }
            }
        }
    }

    public static void CloseRunningProcessAt(string targetExePath)
    {
        if (!File.Exists(targetExePath)) return;
        string fullTarget = Path.GetFullPath(targetExePath);

        try
        {
            foreach (var proc in Process.GetProcesses())
            {
                try
                {
                    if (proc.Id == Environment.ProcessId) continue;
                    string? modPath = proc.MainModule?.FileName;
                    if (modPath is not null && string.Equals(modPath, fullTarget, StringComparison.OrdinalIgnoreCase))
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
    }

    public static void CleanupOldUpgradeFiles(string dir)
    {
        var baseNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ExeFileName,
            UninstallFileName,
        };

        string currentName = Path.GetFileName(Environment.ProcessPath ?? "");
        if (!string.IsNullOrEmpty(currentName))
        {
            baseNames.Add(currentName);
        }

        var candidates = new List<string>
        {
            Path.Combine(dir, ExcelFileName + ".tmp"),
        };

        foreach (string name in baseNames)
        {
            foreach (string suffix in new[] { ".old", ".new", ".tmp" })
            {
                candidates.Add(Path.Combine(dir, name + suffix));
            }
        }

        foreach (string p in candidates)
        {
            if (!File.Exists(p)) continue;

            for (int i = 0; i < 6; i++)
            {
                try
                {
                    File.SetAttributes(p, FileAttributes.Normal);
                    File.Delete(p);
                    break;
                }
                catch
                {
                    Thread.Sleep(200);
                }
            }
        }
    }

    public static bool IsFileLockedByExcel(string path)
    {
        if (!File.Exists(path)) return false;

        string dir = Path.GetDirectoryName(path) ?? "";
        string lockFile = Path.Combine(dir, "~$" + Path.GetFileName(path));
        if (File.Exists(lockFile)) return true;

        try
        {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            return false;
        }
        catch (IOException)
        {
            return true;
        }
        catch
        {
            return false;
        }
    }
}
