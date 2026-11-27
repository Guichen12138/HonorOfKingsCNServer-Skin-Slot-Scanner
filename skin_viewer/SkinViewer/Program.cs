namespace SkinViewer;

internal static class Program
{
    [STAThread]
    static int Main(string[] args)
    {
        string baseDir = AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        AppConfig.CleanupOldUpgradeFiles(baseDir);

        // 支持命令行静默安装（方便自动化验证与批量部署）：Setup.exe --silent-install <dir> [--with-shortcuts]
        int silentIdx = Array.FindIndex(args, a => string.Equals(a, "--silent-install", StringComparison.OrdinalIgnoreCase));
        if (silentIdx >= 0 && silentIdx + 1 < args.Length)
        {
            bool withShortcuts = args.Any(a => string.Equals(a, "--with-shortcuts", StringComparison.OrdinalIgnoreCase));
            try
            {
                SetupForm.PerformInstall(
                    args[silentIdx + 1],
                    createDesktopShortcut: withShortcuts,
                    createStartMenuAndRegistry: withShortcuts);
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.Message);
                return 1;
            }
        }

        ApplicationConfiguration.Initialize();

        bool forceSetup = args.Any(a => string.Equals(a, "--setup", StringComparison.OrdinalIgnoreCase));
        bool forceViewer = args.Any(a => string.Equals(a, "--viewer", StringComparison.OrdinalIgnoreCase));
        bool justUpgraded = args.Any(a => string.Equals(a, "--upgraded", StringComparison.OrdinalIgnoreCase));

        string exeName = Path.GetFileName(Environment.ProcessPath ?? "");
        bool nameIsSetup = exeName.Contains("Setup", StringComparison.OrdinalIgnoreCase)
                        || exeName.Contains("安装", StringComparison.OrdinalIgnoreCase);

        bool hasInstalledFiles = File.Exists(Path.Combine(baseDir, AppConfig.ExcelFileName))
                              || File.Exists(Path.Combine(baseDir, AppConfig.UninstallFileName));

        bool runAsSetup = forceSetup || (!forceViewer && (nameIsSetup || !hasInstalledFiles));

        if (runAsSetup)
        {
            Application.Run(new SetupForm());
        }
        else
        {
            Application.Run(new ViewerForm(justUpgraded));
        }

        return 0;
    }
}
