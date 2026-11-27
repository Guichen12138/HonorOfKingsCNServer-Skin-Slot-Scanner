using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Reflection;

namespace SkinViewer;

internal sealed class SetupForm : Form
{
    private readonly TextBox _txtPath;
    private readonly CheckBox _chkDesktop;
    private readonly CheckBox _chkStartMenu;
    private readonly CheckBox _chkLaunch;
    private readonly ProgressBar _progress;
    private readonly Label _lblStatus;
    private readonly Button _btnInstall;
    private readonly Button _btnBrowse;
    private readonly Button _btnCancel;

    private int S(int px) => LogicalToDeviceUnits(px);

    public SetupForm()
    {
        SuspendLayout();

        Text = $"{AppConfig.AppName} v{AppConfig.CurrentVersion} - 安装向导";
        Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Regular);
        AutoScaleMode = AutoScaleMode.None;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = Color.FromArgb(244, 246, 251);
        ClientSize = new Size(S(640), S(430));

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            Margin = Padding.Empty,
            Padding = Padding.Empty,
        };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 0: 渐变横幅
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); // 1: 主体卡片区
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 2: 底部按钮栏

        // 0) 顶部渐变横幅
        var header = new Panel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Margin = Padding.Empty,
            Padding = new Padding(S(24), S(18), S(24), S(18)),
        };
        header.Paint += (_, e) =>
        {
            if (header.ClientRectangle.Width <= 0 || header.ClientRectangle.Height <= 0) return;
            using var brush = new LinearGradientBrush(
                header.ClientRectangle,
                Color.FromArgb(90, 79, 243),
                Color.FromArgb(56, 168, 255),
                LinearGradientMode.Horizontal);
            e.Graphics.FillRectangle(brush, header.ClientRectangle);
        };

        var headerStack = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            BackColor = Color.Transparent,
            Margin = Padding.Empty,
        };
        headerStack.Controls.Add(new Label
        {
            Text = $"{AppConfig.AppName} 安装向导",
            ForeColor = Color.White,
            BackColor = Color.Transparent,
            Font = new Font("Microsoft YaHei UI", 14.5F, FontStyle.Bold),
            AutoSize = true,
            Margin = new Padding(0, 0, 0, S(4)),
        });
        headerStack.Controls.Add(new Label
        {
            Text = $"版本 v{AppConfig.CurrentVersion} · 一键部署皮肤表格同步查看器、初始皮肤表格与卸载程序",
            ForeColor = Color.FromArgb(228, 232, 255),
            BackColor = Color.Transparent,
            Font = new Font("Microsoft YaHei UI", 9F),
            AutoSize = true,
            Margin = Padding.Empty,
        });
        header.Controls.Add(headerStack);
        root.Controls.Add(header, 0, 0);

        // 1) 主体白色卡片区
        var bodyWrap = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(S(22), S(18), S(22), S(10)),
            Margin = Padding.Empty,
        };

        var card = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.White,
            Padding = new Padding(S(20), S(16), S(20), S(16)),
        };
        card.Paint += (_, e) =>
        {
            using var pen = new Pen(Color.FromArgb(223, 228, 236));
            e.Graphics.DrawRectangle(pen, 0, 0, card.Width - 1, card.Height - 1);
        };

        var cardLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 6,
            BackColor = Color.White,
            Margin = Padding.Empty,
        };
        cardLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        cardLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 0: 路径标题
        cardLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 1: 输入框 + 浏览按钮
        cardLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 2: 复选框组
        cardLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); // 3: 弹性空白
        cardLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 4: 状态文字
        cardLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 5: 进度条

        var lblPathTitle = new Label
        {
            Text = "选择安装文件夹：",
            Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold),
            ForeColor = Color.FromArgb(10, 37, 64),
            AutoSize = true,
            Margin = new Padding(0, 0, 0, S(6)),
        };
        cardLayout.Controls.Add(lblPathTitle, 0, 0);

        var pathRow = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            ColumnCount = 2,
            RowCount = 1,
            Margin = new Padding(0, 0, 0, S(14)),
        };
        pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        _txtPath = new TextBox
        {
            Text = AppConfig.DefaultInstallDir(),
            Dock = DockStyle.Fill,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Font = new Font("Microsoft YaHei UI", 9.5F),
            Margin = new Padding(0, S(2), S(10), S(2)),
        };

        _btnBrowse = new Button
        {
            Text = "浏览...",
            AutoSize = true,
            MinimumSize = new Size(S(86), S(30)),
            FlatStyle = FlatStyle.System,
            Margin = Padding.Empty,
        };
        _btnBrowse.Click += OnBrowseClicked;

        pathRow.Controls.Add(_txtPath, 0, 0);
        pathRow.Controls.Add(_btnBrowse, 1, 0);
        cardLayout.Controls.Add(pathRow, 0, 1);

        var chkStack = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            Margin = new Padding(S(2), 0, 0, S(8)),
        };

        _chkDesktop = new CheckBox
        {
            Text = "创建桌面快捷方式",
            Checked = true,
            AutoSize = true,
            Margin = new Padding(0, 0, 0, S(6)),
        };

        _chkStartMenu = new CheckBox
        {
            Text = "创建开始菜单快捷方式与系统「已安装应用」卸载项",
            Checked = true,
            AutoSize = true,
            Margin = new Padding(0, 0, 0, S(6)),
        };

        _chkLaunch = new CheckBox
        {
            Text = "安装完成后立即启动查看器",
            Checked = true,
            AutoSize = true,
            Margin = Padding.Empty,
        };

        chkStack.Controls.Add(_chkDesktop);
        chkStack.Controls.Add(_chkStartMenu);
        chkStack.Controls.Add(_chkLaunch);
        cardLayout.Controls.Add(chkStack, 0, 2);

        cardLayout.Controls.Add(new Panel { Dock = DockStyle.Fill, Height = S(4) }, 0, 3);

        _lblStatus = new Label
        {
            Text = "安装后目录包含：查看器主程序、hero_skin_scan.xlsx 与 Uninstall.exe",
            ForeColor = Color.FromArgb(66, 84, 102),
            AutoSize = true,
            Margin = new Padding(0, 0, 0, S(6)),
        };
        cardLayout.Controls.Add(_lblStatus, 0, 4);

        _progress = new ProgressBar
        {
            Dock = DockStyle.Top,
            Height = S(12),
            Minimum = 0,
            Maximum = 100,
            Value = 0,
            Margin = Padding.Empty,
        };
        cardLayout.Controls.Add(_progress, 0, 5);

        card.Controls.Add(cardLayout);
        bodyWrap.Controls.Add(card);
        root.Controls.Add(bodyWrap, 0, 1);

        // 2) 底部按钮栏
        var bottomBar = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            FlowDirection = FlowDirection.RightToLeft,
            WrapContents = false,
            Padding = new Padding(S(22), S(4), S(22), S(16)),
            Margin = Padding.Empty,
        };

        _btnCancel = new Button
        {
            Text = "取消",
            AutoSize = true,
            MinimumSize = new Size(S(92), S(36)),
            FlatStyle = FlatStyle.System,
            Margin = Padding.Empty,
        };
        _btnCancel.Click += (_, _) => Close();

        _btnInstall = new Button
        {
            Text = "立即安装",
            AutoSize = true,
            MinimumSize = new Size(S(128), S(36)),
            BackColor = Color.FromArgb(99, 91, 255),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold),
            Cursor = Cursors.Hand,
            Margin = new Padding(0, 0, S(10), 0),
        };
        _btnInstall.FlatAppearance.BorderSize = 0;
        _btnInstall.Click += OnInstallClicked;

        bottomBar.Controls.Add(_btnCancel);
        bottomBar.Controls.Add(_btnInstall);
        root.Controls.Add(bottomBar, 0, 2);

        Controls.Add(root);
        AcceptButton = _btnInstall;
        CancelButton = _btnCancel;

        ResumeLayout(performLayout: true);
    }

    private void OnBrowseClicked(object? sender, EventArgs e)
    {
        using var dlg = new FolderBrowserDialog
        {
            Description = "选择安装目录",
            UseDescriptionForTitle = true,
            SelectedPath = _txtPath.Text,
        };

        if (dlg.ShowDialog(this) == DialogResult.OK && !string.IsNullOrWhiteSpace(dlg.SelectedPath))
        {
            string chosen = dlg.SelectedPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string leaf = Path.GetFileName(chosen);
            if (!leaf.Equals(AppConfig.AppName, StringComparison.OrdinalIgnoreCase)
                && !leaf.Equals("HonorOfKingsSkinViewer", StringComparison.OrdinalIgnoreCase))
            {
                chosen = Path.Combine(chosen, AppConfig.AppName);
            }
            _txtPath.Text = chosen;
        }
    }

    private async void OnInstallClicked(object? sender, EventArgs e)
    {
        string targetDir = _txtPath.Text.Trim();
        if (string.IsNullOrEmpty(targetDir))
        {
            MessageBox.Show(this, "请先选择有效的安装路径。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        _btnInstall.Enabled = false;
        _btnBrowse.Enabled = false;
        _btnCancel.Enabled = false;
        _txtPath.ReadOnly = true;

        try
        {
            bool createDesktop = _chkDesktop.Checked;
            bool createStartMenu = _chkStartMenu.Checked;
            bool launchNow = _chkLaunch.Checked;

            var progress = new Progress<(int pct, string msg)>(p =>
            {
                if (IsDisposed) return;
                _progress.Value = Math.Clamp(p.pct, 0, 100);
                _lblStatus.Text = p.msg;
            });

            string installedExe = await Task.Run(() =>
                PerformInstall(targetDir, createDesktop, createStartMenu, progress));

            if (IsDisposed) return;
            _progress.Value = 100;
            _lblStatus.Text = "安装完成！";

            if (launchNow)
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = installedExe,
                    WorkingDirectory = targetDir,
                    UseShellExecute = true,
                });
            }
            else
            {
                MessageBox.Show(
                    this,
                    $"安装成功！\n\n安装目录：{targetDir}",
                    AppConfig.AppName,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }

            Close();
        }
        catch (Exception ex)
        {
            if (IsDisposed) return;
            _btnInstall.Enabled = true;
            _btnBrowse.Enabled = true;
            _btnCancel.Enabled = true;
            _txtPath.ReadOnly = false;
            _lblStatus.Text = "安装失败：" + ex.Message;

            MessageBox.Show(
                this,
                $"安装过程中出错：\n{ex.Message}\n\n建议更换为非系统保护目录（例如：{AppConfig.DefaultInstallDir()}）后重试。",
                "安装失败",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    public static string PerformInstall(
        string targetDir,
        bool createDesktopShortcut,
        bool createStartMenuAndRegistry,
        IProgress<(int pct, string msg)>? progress = null)
    {
        targetDir = Path.GetFullPath(targetDir);
        progress?.Report((10, "正在创建安装目录……"));
        Directory.CreateDirectory(targetDir);

        string selfExe = Environment.ProcessPath
            ?? throw new InvalidOperationException("无法定位当前安装程序路径。");

        string dstExe = Path.Combine(targetDir, AppConfig.ExeFileName);
        string dstExcel = Path.Combine(targetDir, AppConfig.ExcelFileName);
        string dstUninstall = Path.Combine(targetDir, AppConfig.UninstallFileName);

        progress?.Report((35, "正在释放查看器主程序……"));
        if (!string.Equals(selfExe, dstExe, StringComparison.OrdinalIgnoreCase))
        {
            AppConfig.CloseRunningProcessAt(dstExe);
            for (int attempt = 0; ; attempt++)
            {
                try
                {
                    File.Copy(selfExe, dstExe, overwrite: true);
                    break;
                }
                catch (IOException) when (attempt < 12)
                {
                    Thread.Sleep(200);
                    AppConfig.CloseRunningProcessAt(dstExe);
                }
            }
        }

        var asm = Assembly.GetExecutingAssembly();

        progress?.Report((60, "正在部署 hero_skin_scan.xlsx 数据表……"));
        using (var excelStream = asm.GetManifestResourceStream("Payload.Excel"))
        {
            if (excelStream is not null)
            {
                using var ms = new MemoryStream();
                excelStream.CopyTo(ms);
                byte[] embeddedBytes = ms.ToArray();

                bool shouldWriteExcel = !File.Exists(dstExcel);
                if (!shouldWriteExcel && !AppConfig.IsFileLockedByExcel(dstExcel))
                {
                    try
                    {
                        var localSnap = XlsxParser.Parse(dstExcel);
                        var embeddedSnap = XlsxParser.Parse(embeddedBytes);
                        string localLatest = localSnap.LatestSkin?.LastModified ?? "";
                        string embLatest = embeddedSnap.LatestSkin?.LastModified ?? "";
                        // 仅当安装包内置表比本地旧表更新/更全时才覆盖；若用户已在线同步过更新的数据则保留用户本地表
                        if (string.CompareOrdinal(embLatest, localLatest) > 0
                            || (string.Equals(embLatest, localLatest, StringComparison.Ordinal)
                                && embeddedSnap.SkinCount > localSnap.SkinCount))
                        {
                            shouldWriteExcel = true;
                        }
                    }
                    catch
                    {
                        // 本地旧文件损坏时自动修复为内置初始表
                        shouldWriteExcel = true;
                    }
                }

                if (shouldWriteExcel)
                {
                    string tmpExcel = dstExcel + ".tmp";
                    File.WriteAllBytes(tmpExcel, embeddedBytes);
                    File.Move(tmpExcel, dstExcel, overwrite: true);
                }
            }
        }

        progress?.Report((80, "正在生成卸载程序 Uninstall.exe……"));
        using (var uninstStream = asm.GetManifestResourceStream("Payload.Uninstall"))
        {
            if (uninstStream is not null)
            {
                AppConfig.CloseRunningProcessAt(dstUninstall);
                using var fs = new FileStream(dstUninstall, FileMode.Create, FileAccess.Write, FileShare.None);
                uninstStream.CopyTo(fs);
            }
        }

        progress?.Report((90, "正在配置快捷方式与卸载信息……"));
        if (createStartMenuAndRegistry)
        {
            AppConfig.WriteUninstallRegistry(targetDir);

            string menuDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Programs),
                AppConfig.AppName);
            AppConfig.CreateShortcut(
                Path.Combine(menuDir, $"{AppConfig.AppName}.lnk"),
                dstExe, targetDir, "打开王者荣耀皮肤格子查看器");
            if (File.Exists(dstUninstall))
            {
                AppConfig.CreateShortcut(
                    Path.Combine(menuDir, $"卸载 {AppConfig.AppName}.lnk"),
                    dstUninstall, targetDir, "卸载王者荣耀皮肤格子查看器");
            }
        }

        if (createDesktopShortcut)
        {
            string desktopLnk = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                $"{AppConfig.AppName}.lnk");
            AppConfig.CreateShortcut(desktopLnk, dstExe, targetDir, "打开王者荣耀皮肤格子查看器并同步最新表格");
        }

        progress?.Report((100, "安装完成"));
        return dstExe;
    }
}
