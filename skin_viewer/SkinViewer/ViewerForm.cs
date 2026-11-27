using System.Diagnostics;
using System.Drawing.Drawing2D;

namespace SkinViewer;

internal sealed class ViewerForm : Form
{
    private readonly string _installDir;
    private readonly string _excelPath;
    private readonly CancellationTokenSource _lifetimeCts = new();

    private readonly Label _lblPill;
    private readonly Label _lblLocalStats;
    private readonly Label _lblLatestRecord;

    private readonly Panel _appUpdateBanner;
    private readonly Label _lblAppUpdateInfo;
    private readonly Button _btnUpgradeApp;

    private readonly Button _btnCheck;
    private readonly Button _btnApplyExcel;
    private readonly Button _btnOpenExcel;
    private readonly Button _btnOpenFolder;

    private readonly ProgressBar _progress;
    private readonly TextBox _txtLog;

    private byte[]? _pendingExcelBytes;
    private ViewerUpdateInfo? _pendingAppUpdate;
    private bool _busy;

    private int S(int px) => LogicalToDeviceUnits(px);

    public ViewerForm(bool justUpgraded = false)
    {
        _installDir = AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        _excelPath = Path.Combine(_installDir, AppConfig.ExcelFileName);

        // 启动时顺手刷新已存在卸载注册表中的当前版本号（仅当该项已注册且指向本目录时更新，绝不在免安装目录乱写注册表）
        try
        {
            using var existingKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(AppConfig.UninstallRegSubKey);
            string? regLoc = existingKey?.GetValue("InstallLocation") as string;
            if (regLoc is not null
                && string.Equals(
                    regLoc.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
                    _installDir,
                    StringComparison.OrdinalIgnoreCase))
            {
                AppConfig.WriteUninstallRegistry(_installDir);
            }
        }
        catch { }

        SuspendLayout();

        Text = $"{AppConfig.AppName} v{AppConfig.CurrentVersion}";
        Font = new Font("Microsoft YaHei UI", 9.5F);
        AutoScaleMode = AutoScaleMode.None;
        ClientSize = new Size(S(840), S(620));
        MinimumSize = new Size(S(740), S(520));
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = Color.FromArgb(244, 246, 251);

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(S(18), S(16), S(18), S(12)),
            ColumnCount = 1,
            RowCount = 6,
        };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 0: Header
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 1: App self-update banner
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 2: Stats card
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 3: Action buttons
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); // 4: Log area
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 5: Footer

        // 0) 顶部渐变标题栏
        var header = new Panel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Margin = new Padding(0, 0, 0, S(10)),
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

        var headerLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            BackColor = Color.Transparent,
            Padding = new Padding(S(20), S(14), S(20), S(14)),
            ColumnCount = 2,
            RowCount = 1,
            Margin = Padding.Empty,
        };
        headerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        headerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        var titleStack = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            BackColor = Color.Transparent,
            Margin = Padding.Empty,
        };
        titleStack.Controls.Add(new Label
        {
            Text = $"{AppConfig.AppName}  v{AppConfig.CurrentVersion}",
            ForeColor = Color.White,
            Font = new Font("Microsoft YaHei UI", 14.5F, FontStyle.Bold),
            AutoSize = true,
            Margin = new Padding(0, 0, 0, S(3)),
        });
        titleStack.Controls.Add(new Label
        {
            Text = "从 GitHub 仓库实时同步最新王者荣耀英雄皮肤格子表 · 支持查看器在线自更新",
            ForeColor = Color.FromArgb(228, 232, 255),
            Font = new Font("Microsoft YaHei UI", 9F),
            AutoSize = true,
            Margin = Padding.Empty,
        });

        _lblPill = new Label
        {
            Text = "就绪",
            AutoSize = true,
            Padding = new Padding(S(12), S(5), S(12), S(5)),
            Anchor = AnchorStyles.Right,
            Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold),
            BackColor = Color.FromArgb(238, 241, 246),
            ForeColor = Color.FromArgb(66, 84, 102),
            Margin = new Padding(S(12), 0, 0, 0),
        };

        headerLayout.Controls.Add(titleStack, 0, 0);
        headerLayout.Controls.Add(_lblPill, 1, 0);
        header.Controls.Add(headerLayout);
        root.Controls.Add(header, 0, 0);

        // 1) 查看器程序自更新提示横幅（默认隐藏，发现新版查看器时自动浮现）
        _appUpdateBanner = new Panel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Margin = new Padding(0, 0, 0, S(10)),
            BackColor = Color.FromArgb(238, 240, 255),
            Visible = false,
        };
        _appUpdateBanner.Paint += (_, e) =>
        {
            using var pen = new Pen(Color.FromArgb(180, 174, 255));
            e.Graphics.DrawRectangle(pen, 0, 0, _appUpdateBanner.Width - 1, _appUpdateBanner.Height - 1);
        };
        var bannerLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Padding = new Padding(S(14), S(8), S(14), S(8)),
            ColumnCount = 2,
            RowCount = 1,
            BackColor = Color.Transparent,
            Margin = Padding.Empty,
        };
        bannerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        bannerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        _lblAppUpdateInfo = new Label
        {
            Text = "发现查看器新版本！",
            ForeColor = Color.FromArgb(55, 48, 163),
            Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold),
            AutoSize = true,
            Anchor = AnchorStyles.Left,
            Margin = new Padding(0, 0, S(10), 0),
        };
        _btnUpgradeApp = new Button
        {
            Text = "立即升级查看器",
            AutoSize = true,
            MinimumSize = new Size(S(136), S(32)),
            BackColor = Color.FromArgb(99, 91, 255),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold),
            Cursor = Cursors.Hand,
            Anchor = AnchorStyles.Right,
            Margin = Padding.Empty,
        };
        _btnUpgradeApp.FlatAppearance.BorderSize = 0;
        _btnUpgradeApp.Click += OnUpgradeViewerClicked;

        bannerLayout.Controls.Add(_lblAppUpdateInfo, 0, 0);
        bannerLayout.Controls.Add(_btnUpgradeApp, 1, 0);
        _appUpdateBanner.Controls.Add(bannerLayout);
        root.Controls.Add(_appUpdateBanner, 0, 1);

        // 2) 本地表数据统计卡
        var statsCard = new Panel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Margin = new Padding(0, 0, 0, S(10)),
            BackColor = Color.White,
            Padding = new Padding(S(16), S(12), S(16), S(12)),
        };
        statsCard.Paint += (_, e) =>
        {
            using var pen = new Pen(Color.FromArgb(223, 228, 236));
            e.Graphics.DrawRectangle(pen, 0, 0, statsCard.Width - 1, statsCard.Height - 1);
        };

        var statsStack = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            BackColor = Color.White,
            Margin = Padding.Empty,
        };
        _lblLocalStats = new Label
        {
            Text = "本地表格统计：读取中……",
            Font = new Font("Microsoft YaHei UI", 10.5F, FontStyle.Bold),
            ForeColor = Color.FromArgb(10, 37, 64),
            AutoSize = true,
            Margin = new Padding(0, 0, 0, S(4)),
        };
        _lblLatestRecord = new Label
        {
            Text = "最近入表记录：—",
            Font = new Font("Microsoft YaHei UI", 9F),
            ForeColor = Color.FromArgb(66, 84, 102),
            AutoSize = true,
            Margin = Padding.Empty,
        };
        statsStack.Controls.Add(_lblLocalStats);
        statsStack.Controls.Add(_lblLatestRecord);
        statsCard.Controls.Add(statsStack);
        root.Controls.Add(statsCard, 0, 2);

        // 3) 操作按钮行
        var btnRow = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = true,
            Margin = new Padding(0, 0, 0, S(10)),
        };

        _btnCheck = new Button
        {
            Text = "检查更新",
            AutoSize = true,
            MinimumSize = new Size(S(130), S(38)),
            BackColor = Color.FromArgb(99, 91, 255),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold),
            Cursor = Cursors.Hand,
            Margin = new Padding(0, 0, S(10), S(2)),
        };
        _btnCheck.FlatAppearance.BorderSize = 0;
        _btnCheck.Click += async (_, _) => await CheckAllUpdatesAsync();

        _btnApplyExcel = new Button
        {
            Text = "更新本地 Excel 表格",
            AutoSize = true,
            MinimumSize = new Size(S(172), S(38)),
            BackColor = Color.FromArgb(18, 183, 106),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold),
            Cursor = Cursors.Hand,
            Enabled = false,
            Margin = new Padding(0, 0, S(10), S(2)),
        };
        _btnApplyExcel.FlatAppearance.BorderSize = 0;
        _btnApplyExcel.Click += OnApplyExcelClicked;

        _btnOpenExcel = new Button
        {
            Text = "打开 Excel 表格",
            AutoSize = true,
            MinimumSize = new Size(S(142), S(38)),
            FlatStyle = FlatStyle.System,
            Margin = new Padding(0, 0, S(10), S(2)),
        };
        _btnOpenExcel.Click += OnOpenExcelClicked;

        _btnOpenFolder = new Button
        {
            Text = "打开所在文件夹",
            AutoSize = true,
            MinimumSize = new Size(S(142), S(38)),
            FlatStyle = FlatStyle.System,
            Margin = new Padding(0, 0, 0, S(2)),
        };
        _btnOpenFolder.Click += (_, _) =>
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = _installDir,
                    UseShellExecute = true,
                });
            }
            catch { }
        };

        btnRow.Controls.AddRange(new Control[] { _btnCheck, _btnApplyExcel, _btnOpenExcel, _btnOpenFolder });
        root.Controls.Add(btnRow, 0, 3);

        // 4) 进度条 + 详情输出框
        var logPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            Margin = new Padding(0, 0, 0, S(8)),
        };
        logPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        logPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        logPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

        _progress = new ProgressBar
        {
            Dock = DockStyle.Top,
            Height = S(10),
            Margin = new Padding(0, 0, 0, S(6)),
        };

        _txtLog = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            BackColor = Color.White,
            ForeColor = Color.FromArgb(37, 54, 74),
            Font = new Font("Microsoft YaHei UI", 9.2F),
            Margin = Padding.Empty,
        };

        logPanel.Controls.Add(_progress, 0, 0);
        logPanel.Controls.Add(_txtLog, 0, 1);
        root.Controls.Add(logPanel, 0, 4);

        // 5) 底部栏
        var footer = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            ColumnCount = 2,
            RowCount = 1,
            Margin = Padding.Empty,
        };
        footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        footer.Controls.Add(new Label
        {
            Text = $"目录：{_installDir}",
            ForeColor = Color.FromArgb(102, 112, 133),
            Font = new Font("Microsoft YaHei UI", 8.5F),
            AutoSize = true,
            Anchor = AnchorStyles.Left,
            Margin = Padding.Empty,
        }, 0, 0);

        var linkRepo = new LinkLabel
        {
            Text = "访问 GitHub 数据仓库",
            AutoSize = true,
            Anchor = AnchorStyles.Right,
            Font = new Font("Microsoft YaHei UI", 8.5F),
            Margin = new Padding(S(12), 0, 0, 0),
        };
        linkRepo.LinkClicked += (_, _) =>
        {
            try
            {
                Process.Start(new ProcessStartInfo { FileName = AppConfig.RepoWebUrl, UseShellExecute = true });
            }
            catch { }
        };
        footer.Controls.Add(linkRepo, 1, 0);
        root.Controls.Add(footer, 0, 5);

        Controls.Add(root);
        ResumeLayout(performLayout: true);

        FormClosing += (_, _) =>
        {
            try { _lifetimeCts.Cancel(); } catch { }
        };

        Shown += async (_, _) =>
        {
            RefreshLocalSummary();
            if (justUpgraded)
            {
                AppendLog($"[提示] 查看器程序已成功升级至 v{AppConfig.CurrentVersion}！");
            }
            await CheckAllUpdatesAsync();
        };
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _lifetimeCts.Dispose();
        }
        base.Dispose(disposing);
    }

    private void SetPill(string text, string state)
    {
        if (IsDisposed) return;
        _lblPill.Text = text;
        switch (state)
        {
            case "running":
                _lblPill.BackColor = Color.FromArgb(255, 233, 168);
                _lblPill.ForeColor = Color.FromArgb(90, 67, 0);
                break;
            case "ok":
                _lblPill.BackColor = Color.FromArgb(167, 243, 208);
                _lblPill.ForeColor = Color.FromArgb(6, 64, 44);
                break;
            case "update":
                _lblPill.BackColor = Color.FromArgb(224, 231, 255);
                _lblPill.ForeColor = Color.FromArgb(55, 48, 163);
                break;
            case "error":
                _lblPill.BackColor = Color.FromArgb(254, 205, 201);
                _lblPill.ForeColor = Color.FromArgb(122, 22, 19);
                break;
            default:
                _lblPill.BackColor = Color.FromArgb(238, 241, 246);
                _lblPill.ForeColor = Color.FromArgb(66, 84, 102);
                break;
        }
    }

    private void AppendLog(string line)
    {
        if (IsDisposed) return;
        if (_txtLog.TextLength > 0)
        {
            _txtLog.AppendText(Environment.NewLine);
        }
        _txtLog.AppendText(line);
    }

    private SheetSnapshot? RefreshLocalSummary()
    {
        if (IsDisposed) return null;
        if (!File.Exists(_excelPath))
        {
            _lblLocalStats.Text = "本地表格统计：本地尚未存在 hero_skin_scan.xlsx（点击「检查更新」即可拉取）";
            _lblLatestRecord.Text = "最近入表记录：无";
            return null;
        }

        try
        {
            var snap = XlsxParser.Parse(_excelPath);
            _lblLocalStats.Text = $"本地表格统计：共 {snap.HeroCount} 位英雄 · {snap.SkinCount} 款皮肤";
            if (snap.LatestSkin is { } last)
            {
                string sname = string.IsNullOrEmpty(last.SkinName) ? $"{last.SkinId:D2}槽(待命名)" : last.SkinName;
                _lblLatestRecord.Text = $"最近入表记录：{last.LastModified}  ·  {last.HeroId:D3} {last.HeroName} - {sname}";
            }
            else
            {
                _lblLatestRecord.Text = "最近入表记录：—";
            }
            return snap;
        }
        catch (Exception ex)
        {
            _lblLocalStats.Text = "本地表格读取异常：" + ex.Message;
            return null;
        }
    }

    private async Task CheckAllUpdatesAsync()
    {
        if (_busy || IsDisposed) return;
        _busy = true;
        _btnCheck.Enabled = false;
        _btnApplyExcel.Enabled = false;
        _pendingExcelBytes = null;
        _progress.Style = ProgressBarStyle.Marquee;
        SetPill("检查中……", "running");

        _txtLog.Clear();
        AppendLog($"[{DateTime.Now:HH:mm:ss}] 正在连接 GitHub 仓库检查最新表格与查看器版本……");

        try
        {
            CancellationToken ct = _lifetimeCts.Token;
            var localSnap = RefreshLocalSummary();
            byte[]? localBytes = File.Exists(_excelPath) ? await File.ReadAllBytesAsync(_excelPath, ct) : null;

            var excelTask = UpdateService.FetchLatestExcelAsync(ct);
            var appTask = UpdateService.CheckViewerUpdateAsync(ct);

            await Task.WhenAll(excelTask, appTask);
            if (IsDisposed) return;

            // 1) 处理查看器程序自更新检查结果
            _pendingAppUpdate = await appTask;
            if (_pendingAppUpdate is not null)
            {
                _lblAppUpdateInfo.Text = $"发现查看器新版本 v{_pendingAppUpdate.Version}（当前 v{AppConfig.CurrentVersion}），支持一键在线升级";
                _appUpdateBanner.Visible = true;
                AppendLog($"[程序更新] 发现新版查看器 v{_pendingAppUpdate.Version}（{_pendingAppUpdate.AssetName}）");
                if (!string.IsNullOrEmpty(_pendingAppUpdate.ReleaseNotes))
                {
                    AppendLog("  更新说明：" + _pendingAppUpdate.ReleaseNotes.Replace("\n", "\n  "));
                }
            }
            else
            {
                _appUpdateBanner.Visible = false;
            }

            // 2) 处理 Excel 表格对比结果
            var remote = await excelTask;
            AppendLog($"[{DateTime.Now:HH:mm:ss}] 已通过「{remote.SourceName}」获取线上表格（共 {remote.RemoteSheet.HeroCount} 位英雄 · {remote.RemoteSheet.SkinCount} 款皮肤）");

            bool bytesDiffer = localBytes is null
                || !string.Equals(UpdateService.Sha256Hex(localBytes), UpdateService.Sha256Hex(remote.RemoteBytes), StringComparison.Ordinal);

            if (localSnap is null)
            {
                _pendingExcelBytes = remote.RemoteBytes;
                _btnApplyExcel.Enabled = true;
                SetPill("待下载初始表", "update");
                AppendLog("本地尚无 hero_skin_scan.xlsx，请点击上方绿色「更新本地 Excel 表格」按钮保存。");
            }
            else
            {
                var diff = XlsxParser.Compare(localSnap, remote.RemoteSheet);
                if (!diff.HasAnyChange && !bytesDiffer)
                {
                    SetPill(_pendingAppUpdate is not null ? "程序有更新" : "已是最新", _pendingAppUpdate is not null ? "update" : "ok");
                    AppendLog("结论：本地 Excel 表格与仓库最新版本完全一致，暂无新皮肤更新。");
                }
                else if (!diff.HasAnyChange && bytesDiffer)
                {
                    _pendingExcelBytes = remote.RemoteBytes;
                    _btnApplyExcel.Enabled = true;
                    SetPill("表格有微调", "update");
                    AppendLog("检测到线上 Excel 文件有非皮肤槽位的格式/备注调整，可点击「更新本地 Excel 表格」同步。");
                }
                else
                {
                    _pendingExcelBytes = remote.RemoteBytes;
                    _btnApplyExcel.Enabled = true;
                    SetPill($"发现 {diff.TotalChangeCount} 处更新", "update");

                    AppendLog("============================================================");
                    AppendLog($"检测到仓库表格有更新！（新增皮肤 {diff.Added.Count} · 补充命名 {diff.Named.Count} · 名称修正 {diff.Renamed.Count} · 时间刷新 {diff.TimeUpdated.Count}）");
                    AppendLog("点击上方绿色「更新本地 Excel 表格」按钮即可同步到本地：");
                    AppendLog("------------------------------------------------------------");

                    foreach (var (hid, oldHn, newHn) in diff.HeroNamed)
                    {
                        AppendLog($"  [英雄命名] {hid:D3}：{oldHn} → {newHn}");
                    }
                    foreach (var r in diff.Added)
                    {
                        string nm = string.IsNullOrEmpty(r.SkinName) ? "（待命名）" : r.SkinName;
                        AppendLog($"  [新增皮肤] {r.HeroId:D3} {r.HeroName} · {r.SkinId:D2}槽：{nm}  ({r.LastModified})");
                    }
                    foreach (var (_, r) in diff.Named)
                    {
                        AppendLog($"  [补充命名] {r.HeroId:D3} {r.HeroName} · {r.SkinId:D2}槽 → {r.SkinName}");
                    }
                    foreach (var (l, r) in diff.Renamed)
                    {
                        AppendLog($"  [名称修正] {r.HeroId:D3} {r.HeroName} · {r.SkinId:D2}槽：{l.SkinName} → {r.SkinName}");
                    }
                    foreach (var (l, r) in diff.TimeUpdated)
                    {
                        string nm = string.IsNullOrEmpty(r.SkinName) ? "" : $"（{r.SkinName}）";
                        AppendLog($"  [原画更新] {r.HeroId:D3} {r.HeroName} · {r.SkinId:D2}槽{nm}：{l.LastModified} → {r.LastModified}");
                    }
                    foreach (var l in diff.Removed)
                    {
                        string nm = string.IsNullOrEmpty(l.SkinName) ? "（未命名）" : l.SkinName;
                        AppendLog($"  [移除槽位] {l.HeroId:D3} {l.HeroName} · {l.SkinId:D2}槽：{nm}");
                    }
                    AppendLog("============================================================");
                }
            }
        }
        catch (OperationCanceledException) when (_lifetimeCts.IsCancellationRequested)
        {
            return;
        }
        catch (Exception ex)
        {
            if (IsDisposed) return;
            SetPill("检查失败", "error");
            AppendLog($"[错误] 检查更新失败：{ex.Message}");
        }
        finally
        {
            if (!IsDisposed)
            {
                _progress.Style = ProgressBarStyle.Blocks;
                _progress.Value = 0;
                _btnCheck.Enabled = true;
                _busy = false;
            }
        }
    }

    private void OnApplyExcelClicked(object? sender, EventArgs e)
    {
        if (_pendingExcelBytes is null) return;

        if (AppConfig.IsFileLockedByExcel(_excelPath))
        {
            MessageBox.Show(
                this,
                "检测到 hero_skin_scan.xlsx 正在被 Excel 或 WPS 打开占用。\n\n请先关闭表格窗口，再点击「更新本地 Excel 表格」。",
                "文件被占用",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        string tmpPath = _excelPath + ".tmp";
        try
        {
            File.WriteAllBytes(tmpPath, _pendingExcelBytes);
            File.Move(tmpPath, _excelPath, overwrite: true);

            _pendingExcelBytes = null;
            _btnApplyExcel.Enabled = false;
            RefreshLocalSummary();
            SetPill("已更新至最新", "ok");
            AppendLog($"[{DateTime.Now:HH:mm:ss}] 本地 hero_skin_scan.xlsx 已成功更新至最新版本！");

            var ask = MessageBox.Show(
                this,
                "本地 Excel 表格已更新至最新版本！\n是否立即打开表格查看？",
                "更新完成",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Information);

            if (ask == DialogResult.Yes)
            {
                OnOpenExcelClicked(this, EventArgs.Empty);
            }
        }
        catch (Exception ex)
        {
            try { if (File.Exists(tmpPath)) File.Delete(tmpPath); } catch { }
            MessageBox.Show(
                this,
                $"写入本地表格失败：\n{ex.Message}",
                "更新失败",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private void OnOpenExcelClicked(object? sender, EventArgs e)
    {
        if (!File.Exists(_excelPath))
        {
            MessageBox.Show(
                this,
                "本地还没有 hero_skin_scan.xlsx，请先点击「检查更新」拉取最新表格。",
                "缺少文件",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = _excelPath,
                UseShellExecute = true,
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"打开表格失败：\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private async void OnUpgradeViewerClicked(object? sender, EventArgs e)
    {
        if (_pendingAppUpdate is null || _busy) return;

        var confirm = MessageBox.Show(
            this,
            $"即将在线下载并升级查看器至 v{_pendingAppUpdate.Version}，升级完成后程序会自动重启。\n\n确定立即升级吗？",
            "升级查看器程序",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (confirm != DialogResult.Yes) return;

        _busy = true;
        _btnUpgradeApp.Enabled = false;
        _btnCheck.Enabled = false;
        _progress.Style = ProgressBarStyle.Blocks;
        _progress.Value = 0;
        SetPill("升级程序中……", "running");
        AppendLog($"[{DateTime.Now:HH:mm:ss}] 正在下载查看器更新包 {_pendingAppUpdate.AssetName}……");

        try
        {
            var prog = new Progress<int>(pct =>
            {
                if (!IsDisposed) _progress.Value = Math.Clamp(pct, 0, 100);
            });
            await UpdateService.ApplyViewerSelfUpdateAsync(_pendingAppUpdate, _installDir, prog, _lifetimeCts.Token);
            Application.Exit();
        }
        catch (OperationCanceledException) when (_lifetimeCts.IsCancellationRequested)
        {
            return;
        }
        catch (Exception ex)
        {
            if (IsDisposed) return;
            _busy = false;
            _btnUpgradeApp.Enabled = true;
            _btnCheck.Enabled = true;
            SetPill("升级失败", "error");
            AppendLog($"[错误] 查看器自更新失败：{ex.Message}");
            MessageBox.Show(this, $"查看器升级失败：\n{ex.Message}", "升级失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
