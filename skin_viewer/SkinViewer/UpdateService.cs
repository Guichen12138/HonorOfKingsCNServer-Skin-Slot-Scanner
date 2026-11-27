using System.Diagnostics;
using System.IO.Compression;
using System.Net;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace SkinViewer;

internal sealed record ExcelRemoteCheckResult(
    byte[] RemoteBytes,
    string SourceName,
    string? CommitSha,
    SheetSnapshot RemoteSheet);

internal sealed record ViewerUpdateInfo(
    string Version,
    string DownloadUrl,
    string AssetName,
    string ReleaseNotes);

internal static class UpdateService
{
    private static readonly HttpClient Http = CreateHttpClient();

    private static HttpClient CreateHttpClient()
    {
        var handler = new SocketsHttpHandler
        {
            UseProxy = true,
            AutomaticDecompression = DecompressionMethods.All,
            ConnectTimeout = TimeSpan.FromSeconds(6),
            PooledConnectionLifetime = TimeSpan.FromMinutes(3),
        };
        var client = new HttpClient(handler)
        {
            Timeout = TimeSpan.FromSeconds(30),
        };
        client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 HonorOfKingsSkinViewer/" + AppConfig.CurrentVersion);
        return client;
    }

    public static string Sha256Hex(byte[] data)
    {
        byte[] hash = SHA256.HashData(data);
        return Convert.ToHexString(hash);
    }

    // --------------------------------------------------------
    // 1. 拉取仓库最新 hero_skin_scan.xlsx（多源实时防缓存）
    // --------------------------------------------------------

    public static async Task<ExcelRemoteCheckResult> FetchLatestExcelAsync(CancellationToken ct = default)
    {
        string? latestSha = await TryResolveLatestBranchShaAsync(ct);
        string refOrSha = !string.IsNullOrEmpty(latestSha) ? latestSha : AppConfig.RepoBranch;
        long buster = DateTimeOffset.UtcNow.ToUnixTimeSeconds();

        var sources = new List<(string Name, string Url)>();

        if (!string.IsNullOrEmpty(latestSha))
        {
            // 精确到 commit SHA 的 jsDelivr 地址：永不命中旧分支缓存，且国内直连快
            sources.Add((
                "jsDelivr CDN",
                $"https://cdn.jsdelivr.net/gh/{AppConfig.RepoOwner}/{AppConfig.RepoName}@{latestSha}/{AppConfig.ExcelFileName}"));
        }

        sources.Add((
            "GitHub 镜像源 (gh-proxy)",
            $"https://gh-proxy.com/https://raw.githubusercontent.com/{AppConfig.RepoOwner}/{AppConfig.RepoName}/{refOrSha}/{AppConfig.ExcelFileName}?t={buster}"));

        sources.Add((
            "GitHub 官方源",
            $"https://raw.githubusercontent.com/{AppConfig.RepoOwner}/{AppConfig.RepoName}/{refOrSha}/{AppConfig.ExcelFileName}?t={buster}"));

        sources.Add((
            "jsDelivr 分支源",
            $"https://cdn.jsdelivr.net/gh/{AppConfig.RepoOwner}/{AppConfig.RepoName}@{AppConfig.RepoBranch}/{AppConfig.ExcelFileName}"));

        var errors = new List<string>();

        foreach (var (name, url) in sources)
        {
            ct.ThrowIfCancellationRequested();
            try
            {
                using var req = new HttpRequestMessage(HttpMethod.Get, url);
                req.Headers.CacheControl = new System.Net.Http.Headers.CacheControlHeaderValue { NoCache = true };

                using var resp = await Http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct);
                if (!resp.IsSuccessStatusCode)
                {
                    errors.Add($"{name}: HTTP {(int)resp.StatusCode}");
                    continue;
                }

                byte[] bytes = await resp.Content.ReadAsByteArrayAsync(ct);
                if (!XlsxParser.IsValidXlsx(bytes))
                {
                    errors.Add($"{name}: 返回内容不是有效的 xlsx 文件");
                    continue;
                }

                var snapshot = XlsxParser.Parse(bytes);
                return new ExcelRemoteCheckResult(bytes, name, latestSha, snapshot);
            }
            catch (Exception ex) when (ex is not OperationCanceledException || !ct.IsCancellationRequested)
            {
                errors.Add($"{name}: {ex.Message}");
            }
        }

        throw new InvalidOperationException("所有下载线路均未能连上仓库：\n" + string.Join("\n", errors));
    }

    private static async Task<string?> TryResolveLatestBranchShaAsync(CancellationToken ct)
    {
        // 路线 A：Git info/refs（通过镜像和直连均可实时拿到 HEAD SHA，无 API 频率限制且仅 ~200 字节）
        string infoRefsPath = $"https://github.com/{AppConfig.RepoOwner}/{AppConfig.RepoName}/info/refs?service=git-upload-pack";
        string[] refUrls =
        {
            "https://gh-proxy.com/" + infoRefsPath,
            infoRefsPath,
        };

        foreach (string u in refUrls)
        {
            ct.ThrowIfCancellationRequested();
            try
            {
                using var cts = CancellationTokenSource.CreateLinkedTokenSource(ct);
                cts.CancelAfter(TimeSpan.FromSeconds(5));
                string text = await Http.GetStringAsync(u, cts.Token);
                var m = Regex.Match(text, @"([0-9a-f]{40})\s+refs/heads/" + Regex.Escape(AppConfig.RepoBranch));
                if (m.Success)
                {
                    return m.Groups[1].Value;
                }
            }
            catch (OperationCanceledException) when (ct.IsCancellationRequested)
            {
                throw;
            }
            catch { }
        }

        // 路线 B：GitHub Commits API
        try
        {
            using var cts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            cts.CancelAfter(TimeSpan.FromSeconds(5));
            string api = $"https://api.github.com/repos/{AppConfig.RepoOwner}/{AppConfig.RepoName}/commits/{AppConfig.RepoBranch}";
            string json = await Http.GetStringAsync(api, cts.Token);
            using var doc = JsonDocument.Parse(json);
            if (doc.RootElement.TryGetProperty("sha", out var shaProp))
            {
                return shaProp.GetString();
            }
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested)
        {
            throw;
        }
        catch { }

        return null;
    }

    // --------------------------------------------------------
    // 2. 检查并自更新查看器程序（读取 GitHub Releases 中发布的 SkinViewer_vX.Y.Z.zip/.exe）
    // --------------------------------------------------------

    private static readonly Regex ViewerAssetRegex = new(
        @"(?:SkinViewer|王者荣耀皮肤格子查看器)[_-]?v?(\d+\.\d+\.\d+)\.(zip|exe)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ExpandedAssetHrefRegex = new(
        @"href=""(/[^""]+/releases/download/[^""]+/([^""/]+))""",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static async Task<ViewerUpdateInfo?> CheckViewerUpdateAsync(CancellationToken ct = default)
    {
        if (!Version.TryParse(AppConfig.CurrentVersion, out var currentVer))
        {
            currentVer = new Version(1, 0, 0);
        }

        // 优先走官方 Releases API
        var fromApi = await TryCheckViewerViaApiAsync(currentVer, ct);
        if (fromApi is not null)
        {
            return fromApi;
        }

        // 兜底：若公共网络环境命中 GitHub API 60 次/小时匿名限流，通过镜像读取 releases.atom + expanded_assets
        return await TryCheckViewerViaAtomAsync(currentVer, ct);
    }

    private static async Task<ViewerUpdateInfo?> TryCheckViewerViaApiAsync(Version currentVer, CancellationToken ct)
    {
        string apiUrl = $"https://api.github.com/repos/{AppConfig.RepoOwner}/{AppConfig.RepoName}/releases?per_page=10";
        try
        {
            using var cts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            cts.CancelAfter(TimeSpan.FromSeconds(6));
            string json = await Http.GetStringAsync(apiUrl, cts.Token);
            if (string.IsNullOrWhiteSpace(json)) return null;

            using var doc = JsonDocument.Parse(json);
            if (doc.RootElement.ValueKind != JsonValueKind.Array) return null;

            Version? bestVer = null;
            ViewerUpdateInfo? bestInfo = null;

            foreach (var rel in doc.RootElement.EnumerateArray())
            {
                if (rel.TryGetProperty("draft", out var d) && d.GetBoolean()) continue;

                string notes = rel.TryGetProperty("body", out var b) ? (b.GetString() ?? "") : "";
                if (!rel.TryGetProperty("assets", out var assets) || assets.ValueKind != JsonValueKind.Array)
                    continue;

                foreach (var asset in assets.EnumerateArray())
                {
                    string name = asset.TryGetProperty("name", out var n) ? (n.GetString() ?? "") : "";
                    string dlUrl = asset.TryGetProperty("browser_download_url", out var u) ? (u.GetString() ?? "") : "";
                    if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(dlUrl)) continue;

                    if (TryMatchNewerViewerAsset(name, dlUrl, notes, currentVer, ref bestVer, ref bestInfo))
                    {
                        // 继续遍历以选最高版本
                    }
                }
            }

            return bestInfo;
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested)
        {
            throw;
        }
        catch
        {
            return null;
        }
    }

    private static async Task<ViewerUpdateInfo?> TryCheckViewerViaAtomAsync(Version currentVer, CancellationToken ct)
    {
        string[] bases =
        {
            $"https://gh-proxy.com/https://github.com/{AppConfig.RepoOwner}/{AppConfig.RepoName}",
            $"https://github.com/{AppConfig.RepoOwner}/{AppConfig.RepoName}",
        };

        foreach (string repoBase in bases)
        {
            ct.ThrowIfCancellationRequested();
            try
            {
                using var cts = CancellationTokenSource.CreateLinkedTokenSource(ct);
                cts.CancelAfter(TimeSpan.FromSeconds(6));
                string atom = await Http.GetStringAsync(repoBase + "/releases.atom", cts.Token);
                var tagMatches = Regex.Matches(atom, @"/releases/tag/([^""<\s]+)");
                var tags = tagMatches.Select(m => Uri.UnescapeDataString(m.Groups[1].Value)).Distinct().Take(3).ToList();
                if (tags.Count == 0) continue;

                Version? bestVer = null;
                ViewerUpdateInfo? bestInfo = null;

                foreach (string tag in tags)
                {
                    using var ctsAssets = CancellationTokenSource.CreateLinkedTokenSource(ct);
                    ctsAssets.CancelAfter(TimeSpan.FromSeconds(5));
                    string html = await Http.GetStringAsync(repoBase + "/releases/expanded_assets/" + Uri.EscapeDataString(tag), ctsAssets.Token);
                    foreach (Match m in ExpandedAssetHrefRegex.Matches(html))
                    {
                        string relPath = m.Groups[1].Value;
                        string assetName = Uri.UnescapeDataString(m.Groups[2].Value);
                        string fullUrl = "https://github.com" + relPath;
                        TryMatchNewerViewerAsset(assetName, fullUrl, "", currentVer, ref bestVer, ref bestInfo);
                    }
                }

                return bestInfo;
            }
            catch (OperationCanceledException) when (ct.IsCancellationRequested)
            {
                throw;
            }
            catch { }
        }

        return null;
    }

    private static bool TryMatchNewerViewerAsset(
        string assetName,
        string downloadUrl,
        string notes,
        Version currentVer,
        ref Version? bestVer,
        ref ViewerUpdateInfo? bestInfo)
    {
        // 不把首次安装器 Setup.exe 当成原地自更新包
        if (assetName.Contains("Setup", StringComparison.OrdinalIgnoreCase)
            || assetName.Contains("安装", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var m = ViewerAssetRegex.Match(assetName);
        if (!m.Success) return false;

        if (Version.TryParse(m.Groups[1].Value, out var ver) && ver > currentVer)
        {
            if (bestVer is null || ver > bestVer)
            {
                bestVer = ver;
                bestInfo = new ViewerUpdateInfo(m.Groups[1].Value, downloadUrl, assetName, notes.Trim());
                return true;
            }
        }
        return false;
    }

    public static async Task ApplyViewerSelfUpdateAsync(
        ViewerUpdateInfo info,
        string installDir,
        IProgress<int>? progress = null,
        CancellationToken ct = default)
    {
        string[] downloadUrls =
        {
            "https://gh-proxy.com/" + info.DownloadUrl,
            info.DownloadUrl,
        };

        byte[]? payload = null;
        Exception? lastErr = null;

        foreach (string url in downloadUrls)
        {
            ct.ThrowIfCancellationRequested();
            try
            {
                using var resp = await Http.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, ct);
                resp.EnsureSuccessStatusCode();

                long total = resp.Content.Headers.ContentLength ?? -1L;
                await using var netStream = await resp.Content.ReadAsStreamAsync(ct);
                using var ms = new MemoryStream();
                byte[] buf = new byte[65536];
                long readTotal = 0;
                int read;
                while ((read = await netStream.ReadAsync(buf, ct)) > 0)
                {
                    ms.Write(buf, 0, read);
                    readTotal += read;
                    if (total > 0)
                    {
                        progress?.Report((int)Math.Clamp(readTotal * 100L / total, 0, 98));
                    }
                }
                payload = ms.ToArray();
                if (payload.Length > 100 * 1024) break;
            }
            catch (Exception ex) when (ex is not OperationCanceledException || !ct.IsCancellationRequested)
            {
                lastErr = ex;
                payload = null;
            }
        }

        if (payload is null || payload.Length < 100 * 1024)
        {
            throw new InvalidOperationException("下载查看器更新包失败：" + (lastErr?.Message ?? "文件不完整"));
        }

        await ApplyViewerPackageBytesAsync(payload, info.AssetName, installDir, progress, ct);
    }

    public static async Task ApplyViewerPackageBytesAsync(
        byte[] payload,
        string assetName,
        string installDir,
        IProgress<int>? progress = null,
        CancellationToken ct = default,
        bool restartAfterUpdate = true)
    {
        string currentExe = Environment.ProcessPath ?? Path.Combine(installDir, AppConfig.ExeFileName);
        string newExePath = currentExe + ".new";
        string oldExePath = currentExe + ".old";

        if (assetName.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
        {
            using var zip = new ZipArchive(new MemoryStream(payload), ZipArchiveMode.Read);
            var exeEntry = zip.Entries.FirstOrDefault(e =>
                    string.Equals(e.Name, AppConfig.ExeFileName, StringComparison.OrdinalIgnoreCase))
                ?? zip.Entries.FirstOrDefault(e =>
                    e.Name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)
                    && !e.Name.Equals(AppConfig.UninstallFileName, StringComparison.OrdinalIgnoreCase))
                ?? throw new InvalidDataException("更新压缩包内未找到查看器可执行文件。");

            await using (var es = exeEntry.Open())
            await using (var fs = new FileStream(newExePath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                await es.CopyToAsync(fs, ct);
            }

            // 若更新包内附带了新版 Uninstall.exe，顺带静默更新
            var uninstEntry = zip.Entries.FirstOrDefault(e =>
                string.Equals(e.Name, AppConfig.UninstallFileName, StringComparison.OrdinalIgnoreCase));
            if (uninstEntry is not null)
            {
                string dstUninst = Path.Combine(installDir, AppConfig.UninstallFileName);
                string tmpUninst = dstUninst + ".new";
                try
                {
                    await using (var ues = uninstEntry.Open())
                    await using (var ufs = new FileStream(tmpUninst, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        await ues.CopyToAsync(ufs, ct);
                    }
                    File.Move(tmpUninst, dstUninst, overwrite: true);
                }
                catch
                {
                    try { if (File.Exists(tmpUninst)) File.Delete(tmpUninst); } catch { }
                }
            }
        }
        else
        {
            await File.WriteAllBytesAsync(newExePath, payload, ct);
        }

        // 校验 PE MZ 头
        byte[] head = new byte[2];
        await using (var checkFs = File.OpenRead(newExePath))
        {
            if (await checkFs.ReadAsync(head, ct) < 2 || head[0] != (byte)'M' || head[1] != (byte)'Z')
            {
                try { File.Delete(newExePath); } catch { }
                throw new InvalidDataException("下载的更新文件不是有效的 Windows 可执行程序。");
            }
        }

        progress?.Report(100);

        // Windows 允许对正在运行的 exe 改名为 .old，再把 .new 移入原路径并重启
        try
        {
            if (File.Exists(oldExePath)) File.Delete(oldExePath);
            File.Move(currentExe, oldExePath);
            File.Move(newExePath, currentExe);

            if (restartAfterUpdate)
            {
                var psi = new ProcessStartInfo
                {
                    FileName = currentExe,
                    WorkingDirectory = installDir,
                    UseShellExecute = false,
                };
                psi.ArgumentList.Add("--upgraded");
                Process.Start(psi);
            }
        }
        catch
        {
            if (restartAfterUpdate)
            {
                // 兜底：若被安全软件短暂占用导致改名受阻，等当前进程退出后替换并重拉起
                string cmd = $"/c timeout /t 1 /nobreak >nul & move /y \"{newExePath}\" \"{currentExe}\" >nul & start \"\" \"{currentExe}\" --upgraded";
                Process.Start(new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = cmd,
                    WorkingDirectory = installDir,
                    CreateNoWindow = true,
                    UseShellExecute = false,
                });
            }
            else
            {
                throw;
            }
        }
    }
}
