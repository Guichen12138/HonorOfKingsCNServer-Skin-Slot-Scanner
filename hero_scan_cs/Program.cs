// ============================================================
// hero_scan_cs —— C# 扫描引擎
//
// 被 hero_scan.py 当子进程调用：
//     hero_scan_cs.exe <起始英雄号> <结束英雄号> <并发> <起始槽> <结束槽>
//
// 每完成一个英雄，往 stdout 吐一行 JSON：
//     {"hero_id": 105, "failed": 0,
//      "results": [{"skin_id": 0, "lm": "2020-04-26 11:44:26"}, ...]}
//
// results 只包含实际存在（200 且有 Last-Modified）的格子；
// failed 是这一英雄重试完仍失败的格子数。
// 00 槽 404 的空号英雄输出 {"hero_id": n, "failed": 0, "results": []}。
//
// 扫描语义和 hero_scan.py 的 requests 版逐条对应：
// 重试 3 次、429 全局冷却、过期 404 带 buster 回源确认、
// Last-Modified 转北京时间。
//
// 为什么快：实测（见 speed_test/REPORT.md）这个 CDN 的吞吐上限
// 不在服务端，而在 Python 线程 + GIL；C# 的异步 HttpClient
// 用 512 条 HTTP/1.1 连接能跑到现网的 9~15 倍。
// ============================================================

using System.Diagnostics;
using System.Globalization;
using System.Net;
using System.Net.Http.Headers;

internal static class Program
{
    // --------------------------------------------------------
    // 配置（和 hero_scan.py 保持一致）
    // --------------------------------------------------------

    const int RETRIES = 3;          // 单地址失败后的重试次数
    const double RETRY_DELAY = 0.5; // 每次重试前的等待秒数（按次数递增）
    const int TIMEOUT_MS = 10_000;  // 请求超时
    const int THROTTLE_MS = 1_000;  // 429 全局冷却窗口

    static readonly string BaseUrl = "https://image.smoba.qq.com/Picture/HeroOriginalPainting/";

    // 整轮扫描共用的缓存标记：只在重试时带上，强制回源
    static readonly string CacheBuster =
        DateTimeOffset.UtcNow.ToUnixTimeMilliseconds().ToString();

    // 在 Main 里按并发参数构建（连接数要跟 threads 走）
    static HttpClient Http = null!;

    // 429 全局冷却：任何线程撞到 429 就推后一个窗口，
    // 所有请求发出前先看一眼（对应 hero_scan.py 的 _note_429/_wait_throttle）
    static long _throttleUntilMs;
    static readonly object PrintLock = new();
    static readonly object ThrottleLock = new();

    // --------------------------------------------------------
    // 请求成分统计（跑完打到 stderr，用于定位瓶颈）
    // --------------------------------------------------------

    static long StatRequests;    // 发出的请求总数
    static long StatOk200;       // 200 且带 Last-Modified
    static long StatConfirm404;  // 边缘缓存的 404 → 带 buster 回源确认
    static long StatFinal404;    // 最终判定不存在
    static long StatRetries;     // 429/5xx/无LM/网络抖动的重试
    static long StatFailed;      // 重试完仍失败

    // --------------------------------------------------------
    // 记录类型
    // --------------------------------------------------------

    sealed record SkinResult(int SkinId, string? LastModified, bool Failed)
    {
        public bool Exists => LastModified is not null;
    }

    sealed record HeroResult(int HeroId, List<SkinResult> Results, int Failed);

    static int Main(string[] args)
    {
        int start     = args.Length > 0 ? int.Parse(args[0]) : 1;
        int end       = args.Length > 1 ? int.Parse(args[1]) : 999;
        int threads   = args.Length > 2 ? int.Parse(args[2]) : 512;
        int skinStart = args.Length > 3 ? int.Parse(args[3]) : 0;
        int skinEnd   = args.Length > 4 ? int.Parse(args[4]) : 19;
        string ver    = args.Length > 5 ? args[5] : "h1";

        var version = ver == "h2"
            ? HttpVersion.Version20
            : HttpVersion.Version11;

        Http = new HttpClient(new SocketsHttpHandler
        {
            MaxConnectionsPerServer = threads,
            PooledConnectionLifetime = TimeSpan.FromMinutes(5),
            ConnectTimeout = TimeSpan.FromMilliseconds(TIMEOUT_MS),
        })
        {
            Timeout = TimeSpan.FromMilliseconds(TIMEOUT_MS),
            DefaultRequestVersion = version,
            DefaultVersionPolicy = HttpVersionPolicy.RequestVersionOrLower,
        };

        Http.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0");
        Http.DefaultRequestHeaders.CacheControl = new CacheControlHeaderValue
        {
            NoCache = true,
        };
        Http.DefaultRequestHeaders.Pragma.ParseAdd("no-cache");

        // 连接预热（第 7 个参数传 warm 启用）：
        // 先把 threads 条 TLS 连接建立好再开扫，
        // 免得头一波请求全在等握手
        bool warm = args.Length > 6 && args[6] == "warm";

        if (warm)
        {
            var warmTasks = new Task[threads];
            for (int i = 0; i < threads; i++)
            {
                warmTasks[i] = Task.Run(async () =>
                {
                    try
                    {
                        // 根路径响应很小，整读不伤连接复用
                        using var resp = await Http.GetAsync(
                            BaseUrl, HttpCompletionOption.ResponseHeadersRead);
                        await resp.Content.LoadIntoBufferAsync(64 * 1024);
                    }
                    catch { }
                });
            }
            try { Task.WaitAll(warmTasks); } catch { }
            Console.Error.WriteLine("{\"warm\": true}");
        }

        var sw = Stopwatch.StartNew();

        var tasks = new Task[end - start + 1];
        for (int heroId = start; heroId <= end; heroId++)
        {
            int id = heroId;
            tasks[id - start] = Task.Run(async () =>
            {
                var result = await ScanHeroAsync(id, skinStart, skinEnd);
                Emit(result);
            });
        }

        try
        {
            Task.WaitAll(tasks);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("hero task fault: " + ex.Message);
            return 1;
        }

        sw.Stop();

        // 请求成分汇总（stderr，供瓶颈分析）
        Console.Error.WriteLine(
            $"{{\"elapsed_s\": {sw.Elapsed.TotalSeconds:F2}, " +
            $"\"requests\": {StatRequests}, \"ok200\": {StatOk200}, " +
            $"\"confirm404\": {StatConfirm404}, \"final404\": {StatFinal404}, " +
            $"\"retries\": {StatRetries}, \"failed\": {StatFailed}, " +
            $"\"version\": \"{ver}\", \"threads\": {threads}}}");

        return 0;
    }

    // --------------------------------------------------------
    // 扫描一个英雄：先探测 00 槽，空号直接跳过剩余槽位
    //（对应 hero_scan.py 的 scan_hero）
    // --------------------------------------------------------

    static async Task<HeroResult> ScanHeroAsync(int heroId, int skinStart, int skinEnd)
    {
        var results = new List<SkinResult>();
        int failed = 0;

        var first = await CheckSkinAsync(heroId, skinStart);

        if (first.Exists)
        {
            results.Add(first);
        }
        else if (first.Failed)
        {
            // 00 槽请求失败：不断定英雄不存在，继续扫完整个英雄
            failed++;
        }
        else
        {
            // 00 槽 404：英雄编号不存在，跳过剩余槽位
            return new HeroResult(heroId, results, failed);
        }

        for (int skinId = skinStart + 1; skinId <= skinEnd; skinId++)
        {
            var result = await CheckSkinAsync(heroId, skinId);
            if (result.Exists) results.Add(result);
            else if (result.Failed) failed++;
        }

        return new HeroResult(heroId, results, failed);
    }

    // --------------------------------------------------------
    // 检查单个资源（对应 hero_scan.py 的 check_skin）
    // --------------------------------------------------------

    static async Task<SkinResult> CheckSkinAsync(int heroId, int skinId)
    {
        string url = BaseUrl + $"30{heroId:D3}{skinId:D2}.jpg";

        for (int attempt = 1; attempt <= RETRIES; attempt++)
        {
            await WaitThrottleAsync();

            HttpResponseMessage? response = null;
            try
            {
                // 首次请求不带缓存标记，让边缘节点正常命中；
                // 重试才带上，强制回源拿最新结果
                var requestUri = attempt > 1
                    ? new Uri(url + "?_=" + CacheBuster)
                    : new Uri(url);

                using var request = new HttpRequestMessage(HttpMethod.Get, requestUri);
                Interlocked.Increment(ref StatRequests);
                response = await Http.SendAsync(
                    request, HttpCompletionOption.ResponseHeadersRead);

                int status = (int)response.StatusCode;
                if (status == 200)
                {
                    Interlocked.Increment(ref StatOk200);
                }

                // 429 / 5xx 是临时繁忙或限流，要重试；
                // 其余非 200 才是真的不存在
                if (status == 429 || (status >= 500 && status < 600))
                {
                    response.Dispose();
                    if (attempt < RETRIES)
                    {
                        if (status == 429) Note429();
                        Interlocked.Increment(ref StatRetries);
                        await Task.Delay(DelayMs(attempt));
                        continue;
                    }
                    Interlocked.Increment(ref StatFailed);
                    return new SkinResult(skinId, null, Failed: true);
                }

                if (status != 200)
                {
                    bool stale = MayBeStaleCache(response);
                    response.Dispose();

                    // 「过期的 404」确认：首次 404 且响应可能来自
                    // 边缘缓存时，带 buster 强制回源再确认一次
                    if (status == 404 && attempt == 1 && stale)
                    {
                        Interlocked.Increment(ref StatConfirm404);
                        continue;
                    }
                    Interlocked.Increment(ref StatFinal404);
                    return new SkinResult(skinId, null, Failed: false);
                }

                string? lastModified = ReadLastModified(response);
                response.Dispose();

                // 200 但没有 Last-Modified：响应不完整，重试
                if (lastModified is null)
                {
                    Interlocked.Increment(ref StatRetries);
                    await Task.Delay(DelayMs(attempt));
                    continue;
                }

                return new SkinResult(skinId, lastModified, Failed: false);
            }
            catch (Exception ex)
            {
                // 偶发抖动/超时：等一下再试，
                // 不要因为一次失败就把这条记录整个丢掉
                response?.Dispose();
                if (attempt < RETRIES)
                {
                    Interlocked.Increment(ref StatRetries);
                    await Task.Delay(DelayMs(attempt));
                    continue;
                }
                Interlocked.Increment(ref StatFailed);
                // 最终失败的格子记到 stderr（进 hero_scan_cs_error.log）
                lock (PrintLock)
                {
                    Console.Error.WriteLine(
                        $"{heroId}/{skinId}: {ex.GetType().Name}: {ex.Message}");
                }
                return new SkinResult(skinId, null, Failed: true);
            }
        }

        return new SkinResult(skinId, null, Failed: true);
    }

    // --------------------------------------------------------
    // 工具
    // --------------------------------------------------------

    static string? ReadLastModified(HttpResponseMessage response)
    {
        // 注意：Last-Modified 是实体头，.NET 把它放在 Content.Headers
        // 而不是 response.Headers（就是这里一开始踩的坑：
        // 只查 response.Headers 会“永远拿不到”，全部格子白扫）
        if (response.Content.Headers.TryGetValues("Last-Modified", out var values))
        {
            string? raw = values.FirstOrDefault();
            if (string.IsNullOrEmpty(raw))
            {
                return null;
            }

            // GMT → 北京时间 UTC+8（解析失败保持原样，和 Python 版一致）
            if (DateTimeOffset.TryParse(
                    raw, CultureInfo.InvariantCulture, DateTimeStyles.None,
                    out var parsed))
            {
                return parsed.ToOffset(TimeSpan.FromHours(8))
                             .ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            }

            return raw;
        }

        return null;
    }

    static bool MayBeStaleCache(HttpResponseMessage response)
    {
        // 腾讯 CDN 的标记头是 X-Cache-Lookup，可能多段连在一起：
        // 出现过 MISS 说明刚回源，404 可信；全是 HIT 才需要确认。
        // 其他 CDN 用标准 X-Cache / Age 兜底；
        // 什么线索都没有时保守当成缓存。
        if (response.Headers.TryGetValues("X-Cache-Lookup", out var lookup))
        {
            string l = string.Join(",", lookup);
            if (l.Length > 0)
            {
                return !l.ToUpperInvariant().Contains("MISS");
            }
        }

        if (response.Headers.TryGetValues("X-Cache", out var xCache))
        {
            string x = string.Join(",", xCache).ToUpperInvariant();
            if (x.Contains("MISS")) return false;
            if (x.Contains("HIT")) return true;
        }

        if (response.Headers.TryGetValues("Age", out var ages) &&
            int.TryParse(ages.FirstOrDefault(), NumberStyles.Integer,
                         CultureInfo.InvariantCulture, out int age))
        {
            return age > 0;
        }

        return true;
    }

    static int DelayMs(int attempt)
    {
        // 0.5s * 次数 + 0.1~0.6s 抖动，把重试时刻错开
        double seconds = RETRY_DELAY * attempt
                         + 0.1 + Random.Shared.NextDouble() * 0.5;
        return (int)(seconds * 1000);
    }

    static void Note429()
    {
        lock (ThrottleLock)
        {
            long until = Environment.TickCount64 + THROTTLE_MS;
            if (until > _throttleUntilMs) _throttleUntilMs = until;
        }
    }

    static async Task WaitThrottleAsync()
    {
        while (true)
        {
            long now = Environment.TickCount64;
            long wait;
            lock (ThrottleLock)
            {
                wait = _throttleUntilMs - now;
            }
            if (wait <= 0) return;
            await Task.Delay((int)Math.Min(wait, 500));
        }
    }

    static void Emit(HeroResult hero)
    {
        var line = System.Text.Json.JsonSerializer.Serialize(new
        {
            hero_id = hero.HeroId,
            failed = hero.Failed,
            results = hero.Results.Select(r => new
            {
                skin_id = r.SkinId,
                lm = r.LastModified,
            }).ToArray(),
        });

        lock (PrintLock)
        {
            Console.Out.WriteLine(line);
            Console.Out.Flush();
        }
    }
}
