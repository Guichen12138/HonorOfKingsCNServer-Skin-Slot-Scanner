// ============================================================
// hero_scan_cs —— C# 扫描引擎
//
// 被 hero_scan.py 当子进程调用：
//     hero_scan_cs.exe <起始英雄号> <结束英雄号> <并发> <起始槽> <结束槽> [h1|h2] [warm] [range|full]
//
// 每完成一个英雄，往 stdout 吐一行 JSON：
//     {"hero_id":105,"failed":0,"results":[{"skin_id":0,"lm":"2020-04-26 11:44:26"},...]}
//
// results 只包含实际存在（200/206 且有 Last-Modified）的格子；
// failed 是这一英雄重试完仍失败的格子数。
// 00 槽 404 的空号英雄输出 {"hero_id":n,"failed":0,"results":[]}。
//
// 扫描语义和 hero_scan.py 的 requests 版逐条对应：
// 重试 3 次、429 全局冷却、过期 404 带 buster 回源确认、
// Last-Modified 转北京时间。
//
// 优化要点：
// 1. 00 槽命中后，01~19 槽全部并行探测，由全局 SemaphoreSlim 维持满并发，
//    避免单英雄串行把有效并发从 512 压到 139；
// 2. GET 带 Range: bytes=0-0（兼容 206/200，416 自动回退普通 GET），
//    消除 .NET SocketsHttpHandler 默认 1MB MaxResponseDrainSize
//    对 ~980 张 200KB 原画的后台排空流量（~187MB → ~1KB）并保持长连接复用；
// 3. 启用 EnableMultipleHttp2Connections，支持 HTTP/2 多连接并发；
// 4. 零反射 Utf8JsonWriter 输出 JSON，支持单文件裁剪/压缩发布。
// ============================================================

using System.Buffers;
using System.Diagnostics;
using System.Globalization;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

internal static class Program
{
    // --------------------------------------------------------
    // 配置（和 hero_scan.py 保持一致）
    // --------------------------------------------------------

    const int RETRIES = 3;          // 单地址失败后的重试次数
    const double RETRY_DELAY = 0.5; // 每次重试前的等待秒数（按次数递增）
    const int TIMEOUT_MS = 10_000;  // 请求超时
    const int THROTTLE_MS = 1_000;  // 429 全局冷却窗口
    const int TAIL_TRIGGER_MARGIN = 2; // 命中 >= skinEnd - 2 时触发尾部自适应续探
    const int TAIL_STOP_MISSES = 4;    // 连续 4 个槽位 404 即停（覆盖已知 3 槽空洞）
    const int MAX_DYNAMIC_SKIN = 99;   // 30{hero:D3}{skin:D2}.jpg 最大支持 99 槽

    static readonly string BaseUrl = "https://image.smoba.qq.com/Picture/HeroOriginalPainting/";
    static readonly RangeHeaderValue FirstByteRange = new(0, 0);

    // 整轮扫描共用的缓存标记：只在重试时带上，强制回源
    static readonly string CacheBuster =
        DateTimeOffset.UtcNow.ToUnixTimeMilliseconds().ToString();

    static HttpClient Http = null!;
    static SemaphoreSlim Gate = null!;
    static bool UseRange = true;

    // 429 全局冷却：任何任务撞到 429 就推后一个窗口，
    // 所有请求发出前先看一眼（对应 hero_scan.py 的 _note_429/_wait_throttle）
    static long _throttleUntilMs;
    static readonly object PrintLock = new();
    static readonly object ThrottleLock = new();

    // --------------------------------------------------------
    // 请求成分统计（跑完打到 stderr，用于定位瓶颈）
    // --------------------------------------------------------

    static long StatRequests;    // 发出的请求总数
    static long StatOk200;       // 200/206 且带 Last-Modified
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

    sealed record HeroResult(int HeroId, IReadOnlyList<SkinResult> Results, int Failed);

    static async Task<int> Main(string[] args)
    {
        int start      = args.Length > 0 ? int.Parse(args[0], CultureInfo.InvariantCulture) : 1;
        int end        = args.Length > 1 ? int.Parse(args[1], CultureInfo.InvariantCulture) : 999;
        int threads    = args.Length > 2 ? int.Parse(args[2], CultureInfo.InvariantCulture) : 512;
        int skinStart  = args.Length > 3 ? int.Parse(args[3], CultureInfo.InvariantCulture) : 0;
        int skinEnd    = args.Length > 4 ? int.Parse(args[4], CultureInfo.InvariantCulture) : 19;
        string ver     = args.Length > 5 ? args[5] : "h2";
        bool warm      = args.Length > 6 && args[6] == "warm";
        string rngArg  = args.Length > 7 ? args[7] : "range";

        if (threads < 1) threads = 1;
        UseRange = !string.Equals(rngArg, "full", StringComparison.OrdinalIgnoreCase);

        var version = ver == "h2"
            ? HttpVersion.Version20
            : HttpVersion.Version11;

        Gate = new SemaphoreSlim(threads, threads);

        Http = new HttpClient(new SocketsHttpHandler
        {
            MaxConnectionsPerServer = threads,
            EnableMultipleHttp2Connections = true,
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

        // 连接预热（第 7 个参数传 warm 启用）
        if (warm)
        {
            var warmTasks = new Task[threads];
            for (int i = 0; i < threads; i++)
            {
                warmTasks[i] = WarmOneAsync();
            }
            try { await Task.WhenAll(warmTasks); } catch { }
            Console.Error.WriteLine("{\"warm\": true}");
        }

        var sw = Stopwatch.StartNew();

        int count = end >= start ? end - start + 1 : 0;
        var tasks = new Task[count];
        for (int i = 0; i < count; i++)
        {
            int id = start + i;
            tasks[i] = ScanAndEmitAsync(id, skinStart, skinEnd);
        }

        try
        {
            await Task.WhenAll(tasks);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("hero task fault: " + ex.Message);
            return 1;
        }

        sw.Stop();

        string ts = DateTimeOffset.UtcNow
            .ToOffset(TimeSpan.FromHours(8))
            .ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

        // 请求成分汇总（stderr，供瓶颈分析）
        Console.Error.WriteLine(
            $"{{\"ts\": \"{ts}\", \"elapsed_s\": {sw.Elapsed.TotalSeconds.ToString("F2", CultureInfo.InvariantCulture)}, " +
            $"\"requests\": {StatRequests}, \"ok200\": {StatOk200}, " +
            $"\"confirm404\": {StatConfirm404}, \"final404\": {StatFinal404}, " +
            $"\"retries\": {StatRetries}, \"failed\": {StatFailed}, " +
            $"\"version\": \"{ver}\", \"threads\": {threads}, \"range\": {(UseRange ? "true" : "false")}}}");

        return 0;
    }

    static async Task WarmOneAsync()
    {
        try
        {
            using var resp = await Http.GetAsync(
                BaseUrl, HttpCompletionOption.ResponseHeadersRead);
            await resp.Content.LoadIntoBufferAsync(64 * 1024);
        }
        catch { }
    }

    static async Task ScanAndEmitAsync(int heroId, int skinStart, int skinEnd)
    {
        var result = await ScanHeroAsync(heroId, skinStart, skinEnd);
        Emit(result);
    }

    // --------------------------------------------------------
    // 扫描一个英雄：先探测 00 槽，空号直接跳过；
    // 命中后并行探测剩余槽位（01~19），保持 512 并发全程跑满
    // --------------------------------------------------------

    static async Task<HeroResult> ScanHeroAsync(int heroId, int skinStart, int skinEnd)
    {
        var first = await CheckSkinAsync(heroId, skinStart);

        if (!first.Exists && !first.Failed)
        {
            // 00 槽 404：英雄编号不存在，跳过剩余槽位
            return new HeroResult(heroId, Array.Empty<SkinResult>(), 0);
        }

        int remain = skinEnd > skinStart ? skinEnd - skinStart : 0;
        var results = new List<SkinResult>(remain + 1);
        int failed = 0;
        int maxFound = -1;

        if (first.Exists)
        {
            results.Add(first);
            maxFound = first.SkinId;
        }
        else if (first.Failed)
        {
            failed++;
        }

        if (remain > 0)
        {
            var subTasks = new Task<SkinResult>[remain];
            for (int i = 0; i < remain; i++)
            {
                subTasks[i] = CheckSkinAsync(heroId, skinStart + 1 + i);
            }

            var subResults = await Task.WhenAll(subTasks);
            for (int i = 0; i < subResults.Length; i++)
            {
                var r = subResults[i];
                if (r.Exists)
                {
                    results.Add(r);
                    if (r.SkinId > maxFound) maxFound = r.SkinId;
                }
                else if (r.Failed)
                {
                    failed++;
                }
            }
        }

        // 尾部自适应续探：当有效槽位逼近 skinEnd（如 >= 17 槽）时，
        // 自动向后成批探测直到连续 4 个 404，防止未来突破 19 槽漏扫
        if (maxFound >= skinEnd - TAIL_TRIGGER_MARGIN && skinEnd < MAX_DYNAMIC_SKIN)
        {
            int nextSlot = skinEnd + 1;
            int consecutiveMisses = skinEnd - maxFound;
            int failOnlyBatches = 0;

            while (nextSlot <= MAX_DYNAMIC_SKIN && consecutiveMisses < TAIL_STOP_MISSES)
            {
                int batchSize = Math.Min(
                    TAIL_STOP_MISSES - consecutiveMisses,
                    MAX_DYNAMIC_SKIN - nextSlot + 1);

                var extTasks = new Task<SkinResult>[batchSize];
                for (int i = 0; i < batchSize; i++)
                {
                    extTasks[i] = CheckSkinAsync(heroId, nextSlot + i);
                }

                var extResults = await Task.WhenAll(extTasks);
                bool batchHit = false;
                bool batchFailed = false;

                for (int i = 0; i < extResults.Length; i++)
                {
                    var r = extResults[i];
                    if (r.Exists)
                    {
                        results.Add(r);
                        consecutiveMisses = 0;
                        batchHit = true;
                    }
                    else if (r.Failed)
                    {
                        failed++;
                        batchFailed = true;
                        consecutiveMisses = 0;
                    }
                    else
                    {
                        consecutiveMisses++;
                    }
                }

                nextSlot += batchSize;

                if (batchHit)
                {
                    failOnlyBatches = 0;
                }
                else if (batchFailed)
                {
                    failOnlyBatches++;
                    if (failOnlyBatches >= 2) break;
                }
            }
        }

        return new HeroResult(heroId, results, failed);
    }

    // --------------------------------------------------------
    // 检查单个资源（对应 hero_scan.py 的 check_skin）
    // --------------------------------------------------------

    static async Task<SkinResult> CheckSkinAsync(int heroId, int skinId)
    {
        string url = $"{BaseUrl}30{heroId:D3}{skinId:D2}.jpg";
        bool useRange = UseRange;

        for (int attempt = 1; attempt <= RETRIES; attempt++)
        {
            await WaitThrottleAsync();

            HttpResponseMessage? response = null;
            bool gateHeld = false;

            try
            {
                await Gate.WaitAsync();
                gateHeld = true;

                // 首次请求不带缓存标记，让边缘节点正常命中；
                // 重试才带上，强制回源拿最新结果
                string requestUrl = attempt > 1
                    ? url + "?_=" + CacheBuster
                    : url;

                using var request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
                if (useRange)
                {
                    request.Headers.Range = FirstByteRange;
                }

                Interlocked.Increment(ref StatRequests);
                response = await Http.SendAsync(
                    request, HttpCompletionOption.ResponseHeadersRead);

                int status = (int)response.StatusCode;

                // 极端保护：若遇 0 字节文件导致 416，去掉 Range 立即重试
                if (status == 416 && useRange)
                {
                    response.Dispose();
                    response = null;
                    Gate.Release();
                    gateHeld = false;

                    useRange = false;
                    attempt--;
                    continue;
                }

                // 429 / 5xx 是临时繁忙或限流，要重试；
                // 其余非 200/206 才是真的不存在
                if (status == 429 || (status >= 500 && status < 600))
                {
                    response.Dispose();
                    response = null;
                    Gate.Release();
                    gateHeld = false;

                    if (attempt < RETRIES)
                    {
                        if (status == 429) Note429();
                        Interlocked.Increment(ref StatRetries);
                        await Task.Delay(DelayMs(attempt));
                        continue;
                    }

                    Interlocked.Increment(ref StatFailed);
                    LogError($"{heroId}/{skinId}: HTTP {status} after {RETRIES} attempts");
                    return new SkinResult(skinId, null, Failed: true);
                }

                if (status != 200 && status != 206)
                {
                    bool stale = MayBeStaleCache(response);
                    response.Dispose();
                    response = null;
                    Gate.Release();
                    gateHeld = false;

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
                response = null;
                Gate.Release();
                gateHeld = false;

                // 200/206 但没有 Last-Modified：响应不完整，重试
                if (lastModified is null)
                {
                    if (attempt < RETRIES)
                    {
                        Interlocked.Increment(ref StatRetries);
                        await Task.Delay(DelayMs(attempt));
                        continue;
                    }

                    Interlocked.Increment(ref StatFailed);
                    LogError($"{heroId}/{skinId}: HTTP {status} missing Last-Modified");
                    return new SkinResult(skinId, null, Failed: true);
                }

                Interlocked.Increment(ref StatOk200);
                return new SkinResult(skinId, lastModified, Failed: false);
            }
            catch (Exception ex)
            {
                response?.Dispose();
                if (gateHeld)
                {
                    Gate.Release();
                }

                // 偶发抖动/超时：等一下再试，
                // 不要因为一次失败就把这条记录整个丢掉
                if (attempt < RETRIES)
                {
                    Interlocked.Increment(ref StatRetries);
                    await Task.Delay(DelayMs(attempt));
                    continue;
                }

                Interlocked.Increment(ref StatFailed);
                LogError($"{heroId}/{skinId}: {ex.GetType().Name}: {ex.Message}");
                return new SkinResult(skinId, null, Failed: true);
            }
        }

        return new SkinResult(skinId, null, Failed: true);
    }

    // --------------------------------------------------------
    // 工具
    // --------------------------------------------------------

    static void LogError(string message)
    {
        lock (PrintLock)
        {
            Console.Error.WriteLine(message);
        }
    }

    static string? ReadLastModified(HttpResponseMessage response)
    {
        // Last-Modified 是实体头，.NET 把它放在 Content.Headers 而不是 response.Headers
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
                return !l.Contains("MISS", StringComparison.OrdinalIgnoreCase);
            }
        }

        if (response.Headers.TryGetValues("X-Cache", out var xCache))
        {
            string x = string.Join(",", xCache);
            if (x.Contains("MISS", StringComparison.OrdinalIgnoreCase)) return false;
            if (x.Contains("HIT", StringComparison.OrdinalIgnoreCase)) return true;
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
        var buffer = new ArrayBufferWriter<byte>(256);
        using (var writer = new Utf8JsonWriter(buffer))
        {
            writer.WriteStartObject();
            writer.WriteNumber("hero_id", hero.HeroId);
            writer.WriteNumber("failed", hero.Failed);
            writer.WriteStartArray("results");
            for (int i = 0; i < hero.Results.Count; i++)
            {
                var r = hero.Results[i];
                writer.WriteStartObject();
                writer.WriteNumber("skin_id", r.SkinId);
                writer.WriteString("lm", r.LastModified);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WriteEndObject();
        }

        string line = Encoding.UTF8.GetString(buffer.WrittenSpan);

        lock (PrintLock)
        {
            Console.Out.WriteLine(line);
            Console.Out.Flush();
        }
    }
}
