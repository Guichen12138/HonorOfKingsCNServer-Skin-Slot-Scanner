# -*- coding: utf-8 -*-
import concurrent.futures
import json
import os
import random
import subprocess
import sys
import threading
import time
from datetime import timedelta, timezone
from email.utils import parsedate_to_datetime

# 共用常量与工具：Excel 路径 / 槽位范围 / 表头解析 / URL 拼装 / 官网英雄名
# （common.py 和本脚本在同一目录，Python 会自动找到）
from common import (
    APP_DIR,
    EXCEL_FILE,
    SKIN_END,
    SKIN_START,
    fetch_official_hero_names as _common_fetch_hero_names,
    parse_header_columns,
    skin_image_url,
)

# 海报增量保存是本机自用功能（poster_save.py 不进仓库）：
# 仅在「已命名皮肤的 Last-Modified 更新」时顺带刷新本地海报，
# 没有 poster_save.py 时自动跳过，不影响 GitHub 仓库的独立运行。
try:
    from poster_save import save_poster as _save_poster
except ImportError:
    _save_poster = None


# ============================================================
# 配置
# ============================================================

# 英雄编号范围
HERO_START = 1
HERO_END = 999

# 并发线程（仅在降级到 Python 引擎时使用）
THREADS = 192

# 请求超时时间
TIMEOUT = 10

# 单个地址失败后的重试次数
RETRIES = 3

# 每次重试前的等待秒数（按重试次数递增）
RETRY_DELAY = 0.5

# 缓存标记：只在重试时带上，强制回源拿最新结果
CACHE_BUSTER = str(int(time.time() * 1000))

# C# 扫描引擎：同目录存在 hero_scan_cs.exe 就优先使用，
# 不存在或启动失败自动退回下方的 Python 多线程引擎
CS_EXE = os.path.join(APP_DIR, "hero_scan_cs.exe")
CS_THREADS = 512

# C# 引擎错误日志超过 1MB 自动轮转
CS_LOG_MAX_BYTES = 1024 * 1024

# 尾部自适应续探（与 C# 引擎保持一致，突破 SKIN_END=19 上限永不漏扫）
TAIL_TRIGGER_MARGIN = 2
TAIL_STOP_MISSES = 4
MAX_DYNAMIC_SKIN = 99


# ============================================================
# 懒加载 requests（走 C# 引擎时省掉 ~0.36s 导入耗时）
# ============================================================

_requests_mod = None
_requests_lock = threading.Lock()


def _get_requests():

    global _requests_mod

    if _requests_mod is None:
        with _requests_lock:
            if _requests_mod is None:
                import requests as _req
                _requests_mod = _req

    return _requests_mod


# ============================================================
# Session（Python 后备引擎用）
# ============================================================

thread_local = threading.local()


def get_session():

    if not hasattr(thread_local, "session"):

        requests = _get_requests()
        session = requests.Session()

        session.headers.update({
            "User-Agent": "Mozilla/5.0",
            # 让 CDN 回源校验，不要直接返回边缘节点上的过期副本
            "Cache-Control": "no-cache",
            "Pragma": "no-cache",
        })

        thread_local.session = session

    return thread_local.session


# ============================================================
# 全局 429 冷却（重试风暴的刹车）
# ============================================================

_throttle_lock = threading.Lock()
_throttle_until = 0.0


def _note_429():

    global _throttle_until

    with _throttle_lock:
        _throttle_until = max(
            _throttle_until,
            time.monotonic() + 1.0,
        )


def _wait_throttle():

    while True:
        with _throttle_lock:
            wait = _throttle_until - time.monotonic()

        if wait <= 0:
            return

        time.sleep(min(wait, 0.5))


# ============================================================
# 判断一个 404 响应是不是可能来自边缘节点的旧缓存
# ============================================================

def _may_be_stale_cache(response):

    lookup = response.headers.get("X-Cache-Lookup") or ""

    if lookup:
        return "MISS" not in lookup.upper()

    x_cache = (response.headers.get("X-Cache") or "").upper()

    if x_cache:
        if "MISS" in x_cache:
            return False
        if "HIT" in x_cache:
            return True

    age = response.headers.get("Age")

    try:
        return int(age) > 0
    except (TypeError, ValueError):
        return True


# ============================================================
# 检查单个资源（Python 后备引擎）
# ============================================================

def check_skin(hero_id, skin_id):

    requests = _get_requests()
    url = skin_image_url(hero_id, skin_id)
    use_range = True

    # 用 GET + Range: bytes=0-0：
    # 1. 和普通 GET 共享同一份边缘缓存（避开 HEAD 缓存不同步的坑）；
    # 2. 206 响应正文仅 1 字节，读掉后连接可直接复用（不会下载 200KB 图片也不会销毁连接）；
    # 3. 极端情况遇 0 字节文件返回 416 时自动回退普通 GET。
    attempt = 1
    while attempt <= RETRIES:

        _wait_throttle()

        try:

            session = get_session()
            req_headers = {"Range": "bytes=0-0"} if use_range else None

            response = session.get(
                url,
                timeout=TIMEOUT,
                allow_redirects=True,
                stream=True,
                headers=req_headers,
                params=(
                    {"_": CACHE_BUSTER}
                    if attempt > 1
                    else None
                ),
            )

            try:

                status = response.status_code

                # 0 字节文件触发 416：去掉 Range 立即重发，不消耗重试次数
                if status == 416 and use_range:
                    response.content
                    use_range = False
                    continue

                if status == 429 or 500 <= status < 600:
                    response.content

                    if attempt < RETRIES:
                        if status == 429:
                            _note_429()

                        time.sleep(
                            RETRY_DELAY * attempt
                            + random.uniform(0.1, 0.6)
                        )
                        attempt += 1
                        continue

                    return {
                        "hero_id": hero_id,
                        "skin_id": skin_id,
                        "exists": False,
                        "failed": True,
                    }

                if status not in (200, 206):
                    # 读掉 404 的小错误体，复用长连接
                    response.content

                    if (
                        status == 404
                        and attempt == 1
                        and _may_be_stale_cache(response)
                    ):
                        attempt += 1
                        continue

                    return {
                        "hero_id": hero_id,
                        "skin_id": skin_id,
                        "exists": False,
                    }

                # 206 正文只有 1 字节，读掉即可复用连接；200 不读大图片正文
                if status == 206:
                    response.content

                last_modified = response.headers.get("Last-Modified", "")

                if not last_modified:
                    if attempt < RETRIES:
                        time.sleep(
                            RETRY_DELAY * attempt
                            + random.uniform(0.1, 0.6)
                        )
                        attempt += 1
                        continue

                    return {
                        "hero_id": hero_id,
                        "skin_id": skin_id,
                        "exists": False,
                        "failed": True,
                    }

                try:
                    dt = parsedate_to_datetime(last_modified)
                    # GMT/UTC → 北京时间 UTC+8
                    dt = dt.astimezone(timezone(timedelta(hours=8)))
                    last_modified = dt.strftime("%Y-%m-%d %H:%M:%S")
                except Exception:
                    pass

                return {
                    "hero_id": hero_id,
                    "skin_id": skin_id,
                    "exists": True,
                    "last_modified": last_modified,
                }

            finally:
                response.close()

        except requests.RequestException:

            if attempt < RETRIES:
                time.sleep(
                    RETRY_DELAY * attempt
                    + random.uniform(0.1, 0.6)
                )
                attempt += 1
                continue
            break

    return {
        "hero_id": hero_id,
        "skin_id": skin_id,
        "exists": False,
        "failed": True,
    }


# ============================================================
# 扫描一个英雄（保留独立接口）
# ============================================================

def scan_hero(hero_id):

    results = []
    failed = 0

    first = check_skin(hero_id, SKIN_START)

    if first["exists"]:
        results.append(first)
    elif first.get("failed"):
        failed += 1
    else:
        return hero_id, results, failed

    for skin_id in range(SKIN_START + 1, SKIN_END + 1):
        result = check_skin(hero_id, skin_id)
        if result["exists"]:
            results.append(result)
        elif result.get("failed"):
            failed += 1

    return hero_id, results, failed


# ============================================================
# Python 多线程全量扫描（后备引擎，同样采用 00 槽命中后全并行扇出）
# ============================================================

def scan_with_python(total_heroes):

    all_results = {}
    total_failed = 0
    scan_errors = 0
    skipped_heroes = 0
    completed = 0
    interrupted = False

    remain_slots = list(range(SKIN_START + 1, SKIN_END + 1))

    # 每个命中 00 槽的英雄在并行槽位收齐前暂存结果
    hero_state = {}
    last_print_ts = 0.0
    last_printed_completed = -1

    def _report_progress(force=False):
        nonlocal last_print_ts, last_printed_completed
        if completed == last_printed_completed:
            return
        now = time.monotonic()
        if force or completed == total_heroes or (now - last_print_ts) >= 0.05:
            last_print_ts = now
            last_printed_completed = completed
            print(
                f"\r扫描进度：{completed}/{total_heroes}",
                end="",
                flush=True,
            )

    def _finish_hero(hid, st):
        nonlocal total_failed, completed
        hero_state.pop(hid, None)
        total_failed += st["failed"]
        if st["results"]:
            st["results"].sort(key=lambda x: x["skin_id"])
            all_results[hid] = st["results"]
        completed += 1
        _report_progress()

    def _maybe_extend_or_finish(hid, st, pending_map, exec_pool):
        # 尾部自适应续探：与 C# 引擎保持一致
        if not st.get("in_tail"):
            max_found = max((r["skin_id"] for r in st["results"]), default=-1)
            if max_found < SKIN_END - TAIL_TRIGGER_MARGIN or SKIN_END >= MAX_DYNAMIC_SKIN:
                _finish_hero(hid, st)
                return
            st["in_tail"] = True
            st["next_slot"] = SKIN_END + 1
            st["consec_misses"] = SKIN_END - max_found
            st["fail_only_batches"] = 0
        else:
            batch_res = st.pop("tail_batch", {})
            batch_hit = False
            batch_failed = False
            for sid in sorted(batch_res):
                r = batch_res[sid]
                if r["exists"]:
                    st["results"].append(r)
                    st["consec_misses"] = 0
                    batch_hit = True
                elif r.get("failed"):
                    st["failed"] += 1
                    st["consec_misses"] = 0
                    batch_failed = True
                else:
                    st["consec_misses"] += 1

            if batch_hit:
                st["fail_only_batches"] = 0
            elif batch_failed:
                st["fail_only_batches"] += 1
                if st["fail_only_batches"] >= 2:
                    _finish_hero(hid, st)
                    return

        if st["next_slot"] > MAX_DYNAMIC_SKIN or st["consec_misses"] >= TAIL_STOP_MISSES:
            _finish_hero(hid, st)
            return

        batch_size = min(
            TAIL_STOP_MISSES - st["consec_misses"],
            MAX_DYNAMIC_SKIN - st["next_slot"] + 1,
        )
        start_s = st["next_slot"]
        st["next_slot"] = start_s + batch_size
        st["left"] = batch_size
        st["tail_batch"] = {}
        for sid in range(start_s, start_s + batch_size):
            pending_map[exec_pool.submit(check_skin, hid, sid)] = (hid, "tail", sid)

    executor = concurrent.futures.ThreadPoolExecutor(max_workers=THREADS)

    try:
        # 第一波：并发探测所有英雄的 00 槽
        pending = {
            executor.submit(check_skin, hero_id, SKIN_START): (hero_id, "first", SKIN_START)
            for hero_id in range(HERO_START, HERO_END + 1)
        }

        while pending:
            done_set, _ = concurrent.futures.wait(
                pending,
                return_when=concurrent.futures.FIRST_COMPLETED,
            )

            for fut in done_set:
                hero_id, phase, sid = pending.pop(fut)

                try:
                    res = fut.result()
                except Exception:
                    if phase == "first":
                        scan_errors += 1
                        completed += 1
                        _report_progress()
                    else:
                        st = hero_state.get(hero_id)
                        if st is not None:
                            if phase == "tail":
                                st["tail_batch"][sid] = {"hero_id": hero_id, "skin_id": sid, "exists": False, "failed": True}
                            else:
                                st["failed"] += 1
                            st["left"] -= 1
                            if st["left"] == 0:
                                _maybe_extend_or_finish(hero_id, st, pending, executor)
                    continue

                if phase == "first":
                    if not res["exists"] and not res.get("failed"):
                        # 00 槽 404：空号直接跳过
                        skipped_heroes += 1
                        completed += 1
                        _report_progress()
                        continue

                    results = [res] if res["exists"] else []
                    failed = 1 if res.get("failed") else 0

                    if not remain_slots:
                        st = {"results": results, "failed": failed, "left": 0}
                        hero_state[hero_id] = st
                        _maybe_extend_or_finish(hero_id, st, pending, executor)
                        continue

                    # 00 槽命中/失败：把 01~19 槽全部并入同一线程池并发扫描
                    hero_state[hero_id] = {
                        "results": results,
                        "failed": failed,
                        "left": len(remain_slots),
                    }
                    for s in remain_slots:
                        pending[executor.submit(check_skin, hero_id, s)] = (hero_id, "main", s)
                elif phase == "main":
                    st = hero_state.get(hero_id)
                    if st is None:
                        continue
                    if res["exists"]:
                        st["results"].append(res)
                    elif res.get("failed"):
                        st["failed"] += 1
                    st["left"] -= 1

                    if st["left"] == 0:
                        _maybe_extend_or_finish(hero_id, st, pending, executor)
                else:
                    # phase == "tail"
                    st = hero_state.get(hero_id)
                    if st is None:
                        continue
                    st["tail_batch"][sid] = res
                    st["left"] -= 1
                    if st["left"] == 0:
                        _maybe_extend_or_finish(hero_id, st, pending, executor)

        _report_progress(force=True)

    except KeyboardInterrupt:
        interrupted = True
        # 已经扫出部分结果的英雄也保留，不白跑
        for hid, st in hero_state.items():
            if st["results"]:
                st["results"].sort(key=lambda x: x["skin_id"])
                all_results[hid] = st["results"]
            total_failed += st["failed"]

        executor.shutdown(wait=False, cancel_futures=True)
    else:
        executor.shutdown(wait=True)

    return (
        all_results,
        total_failed,
        scan_errors,
        skipped_heroes,
        interrupted,
        completed,
    )


# ============================================================
# 官网英雄列表（复用 common.py 标准库实现）
# ============================================================

def fetch_official_hero_names():

    return _common_fetch_hero_names(timeout=TIMEOUT, warn_on_error=True)


# ============================================================
# C# 引擎：拉起 hero_scan_cs.exe 并按英雄收结果
# ============================================================

def _rotate_log_if_needed(path, max_bytes=CS_LOG_MAX_BYTES):

    try:
        if os.path.exists(path) and os.path.getsize(path) > max_bytes:
            os.replace(path, path + ".old")
    except OSError:
        pass


def scan_with_cs(total_heroes):

    """
    调用同目录的 hero_scan_cs.exe 扫描。

    返回: (all_results, total_failed, scan_errors, skipped_heroes,
           interrupted, completed, returncode)
    """

    all_results = {}
    total_failed = 0
    scan_errors = 0
    skipped_heroes = 0
    completed = 0
    interrupted = False
    last_print_ts = 0.0

    stderr_path = os.path.join(APP_DIR, "hero_scan_cs_error.log")
    _rotate_log_if_needed(stderr_path)

    with open(stderr_path, "a", encoding="utf-8") as stderr_fh:

        proc = subprocess.Popen(
            [
                CS_EXE,
                str(HERO_START),
                str(HERO_END),
                str(CS_THREADS),
                str(SKIN_START),
                str(SKIN_END),
            ],
            stdout=subprocess.PIPE,
            stderr=stderr_fh,
            text=True,
            encoding="utf-8",
            errors="replace",
            cwd=APP_DIR,
            creationflags=(
                subprocess.CREATE_NO_WINDOW
                if os.name == "nt"
                else 0
            ),
        )

        try:

            for line in proc.stdout:

                line = line.strip()

                if not line:
                    continue

                try:
                    record = json.loads(line)
                except ValueError:
                    scan_errors += 1
                    continue

                if "hero_id" not in record:
                    continue

                hero_id = record["hero_id"]
                failed = record.get("failed", 0)

                results = [
                    {
                        "skin_id": item["skin_id"],
                        "last_modified": item["lm"],
                        "exists": True,
                    }
                    for item in record.get("results", [])
                ]

                completed += 1
                total_failed += failed

                if results:
                    all_results[hero_id] = results
                elif not failed:
                    skipped_heroes += 1

                now = time.monotonic()
                if completed == total_heroes or (now - last_print_ts) >= 0.05:
                    last_print_ts = now
                    print(
                        f"\r扫描进度：{completed}/{total_heroes}",
                        end="",
                        flush=True,
                    )

        except KeyboardInterrupt:
            interrupted = True
            proc.terminate()

        try:
            proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            proc.kill()
            proc.wait()

    returncode = proc.returncode

    if returncode not in (0, None) and not interrupted:
        print()
        print(
            f"注意：C# 引擎异常退出"
            f"（返回码 {returncode}，"
            f"详情见 hero_scan_cs_error.log），"
            f"本次只完成 {completed}/{total_heroes} 个英雄。"
        )

    return (
        all_results,
        total_failed,
        scan_errors,
        skipped_heroes,
        interrupted,
        completed,
        returncode,
    )


# ============================================================
# 并行加载 openpyxl 与 Excel（与扫描重叠，省掉 ~0.57s 串行等待）
# ============================================================

def _load_or_create_excel():

    from openpyxl import Workbook, load_workbook

    if os.path.exists(EXCEL_FILE):
        excel_existed = True
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
    else:
        excel_existed = False
        wb = Workbook()
        ws = wb.active
        ws.title = "皮肤资源"

    # 确保基本表头存在
    ws.cell(row=1, column=1, value="英雄编号")
    ws.cell(row=1, column=2, value="英雄名字")

    skin_name_columns, last_modified_columns = parse_header_columns(ws)

    # 增加缺失的皮肤格子列
    for skin_id in range(SKIN_START, SKIN_END + 1):
        if skin_id not in skin_name_columns:
            new_col = ws.max_column + 1
            ws.cell(row=1, column=new_col, value=f"{skin_id:02d}-皮肤名")
            skin_name_columns[skin_id] = new_col

        if skin_id not in last_modified_columns:
            new_col = ws.max_column + 1
            ws.cell(row=1, column=new_col, value=f"{skin_id:02d}-Last-Modified")
            last_modified_columns[skin_id] = new_col

    # 找已有英雄
    hero_rows = {}

    for row in range(2, ws.max_row + 1):
        hero_id = ws.cell(row=row, column=1).value
        if hero_id is None:
            continue
        try:
            hero_rows[int(hero_id)] = row
        except (TypeError, ValueError):
            pass

    return (
        wb,
        ws,
        excel_existed,
        skin_name_columns,
        last_modified_columns,
        hero_rows,
    )


# ============================================================
# 主程序
# ============================================================

def main():

    print("=" * 70)
    print("王者荣耀 HeroOriginalPainting 增量扫描器")
    print("=" * 70)

    total_heroes = HERO_END - HERO_START + 1

    # ========================================================
    # 官网英雄名 + Excel 载入全部放到后台并行，
    # 扫描引擎第一时间拉起，消除 ~0.9s 串行启动开销
    # ========================================================

    official_names_holder = {}
    official_thread = threading.Thread(
        target=lambda: official_names_holder.update(
            fetch_official_hero_names()
        ),
        daemon=True,
    )
    official_thread.start()

    excel_holder = {}

    def _excel_loader():
        try:
            excel_holder["data"] = _load_or_create_excel()
        except Exception as exc:
            excel_holder["error"] = exc

    excel_thread = threading.Thread(
        target=_excel_loader,
        daemon=True,
    )
    excel_thread.start()

    # ========================================================
    # 引擎选择：有 hero_scan_cs.exe 就用 C#，
    # 没有 / 启动失败 / 零产出崩溃 自动退回 Python 多线程
    # ========================================================

    use_cs = os.path.exists(CS_EXE)
    cs_abnormal = False

    print(
        f"扫描范围：{HERO_START:03d}~{HERO_END:03d} 号、"
        f"皮肤槽 {SKIN_START:02d}~{SKIN_END:02d} · "
        f"扫描引擎："
        + ("C#（hero_scan_cs.exe）" if use_cs else "Python 多线程")
    )

    all_results = {}
    total_failed = 0
    scan_errors = 0
    skipped_heroes = 0
    interrupted = False
    completed = 0

    if use_cs:
        try:
            (
                all_results,
                total_failed,
                scan_errors,
                skipped_heroes,
                interrupted,
                completed,
                cs_rc,
            ) = scan_with_cs(total_heroes)

            if cs_rc not in (0, None) and not interrupted:
                if completed == 0:
                    use_cs = False
                    print("C# 引擎零结果异常退出，自动改用 Python 多线程方案。")
                else:
                    cs_abnormal = True

        except OSError:
            use_cs = False
            print("C# 引擎启动失败，改用 Python 多线程方案。")

    if not use_cs:
        (
            all_results,
            total_failed,
            scan_errors,
            skipped_heroes,
            interrupted,
            completed,
        ) = scan_with_python(total_heroes)

    # 这个换行不能删：进度行用 \r 刷新、末尾没有换行，
    # 少了它下一段文字会和进度文本粘成一行、被当成进度吞掉
    print()

    # 等 Excel 加载完成（绝大多数情况下扫描结束时早已载好）
    excel_thread.join()

    if "error" in excel_holder:
        raise excel_holder["error"]

    (
        wb,
        ws,
        excel_existed,
        skin_name_columns,
        last_modified_columns,
        hero_rows,
    ) = excel_holder["data"]

    if excel_existed:
        print(
            f"发现已有 Excel：{EXCEL_FILE}"
            f"（已有 {len(hero_rows)} 行记录），已有数据将被保留。"
        )
    else:
        print("没有找到 Excel，正在创建新的文件……")

    if interrupted:
        print(
            f"已手动中断：本次只扫完 {completed}/{total_heroes} 个英雄，"
            "已扫完的部分会照常保存，建议之后完整重跑一次。"
        )
    else:
        print("扫描完成，开始更新 Excel……")

    # ========================================================
    # 更新 Excel
    # ========================================================

    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    cell_fill = PatternFill(fill_type="solid", fgColor="E2F0D9")

    new_heroes = 0
    new_records = 0
    updated_records = 0

    # 动态补充超出 SKIN_END 的表头列（配合尾部自适应续探，突破 19 槽上限）
    max_found_skin = max(
        (r["skin_id"] for res in all_results.values() for r in res),
        default=SKIN_END,
    )
    max_table_skin = max(
        SKIN_END,
        max_found_skin,
        max(skin_name_columns, default=SKIN_END),
        max(last_modified_columns, default=SKIN_END),
    )
    added_slot_cols = 0

    for sid in range(SKIN_START, max_table_skin + 1):
        if sid not in skin_name_columns:
            new_col = ws.max_column + 1
            ws.cell(row=1, column=new_col, value=f"{sid:02d}-皮肤名")
            skin_name_columns[sid] = new_col
            added_slot_cols += 1

        if sid not in last_modified_columns:
            new_col = ws.max_column + 1
            ws.cell(row=1, column=new_col, value=f"{sid:02d}-Last-Modified")
            last_modified_columns[sid] = new_col
            added_slot_cols += 1

    # 记录变更明细，汇总输出并用于同步已更新的本地海报
    added_items = []
    updated_items = []

    for hero_id in sorted(all_results):

        results = all_results[hero_id]

        if hero_id in hero_rows:
            row = hero_rows[hero_id]
        else:
            row = ws.max_row + 1
            ws.cell(row=row, column=1, value=hero_id)
            ws.cell(row=row, column=2, value="")
            hero_rows[hero_id] = row
            new_heroes += 1

        for result in results:

            skin_id = result["skin_id"]
            last_modified = result["last_modified"]

            if skin_id not in skin_name_columns or skin_id not in last_modified_columns:
                continue

            name_col = skin_name_columns[skin_id]
            time_col = last_modified_columns[skin_id]

            # 皮肤名：绝对不覆盖已有内容
            skin_name = ws.cell(row=row, column=name_col).value
            has_skin_name = skin_name is not None and str(skin_name).strip() != ""

            old_time = ws.cell(row=row, column=time_col).value

            # Last-Modified：以服务器最新值为准；非字符串单元格也强制自愈为文本
            if last_modified:
                if (
                    old_time is None
                    or not isinstance(old_time, str)
                    or str(old_time).strip() != str(last_modified).strip()
                ):
                    time_cell = ws.cell(
                        row=row,
                        column=time_col,
                        value=last_modified,
                    )
                    time_cell.number_format = "@"

                    name_cell = ws.cell(row=row, column=name_col)
                    time_cell.fill = cell_fill
                    name_cell.fill = cell_fill

                    is_new = old_time is None or str(old_time).strip() == ""
                    if is_new:
                        new_records += 1
                        added_items.append((hero_id, row, skin_id, last_modified))
                    else:
                        updated_records += 1
                        updated_items.append((
                            hero_id,
                            row,
                            skin_id,
                            str(old_time).strip(),
                            last_modified,
                            has_skin_name,
                            str(skin_name).strip() if has_skin_name else "",
                        ))

    # ========================================================
    # 英雄名自动填充（官网英雄列表）
    # 只填空格、只填有 00 槽资源的真实英雄行，已有名字绝不碰
    # ========================================================

    auto_named = 0

    official_thread.join(timeout=5)

    if official_thread.is_alive():
        print("官网英雄名拉取超时，本次跳过自动填充。")
        official_names = {}
    else:
        official_names = dict(official_names_holder)

    first_time_col = (
        last_modified_columns.get(SKIN_START)
        if official_names
        else None
    )

    if official_names and first_time_col is not None:

        for row in range(2, ws.max_row + 1):
            hero_id = ws.cell(row=row, column=1).value
            if hero_id is None:
                continue

            try:
                hero_id = int(hero_id)
            except (TypeError, ValueError):
                continue

            if not ws.cell(row=row, column=first_time_col).value:
                continue

            name_cell = ws.cell(row=row, column=2)
            if name_cell.value is not None and str(name_cell.value).strip() != "":
                continue

            official_name = official_names.get(hero_id)
            if official_name:
                name_cell.value = official_name
                name_cell.fill = cell_fill
                auto_named += 1

    # ========================================================
    # 保存（仅在有变化或新建时设置样式并原子写入）
    # ========================================================

    has_changes = bool(
        new_heroes
        or new_records
        or updated_records
        or auto_named
        or added_slot_cols
    )

    if has_changes or not excel_existed:

        header_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
        header_font = Font(bold=True)
        header_align = Alignment(horizontal="center", vertical="center")

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_align

        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 15

        for skin_id in range(SKIN_START, max_table_skin + 1):
            name_col = skin_name_columns[skin_id]
            time_col = last_modified_columns[skin_id]
            ws.column_dimensions[get_column_letter(name_col)].width = 18
            ws.column_dimensions[get_column_letter(time_col)].width = 22

        ws.freeze_panes = "C2"

        # 清除 Excel 文档属性中的最后修改者姓名，避免泄露本机系统/Office 用户名
        try:
            wb.properties.creator = "openpyxl"
            wb.properties.lastModifiedBy = ""
        except Exception:
            pass

        # 原子写入：先落临时文件再替换，避免写入中途被杀掉损坏原表
        tmp_file = EXCEL_FILE + ".tmp"
        try:
            wb.save(tmp_file)
            os.replace(tmp_file, EXCEL_FILE)
        except Exception:
            try:
                if os.path.exists(tmp_file):
                    os.remove(tmp_file)
            except OSError:
                pass
            raise
    else:
        print("数据无变化，跳过写入 Excel。")

    # ========================================================
    # 变更明细 & 已命名皮肤海报更新同步
    # ========================================================

    def _hero_label(hid, r):
        nm = ws.cell(row=r, column=2).value
        nm_str = str(nm).strip() if nm else ""
        return f"{hid:03d} {nm_str}" if nm_str else f"{hid:03d}"

    if added_items:
        print("—— 本次新增格子 ——")
        for hid, r, sid, lm in added_items:
            print(f"  [新增] {_hero_label(hid, r)} · {sid:02d}槽 ({lm})")

    if updated_items:
        print("—— 本次更新格子 ——")
        for hid, r, sid, old_lm, new_lm, has_name, sname in updated_items:
            tag = f"（{sname}）" if sname else ""
            print(f"  [更新] {_hero_label(hid, r)} · {sid:02d}槽{tag}：{old_lm} → {new_lm}")
            # 已命名皮肤不会被 skin_scan 打开，若本机有 poster_save 则顺带刷新本地原画
            if has_name and _save_poster is not None:
                _save_poster(skin_image_url(hid, sid))

    # ========================================================
    # 统计
    # ========================================================

    print("=" * 70)

    print(
        f"更新完成！新增英雄 {new_heroes} · "
        f"自动填充英雄名 {auto_named} · "
        f"跳过空号英雄 {skipped_heroes}"
    )

    print(
        f"新增 Last-Modified {new_records} · "
        f"更新 Last-Modified {updated_records}"
    )

    print("=" * 70)

    if max_found_skin > SKIN_END:
        print(
            f"动态扩展：已自动续探至 {max_found_skin:02d} 槽并补充表头列。"
        )

    has_errors = bool(total_failed or scan_errors or cs_abnormal)

    if total_failed or scan_errors:
        print(
            f"注意：{total_failed} 个请求、"
            f"{scan_errors} 个英雄任务失败，"
            "这些格子本次没有扫到，建议再完整跑一次补齐。"
        )

    print(
        f"Excel：{EXCEL_FILE}"
        "（已有皮肤名不会被修改；"
        "已有 Last-Modified 以服务器最新值为准，"
        "有更新会自动刷新并高亮）"
    )

    try:
        wb.close()
    except Exception:
        pass

    if has_errors:
        sys.exit(1)


if __name__ == "__main__":

    main()
