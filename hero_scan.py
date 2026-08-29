import requests
import concurrent.futures
import threading
import time
import os
from email.utils import parsedate_to_datetime
from datetime import timezone, timedelta

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter


# ============================================================
# 配置
# ============================================================

EXCEL_FILE = "hero_skin_scan.xlsx"

BASE_URL = "https://image.smoba.qq.com/Picture/HeroOriginalPainting/"

# 英雄编号范围
HERO_START = 1
HERO_END = 999

# 每个英雄扫描 00 ~ 19
SKIN_START = 0
SKIN_END = 19

# 并发线程
#
# 30 → 64：实测 64 再往上（96）没有增益，
# 64 是这台 CDN 的性价比点
THREADS = 64

# 请求超时时间
TIMEOUT = 10


# 单个地址失败后的重试次数
RETRIES = 3

# 每次重试前的等待秒数（按重试次数递增）
RETRY_DELAY = 0.5

# --------------------------------------------------------------------
# 缓存标记
#
# 每一轮扫描使用同一个值，让它不重复、但整轮保持一致。
#
# 作用：这一轮的请求不会命中上一轮留在 CDN 边缘节点上的过期缓存。
#
# 关键点在于「过期的 404」：
# 资源还没上线时扫出来的是 404，边缘节点会把这个 404 缓存下来；
# 等资源真正上传之后，边缘节点可能还在给旧的 404，
# 于是这个最新的资源就一直扫不到，Last-Modified 也就一直写不进去。
# --------------------------------------------------------------------

CACHE_BUSTER = str(int(time.time() * 1000))


# ============================================================
# Session
# ============================================================

thread_local = threading.local()


def get_session():

    if not hasattr(thread_local, "session"):

        session = requests.Session()

        session.headers.update({
            "User-Agent": "Mozilla/5.0",

            # 让 CDN 回源校验，
            # 不要直接返回边缘节点上的过期副本
            "Cache-Control": "no-cache",
            "Pragma": "no-cache"
        })

        thread_local.session = session

    return thread_local.session


# ============================================================
# 检查单个资源
# ============================================================

def check_skin(hero_id, skin_id):

    filename = (
        f"30{hero_id:03d}"
        f"{skin_id:02d}"
        ".jpg"
    )

    url = BASE_URL + filename


    # ================================================================
    # 为什么改成 GET（和 curl -i 一样），不再用 HEAD
    #
    # HEAD 在这个 CDN 上命中的缓存和 GET 不是同一份，
    # 刚更新过的资源经常在 HEAD 上拿到旧的缓存结果，
    # 于是最新的资源扫不出来，Last-Modified 也就一直写不进去。
    #
    # 用 GET 时加 stream=True：
    # 只读取响应头，不下载图片正文，开销和 HEAD 差不多。
    # ================================================================

    for attempt in range(
        1,
        RETRIES + 1
    ):

        try:

            session = get_session()

            response = session.get(
                url,
                timeout=TIMEOUT,
                allow_redirects=True,
                stream=True,
                params={
                    "_": CACHE_BUSTER
                }
            )

            try:

                status = response.status_code


                # ------------------------------------------------
                # 429 / 5xx 是服务器临时繁忙或限流，
                # 不能当成「资源不存在」，要重试
                #
                # 404 之类的才是真的不存在
                # ------------------------------------------------

                if (
                        status == 429
                        or 500 <= status < 600
                ):

                    if attempt < RETRIES:

                        # 读掉小错误体：
                        # 不读的话这个连接会被整个丢弃，
                        # 下一个请求又要重新 TCP+TLS 握手，
                        # 这正是之前扫描慢的主要原因
                        response.content

                        time.sleep(
                            RETRY_DELAY * attempt
                        )

                        continue


                    # 重试次数用完还是失败：
                    # 标记为失败，让统计能看出来

                    response.content

                    return {
                        "hero_id": hero_id,
                        "skin_id": skin_id,
                        "exists": False,
                        "failed": True
                    }


                if status != 200:

                    # 同上：读掉 404 的小错误体
                    # （几百字节的 XML，很便宜），
                    # 让连接可以复用给下一个请求
                    response.content

                    return {
                        "hero_id": hero_id,
                        "skin_id": skin_id,
                        "exists": False
                    }

                last_modified = response.headers.get(
                    "Last-Modified",
                    ""
                )


                # ------------------------------------------------
                # 200 但没有 Last-Modified：
                #
                # 说明这次拿到的响应不完整，
                # 不能当成「存在但没有时间」写进表，
                # 直接重试
                # ------------------------------------------------

                if not last_modified:

                    time.sleep(
                        RETRY_DELAY * attempt
                    )

                    continue


                try:
                    dt = parsedate_to_datetime(last_modified)

                    # GMT/UTC → 北京时间 UTC+8
                    dt = dt.astimezone(
                        timezone(timedelta(hours=8))
                    )

                    last_modified = dt.strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )

                except Exception:
                    pass


                return {
                    "hero_id": hero_id,
                    "skin_id": skin_id,
                    "exists": True,
                    "last_modified": last_modified
                }


            finally:

                # 只取了响应头，
                # 这里必须关掉，否则连接不会释放
                response.close()


        except requests.RequestException:

            # 偶发抖动：等一下再试，
            # 不要因为一次失败就把这条记录整个丢掉
            if attempt < RETRIES:

                time.sleep(
                    RETRY_DELAY * attempt
                )

                continue


    # 重试全部用完还是失败：
    # 标记为失败，让统计能看出来，
    # 提醒用户这一格本次没扫到

    return {
        "hero_id": hero_id,
        "skin_id": skin_id,
        "exists": False,
        "failed": True
    }


# ============================================================
# 扫描一个英雄
# ============================================================

def scan_hero(hero_id):

    results = []

    # 这一英雄有多少个格子重试完还是失败
    failed = 0


    for skin_id in range(
        SKIN_START,
        SKIN_END + 1
    ):

        result = check_skin(
            hero_id,
            skin_id
        )

        if result["exists"]:

            results.append(result)

        elif result.get("failed"):

            failed += 1

    return hero_id, results, failed


# ============================================================
# 主程序
# ============================================================

def main():

    print("=" * 70)
    print("王者荣耀 HeroOriginalPainting 增量扫描器")
    print("=" * 70)
    print()


    # ========================================================
    # 打开 / 创建 Excel
    # ========================================================

    if os.path.exists(EXCEL_FILE):

        print(
            f"发现已有 Excel：{EXCEL_FILE}"
        )

        print(
            "正在读取已有数据……"
        )

        wb = load_workbook(
            EXCEL_FILE
        )

        ws = wb.active

        print(
            "已有数据将被保留，不会初始化覆盖。"
        )

    else:

        print(
            "没有找到 Excel，正在创建新的文件……"
        )

        wb = Workbook()

        ws = wb.active

        ws.title = "皮肤资源"


    # ========================================================
    # 确保基本表头存在
    # ========================================================

    ws.cell(
        row=1,
        column=1,
        value="英雄编号"
    )

    ws.cell(
        row=1,
        column=2,
        value="英雄名字"
    )


    # ========================================================
    # 获取已有列
    # ========================================================

    skin_name_columns = {}
    last_modified_columns = {}


    for col in range(
        1,
        ws.max_column + 1
    ):

        header = ws.cell(
            row=1,
            column=col
        ).value

        if not header:
            continue

        header = str(header)


        # ----------------------------------------------------
        # 皮肤名
        # ----------------------------------------------------

        if header.endswith("-皮肤名"):

            try:

                skin_id = int(
                    header.split("-")[0]
                )

                skin_name_columns[
                    skin_id
                ] = col

            except ValueError:

                pass


        # ----------------------------------------------------
        # Last-Modified
        # ----------------------------------------------------

        elif header.endswith(
            "-Last-Modified"
        ):

            try:

                skin_id = int(
                    header.split("-")[0]
                )

                last_modified_columns[
                    skin_id
                ] = col

            except ValueError:

                pass


    # ========================================================
    # 增加缺失的皮肤格子列
    # ========================================================

    for skin_id in range(
        SKIN_START,
        SKIN_END + 1
    ):

        # ----------------------------------------------------
        # 如果没有这个皮肤名列，就新增
        # ----------------------------------------------------

        if skin_id not in skin_name_columns:

            new_col = ws.max_column + 1

            ws.cell(
                row=1,
                column=new_col,
                value=f"{skin_id:02d}-皮肤名"
            )

            skin_name_columns[
                skin_id
            ] = new_col


        # ----------------------------------------------------
        # 如果没有 Last-Modified 列，就新增
        # ----------------------------------------------------

        if skin_id not in last_modified_columns:

            new_col = ws.max_column + 1

            ws.cell(
                row=1,
                column=new_col,
                value=f"{skin_id:02d}-Last-Modified"
            )

            last_modified_columns[
                skin_id
            ] = new_col


    # ========================================================
    # 找已有英雄
    # ========================================================

    hero_rows = {}


    for row in range(
        2,
        ws.max_row + 1
    ):

        hero_id = ws.cell(
            row=row,
            column=1
        ).value

        if hero_id is None:
            continue

        try:

            hero_id = int(hero_id)

            hero_rows[
                hero_id
            ] = row

        except ValueError:

            pass


    print()

    print(
        f"Excel 中已有 "
        f"{len(hero_rows)} 个英雄。"
    )

    print()

    print(
        f"本次扫描英雄："
        f"{HERO_START:03d} ~ {HERO_END:03d}"
    )

    print(
        f"本次扫描皮肤槽："
        f"{SKIN_START:02d} ~ {SKIN_END:02d}"
    )

    print()

    # ========================================================
    # 开始扫描
    # ========================================================

    total_heroes = (
        HERO_END -
        HERO_START +
        1
    )

    all_results = {}

    # 重试完还是失败的请求数
    total_failed = 0

    # 整个英雄任务崩掉的数量
    scan_errors = 0

    # 是否被 Ctrl+C 手动中断
    interrupted = False


    executor = concurrent.futures.ThreadPoolExecutor(
        max_workers=THREADS
    )

    futures = {
        executor.submit(
            scan_hero,
            hero_id
        ): hero_id
        for hero_id in range(
            HERO_START,
            HERO_END + 1
        )
    }


    completed = 0


    try:

        for future in concurrent.futures.as_completed(
            futures
        ):

            # ------------------------------------------------
            # 单个英雄的任务崩了，
            # 记一笔继续跑，
            # 不能让整轮扫描的结果全部丢掉
            # ------------------------------------------------

            try:

                hero_id, results, failed = future.result()

            except Exception:

                scan_errors += 1

                completed += 1

                continue

            completed += 1

            total_failed += failed


            if results:

                all_results[
                    hero_id
                ] = results


            print(
                f"\r扫描进度："
                f"{completed}/{total_heroes}",
                end="",
                flush=True
            )


    except KeyboardInterrupt:

        # ------------------------------------------------
        # Ctrl+C：取消还没开始的任务，
        # 已扫完的部分照常保存，不白跑
        # ------------------------------------------------

        interrupted = True

        executor.shutdown(
            wait=False,
            cancel_futures=True
        )


    else:

        executor.shutdown(wait=True)


    print()
    print()

    if interrupted:

        print(
            f"已手动中断：本次只扫完 "
            f"{completed}/{total_heroes} 个英雄。"
        )

        print(
            "已扫完的部分会照常保存，"
            "建议之后完整重跑一次。"
        )

    else:

        print(
            "扫描完成，开始更新 Excel……"
        )

    print()


    # ========================================================
    # 更新 Excel
    # ========================================================

    # 和其他已有数据格子一样的填充色
    cell_fill = PatternFill(
        fill_type="solid",
        fgColor="E2F0D9"
    )

    new_heroes = 0

    # 原来 没有 → 这次有：新增
    new_records = 0

    # 原来 有 → 这次时间变了：更新
    updated_records = 0


    for hero_id in sorted(
        all_results
    ):

        results = all_results[
            hero_id
        ]


        # ----------------------------------------------------
        # 判断英雄是否已经存在
        # ----------------------------------------------------

        if hero_id in hero_rows:

            row = hero_rows[
                hero_id
            ]

        else:

            # 新英雄 → 新增一行

            row = ws.max_row + 1

            ws.cell(
                row=row,
                column=1,
                value=hero_id
            )

            # 英雄名字故意留空
            # 给你自己填写

            ws.cell(
                row=row,
                column=2,
                value=""
            )

            hero_rows[
                hero_id
            ] = row

            new_heroes += 1


        # ----------------------------------------------------
        # 更新每一个实际存在的皮肤
        # ----------------------------------------------------

        for result in results:

            skin_id = result[
                "skin_id"
            ]

            last_modified = result[
                "last_modified"
            ]


            name_col = skin_name_columns[
                skin_id
            ]

            time_col = last_modified_columns[
                skin_id
            ]


            # ------------------------------------------------
            # 皮肤名：
            #
            # 绝对不覆盖已有内容
            # ------------------------------------------------

            skin_name = ws.cell(
                row=row,
                column=name_col
            ).value


            # 如果原来就是空的，就保持空
            #
            # 如果已经填写，就完全不碰
            #
            # 所以这里什么都不用做

            old_time = ws.cell(
                row=row,
                column=time_col
            ).value

            # ============================================================
            # Last-Modified：
            # 每次都以服务器最新返回值为准
            # ============================================================

            if last_modified:

                if (
                        old_time is None
                        or str(old_time).strip() != str(last_modified).strip()
                ):
                    time_cell = ws.cell(
                        row=row,
                        column=time_col,
                        value=last_modified
                    )

                    name_cell = ws.cell(
                        row=row,
                        column=name_col
                    )

                    # 出现新的 Last-Modified：
                    # 给 编号-Last-Modified 和 编号-皮肤名 两个格子
                    # 涂上和其他格子一样的颜色
                    time_cell.fill = cell_fill
                    name_cell.fill = cell_fill

                    # 区分「新增」和「更新」两种情况

                    if old_time is None:

                        new_records += 1

                    else:

                        updated_records += 1


    # ========================================================
    # 样式
    # ========================================================

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="D9EAF7"
    )

    header_font = Font(
        bold=True
    )


    for cell in ws[1]:

        cell.fill = header_fill

        cell.font = header_font

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )


    # ========================================================
    # 设置列宽
    # ========================================================

    ws.column_dimensions[
        "A"
    ].width = 12

    ws.column_dimensions[
        "B"
    ].width = 15


    for skin_id in range(
        SKIN_START,
        SKIN_END + 1
    ):

        name_col = skin_name_columns[
            skin_id
        ]

        time_col = last_modified_columns[
            skin_id
        ]


        ws.column_dimensions[
            get_column_letter(
                name_col
            )
        ].width = 18


        ws.column_dimensions[
            get_column_letter(
                time_col
            )
        ].width = 22


    # ========================================================
    # 冻结
    # ========================================================

    ws.freeze_panes = "C2"


    # ========================================================
    # 保存
    # ========================================================

    wb.save(
        EXCEL_FILE
    )


    # ========================================================
    # 统计
    # ========================================================

    print("=" * 70)

    print(
        "更新完成！"
    )

    print("=" * 70)

    print()

    print(
        f"新增英雄："
        f"{new_heroes}"
    )

    print(
        f"新增 Last-Modified："
        f"{new_records}"
    )

    print(
        f"更新 Last-Modified："
        f"{updated_records}"
    )

    print()

    if total_failed or scan_errors:

        print(
            f"注意：{total_failed} 个请求、"
            f"{scan_errors} 个英雄任务失败，"
            f"这些格子本次没有扫到。"
        )

        print(
            "建议再完整跑一次补齐。"
        )

        print()

    print(
        f"Excel：{EXCEL_FILE}"
    )

    print()

    print(
        "已有皮肤名不会被修改。"
    )

    print(
        "已有 Last-Modified 以服务器最新值为准，"
        "有更新会自动刷新并高亮。"
    )

    print()


if __name__ == "__main__":

    main()