import requests
import concurrent.futures
import threading
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
THREADS = 30

# 请求超时时间
TIMEOUT = 10


# ============================================================
# Session
# ============================================================

thread_local = threading.local()


def get_session():

    if not hasattr(thread_local, "session"):

        session = requests.Session()

        session.headers.update({
            "User-Agent": "Mozilla/5.0"
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

    try:

        session = get_session()

        response = session.head(
            url,
            timeout=TIMEOUT,
            allow_redirects=True
        )

        if response.status_code != 200:

            return {
                "hero_id": hero_id,
                "skin_id": skin_id,
                "exists": False
            }

        last_modified = response.headers.get(
            "Last-Modified",
            ""
        )

        if last_modified:
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


    except requests.RequestException:

        return {
            "hero_id": hero_id,
            "skin_id": skin_id,
            "exists": False
        }


# ============================================================
# 扫描一个英雄
# ============================================================

def scan_hero(hero_id):

    results = []

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

    return hero_id, results


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


    with concurrent.futures.ThreadPoolExecutor(
        max_workers=THREADS
    ) as executor:

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


        for future in concurrent.futures.as_completed(
            futures
        ):

            hero_id, results = future.result()

            completed += 1


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


    print()
    print()

    print(
        "扫描完成，开始更新 Excel……"
    )

    print()


    # ========================================================
    # 更新 Excel
    # ========================================================

    new_heroes = 0
    new_resources = 0


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
                    ws.cell(
                        row=row,
                        column=time_col,
                        value=last_modified
                    )

                    new_resources += 1


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
        f"新增资源时间记录："
        f"{new_resources}"
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
        "已有 Last-Modified 不会被覆盖。"
    )

    print()


if __name__ == "__main__":

    main()