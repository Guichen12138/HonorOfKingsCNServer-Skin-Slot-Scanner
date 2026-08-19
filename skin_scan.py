import openpyxl
import webbrowser
import os


# ============================================================
# 配置
# ============================================================

EXCEL_FILE = "hero_skin_scan.xlsx"

BASE_URL = "https://image.smoba.qq.com/Picture/HeroOriginalPainting/"

# 皮肤格子：00 ~ 19
SKIN_START = 0
SKIN_END = 19


# ============================================================
# 检查 Excel
# ============================================================

if not os.path.exists(EXCEL_FILE):

    print(f"找不到 Excel 文件：{EXCEL_FILE}")

    input("按回车退出...")
    exit()


# ============================================================
# 读取 Excel
# ============================================================

wb = openpyxl.load_workbook(
    EXCEL_FILE,
    data_only=True
)

ws = wb.active


# ============================================================
# 找到各个 Last-Modified 和皮肤名列
# ============================================================

last_modified_columns = {}
skin_name_columns = {}

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


    # --------------------------------------------------------
    # Last-Modified
    # --------------------------------------------------------

    if header.endswith("-Last-Modified"):

        try:

            skin_id = int(
                header.split("-")[0]
            )

            last_modified_columns[
                skin_id
            ] = col

        except ValueError:
            pass


    # --------------------------------------------------------
    # 皮肤名
    # --------------------------------------------------------

    elif header.endswith("-皮肤名"):

        try:

            skin_id = int(
                header.split("-")[0]
            )

            skin_name_columns[
                skin_id
            ] = col

        except ValueError:
            pass


# ============================================================
# 读取英雄
# ============================================================

heroes = []


for row in range(
    2,
    ws.max_row + 1
):

    hero_id = ws.cell(
        row=row,
        column=1
    ).value

    hero_name = ws.cell(
        row=row,
        column=2
    ).value


    if hero_id is None:
        continue


    try:

        hero_id = int(
            hero_id
        )

    except ValueError:

        continue


    # ========================================================
    # 检查这个英雄有哪些实际存在的资源
    # ========================================================

    skins_to_open = []


    for skin_id in range(
        SKIN_START,
        SKIN_END + 1
    ):

        last_modified_col = (
            last_modified_columns.get(
                skin_id
            )
        )

        skin_name_col = (
            skin_name_columns.get(
                skin_id
            )
        )


        if (
            last_modified_col is None
            or skin_name_col is None
        ):
            continue


        # ----------------------------------------------------
        # 服务器资源是否存在
        # ----------------------------------------------------

        last_modified = ws.cell(
            row=row,
            column=last_modified_col
        ).value


        if not last_modified:
            continue


        # ----------------------------------------------------
        # 你是否已经填写过皮肤名字
        # ----------------------------------------------------

        skin_name = ws.cell(
            row=row,
            column=skin_name_col
        ).value


        # 有名字 → 已经处理过 → 跳过
        if (
            skin_name is not None
            and str(skin_name).strip() != ""
        ):
            continue


        # ----------------------------------------------------
        # 有资源 + 没填写名字
        # → 需要打开
        # ----------------------------------------------------

        filename = (
            f"30{hero_id:03d}"
            f"{skin_id:02d}"
            ".jpg"
        )

        url = BASE_URL + filename


        skins_to_open.append({
            "skin_id": skin_id,
            "url": url,
            "last_modified": last_modified
        })


    # ========================================================
    # 只有存在“未处理资源”的英雄才加入列表
    # ========================================================

    if skins_to_open:

        heroes.append({
            "row": row,
            "hero_id": hero_id,
            "hero_name": (
                hero_name
                if hero_name
                else "未填写英雄名"
            ),
            "skins": skins_to_open
        })


# ============================================================
# 按英雄编号排序
# ============================================================

heroes.sort(
    key=lambda x: x["hero_id"]
)


# ============================================================
# 显示统计
# ============================================================

print("=" * 70)

print("王者荣耀英雄原画查看器")

print("=" * 70)

print()

print(
    f"Excel 中共有 "
    f"{ws.max_row - 1} 个英雄"
)

print(
    f"还有 {len(heroes)} 个英雄存在未处理的海报"
)

print()

print(
    "每次按一次回车："
    "打开当前英雄尚未填写皮肤名的海报。"
)

print(
    "你处理完当前英雄后，再按一次回车进入下一个英雄。"
)

print()

print("=" * 70)

input(
    "按回车开始..."
)


# ============================================================
# 一个英雄一个英雄处理
# ============================================================

for index, hero in enumerate(
    heroes
):

    hero_id = hero["hero_id"]

    hero_name = hero["hero_name"]

    skins = hero["skins"]


    print()

    print("-" * 70)

    print(
        f"[{index + 1}/{len(heroes)}] "
        f"{hero_id:03d} {hero_name}"
    )

    print()

    print(
        "本次需要查看的格子："
        + " ".join(
            f"{skin['skin_id']:02d}"
            for skin in skins
        )
    )

    print()

    input(
        "按回车打开这些海报..."
    )


    # ========================================================
    # 打开当前英雄尚未处理的海报
    # ========================================================

    print()

    print(
        f"正在打开 "
        f"{hero_id:03d} {hero_name} ..."
    )

    print()


    for skin in skins:

        print(
            f"{skin['skin_id']:02d} → "
            f"{skin['url']}"
        )

        webbrowser.open_new_tab(
            skin["url"]
        )


    print()

    print(
        f"{hero_id:03d} {hero_name} "
        f"的未处理海报已全部打开。"
    )


    # ========================================================
    # 等待用户处理完当前英雄
    # ========================================================

    if index < len(heroes) - 1:

        input(
            "\n处理完当前英雄后，"
            "按回车打开下一个英雄..."
        )


# ============================================================
# 完成
# ============================================================

print()

print("=" * 70)

print(
    "所有未处理的英雄海报已经查看完成。"
)

print("=" * 70)

input(
    "按回车退出..."
)