# -*- coding: utf-8 -*-
"""
hero_scan.py / skin_scan.py 共用的常量和小工具。

本文件必须和两个扫描脚本放在同一目录
（Python 会把脚本所在目录加进 sys.path，直接 import common 即可）。
"""

import os


# ============================================================
# 路径
# ============================================================

# Excel 锚定到脚本所在目录：
# 从任何工作目录运行都能找到同一个文件，
# 不会在别处凭空新建一个空表
APP_DIR = os.path.dirname(os.path.abspath(__file__))

EXCEL_FILE = os.path.join(APP_DIR, "hero_skin_scan.xlsx")


# ============================================================
# 资源地址
# ============================================================

BASE_URL = "https://image.smoba.qq.com/Picture/HeroOriginalPainting/"

# 皮肤格子：00 ~ 19
SKIN_START = 0
SKIN_END = 19


def skin_image_url(hero_id, skin_id):

    """英雄编号 + 皮肤格子 → 原画 jpg 地址"""

    filename = (
        f"30{hero_id:03d}"
        f"{skin_id:02d}"
        ".jpg"
    )

    return BASE_URL + filename


# ============================================================
# 表头解析
# ============================================================

def parse_header_columns(ws):

    """
    扫描第一行表头，返回 (skin_name_columns, last_modified_columns)。

    两个都是 {皮肤格子: 列号}，
    表头格式："{skin_id:02d}-皮肤名" / "{skin_id:02d}-Last-Modified"
    """

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

    return skin_name_columns, last_modified_columns
