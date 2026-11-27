# -*- coding: utf-8 -*-
"""
hero_scan.py / skin_scan.py 共用的常量和小工具。

本文件必须和两个扫描脚本放在同一目录
（Python 会把脚本所在目录加进 sys.path，直接 import common 即可）。
"""

import json
import os
import urllib.request


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
HERO_LIST_URL = "https://pvp.qq.com/web201605/js/herolist.json"

# 默认皮肤格子：00 ~ 19（扫描器会在临近上限时自动向后续探）
SKIN_START = 0
SKIN_END = 19


def skin_image_url(hero_id, skin_id):

    """英雄编号 + 皮肤格子 → 原画 jpg 地址"""

    return f"{BASE_URL}30{hero_id:03d}{skin_id:02d}.jpg"


# ============================================================
# 结束前暂停
# ============================================================

def pause(prompt="按回车退出..."):

    """
    脚本结束前的暂停。

    被 GUI 当子进程运行时（环境变量 SCAN_GUI=1）
    直接跳过——不用再点一次回车；
    直接双击运行时照常停住，避免窗口一闪而过。
    """

    if os.environ.get("SCAN_GUI") == "1":
        return

    try:
        input(prompt)
    except (EOFError, KeyboardInterrupt):
        pass


# ============================================================
# 表头解析（同时支持只读行元组与常规 Worksheet）
# ============================================================

def parse_header_row(header_row):

    """
    从第一行表头序列中解析 (skin_name_idx, last_modified_idx)。

    两个都是 {皮肤格子: 0-based 列索引}，
    表头格式："{skin_id:02d}-皮肤名" / "{skin_id:02d}-Last-Modified"
    """

    skin_name_idx = {}
    last_modified_idx = {}

    for idx, val in enumerate(header_row):
        if not val:
            continue

        header = str(val)

        if header.endswith("-皮肤名"):
            try:
                skin_name_idx[int(header.split("-")[0])] = idx
            except ValueError:
                pass
        elif header.endswith("-Last-Modified"):
            try:
                last_modified_idx[int(header.split("-")[0])] = idx
            except ValueError:
                pass

    return skin_name_idx, last_modified_idx


def parse_header_columns(ws):

    """
    扫描 Worksheet 第一行表头，返回 (skin_name_columns, last_modified_columns)。

    两个都是 {皮肤格子: 1-based 列号}。
    """

    values = [
        ws.cell(row=1, column=col).value
        for col in range(1, ws.max_column + 1)
    ]
    sn_idx, lm_idx = parse_header_row(values)

    return (
        {sid: idx + 1 for sid, idx in sn_idx.items()},
        {sid: idx + 1 for sid, idx in lm_idx.items()},
    )


# ============================================================
# 皮肤名归一化 & 官网英雄/皮肤列表获取（标准库实现，无额外依赖）
# ============================================================

def norm_skin_name(name):

    """归一化皮肤名用于差集比对（忽略不同中点符号与空格差异）"""

    s = str(name).strip()
    for ch in ("·", "●", "•", "・", "-", " ", "\u3000"):
        s = s.replace(ch, "")
    return s


def fetch_herolist_data(timeout=8):

    """拉取官网 herolist.json 并自动处理 UTF-8/GBK 编码，失败时抛出异常"""

    req = urllib.request.Request(
        HERO_LIST_URL,
        headers={"User-Agent": "Mozilla/5.0"},
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        content = resp.read()

    try:
        return json.loads(content.decode("utf-8-sig"))
    except UnicodeDecodeError:
        return json.loads(content.decode("gbk"))


def fetch_official_hero_names(timeout=10, warn_on_error=True):

    """
    拉官网英雄列表，返回 {英雄编号: 官方英雄名}。
    任何失败均返回空 dict，不影响扫描主流程。
    """

    try:
        data = fetch_herolist_data(timeout=timeout)
        names = {}
        for item in data:
            try:
                names[int(item["ename"])] = str(item["cname"]).strip()
            except (KeyError, TypeError, ValueError):
                continue
        return names
    except Exception as exc:
        if warn_on_error:
            print(
                f"官网英雄列表获取失败，"
                f"本次跳过英雄名自动填充：{exc}"
            )
        return {}


def fetch_official_skin_names(timeout=8):

    """
    拉官网 herolist.json 的皮肤名列表，返回 {英雄编号: [皮肤名, ...]}。
    仅在存在未命名皮肤时作为人工填名参考，失败返回空 dict。
    """

    try:
        data = fetch_herolist_data(timeout=timeout)
        result = {}
        for item in data:
            try:
                hid = int(item["ename"])
                raw = str(item.get("skin_name") or "").strip()
                if not raw:
                    continue
                names = [s.strip() for s in raw.split("|") if s.strip()]
                if names:
                    result[hid] = names
            except (KeyError, TypeError, ValueError):
                continue
        return result
    except Exception:
        return {}
