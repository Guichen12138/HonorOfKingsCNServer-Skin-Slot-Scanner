# -*- coding: utf-8 -*-
import json
import os
import urllib.request
import webbrowser

import openpyxl

# 共用常量与工具：Excel 路径 / 槽位范围 / 表头解析 / URL 拼装
# （common.py 和本脚本在同一目录，Python 会自动找到）
from common import (
    EXCEL_FILE,
    SKIN_END,
    SKIN_START,
    pause,
    skin_image_url,
)

# 海报增量保存是本机自用功能（poster_save.py 不进仓库）：
# 仓库里没有这个文件时自动跳过保存，其余功能照常
try:
    from poster_save import save_poster
except ImportError:
    def save_poster(url):
        pass


HERO_LIST_URL = "https://pvp.qq.com/web201605/js/herolist.json"


def _norm_skin_name(name):
    """归一化皮肤名用于差集过滤（忽略不同中点符号和空格差异）"""
    s = str(name).strip()
    for ch in ("·", "●", "•", "・", "-", " ", "\u3000"):
        s = s.replace(ch, "")
    return s


def _fetch_official_skin_names():
    """
    可选拉取官网 herolist.json 的皮肤名列表，返回 {hero_id: [皮肤名, ...]}。
    仅在存在未命名皮肤时调用，作为人工填名参考（绝不自动写入 Excel）。
    """
    try:
        req = urllib.request.Request(
            HERO_LIST_URL,
            headers={"User-Agent": "Mozilla/5.0"},
        )
        with urllib.request.urlopen(req, timeout=6) as resp:
            content = resp.read()

        try:
            data = json.loads(content.decode("utf-8-sig"))
        except UnicodeDecodeError:
            data = json.loads(content.decode("gbk"))

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


def _parse_header_row(header_row):
    """从只读模式的第一行表头元组中解析 {skin_id: 0-based 索引}"""
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


def main():

    # ============================================================
    # 检查 Excel
    # ============================================================

    if not os.path.exists(EXCEL_FILE):
        print(f"找不到 Excel 文件：{EXCEL_FILE}")
        pause()
        return

    # ============================================================
    # 只读流式读取 Excel，并在进入交互前立即释放文件句柄
    # ============================================================

    wb = openpyxl.load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=True,
    )

    heroes = []
    total_skin_count = 0
    total_hero_count = 0

    try:
        ws = wb.active
        row_iter = ws.iter_rows(values_only=True)
        header_row = next(row_iter, None)

        if not header_row:
            print("Excel 表头为空。")
            pause()
            return

        skin_name_idx, last_modified_idx = _parse_header_row(header_row)

        # 英雄数量按「原皮（00 槽）Last-Modified 有没有值」来算：
        # 空号英雄在表里也占一行，但那只是占位，不能算英雄
        first_lm_i = last_modified_idx.get(SKIN_START)

        for row_num, row_vals in enumerate(row_iter, start=2):
            if not row_vals:
                continue

            hero_id = row_vals[0] if len(row_vals) > 0 else None
            hero_name = row_vals[1] if len(row_vals) > 1 else None

            if hero_id is None:
                continue

            try:
                hero_id = int(hero_id)
            except (TypeError, ValueError):
                continue

            # 原皮列是空的 = 空号占位行：不算英雄，也不用再扫它的槽位
            if first_lm_i is not None:
                first_lm = row_vals[first_lm_i] if first_lm_i < len(row_vals) else None
                if not first_lm:
                    continue

            total_hero_count += 1

            skins_to_open = []
            existing_names_norm = set()
            slot0_named = False

            for skin_id in range(SKIN_START, SKIN_END + 1):
                lm_i = last_modified_idx.get(skin_id)
                nm_i = skin_name_idx.get(skin_id)

                if lm_i is None or nm_i is None:
                    continue

                last_modified = row_vals[lm_i] if lm_i < len(row_vals) else None
                if not last_modified:
                    continue

                total_skin_count += 1

                skin_name = row_vals[nm_i] if nm_i < len(row_vals) else None
                if skin_name is not None and str(skin_name).strip() != "":
                    existing_names_norm.add(_norm_skin_name(skin_name))
                    if skin_id == SKIN_START:
                        slot0_named = True
                    continue

                skins_to_open.append({
                    "skin_id": skin_id,
                    "url": skin_image_url(hero_id, skin_id),
                    "last_modified": last_modified,
                })

            if skins_to_open:
                heroes.append({
                    "row": row_num,
                    "hero_id": hero_id,
                    "hero_name": hero_name if hero_name else "未填写英雄名",
                    "skins": skins_to_open,
                    "existing_names_norm": existing_names_norm,
                    "slot0_named": slot0_named,
                })
    finally:
        wb.close()

    # ============================================================
    # 按英雄编号排序
    # ============================================================

    heroes.sort(key=lambda x: x["hero_id"])

    # ============================================================
    # 显示统计
    # ============================================================

    print("=" * 70)
    print(
        f"王者荣耀英雄原画查看器："
        f"Excel 共 {total_hero_count} 个英雄、"
        f"{total_skin_count} 张海报，"
        f"{len(heroes)} 个英雄有待处理"
    )
    print("=" * 70)

    # 没有需要处理的英雄：直接结束（不发任何网络请求）
    if not heroes:
        print("没有需要处理的英雄，所有海报都已填写皮肤名。")
        pause()
        return

    # 有待命名皮肤时，顺手拉一次官网皮肤名列表供人工参考
    official_skins = _fetch_official_skin_names()

    # ============================================================
    # 一个英雄一个英雄处理
    # ============================================================

    interrupted = False

    try:
        for index, hero in enumerate(heroes):

            hero_id = hero["hero_id"]
            hero_name = hero["hero_name"]
            skins = hero["skins"]
            existing_names_norm = hero["existing_names_norm"]

            print(
                f"[{index + 1}/{len(heroes)}] "
                f"{hero_id:03d} {hero_name}  "
                "待查看格子："
                + " ".join(f"{skin['skin_id']:02d}" for skin in skins)
            )

            # 计算官网皮肤名与表中已有皮肤名的差集作为人工参考（不自动填表）
            official_list = official_skins.get(hero_id)
            if official_list:
                # 官网 skin_name 第 0 项是英雄称号（对应 00 原皮槽），
                # 若表里 00 槽已命名（常填为英雄名或原皮），则跳过第 0 项减少误报
                pool = official_list[1:] if hero["slot0_named"] else official_list
                candidates = [
                    name for name in pool
                    if _norm_skin_name(name) not in existing_names_norm
                ]
                if candidates:
                    print("官网候选皮肤名参考：" + " / ".join(candidates))

            for skin in skins:
                print(f"{skin['skin_id']:02d} → {skin['url']}")

                # 打开浏览器前把这张海报的全画幅原图存进海报文件夹
                save_poster(skin["url"])

                webbrowser.open_new_tab(skin["url"])

            print(f"{hero_name} 的 {len(skins)} 张未处理海报已打开。")

            if index < len(heroes) - 1:
                print(
                    "在 Excel 里填好皮肤名，"
                    "然后按回车（GUI 里点「继续」）进入下一个英雄。"
                )
                input()

    except KeyboardInterrupt:
        interrupted = True
        print("已手动中断，已填写的内容保留在 Excel 里。")

    # ============================================================
    # 完成
    # ============================================================

    if interrupted:
        print("本次处理被中断，下次运行会从还没填写的格子继续。")
    else:
        print(f"全部 {len(heroes)} 个英雄的海报已查看完成。")

    pause()


if __name__ == "__main__":
    main()
