# -*- coding: utf-8 -*-
"""
纯白图（白底图）识别 · Excel 批量检测
================================================================
读取 Excel 中存放图片 URL 的列，用像素分析核心（white_detect）逐张判定
是否为纯白图 / 纯色空图，把结果写回一份新的 Excel。

特性：
  - 自动识别 URL 列（含「图片 / 图 / url / 链接 / 主图 / 详情图」等关键字）
  - 单格支持多个 URL（逗号 / 换行 / 空格分隔）
  - 多线程并发检测（--workers）
  - 断点续跑（--resume，进度存 sidecar .progress.json）
  - 行级 + 全局汇总（是否含纯白图、纯白数、明细、统计）

依赖：openpyxl、numpy、Pillow、requests
用法：
  python batch_excel.py INPUT.xlsx
  python batch_excel.py INPUT.xlsx --output result.xlsx --cols 图片地址,详情图地址
  python batch_excel.py INPUT.xlsx --limit 5          # 先试前 5 行
  python batch_excel.py INPUT.xlsx --list-cols        # 只看识别到哪些列
================================================================
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

# ── 引入识别核心 white_detect ──
try:
    from white_detect import detect_from_url
except ImportError:
    _HERE = os.path.dirname(os.path.abspath(__file__))
    for _p in (
        _HERE,
        os.path.join(_HERE, "scripts"),
        r"C:\Users\Lenovo\WorkBuddy\white-image-detect",
    ):
        if os.path.isdir(_p) and os.path.exists(os.path.join(_p, "white_detect.py")):
            sys.path.insert(0, _p)
            try:
                from white_detect import detect_from_url  # type: ignore
                break
            except ImportError:
                continue

import openpyxl

# URL 列自动识别关键字
URL_HINTS = ("图片", "图", "url", "链接", "link", "image", "主图", "详情图", "pic", "img")
# 单格分隔符（逗号 / 换行 / 全角逗号 / 分号 / 空白）
_SPLIT_CHARS = [",", "\n", "，", ";", "；", "\t", " "]


def split_urls(cell) -> list[str]:
    """把单元格内容拆成多个 URL。"""
    if cell is None:
        return []
    text = str(cell).strip()
    if not text:
        return []
    for ch in _SPLIT_CHARS:
        if ch in text:
            text = text.replace(ch, "\n")
    out = []
    for u in text.split("\n"):
        u = u.strip()
        if u and u.lower().startswith(("http://", "https://")):
            out.append(u)
    # 去重保序
    seen, dedup = set(), []
    for u in out:
        if u not in seen:
            seen.add(u)
            dedup.append(u)
    return dedup


def detect_col_name(header: list) -> list[str]:
    """扫描表头，返回识别出的 URL 列名列表。"""
    cols = []
    for name in header:
        if name is None:
            continue
        s = str(name).lower()
        if any(h in s for h in URL_HINTS):
            cols.append(str(name))
    return cols


def short_label(res: dict) -> str:
    if not res.get("success"):
        return "失败"
    return "纯白图" if res.get("is_white_image") else "正常图"


def main():
    ap = argparse.ArgumentParser(description="纯白图识别 · Excel 批量检测")
    ap.add_argument("input", help="输入 Excel 路径")
    ap.add_argument("--output", default=None, help="输出 Excel 路径（默认 INPUT_result.xlsx）")
    ap.add_argument("--cols", default=None, help="指定 URL 列名，逗号分隔；不填则自动识别")
    ap.add_argument("--sheet", default=None, help="工作表名或索引（默认第一个）")
    ap.add_argument("--workers", type=int, default=8, help="并发线程数（默认 8）")
    ap.add_argument("--threshold", type=float, default=0.99, help="近白像素占比阈值（默认 0.99）")
    ap.add_argument("--max-colors", type=int, default=3, help="颜色种类上限（默认 3）")
    ap.add_argument("--limit", type=int, default=0, help="只处理前 N 行数据（测试用）")
    ap.add_argument("--resume", action="store_true", help="断点续跑（读取已有进度）")
    ap.add_argument("--list-cols", action="store_true", help="只打印识别到的 URL 列并退出")
    args = ap.parse_args()

    if not os.path.exists(args.input):
        print(f"[错误] 输入文件不存在: {args.input}")
        sys.exit(1)

    # ── 读源表 ──
    wb = openpyxl.load_workbook(args.input, read_only=True, data_only=True)
    ws = wb[args.sheet] if args.sheet else wb[wb.sheetnames[0]]
    all_rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not all_rows:
        print("[错误] 表格为空")
        sys.exit(1)
    header = [str(c) if c is not None else "" for c in all_rows[0]]
    data_rows = all_rows[1:]

    # ── 选定 URL 列 ──
    if args.cols:
        selected = [c.strip() for c in args.cols.split(",") if c.strip()]
        missing = [c for c in selected if c not in header]
        if missing:
            print(f"[错误] 指定列不存在: {missing}；实际表头: {header}")
            sys.exit(1)
    else:
        selected = detect_col_name(header)
        if not selected:
            print(f"[错误] 未自动识别到 URL 列。表头: {header}\n可用 --cols 手动指定。")
            sys.exit(1)

    if args.list_cols:
        print("识别到的 URL 列：")
        for c in selected:
            print("  -", c)
        print(f"\n共 {len(data_rows)} 行数据。")
        sys.exit(0)

    if args.limit > 0:
        data_rows = data_rows[: args.limit]

    out_path = args.output or (os.path.splitext(args.input)[0] + "_result.xlsx")
    prog_path = out_path + ".progress.json"

    # ── 续跑加载 ──
    results: dict[int, dict] = {}
    if args.resume and os.path.exists(prog_path):
        try:
            with open(prog_path, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            results = {int(k): v for k, v in loaded.get("results", {}).items()}
            print(f"[续跑] 已加载 {len(results)} 行进度")
        except Exception as e:
            print(f"[续跑] 进度读取失败，重新开始: {e}")

    # 待处理行（源行号从 2 开始）
    todo = [i for i in range(2, len(data_rows) + 2) if i not in results]

    # 收集所有检测任务
    tasks = []  # (src_row_idx, col_name, url)
    for r_idx in todo:
        src = data_rows[r_idx - 2]
        for col in selected:
            ci = header.index(col)
            cell = src[ci] if ci < len(src) else None
            for url in split_urls(cell):
                tasks.append((r_idx, col, url))

    print(f"[信息] 输入: {args.input}")
    print(f"[信息] URL 列: {selected}")
    print(f"[信息] 待处理数据行: {len(todo)}，检测任务(URL)总数: {len(tasks)}")
    print(f"[信息] 并发: {args.workers} | 阈值: {args.threshold} | 颜色上限: {args.max_colors}")

    lock = threading.Lock()
    t0 = time.time()

    def persist():
        tmp = prog_path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump({"selected_cols": selected, "results": results}, f,
                      ensure_ascii=False)
        os.replace(tmp, prog_path)

    if tasks:
        with ThreadPoolExecutor(max_workers=args.workers) as ex:
            fut_map = {}
            for r_idx, col, url in tasks:
                f = ex.submit(detect_from_url, url,
                              white_threshold=args.threshold,
                              max_color_variety=args.max_colors)
                fut_map[f] = (r_idx, col, url)
            done = 0
            for f in as_completed(fut_map):
                r_idx, col, url = fut_map[f]
                res = f.result()
                with lock:
                    row = results.setdefault(r_idx, {"cols": {}})
                    colres = row["cols"].setdefault(col, {"per": []})
                    colres["per"].append({
                        "url": url,
                        "success": res.get("success"),
                        "is_white_image": res.get("is_white_image"),
                        "category": res.get("category"),
                        "white_ratio": res.get("white_ratio"),
                        "reason": res.get("reason"),
                        "error": res.get("error"),
                    })
                    done += 1
                    if done % 50 == 0:
                        print(f"  ... 已完成 {done}/{len(tasks)} 个 URL 检测")
                        persist()
            persist()
    else:
        print("[信息] 无待处理任务（可能已全部完成）")

    # ── 汇总每个源行的判定 ──
    for r_idx, row in results.items():
        total = 0
        white_total = 0
        has_white = False
        for col, colres in row.get("cols", {}).items():
            per = colres["per"]
            wc = sum(1 for p in per if p.get("is_white_image"))
            colres["white_count"] = wc
            colres["has_white"] = wc > 0
            colres["detail"] = "、".join(short_label(p) for p in per) or "无图"
            total += len(per)
            white_total += wc
            if wc > 0:
                has_white = True
        row["overall"] = {
            "total": total, "white_total": white_total, "has_white": has_white,
        }

    # ── 写出 Excel ──
    out_appended = []
    for col in selected:
        out_appended += [f"{col}·判定", f"{col}·纯白数", f"{col}·明细"]
    out_appended += ["是否含纯白图", "纯白图总数", "图片总数"]

    out_wb = openpyxl.Workbook()
    out_ws = out_wb.active
    out_ws.title = "检测结果"
    out_ws.append(list(header) + out_appended)

    white_rows = 0
    white_imgs = 0
    normal_imgs = 0
    fail_imgs = 0

    for i, src in enumerate(data_rows):
        r_idx = i + 2
        row_res = results.get(r_idx, {})
        vals = list(src)
        # 补齐长度
        while len(vals) < len(header):
            vals.append(None)
        for col in selected:
            colres = row_res.get("cols", {}).get(col, {})
            vals.append("是" if colres.get("has_white") else ("否" if "per" in colres else "—"))
            vals.append(colres.get("white_count", "—"))
            vals.append(colres.get("detail", "—"))
        ov = row_res.get("overall", {})
        vals.append("是" if ov.get("has_white") else ("否" if ov else "—"))
        vals.append(ov.get("white_total", "—"))
        vals.append(ov.get("total", "—"))
        out_ws.append(vals)

        if ov:
            if ov.get("has_white"):
                white_rows += 1
            white_imgs += ov.get("white_total", 0)
            for colres in row_res.get("cols", {}).values():
                for p in colres.get("per", []):
                    if not p.get("success"):
                        fail_imgs += 1
                    elif not p.get("is_white_image"):
                        normal_imgs += 1

    # 汇总 sheet
    sum_ws = out_wb.create_sheet("汇总")
    elapsed = round(time.time() - t0, 1)
    rows_done = len(results)
    summary = [
        ("纯白图识别 · 批量检测汇总", ""),
        ("输入文件", args.input),
        ("输出文件", out_path),
        ("URL 列", ", ".join(selected)),
        ("阈值(white_threshold)", args.threshold),
        ("颜色上限(max_color_variety)", args.max_colors),
        ("数据总行数", len(data_rows)),
        ("已处理行数", rows_done),
        ("含纯白图的行数", white_rows),
        ("纯白图数量(张)", white_imgs),
        ("正常图数量(张)", normal_imgs),
        ("检测失败数量(张)", fail_imgs),
        ("耗时(秒)", elapsed),
        ("生成时间", time.strftime("%Y-%m-%d %H:%M:%S")),
    ]
    for k, v in summary:
        sum_ws.append([k, v])
    sum_ws.column_dimensions["A"].width = 30
    sum_ws.column_dimensions["B"].width = 60

    out_wb.save(out_path)
    print(f"\n[完成] 结果已写出: {out_path}")
    print(f"[完成] 已处理 {rows_done} 行 | 含纯白图 {white_rows} 行 | "
          f"纯白图 {white_imgs} 张 | 正常图 {normal_imgs} 张 | 失败 {fail_imgs} 张")
    print(f"[完成] 耗时 {elapsed}s")


if __name__ == "__main__":
    main()
