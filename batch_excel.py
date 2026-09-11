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
  - 纯像素分析，不依赖 AI 模型 / 不调用 Dify / 不消耗 token

依赖：openpyxl、numpy、Pillow、requests
用法：
  python batch_excel.py INPUT.xlsx
  python batch_excel.py INPUT.xlsx --output result.xlsx --cols 图片地址,详情图地址
  python batch_excel.py INPUT.xlsx --limit 5          # 先试前 5 行
  python batch_excel.py INPUT.xlsx --list-cols        # 只看识别到哪些列

编程接口：run_batch(input, output=None, cols=None, ...) -> summary dict
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
        u = u.strip().strip('"').strip("'")
        if u and _looks_like_url(u):
            out.append(u)
    # 去重保序
    seen, dedup = set(), []
    for u in out:
        if u not in seen:
            seen.add(u)
            dedup.append(u)
    return dedup


def _looks_like_url(u: str) -> bool:
    """判断是否为可访问的图片 URL（http/https/file）。"""
    ul = u.lower()
    return ul.startswith(("http://", "https://", "file:///", "file:\\"))


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


def run_batch(input, output=None, cols=None, sheet=None, workers=8,
              threshold=0.99, max_colors=3, limit=0, resume=False,
              rules=("any", "all"), rule=None, progress_cb=None,
              offset=0, persist_every=300, cancel_check=None) -> dict:
    """
    批量检测 Excel 中的图片 URL，结果写回 output Excel。

    参数：
      input       输入 Excel 路径
      output      输出 Excel 路径（默认 INPUT_空白图结果.xlsx）
      cols        指定 URL 列名（逗号分隔）；None 则自动识别
      sheet       工作表名/索引；None 取第一个
      workers     并发线程数
      threshold   近白像素占比阈值
      max_colors  颜色种类上限
      limit       只处理前 N 行（0=全部）
      offset      跳过前 N 行数据（分批跑第 i 批时传 offset=i*chunk_size）
      resume      断点续跑（读取已有 .progress.json）
      persist_every 运行中每完成 N 个 URL 就落盘一次进度（默认 300；0=只在结束时落盘）
                  另：距上次落盘超过 60 秒也会落盘。用于长任务防丢进度。
      cancel_check 可选回调，返回 True 表示要求停止；每完成一个任务检查一次。
                  取消时：未开始的任务被取消，进度落盘，返回 summary 带 cancelled=True
                  （不写完整 Excel，可改用 export_from_progress 导出已完成部分）
      rules       要输出的判定维度集合，如 ("any","all") / ("any",) / ("all",)
                  默认 ("any","all") —— 一次扫描同时输出两种维度
      rule        【向后兼容】旧版单值规则（"any" / "all"），等价于 rules=(rule,)
                  若同时给了 rules 和 rule，以 rules 为准
      progress_cb 回调(done, total, msg)，用于 GUI/日志刷新

    判定语义：
      any  — 任意一张图片为空白，整行判为不合规（白图数>0）
      all  — 全部图片均为空白，整行才判为不合规（白图数==图片总数>0）

    返回 summary dict：
      {output, total_rows, rows_done, offset,
       noncompliant_rows_any, noncompliant_rows_all,
       white_imgs, normal_imgs, fail_imgs, elapsed,
       selected_cols, rules}
    """
    if not os.path.exists(input):
        raise FileNotFoundError(f"输入文件不存在: {input}")

    # ── 读源表 ──
    wb = openpyxl.load_workbook(input, read_only=True, data_only=True)
    ws = wb[sheet] if sheet else wb[wb.sheetnames[0]]
    all_rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not all_rows:
        raise ValueError("表格为空")
    header = [str(c) if c is not None else "" for c in all_rows[0]]
    data_rows = all_rows[1:]

    # ── 选定 URL 列 ──
    if cols:
        selected = [c.strip() for c in cols.split(",") if c.strip()]
        missing = [c for c in selected if c not in header]
        if missing:
            raise ValueError(f"指定列不存在: {missing}；实际表头: {header}")
    else:
        selected = detect_col_name(header)
        if not selected:
            raise ValueError(f"未自动识别到 URL 列。表头: {header}；可用 cols 手动指定。")

    # ── 分批区间（offset / limit），行号始终映射回源表行号 ──
    total_data_rows = len(data_rows)
    offset = max(0, int(offset or 0))
    if offset > 0:
        data_rows = data_rows[offset:]
    if limit and limit > 0:
        data_rows = data_rows[:limit]
    # 本批第一行数据在源表中的行号（表头占第 1 行，故 +2）
    row_base = offset + 2

    out_path = output or (os.path.splitext(input)[0] + "_空白图结果.xlsx")
    prog_path = out_path + ".progress.json"

    # ── 续跑加载 ──
    results: dict[int, dict] = {}
    if resume and os.path.exists(prog_path):
        try:
            with open(prog_path, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            results = {int(k): v for k, v in loaded.get("results", {}).items()}
            if progress_cb:
                progress_cb(0, 0, f"[续跑] 已加载 {len(results)} 行进度")
        except Exception as e:
            if progress_cb:
                progress_cb(0, 0, f"[续跑] 进度读取失败，重新开始: {e}")

    # 待处理行（源行号：row_base 起，保证分批时行号与源表一致，续跑不串批）
    todo = [i for i in range(row_base, row_base + len(data_rows)) if i not in results]

    # 收集所有检测任务
    tasks = []  # (src_row_idx, col_name, url)
    for r_idx in todo:
        src = data_rows[r_idx - row_base]
        for col in selected:
            ci = header.index(col)
            cell = src[ci] if ci < len(src) else None
            for url in split_urls(cell):
                tasks.append((r_idx, col, url))

    if progress_cb:
        batch_tag = f" | 批次区间: 第{offset+1}-{offset+len(data_rows)}行" if offset else ""
        progress_cb(0, len(tasks),
                    f"[信息] 输入:{input} | URL列:{selected} | 待处理行:{len(todo)} | 任务数:{len(tasks)}{batch_tag}")

    lock = threading.Lock()
    t0 = time.time()

    def persist():
        tmp = prog_path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump({"selected_cols": selected, "results": results}, f,
                      ensure_ascii=False)
        os.replace(tmp, prog_path)

    cancelled = False
    if tasks:
        ex = ThreadPoolExecutor(max_workers=workers)
        try:
            fut_map = {}
            for r_idx, col, url in tasks:
                f = ex.submit(detect_from_url, url,
                              white_threshold=threshold,
                              max_color_variety=max_colors)
                fut_map[f] = (r_idx, col, url)
            done = 0
            _t_last = [time.time()]
            for f in as_completed(fut_map):
                # ── 取消检查：收到停止指令立即跳出，未开始的任务会被取消 ──
                if cancel_check and cancel_check():
                    cancelled = True
                    break
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
                    # 运行中定期落盘：每 persist_every 个任务、或每 60 秒写一次
                    # 目的 —— 长任务若中途崩溃/被关闭，已完成的判定结果不丢，可用 --resume 续跑
                    if persist_every and (
                        done % int(persist_every) == 0
                        or (time.time() - _t_last[0]) >= 60
                    ):
                        _t_last[0] = time.time()
                        persist()
                    if progress_cb and done % 5 == 0:
                        progress_cb(done, len(tasks),
                                    f"已完成 {done}/{len(tasks)} 个 URL 检测")

            if cancelled:
                # 取消队列中尚未开始的任务（正在执行的 workers 个会自然跑完，通常几秒内）
                ex.shutdown(wait=False, cancel_futures=True)
                with lock:
                    persist()          # 保存已完成的，便于后续 --resume 或导出
                if progress_cb:
                    progress_cb(done, len(tasks),
                                f"[已停止] 收到停止指令，已完成 {done}/{len(tasks)} 个 URL，"
                                f"进度已保存")
            else:
                if progress_cb:
                    progress_cb(len(tasks), len(tasks), "检测完成，正在写回 Excel…")
                persist()
        finally:
            ex.shutdown(wait=False)
    else:
        if progress_cb:
            progress_cb(0, 0, "[信息] 无待处理任务（可能已全部完成）")

    if cancelled:
        # 被取消：不写完整 Excel（未完成行无意义），直接返回已完成统计
        rows_done = len(results)
        non_any = sum(1 for r in results.values()
                      if (r.get("cols") and
                          sum(1 for c in r["cols"].values()
                              for p in c.get("per", []) if p.get("is_white_image")) > 0))
        return {
            "cancelled": True,
            "output": out_path,
            "total_rows": len(data_rows),
            "source_total_rows": total_data_rows,
            "offset": offset,
            "rows_done": rows_done,
            "noncompliant_rows_any": non_any,
            "noncompliant_rows_all": 0,
            "white_rows": non_any,
            "noncompliant_rows": non_any,
            "white_imgs": 0, "normal_imgs": 0, "fail_imgs": 0,
            "elapsed": round(time.time() - t0, 1),
            "selected_cols": selected,
            "rules": list(rules or ("any", "all")),
            "progress_path": prog_path,
        }

    # ── 汇总每个源行的判定 ──
    # 向后兼容：旧调用者可能传 rule="any"，归一成 rules 元组
    if rules is None or (isinstance(rules, (list, tuple)) and len(rules) == 0):
        if rule in ("any", "all"):
            rules = (rule,)
        else:
            rules = ("any", "all")
    rules = tuple(r for r in rules if r in ("any", "all")) or ("any", "all")

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
        # 同时计算两个维度的判定结果（一次扫描，两种结论并存）
        #   any：任意一张空白即不合规（white_total > 0）
        #   all：全部图片均为空白才算不合规（total>0 且 white_total==total）
        any_noncompliant = (white_total > 0)
        all_noncompliant = (total > 0 and white_total == total)
        row["overall"] = {
            "total": total, "white_total": white_total,
            "has_white": has_white,
            "noncompliant_any": any_noncompliant,
            "noncompliant_all": all_noncompliant,
            # 兼容旧 summary 字段
            "noncompliant": any_noncompliant,
        }

    # ── 写出 Excel ──
    out_appended = []
    for col in selected:
        out_appended += [f"{col}·判定", f"{col}·纯白数", f"{col}·明细"]
    # 双维度结果列（顺序 = rules 顺序）
    rule_col_names = []
    for r in rules:
        if r == "any":
            rule_col_names.append("一张即不合规")
        else:
            rule_col_names.append("全部空白才不合规")
    out_appended += rule_col_names + ["纯白图总数", "图片总数"]

    out_wb = openpyxl.Workbook()
    out_ws = out_wb.active
    out_ws.title = "检测结果"
    out_ws.append(list(header) + out_appended)

    noncompliant_any_rows = 0
    noncompliant_all_rows = 0
    white_imgs = 0
    normal_imgs = 0
    fail_imgs = 0

    for i, src in enumerate(data_rows):
        r_idx = i + row_base
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
        # 按 rules 顺序写每一维度
        for r in rules:
            key = "noncompliant_any" if r == "any" else "noncompliant_all"
            if ov:
                vals.append("是" if ov.get(key) else "否")
            else:
                vals.append("—")
        vals.append(ov.get("white_total", "—"))
        vals.append(ov.get("total", "—"))
        out_ws.append(vals)

        if ov:
            if ov.get("noncompliant_any"):
                noncompliant_any_rows += 1
            if ov.get("noncompliant_all"):
                noncompliant_all_rows += 1
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
        ("输入文件", input),
        ("输出文件", out_path),
        ("URL 列", ", ".join(selected)),
        ("判定维度（同时）", "、".join(
            "任意一张空白即不合规" if r == "any" else "全部图片均为空白才算不合规"
            for r in rules)),
        ("阈值(white_threshold)", threshold),
        ("颜色上限(max_color_variety)", max_colors),
        ("源表数据总行数", total_data_rows),
    ]
    if offset or (limit and limit > 0):
        summary.append(("本批区间", f"第 {offset+1} – {offset+len(data_rows)} 行"))
    summary += [
        ("本批处理行数", len(data_rows)),
        ("已处理行数", rows_done),
        ("不合规行数（一张即不合规 · any）", noncompliant_any_rows),
        ("不合规行数（全部空白才不合规 · all）", noncompliant_all_rows),
        ("纯白图数量(张)", white_imgs),
        ("正常图数量(张)", normal_imgs),
        ("检测失败数量(张)", fail_imgs),
        ("耗时(秒)", elapsed),
        ("生成时间", time.strftime("%Y-%m-%d %H:%M:%S")),
    ]
    for k, v in summary:
        sum_ws.append([k, v])
    sum_ws.column_dimensions["A"].width = 38
    sum_ws.column_dimensions["B"].width = 60

    out_wb.save(out_path)

    return {
        "output": out_path,
        "total_rows": len(data_rows),
        "source_total_rows": total_data_rows,
        "offset": offset,
        "rows_done": rows_done,
        # 双维度统计
        "noncompliant_rows_any": noncompliant_any_rows,
        "noncompliant_rows_all": noncompliant_all_rows,
        # 兼容旧 summary 字段（用 any 作为主维度，与旧版语义一致）
        "white_rows": noncompliant_any_rows,
        "noncompliant_rows": noncompliant_any_rows,
        "white_imgs": white_imgs,
        "normal_imgs": normal_imgs,
        "fail_imgs": fail_imgs,
        "elapsed": elapsed,
        "selected_cols": selected,
        "rules": list(rules),
    }


def count_data_rows(input, sheet=None) -> int:
    """只读表头，返回数据行数（不含表头）。用于分批前确定总批次数。"""
    wb = openpyxl.load_workbook(input, read_only=True, data_only=True)
    try:
        ws = wb[sheet] if sheet else wb[wb.sheetnames[0]]
        n = 0
        for i, _ in enumerate(ws.iter_rows(values_only=True)):
            if i == 0:
                continue
            n += 1
        return n
    finally:
        wb.close()


def merge_chunks(chunk_paths, out_path, rules=("any", "all"), input_name="",
                 threshold=0.99, max_colors=3, elapsed=None, cols_label=""):
    """
    把多个批次的结果 Excel 合并成一个全量 Excel。

    - "检测结果" 表：取首个批次的表头，依次追加各批次的数据行
    - "汇总" 表：重新按全量统计（不合规行数 / 纯白图数 / 正常图数 / 失败数）

    返回 dict：{output, rows, noncompliant_rows_any, noncompliant_rows_all,
               white_imgs, normal_imgs, fail_imgs, chunks}
    """
    if not chunk_paths:
        raise ValueError("没有可合并的批次文件")

    rules = tuple(rules) or ("any", "all")
    out_wb = openpyxl.Workbook()
    out_ws = out_wb.active
    out_ws.title = "检测结果"

    header = None
    total_rows = 0
    non_any = 0
    non_all = 0
    white_imgs = 0
    normal_imgs = 0
    fail_imgs = 0

    for path in chunk_paths:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            ws = wb["检测结果"] if "检测结果" in wb.sheetnames else wb[wb.sheetnames[0]]
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue
            hdr = [str(c) if c is not None else "" for c in rows[0]]
            if header is None:
                header = hdr
                out_ws.append(hdr)
            # 定位关键列（用列名找，兼容不同批次列数一致的情况）
            def col_of(name):
                return hdr.index(name) if name in hdr else None
            i_any = col_of("一张即不合规")
            i_all = col_of("全部空白才不合规")
            i_white = col_of("纯白图总数")
            i_total = col_of("图片总数")
            for r in rows[1:]:
                if r is None or all(v is None for v in r):
                    continue
                out_ws.append(list(r))
                total_rows += 1
                if i_any is not None and str(r[i_any]) == "是":
                    non_any += 1
                if i_all is not None and str(r[i_all]) == "是":
                    non_all += 1
                if i_white is not None and isinstance(r[i_white], (int, float)):
                    white_imgs += int(r[i_white])
                if i_total is not None and isinstance(r[i_total], (int, float)):
                    normal_imgs += int(r[i_total])
        finally:
            wb.close()

    # 正常图 = 图片总数 - 纯白图 - 失败（失败数从各批汇总表累加）
    for path in chunk_paths:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            if "汇总" not in wb.sheetnames:
                continue
            ws = wb["汇总"]
            for r in ws.iter_rows(values_only=True):
                if r and str(r[0]) == "检测失败数量(张)" and isinstance(r[1], (int, float)):
                    fail_imgs += int(r[1])
        finally:
            wb.close()
    normal_imgs = max(0, normal_imgs - white_imgs - fail_imgs)

    # 汇总表
    sum_ws = out_wb.create_sheet("汇总")
    summary = [
        ("纯白图识别 · 全量检测汇总（分批合并）", ""),
        ("输入文件", input_name or ""),
        ("输出文件", out_path),
        ("URL 列", cols_label),
        ("判定维度（同时）", "、".join(
            "任意一张空白即不合规" if r == "any" else "全部图片均为空白才算不合规"
            for r in rules)),
        ("阈值(white_threshold)", threshold),
        ("颜色上限(max_color_variety)", max_colors),
        ("批次数", len(chunk_paths)),
        ("数据总行数", total_rows),
        ("不合规行数（一张即不合规 · any）", non_any),
        ("不合规行数（全部空白才不合规 · all）", non_all),
        ("纯白图数量(张)", white_imgs),
        ("正常图数量(张)", normal_imgs),
        ("检测失败数量(张)", fail_imgs),
        ("生成时间", time.strftime("%Y-%m-%d %H:%M:%S")),
    ]
    if elapsed is not None:
        summary.insert(-1, ("总耗时(秒)", round(elapsed, 1)))
    for k, v in summary:
        sum_ws.append([k, v])
    sum_ws.column_dimensions["A"].width = 38
    sum_ws.column_dimensions["B"].width = 60

    out_wb.save(out_path)
    return {
        "output": out_path,
        "rows": total_rows,
        "noncompliant_rows_any": non_any,
        "noncompliant_rows_all": non_all,
        "white_imgs": white_imgs,
        "normal_imgs": normal_imgs,
        "fail_imgs": fail_imgs,
        "chunks": len(chunk_paths),
    }


def run_batch_chunked(input, output=None, chunk_size=1000, cols=None, sheet=None,
                      workers=8, threshold=0.99, max_colors=3, rules=("any", "all"),
                      resume=False, limit=0, progress_cb=None, chunk_cb=None,
                      persist_every=300, cancel_check=None) -> dict:
    """
    分批执行：数据超过 chunk_size 行时自动拆成多批，每批产出独立 Excel；
    每批完成即通过 chunk_cb 回调（GUI 可即时下载），全部完成后合并成全量 Excel。

    参数（其余同 run_batch）：
      chunk_size  每批行数上限（默认 1000；<=0 或 >=总行数时不分批）
      chunk_cb    回调(chunk_index, chunk_total, summary, chunk_path)
                  —— 每批完成时触发，可立即读取 chunk_path 提供下载
      progress_cb 回调(done, total, msg) —— 转发当前批的进度（消息带「第 i/N 批」前缀）

    返回 dict：{output, chunk_paths, total_rows, rows_done,
               noncompliant_rows_any, noncompliant_rows_all,
               white_imgs, normal_imgs, fail_imgs, elapsed, chunk_size, chunks, rules}
    """
    if not os.path.exists(input):
        raise FileNotFoundError(f"输入文件不存在: {input}")

    rules = tuple(r for r in (rules or ("any", "all")) if r in ("any", "all")) or ("any", "all")
    total_rows = count_data_rows(input, sheet=sheet)
    limit = int(limit or 0)
    work_rows = min(total_rows, limit) if limit > 0 else total_rows

    chunk_size = int(chunk_size or 0)
    if chunk_size <= 0 or chunk_size >= max(1, work_rows):
        total_chunks = 1
        chunk_size_eff = work_rows or 1
    else:
        total_chunks = (work_rows + chunk_size - 1) // chunk_size
        chunk_size_eff = chunk_size

    out_path = output or (os.path.splitext(input)[0] + "_空白图结果.xlsx")
    stem, ext = os.path.splitext(out_path)
    out_dir = os.path.dirname(out_path) or "."
    t0 = time.time()

    chunk_paths = []
    last_summary = {}
    cancelled = False
    done_rows_acc = 0

    def _emit(d, t, m):
        if progress_cb:
            progress_cb(d, t, m)

    for ci in range(total_chunks):
        off = ci * chunk_size_eff
        n = min(chunk_size_eff, work_rows - off)
        # 批次之间也检查一次取消信号（避免上一批刚结束又启动下一批）
        if ci > 0 and cancel_check and cancel_check():
            cancelled = True
            if progress_cb:
                progress_cb(0, 0, f"[已停止] 在开始第 {ci+1}/{total_chunks} 批前收到停止指令")
            break
        if total_chunks == 1:
            cpath = out_path
        else:
            cpath = os.path.join(out_dir, f"{os.path.basename(stem)}_批次{ci+1}{ext}")

        def _cb(d, t, m, _ci=ci):
            _emit(d, t, f"[第 {_ci+1}/{total_chunks} 批] {m}" if total_chunks > 1 else m)

        if total_chunks > 1 and progress_cb:
            progress_cb(0, 0, f"=== 开始第 {ci+1}/{total_chunks} 批（源表第 {off+1}-{off+n} 行）===")

        s = run_batch(
            input, output=cpath, cols=cols, sheet=sheet, workers=workers,
            threshold=threshold, max_colors=max_colors,
            limit=n, offset=off, resume=resume, rules=rules, progress_cb=_cb,
            persist_every=persist_every, cancel_check=cancel_check,
        )
        last_summary = s
        if s.get("cancelled"):
            # 该批未写出 Excel，不计入 chunk_paths；但已完成的行数要计入统计
            cancelled = True
            done_rows_acc += s.get("rows_done", 0)
            if progress_cb:
                progress_cb(0, 0, f"[已停止] 第 {ci+1}/{total_chunks} 批已取消"
                                  f"（该批已完成 {s.get('rows_done', 0)} 行，未写出文件）")
            break
        done_rows_acc += s.get("rows_done", 0)
        chunk_paths.append(cpath)
        if chunk_cb:
            chunk_cb(ci + 1, total_chunks, s, cpath)

    elapsed = round(time.time() - t0, 1)

    if cancelled:
        # 被取消：只交付「已完整跑完的批次」，不合并全量
        result = {
            "cancelled": True,
            "output": None,
            "chunk_paths": chunk_paths,
            "completed_chunks": len(chunk_paths),
            "total_chunks": total_chunks,
            "total_rows": total_rows,
            "rows_done": done_rows_acc,
            "noncompliant_rows_any": 0, "noncompliant_rows_all": 0,
            "white_imgs": 0, "normal_imgs": 0, "fail_imgs": 0,
            "elapsed": elapsed,
            "chunk_size": chunk_size_eff, "chunks": len(chunk_paths),
            "rules": list(rules),
        }
        return result

    if total_chunks == 1:
        # 单批：直接复用该批 summary（已是全量）
        result = {
            "output": chunk_paths[0], "chunk_paths": chunk_paths,
            "total_rows": total_rows, "rows_done": last_summary.get("rows_done", work_rows),
            "noncompliant_rows_any": last_summary.get("noncompliant_rows_any", 0),
            "noncompliant_rows_all": last_summary.get("noncompliant_rows_all", 0),
            "white_imgs": last_summary.get("white_imgs", 0),
            "normal_imgs": last_summary.get("normal_imgs", 0),
            "fail_imgs": last_summary.get("fail_imgs", 0),
            "elapsed": elapsed,
            "chunk_size": chunk_size_eff, "chunks": 1, "rules": list(rules),
        }
    else:
        agg = merge_chunks(chunk_paths, out_path, rules=rules, input_name=input,
                           threshold=threshold, max_colors=max_colors, elapsed=elapsed)
        result = {
            "output": out_path, "chunk_paths": chunk_paths,
            "total_rows": total_rows, "rows_done": agg["rows"],
            "noncompliant_rows_any": agg["noncompliant_rows_any"],
            "noncompliant_rows_all": agg["noncompliant_rows_all"],
            "white_imgs": agg["white_imgs"], "normal_imgs": agg["normal_imgs"],
            "fail_imgs": agg["fail_imgs"], "elapsed": elapsed,
            "chunk_size": chunk_size_eff, "chunks": total_chunks, "rules": list(rules),
        }
    # 兼容字段
    result["white_rows"] = result["noncompliant_rows_any"]
    result["noncompliant_rows"] = result["noncompliant_rows_any"]
    return result


def export_from_progress(input, progress_path, output=None, rules=("any", "all"),
                         cols=None, sheet=None, only_done=True) -> dict:
    """
    从进度文件（*.progress.json）导出「当前已完成部分」的结果 Excel。

    用途：长任务跑到一半也能先把已完成的判定结果取走，不必等全部跑完。

    only_done=True（默认）：
        只导出「已完成」的行，并在首列加「源表行号」便于回溯 ——
        大表运行中导出为秒级（不写上万行「未处理」占位行，避免与检测线程抢 GIL 卡死）
    only_done=False：
        输出全表，未完成的行判定列标为「未处理」（行号与源表完全对齐，仅建议任务结束后用）

    返回 {output, total_rows, rows_done, pending_rows,
          noncompliant_rows_any, noncompliant_rows_all,
          white_imgs, normal_imgs, fail_imgs, rules, selected_cols}
    """
    if not os.path.exists(progress_path):
        raise FileNotFoundError(f"进度文件不存在: {progress_path}")
    if not os.path.exists(input):
        raise FileNotFoundError(f"输入文件不存在: {input}")

    rules = tuple(r for r in (rules or ("any", "all")) if r in ("any", "all")) or ("any", "all")

    with open(progress_path, "r", encoding="utf-8") as f:
        prog = json.load(f)
    results = {int(k): v for k, v in (prog.get("results") or {}).items()}

    # ── 流式读源表：only_done 模式只保留「已完成」的行，其余直接丢弃 ──
    wb = openpyxl.load_workbook(input, read_only=True, data_only=True)
    ws = wb[sheet] if sheet else wb[wb.sheetnames[0]]
    it = ws.iter_rows(values_only=True)
    first = next(it, None)
    if first is None:
        wb.close()
        raise ValueError("表格为空")
    header = [str(c) if c is not None else "" for c in first]

    if only_done:
        done_src: dict[int, tuple] = {}
        total_data_rows = 0
        for i, src in enumerate(it):
            total_data_rows += 1
            r_idx = i + 2
            row_res = results.get(r_idx)
            if row_res and row_res.get("cols"):
                done_src[r_idx] = src
        data_rows = None
    else:
        done_src = None
        data_rows = list(it)
        total_data_rows = len(data_rows)
    wb.close()

    if cols:
        selected = [c.strip() for c in cols.split(",") if c.strip()]
    else:
        sc = prog.get("selected_cols")
        selected = list(sc) if sc else detect_col_name(header)
    if not selected:
        raise ValueError("无法确定 URL 列（进度文件无 selected_cols，且未自动识别到）")

    out_path = output or (os.path.splitext(input)[0] + "_已完成部分.xlsx")

    # ── 逐行算判定 ──
    for r_idx, row in results.items():
        total = 0
        white_total = 0
        has_white = False
        for col, colres in row.get("cols", {}).items():
            per = colres.get("per", [])
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
            "noncompliant_any": (white_total > 0),
            "noncompliant_all": (total > 0 and white_total == total),
        }

    # ── 写 Excel ──
    out_appended = []
    for col in selected:
        out_appended += [f"{col}·判定", f"{col}·纯白数", f"{col}·明细"]
    for r in rules:
        out_appended.append("一张即不合规" if r == "any" else "全部空白才不合规")
    out_appended += ["纯白图总数", "图片总数"]

    out_wb = openpyxl.Workbook()
    out_ws = out_wb.active
    out_ws.title = "检测结果"

    non_any = non_all = 0
    white_imgs = normal_imgs = fail_imgs = 0
    rows_done = 0

    if only_done:
        # 只写已完成行，首列「源表行号」
        out_ws.append(["源表行号"] + list(header) + out_appended)
        for r_idx in sorted(done_src):
            src = done_src[r_idx]
            row_res = results.get(r_idx)
            vals = [r_idx]
            v = list(src)
            while len(v) < len(header):
                v.append(None)
            vals += v
            rows_done += 1
            for col in selected:
                colres = row_res.get("cols", {}).get(col, {})
                vals.append("是" if colres.get("has_white") else ("否" if "per" in colres else "—"))
                vals.append(colres.get("white_count", "—"))
                vals.append(colres.get("detail", "—"))
            ov = row_res.get("overall", {})
            for r in rules:
                key = "noncompliant_any" if r == "any" else "noncompliant_all"
                vals.append("是" if ov.get(key) else "否")
            vals.append(ov.get("white_total", 0))
            vals.append(ov.get("total", 0))
            out_ws.append(vals)

            if ov.get("noncompliant_any"):
                non_any += 1
            if ov.get("noncompliant_all"):
                non_all += 1
            white_imgs += ov.get("white_total", 0)
            for colres in row_res.get("cols", {}).values():
                for p in colres.get("per", []):
                    if not p.get("success"):
                        fail_imgs += 1
                    elif not p.get("is_white_image"):
                        normal_imgs += 1
        pending = total_data_rows - rows_done
    else:
        out_ws.append(list(header) + out_appended)
        for i, src in enumerate(data_rows):
            r_idx = i + 2
            row_res = results.get(r_idx)
            vals = list(src)
            while len(vals) < len(header):
                vals.append(None)
            if not row_res or not row_res.get("cols"):
                # 尚未处理的行
                for _col in selected:
                    vals += ["未处理", "—", "—"]
                for _r in rules:
                    vals.append("未处理")
                vals += ["—", "—"]
                out_ws.append(vals)
                continue
            rows_done += 1
            for col in selected:
                colres = row_res.get("cols", {}).get(col, {})
                vals.append("是" if colres.get("has_white") else ("否" if "per" in colres else "—"))
                vals.append(colres.get("white_count", "—"))
                vals.append(colres.get("detail", "—"))
            ov = row_res.get("overall", {})
            for r in rules:
                key = "noncompliant_any" if r == "any" else "noncompliant_all"
                vals.append("是" if ov.get(key) else "否")
            vals.append(ov.get("white_total", 0))
            vals.append(ov.get("total", 0))
            out_ws.append(vals)

            if ov.get("noncompliant_any"):
                non_any += 1
            if ov.get("noncompliant_all"):
                non_all += 1
            white_imgs += ov.get("white_total", 0)
            for colres in row_res.get("cols", {}).values():
                for p in colres.get("per", []):
                    if not p.get("success"):
                        fail_imgs += 1
                    elif not p.get("is_white_image"):
                        normal_imgs += 1
        pending = total_data_rows - rows_done

    sum_ws = out_wb.create_sheet("汇总")
    summary = [
        ("纯白图识别 · 已完成部分导出（中途快照）", ""),
        ("导出模式", "仅已完成行" if only_done else "全表（未完成行标「未处理」）"),
        ("输入文件", input),
        ("进度文件", progress_path),
        ("输出文件", out_path),
        ("URL 列", ", ".join(selected)),
        ("数据总行数", total_data_rows),
        ("已完成行数", rows_done),
        ("未处理行数", pending),
        ("不合规行数（一张即不合规 · any）", non_any),
        ("不合规行数（全部空白才不合规 · all）", non_all),
        ("纯白图数量(张)", white_imgs),
        ("正常图数量(张)", normal_imgs),
        ("检测失败数量(张)", fail_imgs),
        ("生成时间", time.strftime("%Y-%m-%d %H:%M:%S")),
    ]
    for k, v in summary:
        sum_ws.append([k, v])
    sum_ws.column_dimensions["A"].width = 38
    sum_ws.column_dimensions["B"].width = 60
    out_wb.save(out_path)

    return {
        "output": out_path,
        "total_rows": total_data_rows,
        "rows_done": rows_done,
        "pending_rows": pending,
        "noncompliant_rows_any": non_any,
        "noncompliant_rows_all": non_all,
        "white_imgs": white_imgs,
        "normal_imgs": normal_imgs,
        "fail_imgs": fail_imgs,
        "rules": list(rules),
        "selected_cols": selected,
    }


def main():
    ap = argparse.ArgumentParser(description="纯白图识别 · Excel 批量检测")
    ap.add_argument("input", help="输入 Excel 路径")
    ap.add_argument("--output", default=None, help="输出 Excel 路径")
    ap.add_argument("--cols", default=None, help="指定 URL 列名，逗号分隔；不填则自动识别")
    ap.add_argument("--sheet", default=None, help="工作表名或索引（默认第一个）")
    ap.add_argument("--workers", type=int, default=8, help="并发线程数（默认 8）")
    ap.add_argument("--threshold", type=float, default=0.99, help="近白像素占比阈值（默认 0.99）")
    ap.add_argument("--max-colors", type=int, default=3, help="颜色种类上限（默认 3）")
    ap.add_argument("--limit", type=int, default=0, help="只处理前 N 行数据（测试用）")
    ap.add_argument("--chunk-size", type=int, default=1000,
                    help="分批行数上限（默认 1000；超过则分批跑并逐批产出 Excel，最后合并全量）")
    ap.add_argument("--resume", action="store_true", help="断点续跑（读取已有进度）")
    ap.add_argument("--list-cols", action="store_true", help="只打印识别到的 URL 列并退出")
    ap.add_argument("--rule", default="any,all",
                    help="不合规维度（逗号分隔，可多选）：any=一张即不合规 / all=全部空白才不合规；"
                         "默认 any,all（两种维度同时输出）")
    ap.add_argument("--persist-every", type=int, default=300,
                    help="运行中每完成 N 个 URL 落盘一次进度（默认 300；0=只在结束时落盘）")
    ap.add_argument("--export-partial", default=None, metavar="PROGRESS_JSON",
                    help="不跑检测：从指定的 *.progress.json 导出「已完成部分」结果 Excel")
    args = ap.parse_args()

    # 解析 --rule 字符串（如 "any,all" / "any" / "all"）
    rules = tuple(r.strip() for r in args.rule.split(",") if r.strip() in ("any", "all"))
    if not rules:
        rules = ("any", "all")

    if not os.path.exists(args.input):
        print(f"[错误] 输入文件不存在: {args.input}")
        sys.exit(1)

    # ── 仅列出识别到的 URL 列 ──
    if args.list_cols:
        wb = openpyxl.load_workbook(args.input, read_only=True, data_only=True)
        ws = wb[args.sheet] if args.sheet else wb[wb.sheetnames[0]]
        first = next(ws.iter_rows(values_only=True), None)
        wb.close()
        header = [str(c) if c is not None else "" for c in (first or [])]
        selected = detect_col_name(header)
        n = count_data_rows(args.input, sheet=args.sheet)
        print("识别到的 URL 列：")
        for c in selected:
            print("  -", c)
        print(f"\n共 {n} 行数据。")
        if args.chunk_size and n > args.chunk_size:
            import math as _m
            print(f"按 --chunk-size {args.chunk_size} 将拆成 "
                  f"{_m.ceil(n / args.chunk_size)} 批执行。")
        sys.exit(0)

    # ── 从进度文件导出「已完成部分」（不跑检测）──
    if args.export_partial:
        try:
            sp = export_from_progress(
                args.input, args.export_partial, output=args.output,
                rules=rules, cols=args.cols, sheet=args.sheet)
        except Exception as e:
            print(f"[错误] {e}")
            sys.exit(1)
        print(f"\n[完成] 已完成部分已导出: {sp['output']}")
        print(f"[完成] 已完成 {sp['rows_done']} 行 / 未处理 {sp['pending_rows']} 行 "
              f"（共 {sp['total_rows']} 行）")
        print(f"[完成] any 不合规 {sp['noncompliant_rows_any']} 行 | "
              f"all 不合规 {sp['noncompliant_rows_all']} 行 | "
              f"纯白图 {sp['white_imgs']} 张 | 正常图 {sp['normal_imgs']} 张 | "
              f"失败 {sp['fail_imgs']} 张")
        sys.exit(0)

    try:
        def _chunk_done(ci, ct, s, path):
            print(f"[批次完成] 第 {ci}/{ct} 批 → {path} "
                  f"（any {s['noncompliant_rows_any']} 行 / all {s['noncompliant_rows_all']} 行）",
                  flush=True)

        s = run_batch_chunked(
            args.input, output=args.output, cols=args.cols, sheet=args.sheet,
            workers=args.workers, threshold=args.threshold, max_colors=args.max_colors,
            limit=args.limit, chunk_size=args.chunk_size, resume=args.resume, rules=rules,
            progress_cb=lambda d, t, m: print(m),
            chunk_cb=_chunk_done,
        )
    except Exception as e:
        print(f"[错误] {e}")
        sys.exit(1)

    print(f"\n[完成] 全量结果: {s['output']}")
    if s.get("chunks", 1) > 1:
        print(f"[完成] 共 {s['chunks']} 批（每批 {s['chunk_size']} 行），分批文件：")
        for p in s.get("chunk_paths", []):
            print(f"        - {p}")
    rules_label = "、".join(
        "任意一张空白即不合规" if r == "any" else "全部图片均为空白才算不合规"
        for r in s["rules"])
    print(f"[完成] 判定维度: {rules_label}")
    print(f"[完成] 已处理 {s['rows_done']} 行 | "
          f"any 不合规 {s['noncompliant_rows_any']} 行 | "
          f"all 不合规 {s['noncompliant_rows_all']} 行 | "
          f"纯白图 {s['white_imgs']} 张 | 正常图 {s['normal_imgs']} 张 | 失败 {s['fail_imgs']} 张")
    print(f"[完成] 耗时 {s['elapsed']}s")


if __name__ == "__main__":
    main()
