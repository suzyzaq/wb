# -*- coding: utf-8 -*-
"""
纯白图（白底图）识别核心
====================================================================
基于像素分析的纯白图 / 空图检测，不依赖任何 AI 模型、不调用 Dify、
不需要 GPU。毫秒级响应，可在任意普通服务器/容器运行。

判定逻辑：
  - 近白像素占比（RGB 均 > 240）≥ white_threshold（默认 0.99）
    → 纯白图
  - 颜色种类极少（量化后 ≤ max_color_variety）且整体浅色
    → 纯色图 / 疑似空图
  - 其余 → 正常图（有商品主体等真实内容）

输出结构化 JSON，便于接入 Coze / Dify / 钉钉 AI 表格 等工作流。

依赖：numpy、Pillow、requests（仅 URL 模式需要）
====================================================================
"""

from __future__ import annotations

import base64
import io
import time
from typing import Optional

import numpy as np
from PIL import Image

try:
    import requests
except Exception:  # pragma: no cover
    requests = None

# URL 模式下载头（伪装浏览器，部分 CDN 会拦截无 UA 请求）
DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Referer": "https://img.officemate.cn/",
}

# 下载超时（连接 10s，读取 30s）
URL_TIMEOUT = (10, 30)


# ====================================================================
# 核心：字节流分析
# ====================================================================
def analyze_image_bytes(
    data: bytes,
    white_threshold: float = 0.99,
    max_color_variety: int = 3,
) -> dict:
    """
    分析图片字节流，判定是否为纯白图 / 空图。

    返回结构化 dict（success=True 时含判定结果）。
    """
    t0 = time.time()
    if not data:
        return {
            "success": False,
            "error": "empty_image_data",
            "reason": "图片数据为空",
            "processing_time": round(time.time() - t0, 4),
        }

    try:
        img = Image.open(io.BytesIO(data))
        img_format = img.format or "UNKNOWN"
        img = img.convert("RGB")
        arr = np.array(img)
        h, w = arr.shape[:2]
    except Exception as e:
        return {
            "success": False,
            "error": "decode_failed",
            "reason": f"图片解码失败: {e}",
            "processing_time": round(time.time() - t0, 4),
        }

    try:
        # 近白像素（RGB 均 > 240）
        white_mask = np.all(arr > 240, axis=2)
        white_ratio = float(white_mask.mean())

        # 颜色种类（量化到 16 级，抑制噪点）
        reshaped = arr.reshape(-1, 3)
        quantized = (reshaped // 16).astype(np.uint8)
        unique_colors = np.unique(quantized, axis=0)
        n_colors = int(len(unique_colors))

        # 主色（量化后出现次数最多的颜色 ×16 还原，取区间中值）
        vals, counts = np.unique(quantized, axis=0, return_counts=True)
        dominant_q = vals[np.argmax(counts)]
        dominant_color = [int(dominant_q[0] * 16 + 8),
                           int(dominant_q[1] * 16 + 8),
                           int(dominant_q[2] * 16 + 8)]

        # ── 判定 ──
        if white_ratio >= white_threshold:
            is_white = True
            category = "纯白图"
            confidence = round(min(1.0, white_ratio + 0.01), 4)
            reason = (
                f"图片为纯色/接近纯色浅色图（近白像素占比 {white_ratio*100:.1f}%），"
                f"判定为纯白图"
            )
        elif n_colors <= max_color_variety and white_ratio >= 0.9:
            is_white = True
            category = "纯色图"
            confidence = round(min(1.0, 0.85 + 0.1 * (1 - white_ratio)), 4)
            reason = (
                f"图片颜色种类极少（{n_colors} 种）且整体浅色（近白占比 {white_ratio*100:.1f}%），"
                f"疑似纯色空图/白底图"
            )
        elif n_colors <= max_color_variety:
            # 单色但非浅色（如纯红/纯灰背景）
            is_white = True
            category = "纯色图"
            confidence = round(0.8, 4)
            reason = (
                f"图片为单一颜色背景（{n_colors} 种颜色，主色 RGB={dominant_color}），"
                f"无商品主体内容，判定为纯色空图"
            )
        else:
            is_white = False
            category = "正常图"
            confidence = round(min(1.0, 0.7 + 0.03 * n_colors), 4)
            reason = (
                f"图片含丰富内容（颜色种类 {n_colors} 种，近白占比 {white_ratio*100:.1f}%），"
                f"判定为正常图（含商品主体）"
            )

        return {
            "success": True,
            "is_white_image": is_white,
            "category": category,
            "white_ratio": round(white_ratio, 4),
            "color_variety": n_colors,
            "dominant_color": dominant_color,
            "image_size": [int(w), int(h)],
            "format": img_format,
            "confidence": confidence,
            "reason": reason,
            "processing_time": round(time.time() - t0, 4),
        }
    except Exception as e:
        return {
            "success": False,
            "error": "analyze_failed",
            "reason": f"像素分析失败: {e}",
            "processing_time": round(time.time() - t0, 4),
        }


# ====================================================================
# 三种输入适配
# ====================================================================
def detect_from_base64(image_base64: str, **kwargs) -> dict:
    """Base64（支持 data:image/xxx;base64, 前缀）"""
    if not image_base64:
        return {"success": False, "error": "empty_base64", "reason": "Base64 数据为空"}
    try:
        if "," in image_base64 and image_base64.strip().startswith("data:"):
            image_base64 = image_base64.split(",", 1)[1]
        data = base64.b64decode(image_base64)
    except Exception as e:
        return {"success": False, "error": "base64_decode_failed",
                "reason": f"Base64 解码失败: {e}"}
    return analyze_image_bytes(data, **kwargs)


def detect_from_url(image_url: str, **kwargs) -> dict:
    """图片 URL（http/https 直接下载；file:// 本地直读）"""
    if not image_url:
        return {"success": False, "error": "empty_url", "reason": "图片 URL 为空"}
    # 本地文件 file:// → 直接读取，便于本地用例/测试/容器场景
    if image_url.lower().startswith("file:"):
        try:
            from urllib.parse import unquote, urlparse
            p = urlparse(image_url)
            local = unquote(p.path)
            if local.startswith("/") and len(local) > 2 and local[2] == ":":
                local = local[1:]  # Windows: /C:/foo → C:/foo
            with open(local, "rb") as f:
                data = f.read()
            if not data:
                return {"success": False, "error": "empty_file",
                        "reason": "本地文件为空"}
            return analyze_image_bytes(data, **kwargs)
        except Exception as e:
            return {"success": False, "error": "local_file_error",
                    "reason": f"本地文件读取失败: {e}"}
    if requests is None:
        return {"success": False, "error": "missing_requests",
                "reason": "缺少 requests 模块，无法下载 URL"}
    try:
        resp = requests.get(
            image_url, headers=DEFAULT_HEADERS, timeout=URL_TIMEOUT
        )
        if resp.status_code != 200 or not resp.content:
            return {
                "success": False,
                "error": "download_failed",
                "reason": f"图片下载失败（HTTP {resp.status_code}）",
            }
        return analyze_image_bytes(resp.content, **kwargs)
    except Exception as e:
        return {"success": False, "error": "download_exception",
                "reason": f"图片下载异常: {e}"}


def detect_from_file(file_bytes: bytes, **kwargs) -> dict:
    """文件字节流（上传）"""
    return analyze_image_bytes(file_bytes, **kwargs)


# ====================================================================
# 批量（供 Excel 批处理复用）
# ====================================================================
def detect_batch(urls: list[str], white_threshold: float = 0.99,
                 max_color_variety: int = 3) -> list[dict]:
    """批量检测一组 URL，返回与输入等长的判定列表。"""
    results = []
    for u in urls:
        r = detect_from_url(u, white_threshold=white_threshold,
                            max_color_variety=max_color_variety)
        r["image_url"] = u
        results.append(r)
    return results


if __name__ == "__main__":
    # 快速自测
    import sys
    if len(sys.argv) > 1:
        url = sys.argv[1]
        print(detect_from_url(url))
    else:
        print("用法: python white_detect.py <image_url>")
