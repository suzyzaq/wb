# -*- coding: utf-8 -*-
"""
纯白图识别 API 服务（FastAPI）
====================================================================
端点：
  POST /api/v1/detect           文件上传
  POST /api/v1/detect/url       URL 输入（JSON）
  POST /api/v1/detect/base64    Base64 输入（JSON）
  GET  /api/v1/health          健康检查
  GET  /api/v1/guide            Coze / Dify 接入指南（网页）
  GET  /api/v1/openapi.yaml    OpenAPI Schema（供插件导入）

特点：
  - 纯像素分析，不依赖 AI 模型 / Dify / GPU
  - 毫秒级响应，轻量可部署到任意云服务器 / 容器
  - 响应结构对齐 RMBG-2.0 插件，便于统一接入工作流
====================================================================
"""

from __future__ import annotations

import os
import time
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, File, UploadFile, Form
from fastapi.responses import JSONResponse, PlainTextResponse, Response
import uvicorn

from white_detect import detect_from_url, detect_from_base64, detect_from_file

ROOT = Path(__file__).parent
SERVICE_NAME = "纯白图识别 API"
VERSION = "1.0.0"
PORT = int(os.getenv("WHITE_DETECT_PORT", "7861"))

app = FastAPI(
    title=SERVICE_NAME,
    description="基于像素分析的纯白图/空图识别服务（不依赖 AI 模型，轻量可云端部署）",
    version=VERSION,
)


# ====================================================================
# 健康检查
# ====================================================================
@app.get("/api/v1/health")
def health():
    return {
        "status": "ok",
        "service": SERVICE_NAME,
        "version": VERSION,
        "engine": "pixel-analysis (numpy + Pillow)",
        "model_required": False,
    }


# ====================================================================
# 核心检测端点
# ====================================================================
@app.post("/api/v1/detect")
async def detect_file(
    file: UploadFile = File(...),
    white_threshold: float = Form(0.99),
    max_color_variety: int = Form(3),
):
    """文件上传检测"""
    try:
        data = await file.read()
    except Exception as e:
        return JSONResponse(
            status_code=400,
            content={"success": False, "error": "read_failed", "reason": str(e)},
        )
    result = detect_from_file(
        data,
        white_threshold=white_threshold,
        max_color_variety=max_color_variety,
    )
    return result


@app.post("/api/v1/detect/url")
def detect_url(payload: dict):
    """URL 检测"""
    image_url = payload.get("image_url", "")
    white_threshold = float(payload.get("white_threshold", 0.99))
    max_color_variety = int(payload.get("max_color_variety", 3))
    result = detect_from_url(
        image_url,
        white_threshold=white_threshold,
        max_color_variety=max_color_variety,
    )
    return result


@app.post("/api/v1/detect/base64")
def detect_base64(payload: dict):
    """Base64 检测"""
    image_base64 = payload.get("image_base64", "")
    white_threshold = float(payload.get("white_threshold", 0.99))
    max_color_variety = int(payload.get("max_color_variety", 3))
    result = detect_from_base64(
        image_base64,
        white_threshold=white_threshold,
        max_color_variety=max_color_variety,
    )
    return result


# ====================================================================
# OpenAPI Schema
# ====================================================================
@app.get("/api/v1/openapi.yaml", response_class=PlainTextResponse)
def openapi_yaml():
    yaml_path = ROOT / "openapi.yaml"
    if yaml_path.exists():
        return yaml_path.read_text(encoding="utf-8")
    # 兜底：用 FastAPI 自动生成的 JSON 转简易说明
    return PlainTextResponse(
        "# openapi.yaml 未找到，请确认文件已随服务部署",
        status_code=404,
    )


# ====================================================================
# Coze / Dify 接入指南（网页）
# ====================================================================
GUIDE_HTML = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>纯白图识别 API · 接入指南</title>
<style>
  body{{font-family:-apple-system,Segoe UI,Roboto,'Microsoft YaHei',sans-serif;
       max-width:860px;margin:40px auto;padding:0 24px;color:#1f2937;line-height:1.7}}
  h1{{color:#111827;border-bottom:3px solid #2563eb;padding-bottom:8px}}
  h2{{color:#1d4ed8;margin-top:32px}}
  code{{background:#f1f5f9;padding:2px 6px;border-radius:4px;font-size:13px}}
  pre{{background:#0f172a;color:#e2e8f0;padding:16px;border-radius:8px;overflow:auto;font-size:13px}}
  .card{{background:#eff6ff;border:1px solid #bfdbfe;border-radius:10px;padding:16px 20px;margin:16px 0}}
  table{{border-collapse:collapse;width:100%;margin:12px 0}}
  th,td{{border:1px solid #cbd5e1;padding:8px 10px;text-align:left;font-size:14px}}
  th{{background:#f8fafc}}
  .tag{{display:inline-block;background:#2563eb;color:#fff;border-radius:999px;
       padding:2px 10px;font-size:12px;margin-right:6px}}
</style>
</head>
<body>
<h1>🔲 纯白图识别 API · 接入指南</h1>
<p>本服务基于<strong>像素分析</strong>识别纯白图 / 空图，<strong>不依赖 AI 模型</strong>，
毫秒级响应，可云端部署后接入 Coze / Dify / 钉钉 AI 表格 等工作流。</p>

<div class="card">
  <span class="tag">判定结果</span>
  <code>is_white_image = true</code> → 纯白图 / 纯色空图（应标记/退回）<br>
  <code>is_white_image = false</code> → 正常图（含商品主体）
</div>

<h2>① 端点一览</h2>
<table>
<tr><th>方法</th><th>路径</th><th>说明</th></tr>
<tr><td>POST</td><td><code>/api/v1/detect</code></td><td>文件上传（form-data: file）</td></tr>
<tr><td>POST</td><td><code>/api/v1/detect/url</code></td><td>URL 输入（JSON: image_url）</td></tr>
<tr><td>POST</td><td><code>/api/v1/detect/base64</code></td><td>Base64 输入（JSON: image_base64）</td></tr>
<tr><td>GET</td><td><code>/api/v1/health</code></td><td>健康检查</td></tr>
</table>

<h2>② 请求示例（URL 模式）</h2>
<pre>curl -X POST {base}/api/v1/detect/url \\
  -H "Content-Type: application/json" \\
  -d '{{"image_url":"https://img.officemate.cn/aliyun-oss/goods/xxx.jpg"}}'</pre>

<h2>③ 响应示例</h2>
<pre>{{
  "success": true,
  "is_white_image": true,
  "category": "纯白图",
  "white_ratio": 1.0,
  "color_variety": 1,
  "dominant_color": [248, 248, 248],
  "image_size": [800, 800],
  "format": "JPEG",
  "confidence": 0.99,
  "reason": "图片为纯色/接近纯色浅色图（近白像素占比 100.0%），判定为纯白图",
  "processing_time": 0.12
}}</pre>

<h2>④ 接入 Coze / Dify</h2>
<ol>
  <li>将本服务通过 <code>start_public.bat</code> 启动 ngrok 公网隧道，获得公域地址。</li>
  <li>把公域地址填入 <code>openapi.yaml</code> 的 <code>servers.url</code>。</li>
  <li><strong>Coze</strong>：插件 → 创建插件 → 导入 OpenAPI Schema → 选择本文件。</li>
  <li><strong>Dify</strong>：工具 → 自定义 → 创建 → 导入 OpenAPI Schema。</li>
</ol>

<div class="card">
  <strong>参数说明</strong>
  <ul>
    <li><code>white_threshold</code>（默认 0.99）：近白像素占比阈值，≥ 此值判为纯白图。</li>
    <li><code>max_color_variety</code>（默认 3）：颜色种类上限，≤ 此值且整体浅色判为纯色空图。</li>
  </ul>
</div>

<p style="color:#64748b;font-size:13px;margin-top:40px">
Generated by 纯白图识别 API · {version}</p>
</body>
</html>
"""


@app.get("/api/v1/guide")
def guide():
    base = os.getenv("PUBLIC_BASE_URL", f"http://localhost:{PORT}")
    html = GUIDE_HTML.format(base=base, version=VERSION)
    return Response(content=html, media_type="text/html")


# ====================================================================
# 根路径
# ====================================================================
@app.get("/")
def root():
    return {
        "service": SERVICE_NAME,
        "version": VERSION,
        "docs": "/docs",
        "guide": "/api/v1/guide",
        "health": "/api/v1/health",
    }


if __name__ == "__main__":
    print(f"🔲 {SERVICE_NAME} v{VERSION} 启动中... (port {PORT})")
    uvicorn.run(app, host="0.0.0.0", port=PORT, log_level="info")
