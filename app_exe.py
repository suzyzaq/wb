# -*- coding: utf-8 -*-
"""
纯白图（白底图）识别 · 单机版 EXE 入口
========================================================
基于像素分析的纯白图 / 空图检测，不依赖 AI 模型、不调用 Dify、不需要 GPU。
双击运行即启动本地 HTTP 服务 + 内置网页 UI，也可作为 HTTP API 插件调用。

端点（默认端口 7861，可用环境变量 PORT 或命令行 --port 覆盖）：
  GET  /                      网页 UI
  GET  /api/v1/health        健康检查
  POST /api/v1/detect/url    {image_url}              URL 检测
  POST /api/v1/detect/base64 {image_base64}          Base64 检测（也用于网页上传）
  POST /api/v1/detect/file   文件上传（multipart）     文件检测

依赖：white_detect.py（同目录）、numpy、Pillow、requests（URL 模式可选）
========================================================
"""

from __future__ import annotations

import os
import sys
import json
import webbrowser
import threading
import email
from email import policy as email_policy
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

import white_detect

# ── 配置 ──
PORT = int(os.getenv("PORT", "7861"))
HOST = "0.0.0.0"


# ====================================================================
# 网页 UI（内联，无需外部资源）
# ====================================================================
PAGE_HTML = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>纯白图识别 · 白底图检测</title>
<style>
  * { box-sizing: border-box; }
  body { margin:0; font-family:-apple-system,"Microsoft YaHei",sans-serif;
         background:#f5f6fa; color:#222; }
  .wrap { max-width:880px; margin:0 auto; padding:24px 16px 60px; }
  h1 { font-size:22px; margin:0 0 4px; }
  .sub { color:#888; font-size:13px; margin-bottom:18px; }
  .tabs { display:flex; gap:8px; margin-bottom:14px; flex-wrap:wrap; }
  .tabs button { border:1px solid #d0d3dc; background:#fff; padding:8px 16px;
                 border-radius:8px; cursor:pointer; font-size:14px; color:#444; }
  .tabs button.active { background:#3b6ef0; color:#fff; border-color:#3b6ef0; }
  .panel { background:#fff; border:1px solid #e3e6ee; border-radius:12px;
           padding:18px; margin-bottom:16px; }
  .panel.hidden { display:none; }
  input[type=text], textarea { width:100%; padding:10px 12px; border:1px solid #d0d3dc;
           border-radius:8px; font-size:14px; outline:none; }
  textarea { min-height:90px; resize:vertical; font-family:monospace; }
  .row { display:flex; gap:10px; align-items:center; }
  .row input { flex:1; }
  button.go { background:#3b6ef0; color:#fff; border:none; padding:10px 22px;
              border-radius:8px; cursor:pointer; font-size:14px; }
  button.go:disabled { background:#a9bdf3; cursor:wait; }
  .filezone { border:2px dashed #c5cbe0; border-radius:10px; padding:22px;
              text-align:center; color:#777; cursor:pointer; }
  .filezone:hover { border-color:#3b6ef0; color:#3b6ef0; }
  #preview { max-width:160px; max-height:160px; margin-top:12px; border-radius:8px;
             display:none; border:1px solid #eee; }
  .result { background:#fff; border:1px solid #e3e6ee; border-radius:12px;
            padding:18px; min-height:60px; }
  .badge { display:inline-block; padding:4px 12px; border-radius:20px;
           font-size:14px; font-weight:600; margin-bottom:10px; }
  .badge.white { background:#fff3cd; color:#8a6d00; }
  .badge.normal { background:#d4edda; color:#1a6b34; }
  .badge.error { background:#f8d7da; color:#842029; }
  pre { background:#0f172a; color:#e2e8f0; padding:14px; border-radius:8px;
        overflow:auto; font-size:12.5px; line-height:1.5; margin:0; }
  .meta { color:#666; font-size:13px; margin:8px 0; }
  .foot { margin-top:24px; font-size:12px; color:#aaa; text-align:center; }
  code { background:#eef1f7; padding:1px 5px; border-radius:4px; }
</style>
</head>
<body>
<div class="wrap">
  <h1>纯白图（白底图）识别</h1>
  <div class="sub">像素级检测 · 不依赖 AI 模型 / 不上传服务器 · 双击即可本地运行</div>

  <div class="tabs">
    <button data-tab="url" class="active" onclick="switchTab('url')">图片 URL</button>
    <button data-tab="file" onclick="switchTab('file')">上传文件</button>
    <button data-tab="b64" onclick="switchTab('b64')">Base64</button>
  </div>

  <div class="panel" id="panel-url">
    <div class="row">
      <input type="text" id="urlInput" placeholder="粘贴图片 URL，如 https://.../xxx.jpg">
      <button class="go" onclick="detectUrl()">检测</button>
    </div>
  </div>

  <div class="panel hidden" id="panel-file">
    <div class="filezone" onclick="document.getElementById('fileInput').click()">
      点击选择图片文件（jpg / png / webp 等）
      <img id="preview" alt="预览">
    </div>
    <input type="file" id="fileInput" accept="image/*" style="display:none"
           onchange="onFile(this)">
  </div>

  <div class="panel hidden" id="panel-b64">
    <textarea id="b64Input" placeholder="粘贴 Base64 图片数据（可带 data:image/...;base64, 前缀）"></textarea>
    <div style="margin-top:10px"><button class="go" onclick="detectB64()">检测</button></div>
  </div>

  <div class="result" id="result">
    <span style="color:#aaa">检测结果将显示在这里。</span>
  </div>

  <div class="foot">
    API：<code>POST /api/v1/detect/url</code> · <code>/api/v1/detect/base64</code> ·
    <code>/api/v1/detect/file</code> &nbsp;|&nbsp; 健康检查：<code>GET /api/v1/health</code>
  </div>
</div>

<script>
const $ = id => document.getElementById(id);
function switchTab(t){
  document.querySelectorAll('.tabs button').forEach(b=>b.classList.remove('active'));
  document.querySelector(`[data-tab="${t}"]`).classList.add('active');
  ['url','file','b64'].forEach(p=>$(`panel-${p}`).classList.add('hidden'));
  $(`panel-${t}`).classList.remove('hidden');
}
function render(res, thumbUrl){
  const box = $('result');
  if(!res.success){
    box.innerHTML = `<span class="badge error">检测失败</span>
      <div class="meta">${res.reason||res.error||'未知错误'}</div>`;
    return;
  }
  const cls = res.is_white_image ? 'white' : 'normal';
  const label = res.is_white_image ? '✅ 纯白图 / 白底图' : '🖼️ 正常图（含内容）';
  let extra = '';
  if(thumbUrl) extra = `<img src="${thumbUrl}" style="max-width:160px;max-height:160px;border-radius:8px;margin-bottom:10px;display:block">`;
  box.innerHTML = `${extra}<span class="badge ${cls}">${label}</span>
    <div class="meta">类型：<b>${res.category}</b> ｜ 近白占比：${(res.white_ratio*100).toFixed(1)}%
      ｜ 颜色种类：${res.color_variety} ｜ 主色 RGB(${res.dominant_color.join(',')})
      ｜ 尺寸：${res.image_size.join('×')} ｜ 格式：${res.format}
      ｜ 置信度：${res.confidence}</div>
    <div class="meta">${res.reason}</div>
    <pre>${JSON.stringify(res, null, 2)}</pre>`;
}
async function detectUrl(){
  const u = $('urlInput').value.trim(); if(!u){ alert('请输入图片 URL'); return; }
  const btn = event.target; btn.disabled = true;
  try{
    const r = await fetch('/api/v1/detect/url',{method:'POST',
      headers:{'Content-Type':'application/json'}, body:JSON.stringify({image_url:u})});
    render(await r.json(), u);
  }catch(e){ $('result').innerHTML = '<span class="badge error">请求失败：'+e+'</span>'; }
  finally{ btn.disabled = false; }
}
async function onFile(input){
  const f = input.files[0]; if(!f) return;
  const reader = new FileReader();
  reader.onload = async () => {
    $('preview').src = reader.result; $('preview').style.display='block';
    const b64 = reader.result.split(',')[1];
    const btn = document.querySelector('#panel-file .go') || {};
    try{
      const r = await fetch('/api/v1/detect/base64',{method:'POST',
        headers:{'Content-Type':'application/json'}, body:JSON.stringify({image_base64:b64})});
      render(await r.json(), reader.result);
    }catch(e){ $('result').innerHTML='<span class="badge error">请求失败：'+e+'</span>'; }
  };
  reader.readAsDataURL(f);
}
async function detectB64(){
  const b = $('b64Input').value.trim(); if(!b){ alert('请输入 Base64 数据'); return; }
  try{
    const r = await fetch('/api/v1/detect/base64',{method:'POST',
      headers:{'Content-Type':'application/json'}, body:JSON.stringify({image_base64:b})});
    render(await r.json());
  }catch(e){ $('result').innerHTML='<span class="badge error">请求失败：'+e+'</span>'; }
}
</script>
</body>
</html>"""


# ====================================================================
# HTTP 处理
# ====================================================================
class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        # 简化日志，避免刷屏
        sys.stdout.write("[%s] %s\n" % (self.address_string(), fmt % args))
        sys.stdout.flush()

    def _send(self, code: int, body: bytes, content_type: str = "application/json"):
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def _send_json(self, code: int, obj):
        self._send(code, json.dumps(obj, ensure_ascii=False).encode("utf-8"),
                   "application/json; charset=utf-8")

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path in ("/", "/index.html"):
            self._send(200, PAGE_HTML.encode("utf-8"), "text/html; charset=utf-8")
        elif parsed.path == "/api/v1/health":
            self._send_json(200, {"success": True, "status": "ok",
                                  "service": "white-image-detect", "version": "1.0.0"})
        else:
            self._send_json(404, {"success": False, "error": "not_found",
                                  "reason": f"路径不存在: {parsed.path}"})

    def do_POST(self):
        parsed = urlparse(self.path)
        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length) if length else b""

        try:
            if parsed.path == "/api/v1/detect/url":
                payload = json.loads(raw or b"{}")
                res = white_detect.detect_from_url(payload.get("image_url", ""))
            elif parsed.path == "/api/v1/detect/base64":
                payload = json.loads(raw or b"{}")
                res = white_detect.detect_from_base64(payload.get("image_base64", ""))
            elif parsed.path == "/api/v1/detect/file":
                res = self._handle_file(raw, self.headers.get("Content-Type", ""))
            else:
                return self._send_json(404, {"success": False, "error": "not_found",
                                             "reason": f"路径不存在: {parsed.path}"})
        except Exception as e:
            return self._send_json(500, {"success": False, "error": "server_error",
                                         "reason": f"服务端异常: {e}"})

        self._send_json(200, res)

    def _handle_file(self, raw: bytes, content_type: str):
        """用 email 解析 multipart，取第一个文件 part。"""
        try:
            msg = email.message_from_bytes(
                b"Content-Type: " + content_type.encode("utf-8") + b"\r\n\r\n" + raw,
                policy=email_policy.default,
            )
            for part in msg.walk():
                disp = part.get_content_disposition()
                if disp in ("attachment", "inline") or part.get_filename():
                    data = part.get_payload(decode=True)
                    if data:
                        return white_detect.detect_from_file(data)
            # 退化：尝试整个 body 直接当图片
            if raw:
                return white_detect.detect_from_file(raw)
        except Exception as e:
            return {"success": False, "error": "parse_failed",
                    "reason": f"multipart 解析失败: {e}"}
        return {"success": False, "error": "no_file",
                "reason": "未在请求体中找到图片文件"}


# ====================================================================
# 启动
# ====================================================================
def open_browser_later(url: str, delay: float = 1.2):
    def _open():
        try:
            webbrowser.open(url)
        except Exception:
            pass
    threading.Timer(delay, _open).start()


def main():
    args = sys.argv[1:]
    for i, a in enumerate(args):
        if a in ("--port", "-p") and i + 1 < len(args):
            global PORT
            PORT = int(args[i + 1])

    server = ThreadingHTTPServer((HOST, PORT), Handler)
    url = f"http://localhost:{PORT}/"
    print("=" * 52)
    print("  纯白图（白底图）识别 · 单机版")
    print("  服务地址:", url)
    print("  按 Ctrl+C 停止")
    print("=" * 52)
    open_browser_later(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n正在停止服务…")
        server.shutdown()
        sys.exit(0)


if __name__ == "__main__":
    main()
