# -*- coding: utf-8 -*-
"""
纯白图识别服务启动器 — 三种模式

  local    — FastAPI 本地服务（自带 /docs 调试页）
  public   — FastAPI + ngrok 公域隧道（可分享给任何人 / 接入 Coze Dify）

双击 start_public.bat 即可发布到公网。
"""
import subprocess
import sys
import time
import os
import webbrowser
from pathlib import Path

# 项目根目录
ROOT = Path(__file__).parent
# 优先使用受管 Python（已含 fastapi/uvicorn/numpy/PIL）；如需隔离可改为 venv/Scripts/python.exe
PYTHON = os.getenv(
    "WHITE_DETECT_PYTHON",
    r"C:\Users\Lenovo\.workbuddy\binaries\python\versions\3.13.12\python.exe",
)
SERVER = ROOT / "server.py"
NGROK_PATH = Path(r"C:\Users\Lenovo\AppData\Local\Microsoft\WindowsApps\ngrok.exe")
PORT = int(os.getenv("WHITE_DETECT_PORT", "7861"))


def wait_for_server(timeout=60):
    import urllib.request
    for _ in range(timeout):
        try:
            resp = urllib.request.urlopen(f"http://localhost:{PORT}/api/health", timeout=2)
            if resp.status == 200:
                return True
        except Exception:
            pass
        time.sleep(1)
    return False


def start_local():
    print("=" * 55)
    print("  🔲 纯白图识别服务 - 本地模式")
    print("=" * 55)
    print()
    print("[1/1] 启动 FastAPI 服务器...")
    proc = subprocess.Popen([PYTHON, str(SERVER)], cwd=str(ROOT))
    if wait_for_server():
        url = f"http://localhost:{PORT}"
        print()
        print("=" * 55)
        print("  ✅ 服务启动成功!")
        print(f"  🌐 本地地址: {url}")
        print("=" * 55)
        print()
        print("  ── 端点 ──")
        print(f"    POST /api/v1/detect          (文件上传)")
        print(f"    POST /api/v1/detect/url      (URL 输入)")
        print(f"    POST /api/v1/detect/base64   (Base64 输入)")
        print(f"    GET  /api/v1/guide           (接入指南)")
        print(f"    GET  /api/v1/health          (健康检查)")
        print(f"    📖 调试文档: {url}/docs")
        print()
        webbrowser.open(url)
        try:
            proc.wait()
        except KeyboardInterrupt:
            proc.terminate()
    else:
        print("  ❌ 服务器启动失败，请检查日志")


def start_public():
    print("=" * 55)
    print("  🔲 纯白图识别服务 - 公域模式（ngrok）")
    print("=" * 55)
    print()

    if not NGROK_PATH.exists():
        print("  ❌ 未找到 ngrok，请先安装: winget install ngrok.ngrok")
        return
    ngrok_cfg = Path(os.environ.get("LOCALAPPDATA", "")) / "ngrok" / "ngrok.yml"
    if not ngrok_cfg.exists():
        print("  ⚠️ ngrok 未配置 authtoken，请先运行: ngrok config add-authtoken <你的token>")
        return

    print("[1/3] 启动 FastAPI 服务器...")
    server_proc = subprocess.Popen([PYTHON, str(SERVER)], cwd=str(ROOT))
    if not wait_for_server():
        print("  ❌ 服务器启动失败")
        server_proc.terminate()
        return
    print(f"  ✅ 本地服务就绪: http://localhost:{PORT}")

    print("[2/3] 启动 ngrok 公域隧道...")
    ngrok_env = os.environ.copy()
    for k in ("http_proxy", "https_proxy", "HTTP_PROXY", "HTTPS_PROXY"):
        ngrok_env[k] = ""
    ngrok_proc = subprocess.Popen(
        [str(NGROK_PATH), "http", str(PORT)],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        env=ngrok_env,
    )

    print("[3/3] 获取公域链接...")
    import urllib.request, json
    public_url = None
    for _ in range(30):
        try:
            resp = urllib.request.urlopen("http://localhost:4040/api/tunnels", timeout=2)
            data = json.loads(resp.read())
            if data.get("tunnels"):
                public_url = data["tunnels"][0]["public_url"]
                break
        except Exception:
            pass
        time.sleep(1)

    if public_url:
        os.environ["PUBLIC_BASE_URL"] = public_url
        print()
        print("=" * 55)
        print("  ✅ 公域服务启动成功!")
        print("=" * 55)
        print()
        print(f"  🔗 公域地址: {public_url}")
        print("  (任何人可访问，可接入 Coze / Dify)")
        print()
        print("  ── 接入 ──")
        print(f"    📖 接入指南: {public_url}/api/v1/guide")
        print(f"    📋 OpenAPI:  {public_url}/api/v1/openapi.yaml")
        print(f"    ❤️  健康检查: {public_url}/api/v1/health")
        print()
        print("  按 Ctrl+C 停止服务")
        webbrowser.open(public_url)
        try:
            server_proc.wait()
        except KeyboardInterrupt:
            server_proc.terminate()
            ngrok_proc.terminate()
    else:
        print("  ❌ ngrok 隧道创建失败，本地地址仍可用")
        webbrowser.open(f"http://localhost:{PORT}")
        try:
            server_proc.wait()
        except KeyboardInterrupt:
            server_proc.terminate()
            ngrok_proc.terminate()


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "local"
    if mode == "public":
        start_public()
    else:
        start_local()
