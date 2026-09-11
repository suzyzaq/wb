# 纯白图识别服务 · 服务器部署指南（技术交接）

> 本服务用于**识别纯白图 / 纯色空图**（电商主图空拍、白底占位图等），
> 基于**像素分析**，**不依赖 AI 模型、不需要 GPU、不需要联网到第三方模型服务**，
> 仅依赖 Python 标准栈（FastAPI + numpy + Pillow），可部署到任意 Linux 服务器或容器。

---

## 一、交付物清单

```
white-image-detect/server-package/
├── server.py              # FastAPI 服务入口（已支持 WHITE_DETECT_PORT 环境变量）
├── white_detect.py        # 像素识别核心（纯函数，无外部依赖）
├── openapi.yaml           # OpenAPI Schema，供 Coze / Dify 插件导入
├── requirements.txt       # Python 依赖
├── Dockerfile             # 镜像构建
├── docker-compose.yml     # 容器编排（含健康检查）
├── deploy.sh              # 一键构建+启动（Docker）
├── white-detect.service   # systemd 单元（裸机 Python 部署用）
└── SERVER_DEPLOY_GUIDE.md # 本文档
```

> ⚠️ 本包**不含** Windows 版 exe、ngrok 隧道等开发/调试工具，仅生产所需文件。

---

## 二、两种部署方式（二选一）

### 方式 A：Docker 部署（推荐，最省心）

```bash
# 1) 上传整个 server-package 目录到服务器
# 2) 进入目录
cd server-package

# 3) 一键构建 + 启动（默认端口 7861）
./deploy.sh
# 或自定义端口：
PORT=9000 ./deploy.sh
```

启动后验证：
```bash
curl http://localhost:7861/api/v1/health
# {"status":"ok","service":"纯白图识别 API","version":"1.0.0",
#  "engine":"pixel-analysis (numpy + Pillow)","model_required":false}
```

### 方式 B：裸机 Python + systemd（无 Docker 环境）

```bash
# 1) 进入目录
cd /opt/white-detect

# 2) 建虚拟环境并装依赖
python3 -m venv venv
venv/bin/pip install -r requirements.txt

# 3) 注册开机自启
cp white-detect.service /etc/systemd/system/
systemctl daemon-reload
systemctl enable --now white-detect

# 4) 验证
curl http://localhost:7861/api/v1/health
```

---

## 三、配置项（环境变量）

| 变量 | 默认值 | 说明 |
|------|--------|------|
| `WHITE_DETECT_PORT` | `7861` | 服务监听端口 |
| `PUBLIC_BASE_URL` | 空 | 仅用于 `/api/v1/guide` 页面回显域名，不影响判定逻辑 |

> 端口同时受 `docker-compose.yml` 的 `ports` 与 `deploy.sh` 的 `PORT` 控制，三者保持一致即可。

---

## 四、API 端点

| 方法 | 路径 | 入参 | 说明 |
|------|------|------|------|
| POST | `/api/v1/detect` | form-data `file` | 上传图片文件 |
| POST | `/api/v1/detect/url` | JSON `{"image_url":"..."}` | 传图片 URL |
| POST | `/api/v1/detect/base64` | JSON `{"image_base64":"data:image/png;base64,..."}` | 传 Base64 |
| GET | `/api/v1/health` | — | 健康检查 |
| GET | `/api/v1/guide` | — | 接入指南网页 |
| GET | `/api/v1/openapi.yaml` | — | OpenAPI Schema |

### 关键响应字段

```json
{
  "success": true,
  "is_white_image": true,          // true = 纯白图/纯色空图（应退回/标记）
  "category": "纯白图",             // 纯白图 | 纯色空图 | 正常图
  "white_ratio": 1.0,              // 近白像素占比
  "color_variety": 1,              // 颜色种类数
  "dominant_color": [248,248,248], // 主色（RGB）
  "image_size": [800,800],
  "format": "JPEG",
  "confidence": 0.99,
  "reason": "图片为纯色/接近纯色浅色图...",
  "processing_time": 0.12
}
```

### 可调参数（每个检测接口都支持）

- `white_threshold`（默认 `0.99`）：近白像素占比阈值，≥ 此值判为纯白图。
- `max_color_variety`（默认 `3`）：颜色种类上限，≤ 此值且整体浅色判为纯色空图。

---

## 五、接入工作流（Coze / Dify）

1. 服务部署后，拿到服务器公网地址，例如 `https://white-detect.your-domain.com`。
2. 编辑 `openapi.yaml`，把 `servers.url` 改为上面的公网地址。
3. **Coze**：插件 → 创建插件 → 导入 OpenAPI Schema（选本文件）→ 得到 `detectByUrl` 等工具。
4. **Dify**：工具 → 自定义 API 工具 → 导入 OpenAPI Schema。
5. 工作流中调用 `detectByUrl`，根据 `is_white_image` 布尔值分支（true=退回）。

> 若服务前置了 Nginx 反向代理，注意放行 `/api/v1/` 路径与 POST 方法，并建议加 IP 白名单或鉴权。

---

## 六、运维要点

- **资源**：纯 CPU，单实例 4 worker 即可支撑日常批量；内存占用 < 200MB。
- **出网**：检测 URL 模式时需要服务器能访问图片源站（如 `img.officemate.cn`），请保证出网白名单/防火墙放行 HTTPS。
- **并发**：CPU 密集型，建议 worker 数 = CPU 核数；批量任务优先走客户端脚本并发调用。
- **监控**：用 `/api/v1/health` 做存活探针（Docker 已内置 healthcheck）。
- **升级**：替换 `server.py` / `white_detect.py` 后重建镜像或 `systemctl restart white-detect`。

---

## 七、本地 Excel 批量（离线重 IO 场景）

若需对大量 Excel 中的图片 URL 批量判定，建议在**本地**用配套脚本
`batch_excel.py`（见 WorkBuddy 技能 `white-image-detect`）跑完，
云端服务只承接在线单图/小批量判定与对话入口，避免大文件跨网上传。
