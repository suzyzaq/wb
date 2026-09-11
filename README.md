# 纯白图识别云插件（White Image Detect Plugin）

基于**像素分析**的纯白图 / 空图识别服务，**不依赖任何 AI 模型、不调用 Dify、不需要 GPU**。
毫秒级响应，可部署到任意云服务器 / 容器，作为「插件」接入 Coze / Dify / 钉钉 AI 表格 等工作流。

> 与 RMBG-2.0（白底图**生成**）配套：本项目是白底图**识别**（检测图是否为空图/纯白图）。

---

## 能力

| 判定 | 含义 | `is_white_image` |
|------|------|------------------|
| 纯白图 | 近白像素 ≥ 阈值（默认 99%） | `true` |
| 纯色图 | 颜色种类极少（≤3）且整体浅色 / 单色背景 | `true` |
| 正常图 | 含商品主体等真实内容 | `false` |

在电商图片审核场景中，`is_white_image=true` 的图片应被标记 / 退回（白底图无法展示商品）。

---

## 目录结构

```
white-image-detect/
├── white_detect.py      # 识别核心（自包含，可单独 import 复用）
├── server.py            # FastAPI 服务
├── launcher.py          # 三模式启动器（local / public）
├── openapi.yaml         # Coze / Dify 可导入的 Schema
├── start_public.bat     # 一键发布公网（ngrok）
├── requirements.txt     # 依赖清单
└── README.md
```

---

## 本地运行

```bash
# 1. 启动本地服务（默认端口 7861）
python launcher.py local
# 或：python server.py

# 2. 浏览器打开
#    调试文档:  http://localhost:7861/docs
#    接入指南:  http://localhost:7861/api/v1/guide
#    健康检查:  http://localhost:7861/api/v1/health
```

---

## 部署到云端（公网）

```bash
# 一键启动 FastAPI + ngrok 公域隧道
start_public.bat
# 或：python launcher.py public
```

启动后终端会打印公域地址（如 `https://xxxx.ngrok-free.dev`），任何人可访问。

> 前置：ngrok 已安装且 `ngrok config add-authtoken <token>` 已配置（本机已就绪）。

---

## API 端点

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/api/v1/detect` | 文件上传（form-data: `file`） |
| POST | `/api/v1/detect/url` | URL 输入（JSON: `image_url`） |
| POST | `/api/v1/detect/base64` | Base64 输入（JSON: `image_base64`） |
| GET | `/api/v1/health` | 健康检查 |
| GET | `/api/v1/guide` | Coze / Dify 接入指南网页 |
| GET | `/api/v1/openapi.yaml` | OpenAPI Schema |

### 请求示例（URL 模式）

```bash
curl -X POST http://localhost:7861/api/v1/detect/url \
  -H "Content-Type: application/json" \
  -d '{"image_url":"https://img.officemate.cn/aliyun-oss/goods/xxx.jpg"}'
```

### 响应示例

```json
{
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
}
```

### 可调参数

- `white_threshold`（默认 `0.99`）：近白像素占比阈值，≥ 此值判为纯白图。
- `max_color_variety`（默认 `3`）：颜色种类上限，≤ 此值且整体浅色判为纯色空图。

---

## 接入 Coze / Dify

1. 启动公网隧道（`start_public.bat`）获得公域地址。
2. 将公域地址填入 `openapi.yaml` 的 `servers.url`。
3. **Coze**：插件 → 创建插件 → 导入 OpenAPI Schema → 选择本文件。
4. **Dify**：工具 → 自定义 → 创建 → 导入 OpenAPI Schema。

---

## 复用核心模块

`white_detect.py` 可脱离服务单独使用：

```python
from white_detect import detect_from_url, detect_from_base64, detect_from_file

r = detect_from_url("https://.../a.jpg")
print(r["is_white_image"], r["category"], r["reason"])
```

适合在 Excel 批处理脚本、现有审核流水线中直接 `import` 调用。

---

## 与现有图片审核项目的关系

- 原项目 `audit_images.py` 的 `pixel_is_blank_white()` 已抽离为本插件的 `white_detect.py`，
  逻辑保持一致（纯白图 ≥99% 近白 / 纯色空图判定）。
- 后续在监控台跑审核时，白底图识别即调用本插件的同一套像素逻辑，可本地库调用或走云端 API。
