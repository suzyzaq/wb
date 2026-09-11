# 空白图（白底图）批量检测工具

> White Image Detect —— 基于**像素分析**的纯白图 / 纯色空图识别。
> **零 token、零环境依赖、完全离线**：不调 AI 模型、不调 Dify、不需要 GPU。

在电商图片审核场景中，白底图 / 纯色空图无法展示商品，应被标记或退回。
本工具负责把这类图**批量识别出来**（与 RMBG-2.0 这类「白底图**生成**」工具互补）。

---

## 能力

| 判定 | 含义 | `is_white_image` |
|------|------|------------------|
| 纯白图 | 近白像素占比 ≥ 阈值（默认 99%） | `true` |
| 纯色图 | 颜色种类极少（≤3）且整体浅色 / 单色背景 | `true` |
| 正常图 | 含商品主体等真实内容 | `false` |

### 双维度不合规判定（可同时输出）

| 维度 | 规则 | 特点 |
|------|------|------|
| **B · 一张即不合规** | 该行**任意一张**图为空白 → 不合规 | 最严格，多抓 |
| **A · 全部空白才不合规** | 该行**所有图都空白** → 不合规 | 最宽松，少漏 |

一次扫描把两种维度**同时**写入 Excel 不同列，互不影响，便于按口径决策。

### 大表不阻塞

- **超过 1000 行自动分批**，每批完成即可下载该批结果
- **运行中定期落盘**（每 300 个 URL / 60 秒），中断不丢进度，支持 `--resume` 续跑
- **运行中可随时导出「已完成部分」**，不必等全部跑完

---

## 三种使用形态

### 1. 一键启动插件（推荐 · 零安装）

`WhiteImageExcel.exe` —— PyInstaller 单文件，内置 Python + 依赖，**双击即用**：

- 双击 → 自动打开浏览器 → 卡片式三步向导（选 Excel → 勾字段与维度 → 开始）
- 结果写回 Excel（「检测结果」+「汇总」两个工作表）
- 每批完成即可下载；全部完成合并全量；运行中还能随时导出已完成部分

> 打包：`python -m PyInstaller --onefile --noconsole --name WhiteImageExcel --hidden-import numpy --hidden-import PIL --hidden-import openpyxl --hidden-import requests white_excel_app.py`

### 2. 命令行批处理

```bash
# 自动识别 URL 列，双维度同时输出
python batch_excel.py 商品.xlsx

# 指定图片列 / 分批上限 / 断点续跑
python batch_excel.py 商品.xlsx --cols 主图,详情图 --chunk-size 1000 --resume

# 只输出单一维度 / 只看会识别到哪些列
python batch_excel.py 商品.xlsx --rule all
python batch_excel.py 商品.xlsx --list-cols

# 从进度文件导出「已完成部分」（不跑检测）
python batch_excel.py 商品.xlsx --export-partial 商品_空白图结果.xlsx.progress.json -o 部分.xlsx
```

### 3. HTTP 服务 / Coze·Dify 插件

```bash
python launcher.py local      # 本地服务，默认端口 7861
start_public.bat              # FastAPI + ngrok 公网隧道
```

端点为 Coze / Dify 可直接导入的 OpenAPI Schema，详见 [COZE_WORKFLOW_GUIDE.md](COZE_WORKFLOW_GUIDE.md)。

---

## 快速开始

```bash
# 1) 安装依赖（仅命令行 / 服务模式需要；exe 形态无需任何安装）
pip install -r requirements.txt

# 2) 命令行跑一份 Excel
python batch_excel.py 商品.xlsx

# 或 启动 HTTP 服务
python launcher.py local
```

---

## Excel 输出说明

**「检测结果」表**：原表所有列 + 每行追加

- 每个图片列：`{列名}·判定`（是/否）、`{列名}·纯白数`、`{列名}·明细`
- 行级双维度：`一张即不合规`（B）、`全部空白才不合规`（A）
- 行级汇总：`纯白图总数`、`图片总数`

**「汇总」表**：输入/输出文件、URL 列、判定维度、数据总行数、本批区间（分批时）、
两种维度的不合规行数、纯白图/正常图/失败张数、耗时、生成时间。

> 分批跑时还会产出 `{名}_批次i.xlsx`，全部完成后合并为 `{名}_全量.xlsx`。

### 双维度差异示例

| 行内容 | 图片总数 | 纯白数 | 一张即不合规 | 全部空白才不合规 |
|--------|----------|--------|--------------|------------------|
| 主图 + 详情图 都白 | 2 | 2 | 是 | 是 |
| 主图白 / 详情图正常 | 2 | 1 | 是 | 否 |
| 主图 + 详情图 都正常 | 2 | 0 | 否 | 否 |

---

## 目录结构

```
white-image-detect/
├── white_detect.py              # 识别核心（自包含，可单独 import）
├── batch_excel.py               # Excel 批量检测（分批 / 合并 / 中途导出 / 续跑）
├── white_excel_app.py           # 一键启动插件（HTML 向导 + 无界面双模式 + 打包入口）
├── app_exe.py                   # 单机版 HTTP 服务 exe 入口
├── server.py                    # FastAPI 服务
├── launcher.py                  # 三模式启动器（local / public）
├── openapi.yaml                 # Coze / Dify 可导入的 Schema
├── start_public.bat             # 一键发布公网（ngrok）
├── requirements.txt             # 依赖清单
├── server-package/              # 服务器部署包（Docker / systemd / 部署指南）
├── COZE_WORKFLOW_GUIDE.md       # Coze / Dify 工作流接入指南
└── 空白图检测Excel插件说明.md      # 插件完整使用说明
```

---

## HTTP API

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/api/v1/detect` | 文件上传（form-data: `file`） |
| POST | `/api/v1/detect/url` | URL 输入（JSON: `image_url`） |
| POST | `/api/v1/detect/base64` | Base64 输入（JSON: `image_base64`） |
| GET | `/api/v1/health` | 健康检查 |
| GET | `/api/v1/guide` | Coze / Dify 接入指南网页 |
| GET | `/api/v1/openapi.yaml` | OpenAPI Schema |

### 请求示例

```bash
curl -X POST http://localhost:7861/api/v1/detect/url \
  -H "Content-Type: application/json" \
  -d '{"image_url":"https://example.com/goods/xxx.jpg"}'
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

- `white_threshold`（默认 `0.99`）：近白像素占比阈值，≥ 此值判为纯白图
- `max_color_variety`（默认 `3`）：颜色种类上限，≤ 此值且整体浅色判为纯色空图

---

## 部署到服务器

`server-package/` 提供开箱即用的部署材料：

```bash
cd server-package
docker compose up -d          # 或参考 SERVER_DEPLOY_GUIDE.md 用 systemd
```

内含：`Dockerfile`、`docker-compose.yml`、`white-detect.service`、`deploy.sh`、部署指南。

---

## 复用核心模块

```python
from white_detect import detect_from_url, detect_from_base64, detect_from_file

r = detect_from_url("https://example.com/a.jpg")   # 也支持 file:///C:/path/a.jpg
print(r["is_white_image"], r["category"], r["reason"])

# 批量（含分批）
from batch_excel import run_batch, run_batch_chunked, export_from_progress
run_batch("表.xlsx", rules=("any", "all"))
```

---

## 输入要求

- 标准 `.xlsx` / `.xls`，首行为表头
- 图片地址列可**自定义**（界面勾选或 `--cols`）；未指定时自动识别含
  「图片 / 图 / url / 链接 / 主图 / 详情图 / image / img」的列
- 单元格支持多个 URL（逗号 / 换行 / 分号 / 空格分隔，同 URL 自动去重）
- 图片需可访问：`http(s)://` 走网络下载，`file:///C:/...` 直接读本地文件

---

## 许可

内部工具，按需使用。
