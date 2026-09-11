# 纯白图识别 · Coze 工作流搭建指南

本指南教你把本地部署的 **纯白图识别插件（white-image-detect）** 接入 **Coze**，搭成一个可批量/单图判定「是否为白底图」的工作流，并挂到智能体里用自然语言触发。

> 判定含义：`is_white_image = true` → 纯白图 / 纯色空图（电商里应标记或退回）；`false` → 正常图（含商品主体）。

---

## 0. 前置条件（本机已满足）

| 项目 | 状态 | 说明 |
|------|------|------|
| white-image-detect 插件 | ✅ 已部署 | `white-image-detect/` 目录，FastAPI 端口 7861 |
| FastAPI / Uvicorn | ✅ 已就绪 | 受管 Python 3.13.12 自带 |
| ngrok | ✅ 已安装并配置 authtoken | `C:\Users\Lenovo\AppData\Local\Microsoft\WindowsApps\ngrok.exe` |
| Coze 账号 | 需自备 | coze.cn（国内）或 coze.com |

---

## 1. 拿到公网地址（必须，Coze 只能访问公网 HTTPS）

Coze 云端会**服务端去下载图片**，所以本地 `127.0.0.1` 不行，必须有一个公网 HTTPS 地址。

1. 打开 `white-image-detect/` 目录，**双击 `start_public.bat`**（或命令行 `python launcher.py public`）。
2. 终端会依次启动 FastAPI（本地 7861）和 ngrok 隧道，最后打印：

   ```
   🔗 公域地址: https://xxxx.ngrok-free.dev
   ```

3. 复制这个地址备用。

> ⚠️ **ngrok free 版每次重启地址都会变**。地址变了之后，Coze 里要么重新导入插件，要么改 HTTP 节点的 URL。若要稳定地址，见文末「进阶：部署到云拿固定域名」。

**本地自测（可选但建议）**：拿到地址后先在本地用 curl 验证一遍，确认 API 通：

```bash
curl -X POST https://你的ngrok地址/api/v1/detect/url \
  -H "Content-Type: application/json" \
  -d '{"image_url":"https://img.officemate.cn/aliyun-oss/goods/一张白图.jpg"}'
```

返回示例（含 `is_white_image`）：

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

---

## 2.（推荐）把 API 做成 Coze 插件（一次导入，多处复用）

做成插件后，工作流里直接用「插件节点」调用，不用每次手写 HTTP 配置。

1. **改 Schema 里的地址**：用记事本打开 `white-image-detect/openapi.yaml`，把第 28 行

   ```yaml
   - url: https://your-ngrok-url.ngrok-free.dev
   ```

   改成你第 1 步拿到的真实地址，例如：

   ```yaml
   - url: https://xxxx.ngrok-free.dev
   ```

2. 进入 **Coze 控制台 → 插件 → 创建插件 → 选择「导入 OpenAPI Schema」**，上传改好的 `openapi.yaml`。
3. 导入成功后，插件里会出现 4 个工具：
   - `detectByUrl` —— 传图片 URL（**工作流里优先用这个**）
   - `detectByFile` —— 上传文件（Coze 插件对 multipart 支持有限，尽量不用）
   - `detectByBase64` —— 传 Base64
   - `healthCheck` —— 健康检查
4. 在插件页点「测试」，用一张白图 URL 跑一次 `detectByUrl`，确认返回 `is_white_image=true`。

---

## 3.（最简）不导插件，直接用「HTTP 请求」节点

如果你只想在某个工作流里用一次，可以跳过第 2 步，直接在 Coze 工作流里加一个 HTTP 节点：

- **节点类型**：HTTP 请求
- **方法**：POST
- **URL**：`https://你的ngrok地址/api/v1/detect/url`
- **Headers**：`Content-Type: application/json`
- **Body（JSON）**：

  ```json
  {
    "image_url": "{{开始.image_url}}"
  }
  ```

后续节点用变量选择器点选该节点的响应即可（一般挂在节点的 `response` / 解析后的 JSON 字段上）。

> 两种方式的区别：插件方式可复用、可分享给团队；HTTP 节点方式零配置、改 URL 更灵活。新手推荐**先 HTTP 节点**跑通，再决定是否升级成插件。

---

## 4. 搭建工作流（单图判定）

目标：用户给一个图片 URL → 工作流判定是否为白底图 → 返回结论。

### 节点编排

| 顺序 | 节点 | 配置要点 |
|------|------|----------|
| 1 | **开始** | 添加一个输入参数 `image_url`（类型 String，必填） |
| 2 | **插件 / HTTP 节点** | 调用 `detectByUrl`（或 HTTP 节点 POST `/api/v1/detect/url`），入参 `image_url = {{开始.image_url}}` |
| 3 | **条件 / 代码节点** | 判断 `is_white_image` 是否为 `true` |
| 4 | **结束** | 输出「结论」「category」「white_ratio」「reason」 |

### 第 3 步「判断」节点写法（二选一）

**方式 A：条件节点（最简单）**
- 条件：`{{第2步节点.is_white_image}} == true`
- 分支「是」→ 接一个「消息/结束」节点，文本：`⚠️ 纯白图/白底图（{{category}}），近白占比 {{white_ratio}}，建议退回。原因：{{reason}}`
- 分支「否」→ 文本：`✅ 正常图（含商品主体），可上架。`

**方式 B：代码节点（更灵活，可做阈值/汇总）**
```python
# 输入：data["resp"] = 第2步节点的完整返回
r = data.get("resp", {})
if not r.get("success"):
    return {"结论": "检测失败", "详情": r.get("reason", "")}
if r.get("is_white_image"):
    return {"结论": "纯白图/白底图", "建议": "退回",
            "明细": f"{r.get('category')} | 近白{r.get('white_ratio')} | {r.get('reason')}"}
return {"结论": "正常图", "建议": "可上架", "明细": r.get("reason")}
```

### 结束节点输出示例

```
结论：纯白图/白底图
建议：退回
明细：纯白图 | 近白1.0 | 图片为纯色/接近纯色浅色图（近白像素占比 100.0%），判定为纯白图
```

---

## 5. 批量 / Excel 场景

### 多 URL 批量（推荐用「批处理」节点）

1. **开始**节点输入 `image_urls`（类型 Array of String）。
2. 加一个 **批处理节点**，对数组里每个 `url` 调用 `detectByUrl`。
3. 批处理内部可接第 4 节的条件判断。
4. **结束**节点汇总：纯白图数量、命中明细列表。

> 这样一次对话就能把一批商品主图全判完。

### Excel 批量

Coze 工作流**不直接读 xlsx**。两条路：

- **路线 1（推荐，本地 + 云端配合）**：在本机用 Python 脚本（核心 `white_detect.detect_batch`）把 Excel 里的 URL 列批量跑完、结果写回新 Excel。云端 Coze 工作流只负责「单图 API 服务」。
- **路线 2（全在 Coze）**：在「代码节点」里用 Python + `openpyxl` 解析上传的 Excel，把每行 URL 取出成数组，再进批处理节点。**受限**：Coze 代码沙箱能否装 `openpyxl` 看平台，不一定稳定。

> 结论：Excel 这种「重 IO / 大数据量」任务放本地脚本更稳；Coze 工作流负责「在线单图/小批量判定 + 对话入口」。

---

## 6. 挂到智能体里用

1. **Coze → 创建一个智能体**。
2. 在智能体的「工作流 / 技能」里，添加第 4 节搭好的工作流。
3. 写一句开场提示，例如：*「你是一个电商图片审核助手，当用户发来图片 URL 时，调用纯白图识别工作流判断是否为白底图，并给出退回/上架建议。」*
4. 发布智能体。之后用户直接发 `https://.../xxx.jpg` 就能得到判定。

---

## 7. 排错清单

| 现象 | 原因 | 解决 |
|------|------|------|
| Coze 调不通 / 超时 | ngrok 地址变了 | 重新 `start_public.bat` 拿新地址，改插件 URL 或 HTTP 节点 URL |
| `success=false, error=download_failed` | 图片 URL 不可公网访问 / 防盗链 | 换可公网访问的源图；或改用 `detectByBase64` 把图传进去 |
| officemate 图片 403 | 缺 UA/Referer | 本服务已带浏览器 UA + Referer，一般可过；仍 403 就换源图或本地上传 |
| 插件导入报错 | OpenAPI 版本/字段问题 | 确认 `openapi.yaml` 的 `servers.url` 已改成真实地址；Coze 优先 3.0，本文件已是 3.0 |
| 免费 ngrok 限额/断线 | free 版限制 | 生产环境部署到云（见下） |

---

## 8. 进阶：部署到云拿固定域名（生产建议）

ngrok free 不稳定，正式用建议部署一次、拿固定 HTTPS 域名：

1. 把 `white-image-detect/` 整个目录传到任意云服务器（需装 `fastapi uvicorn numpy Pillow requests`）。
2. 用 `uvicorn server:app --host 0.0.0.0 --port 7861` 常驻；前面用 **nginx / caddy** 反代并配 HTTPS 证书。
3. 把固定域名（如 `https://white.your-domain.com`）填进 `openapi.yaml` 的 `servers.url`。
4. Coze 里导入一次插件，永久可用，不再受 ngrok 重启影响。

---

## 附：本插件 API 速查

| 方法 | 路径 | 入参 | 适用 |
|------|------|------|------|
| POST | `/api/v1/detect/url` | `{image_url, white_threshold?, max_color_variety?}` | ✅ 工作流首选 |
| POST | `/api/detect/base64` | `{image_base64}` | 图已在本地的场景 |
| POST | `/api/v1/detect` | `file`（multipart） | Coze 插件支持有限，慎用 |
| GET | `/api/v1/health` | — | 探活 |

可调参数：`white_threshold`（默认 0.99，近白像素占比阈值）、`max_color_variety`（默认 3，颜色种类上限）。
