# 纯白图识别服务 · Docker 镜像
# 纯 CPU / 像素分析，无需 GPU、无需 AI 模型
FROM python:3.11-slim

LABEL maintainer="ops" \
      description="纯白图/空图识别 API（像素分析，无模型依赖）"

WORKDIR /app

# 先装依赖（利用层缓存）
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 再拷贝业务代码
COPY server.py white_detect.py openapi.yaml ./

# 运行时可通过环境变量覆盖
ENV WHITE_DETECT_PORT=7861
EXPOSE 7861

# 多 worker 提升并发（CPU 密集型以进程并行）
CMD ["sh", "-c", "uvicorn server:app --host 0.0.0.0 --port ${WHITE_DETECT_PORT} --workers 4"]
