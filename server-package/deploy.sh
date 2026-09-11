#!/usr/bin/env bash
# 纯白图识别服务 · 一键部署脚本（Linux / Docker）
# 用法：
#   ./deploy.sh            # 构建并后台启动（端口 7861）
#   PORT=9000 ./deploy.sh  # 指定端口
set -e

PORT="${PORT:-7861}"
IMAGE="white-image-detect:1.0.0"

echo "==> 构建镜像 $IMAGE"
docker build -t "$IMAGE" .

echo "==> 启动容器（端口映射 $PORT:7861）"
docker rm -f white-detect 2>/dev/null || true
docker run -d --name white-detect --restart unless-stopped \
  -p "${PORT}:7861" \
  -e WHITE_DETECT_PORT=7861 \
  "$IMAGE"

echo "==> 等待健康检查"
for i in $(seq 1 20); do
  if curl -fsS "http://localhost:${PORT}/api/v1/health" >/dev/null 2>&1; then
    echo "✅ 服务已就绪：http://localhost:${PORT}"
    curl -fsS "http://localhost:${PORT}/api/v1/health"
    exit 0
  fi
  sleep 2
done
echo "❌ 健康检查失败，请查看 docker logs white-detect"
docker logs white-detect
exit 1
