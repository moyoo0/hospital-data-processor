#!/bin/bash

# 配置
IMAGE_NAME="moyoo0/hospital-data-app"
TAG="v2"

echo "开始构建并推送 Docker 镜像: ${IMAGE_NAME}:${TAG}..."

# 使用 buildx 构建 AMD64 架构镜像并推送
# 这将确保镜像可以在标准的 Linux 服务器上运行
docker buildx build --platform linux/amd64 -t "${IMAGE_NAME}:${TAG}" --push .

if [ $? -eq 0 ]; then
    echo "=================================================="
    echo "构建并推送成功！"
    echo "服务器部署命令:"
    echo "docker pull ${IMAGE_NAME}:${TAG}"
    echo "docker stop hospital-app || true"
    echo "docker rm hospital-app || true"
    echo "docker run -d -p 5010:5010 --name hospital-app --restart always ${IMAGE_NAME}:${TAG}"
    echo "=================================================="
else
    echo "构建失败，请检查 Docker 环境或登录状态。"
    exit 1
fi
