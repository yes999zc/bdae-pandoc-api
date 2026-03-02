#!/bin/bash
# ════════════════════════════════════════════════════════
# pandoc-api 容器启动脚本
# 每次重启都必须用此脚本，以确保挂载目录生效
# ════════════════════════════════════════════════════════

# 停止并删除旧容器（忽略不存在的错误）
docker rm -f pandoc-api 2>/dev/null

# 启动新容器
# --network docker_default  : 加入 Dify 同一网络，容器间可互访
# -p 5050:5050              : 映射端口，Mac 本地可访问
# -v 挂载                   : 本地目录映射到容器，reference.docx 修改后即时生效
docker run -d --name pandoc-api \
  --network docker_default \
  -p 5050:5050 \
  -v ~/AI_Workspace/07_Pandoc-api:/app/templates \
  pandoc-api

# 推入最新 app.py（确保代码是最新版本）
sleep 2
docker cp ~/AI_Workspace/07_Pandoc-api/app.py pandoc-api:/app/app.py
docker restart pandoc-api

echo "✅ pandoc-api started"
echo "   健康检查: http://localhost:5050/health"
