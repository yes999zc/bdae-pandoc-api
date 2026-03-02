# ── 基础镜像 ─────────────────────────────────────────────────────────────────
FROM python:3.11-slim

# ── 系统依赖 ─────────────────────────────────────────────────────────────────
# pandoc      : Markdown → DOCX 核心转换工具
# fonts-liberation : 提供 Arial 等基础字体
RUN apt-get update && apt-get install -y \
    pandoc \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

# ── Python 依赖 ───────────────────────────────────────────────────────────────
# flask       : Web 框架，提供 HTTP 接口
# python-docx : DOCX 后处理（样式、颜色、表格）
RUN pip install --no-cache-dir flask python-docx

# ── 工作目录 ─────────────────────────────────────────────────────────────────
WORKDIR /app
COPY app.py .

# ── 启动服务 ─────────────────────────────────────────────────────────────────
CMD ["python", "app.py"]
