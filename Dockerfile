FROM python:3.11-slim

# ── 系统依赖：pandoc + 字体（含中文）──────────────────────────────────────
RUN apt-get update && apt-get install -y --no-install-recommends \
    pandoc \
    fonts-liberation \
    fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

# ── Python 依赖 ────────────────────────────────────────────────────────────
RUN pip install --no-cache-dir flask python-docx

# ── 工作目录 ───────────────────────────────────────────────────────────────
WORKDIR /app

COPY app.py .
COPY create_reference.py .

# ── 启动：先生成 reference.docx，再启动 Flask ─────────────────────────────
CMD python app.py
