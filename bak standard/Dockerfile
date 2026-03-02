FROM python:3.11-slim

# 安装 pandoc + 字体
RUN apt-get update && apt-get install -y \
    pandoc \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

# 安装 Python 依赖
RUN pip install --no-cache-dir flask python-docx

WORKDIR /app

COPY app.py .
COPY create_reference.py .

# 启动时先生成 reference.docx，再启动 Flask
CMD python create_reference.py && python app.py
