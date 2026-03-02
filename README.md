# ESA Report Pandoc 转换服务

将 Markdown 格式的 ESA（环境尽职调查）报告转换为 Word (.docx) 格式。

## 功能

- Pandoc Markdown → DOCX 转换
- 表格自动撑满页面宽度
- HIGH/MED/LOW 风险等级单元格染色
- 表头行黑底白字
- 可选章节自动编号

## 环境变量

| 变量 | 说明 | 默认值 |
|------|------|--------|
| `PUBLIC_BASE_URL` | 对外可访问的服务地址 | `http://localhost:5050` |

## API

### POST /convert

请求体：
```json
{
  "markdown": "# 报告标题\n\n内容...",
  "filename": "ESA_Report.docx",
  "number_sections": false
}
```

响应：
```json
{
  "success": true,
  "download_url": "http://localhost:5050/files/ESA_Report_abc123.docx",
  "filename": "ESA_Report_abc123.docx"
}
```

### GET /files/{filename}

下载生成的 DOCX 文件。

## Docker 运行

```bash
docker build -t esa-pandoc .
docker run -p 5050:5050 -e PUBLIC_BASE_URL=http://your-domain:5050 esa-pandoc
```

## 本地运行

```bash
pip install -r requirements.txt
python app.py
```

## 与 Dify 集成

Dify HTTP 节点配置：
- URL: `http://pandoc-api:5050/convert`
- Method: `POST`
- Body:
  ```json
  {
    "markdown": "{{#report_merge.final_report#}}",
    "filename": "{{#start.report_no#}}_ESA_Report.docx"
  }
  ```
