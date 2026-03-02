#!/bin/bash
# ════════════════════════════════════════════════════════
# 快速测试脚本：验证转换服务是否正常
# ════════════════════════════════════════════════════════

echo "正在测试转换服务..."

curl -s -X POST http://localhost:5050/convert \
  -H "Content-Type: application/json" \
  -d '{
    "markdown": "# 1. Executive Summary\n\n## 1.1 Overview\n\nThis is a test report generated to verify the conversion service.\n\n## 1.2 Risk Assessment Summary\n\n**Table 1-1 Risk Assessment Summary**\n\n| No | Risk | Category | Findings | Suggestion | Must Cost (USD) | Potential Risk Cost (USD) | Ranking |\n|---|---|---|---|---|---|---|---|\n| 1 | Soil Contamination | Environmental | PHCs detected | Further investigation | 50,000 | 500,000 | HIGH |\n| 2 | Groundwater | Environmental | Monitor required | Install wells | 20,000 | 200,000 | MED |\n| 3 | Minor spill | Environmental | Isolated area | Clean up | 5,000 | 10,000 | LOW |\n\n# 2. Site Description\n\n## 2.1 Property Overview\n\nSite located at test location.\n\n### 2.1.1 Current Conditions\n\nDetails of current site conditions.",
    "report_no": "ESA-TEST-001",
    "property_name": "BDAE Template Test",
    "filename": "ESA-TEST-001_Report.docx"
  }' | python3 -m json.tool

echo ""
echo "复制 download_url 到浏览器下载文件，检查样式是否正确"
