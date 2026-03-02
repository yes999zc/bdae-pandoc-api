"""
ESA Report MD → DOCX with Page Breaks
在章节之间自动添加分页符
"""

from flask import Flask, request, jsonify, send_from_directory
import subprocess, tempfile, os, uuid, traceback

app = Flask(__name__)

OUTPUT_DIR = '/tmp/pandoc_outputs'
REFERENCE_DOCX = '/app/templates/reference.docx'
PUBLIC_BASE_URL = os.environ.get('PUBLIC_BASE_URL', 'http://localhost:5050').rstrip('/')

os.makedirs(OUTPUT_DIR, exist_ok=True)

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'reference_doc': os.path.exists(REFERENCE_DOCX)})

@app.route('/convert', methods=['POST'])
def convert():
    try:
        payload = request.get_json(force=True)
        md_content = payload.get('markdown', '')
        filename = payload.get('filename', 'ESA_Report.docx')
        number_sections = bool(payload.get('number_sections', False))

        if not md_content:
            return jsonify({'success': False, 'error': 'markdown field is empty'}), 400

        file_id = str(uuid.uuid4())[:8]
        safe_name = f"{filename.replace('.docx', '')}_{file_id}.docx"
        docx_path = os.path.join(OUTPUT_DIR, safe_name)

        # Step 1: 写临时 Markdown
        with tempfile.NamedTemporaryFile(suffix='.md', mode='w', encoding='utf-8', delete=False) as f:
            f.write(md_content)
            md_path = f.name

        # Step 2: Pandoc 转换
        cmd = [
            'pandoc',
            md_path,
            '-o', docx_path,
            '--from=markdown',
            '--to=docx',
        ]

        if os.path.exists(REFERENCE_DOCX):
            cmd += [f'--reference-doc={REFERENCE_DOCX}']

        if number_sections:
            cmd.append('--number-sections')

        result = subprocess.run(cmd, capture_output=True, text=True)
        os.unlink(md_path)

        if result.returncode != 0:
            return jsonify({
                'success': False,
                'error': f'Pandoc error: {result.stderr}'
            }), 500

        # Step 3: 添加分页符和样式处理
        add_page_breaks(docx_path)

        download_url = f'{PUBLIC_BASE_URL}/files/{safe_name}'
        return jsonify({
            'success': True,
            'download_url': download_url,
            'filename': safe_name,
        })

    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e),
            'trace': traceback.format_exc()
        }), 500

@app.route('/files/<path:filename>')
def download(filename):
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)

def add_page_breaks(docx_path: str):
    """在章节标题前添加分页符"""
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    
    doc = Document(docx_path)
    
    # 遍历所有段落，在 Heading 1 前添加分页符（除了第一个）
    first_heading = True
    for para in doc.paragraphs:
        if para.style.name.startswith('Heading 1'):
            if not first_heading:
                # 在此段落前添加分页符
                pPr = para._p.get_or_add_pPr()
                pageBreakBefore = OxmlElement('w:pageBreakBefore')
                pageBreakBefore.set(qn('w:val'), 'on')
                pPr.append(pageBreakBefore)
            else:
                first_heading = False
    
    doc.save(docx_path)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5050, debug=False)
