"""
DocuAI - Web App (Flask)
========================
واجهة ويب بسيطة لرفع مستند واحد أو أكثر ومعاينة النتائج المستخرجة مباشرة
في المتصفح، مع إمكانية تنزيل النتائج كملف Excel.

التشغيل:
    pip install -r requirements.txt
    python app.py
    ثم افتح المتصفح على: http://127.0.0.1:5000
"""

import os
import tempfile
from pathlib import Path

from flask import Flask, request, render_template_string, send_file

from docuai.extractor import extract_fields
from docuai.readers import read_any
from docuai.cli import export_to_excel

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB

UPLOAD_DIR = Path(tempfile.gettempdir()) / "docuai_uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

LAST_RESULTS = []  # يُستخدم فقط لتفعيل زر "تنزيل Excel" بعد آخر معالجة (تجريبي)

PAGE_TEMPLATE = """
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<title>DocuAI - نظام إدارة مستندات المشاريع الذكي</title>
<style>
  body { font-family: 'Segoe UI', Tahoma, sans-serif; background:#f4f6f9; margin:0; padding:0; }
  header { background:#1f4e78; color:#fff; padding:24px 40px; }
  header h1 { margin:0; font-size:24px; }
  header p { margin:4px 0 0; opacity:.85; font-size:14px; }
  .container { max-width:1000px; margin:30px auto; padding:0 20px; }
  .card { background:#fff; border-radius:10px; padding:24px; box-shadow:0 2px 8px rgba(0,0,0,.06); margin-bottom:24px; }
  input[type=file] { display:block; margin:12px 0; }
  button { background:#1f4e78; color:#fff; border:none; padding:10px 22px; border-radius:6px; cursor:pointer; font-size:15px; }
  button:hover { background:#163a5c; }
  table { width:100%; border-collapse:collapse; margin-top:16px; }
  th, td { padding:10px 12px; border-bottom:1px solid #eee; text-align:right; font-size:14px; }
  th { background:#f0f3f7; color:#1f4e78; }
  .badge { padding:3px 10px; border-radius:12px; font-size:12px; color:#fff; }
  .High { background:#2e7d32; } .Medium { background:#f9a825; } .Low { background:#c62828; }
  footer { text-align:center; color:#888; font-size:12px; margin:30px 0; }
</style>
</head>
<body>
<header>
  <h1>📄 DocuAI — Intelligent Project Document Management System</h1>
  <p>استخراج تلقائي لرقم المستند، النوع، التاريخ، Revision، Status، والجهة المرسلة</p>
</header>
<div class="container">
  <div class="card">
    <form method="POST" enctype="multipart/form-data">
      <label>ارفع ملف أو أكثر (PDF / DOCX / TXT):</label>
      <input type="file" name="files" multiple required>
      <button type="submit">تحليل المستندات</button>
    </form>
  </div>

  {% if results %}
  <div class="card">
    <h3>نتائج التحليل ({{ results|length }} مستند)</h3>
    <table>
      <tr>
        <th>اسم الملف</th><th>رقم المستند</th><th>النوع</th><th>التاريخ</th>
        <th>Revision</th><th>Status</th><th>الجهة المرسلة</th><th>الثقة</th>
      </tr>
      {% for r in results %}
      <tr>
        <td>{{ r.source_file }}</td>
        <td>{{ r.document_number or '—' }}</td>
        <td>{{ r.document_type or '—' }}</td>
        <td>{{ r.date or '—' }}</td>
        <td>{{ r.revision or '—' }}</td>
        <td>{{ r.status or '—' }}</td>
        <td>{{ r.sender or '—' }}</td>
        <td><span class="badge {{ r.confidence }}">{{ r.confidence }}</span></td>
      </tr>
      {% endfor %}
    </table>
    <form method="GET" action="/download">
      <button type="submit">⬇️ تنزيل النتائج (Excel)</button>
    </form>
  </div>
  {% endif %}
</div>
<footer>DocuAI Prototype © 2026</footer>
</body>
</html>
"""


@app.route("/", methods=["GET", "POST"])
def index():
    global LAST_RESULTS
    results = []
    if request.method == "POST":
        files = request.files.getlist("files")
        for f in files:
            if not f.filename:
                continue
            save_path = UPLOAD_DIR / f.filename
            f.save(save_path)
            try:
                text = read_any(str(save_path))
                extracted = extract_fields(text, source_file=f.filename)
                results.append(extracted)
            except Exception as e:
                print(f"خطأ في معالجة {f.filename}: {e}")
        LAST_RESULTS = results

    return render_template_string(PAGE_TEMPLATE, results=results)


@app.route("/download")
def download():
    if not LAST_RESULTS:
        return "لا توجد نتائج بعد. الرجاء رفع مستندات أولاً.", 400
    output_path = UPLOAD_DIR / "docuai_results.xlsx"
    export_to_excel(LAST_RESULTS, str(output_path))
    return send_file(output_path, as_attachment=True, download_name="docuai_results.xlsx")


if __name__ == "__main__":
    app.run(debug=True, port=5000)
