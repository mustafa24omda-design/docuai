"""
DocuAI - Intelligent Project Document Management System
=========================================================
نسخة مدمجة بالكامل في ملف واحد (app.py) — بدون الحاجة لمجلدات فرعية،
لتسهيل الرفع عبر واجهة GitHub الويب والنشر المباشر على Render.

يحتوي هذا الملف على:
1. محرك الاستخراج (Extraction Engine) بالـ Regex.
2. قارئات الملفات (PDF / DOCX / TXT).
3. واجهة ويب (Flask) لرفع المستندات ومعاينة/تنزيل النتائج.
"""

import os
import re
import socket
import tempfile
import threading
import webbrowser
from pathlib import Path
from datetime import datetime
from dataclasses import dataclass, asdict
from typing import Optional

# Data Analytics
try:
    import pandas as pd
except ImportError:
    pd = None



# =============================================================================
# 1) محرك الاستخراج (Extraction Engine)
# =============================================================================

DOC_NUMBER_PATTERNS = [
    r"(?:Document\s*(?:No\.?|Number|#)\s*[:\-]?\s*)([A-Z0-9][A-Z0-9\-\/\.]{3,30})",
    r"(?:Doc\.?\s*(?:No\.?|#)\s*[:\-]?\s*)([A-Z0-9][A-Z0-9\-\/\.]{3,30})",
    r"(?:Ref(?:erence)?\.?\s*(?:No\.?)?\s*[:\-]?\s*)([A-Z0-9][A-Z0-9\-\/\.]{3,30})",
    r"\b([A-Z]{2,6}-[A-Z0-9]{2,10}-\d{2,6}(?:-[A-Z0-9]{1,4})?)\b",
]

DOC_TYPE_KEYWORDS = {
    "Transmittal": ["transmittal", "خطاب إحالة", "إحالة مستند"],
    "RFI": ["request for information", "rfi", "طلب معلومات"],
    "Submittal": ["submittal", "material submittal", "shop drawing submittal"],
    "Shop Drawing": ["shop drawing", "مخطط تنفيذي"],
    "Method Statement": ["method statement", "منهجية تنفيذ"],
    "Inspection Request": ["inspection request", "ir ", "طلب معاينة"],
    "Non-Conformance Report": ["non-conformance", "ncr", "تقرير عدم مطابقة"],
    "Minutes of Meeting": ["minutes of meeting", "mom", "محضر اجتماع"],
    "Design Drawing": ["design drawing", "مخطط تصميمي"],
    "Correspondence": ["letter", "correspondence", "خطاب", "مراسلة"],
    "Progress Report": ["progress report", "تقرير سير العمل"],
    "Invoice": ["invoice", "فاتورة"],
}

DATE_PATTERNS = [
    r"\b(\d{4}[-/]\d{1,2}[-/]\d{1,2})\b",
    r"\b(\d{1,2}[-/]\d{1,2}[-/]\d{4})\b",
    r"\b(\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4})\b",
    r"\b([A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})\b",
]

DATE_FORMATS = [
    "%Y-%m-%d", "%Y/%m/%d",
    "%d-%m-%Y", "%d/%m/%Y",
    "%d %B %Y", "%d %b %Y",
    "%B %d, %Y", "%b %d, %Y", "%B %d %Y", "%b %d %Y",
]

REVISION_PATTERNS = [
    r"(?:Revision|Rev\.?)\s*[:\-]?\s*([A-Z]?\d{1,3}|[A-Z])\b",
    r"(?:مراجعة|مراجعه|النسخة)\s*[:\-]?\s*(\d{1,3})",
]

STATUS_KEYWORDS = {
    "Approved": ["approved", "معتمد"],
    "Approved as Noted": ["approved as noted", "معتمد مع ملاحظات"],
    "Rejected": ["rejected", "not approved", "مرفوض"],
    "Revise and Resubmit": ["revise and resubmit", "أعد التقديم"],
    "For Review": ["for review", "under review", "قيد المراجعة"],
    "For Approval": ["for approval", "للاعتماد"],
    "For Information": ["for information", "for your information", "للعلم"],
    "Closed": ["closed", "مغلق"],
    "Open": ["open", "مفتوح"],
    "Pending": ["pending", "معلق"],
}

SENDER_PATTERNS = [
    r"(?:From|Sender|Originator|Issued\s*By)\s*[:\-]\s*([A-Za-z0-9&.,\-\s]{2,60})(?:\n|$)",
    r"(?:من|المرسل|جهة الإصدار)\s*[:\-]\s*([\u0600-\u06FF0-9A-Za-z&.,\-\s]{2,60})(?:\n|$)",
]


@dataclass
class ExtractedDocument:
    source_file: str
    document_number: Optional[str] = None
    document_type: Optional[str] = None
    date: Optional[str] = None
    revision: Optional[str] = None
    status: Optional[str] = None
    sender: Optional[str] = None
    confidence: Optional[str] = None

    def to_dict(self):
        return asdict(self)


def _first_match(patterns, text, flags=re.IGNORECASE):
    for pat in patterns:
        m = re.search(pat, text, flags)
        if m:
            return m.group(1).strip().rstrip(".,;")
    return None


def _extract_document_type(text: str) -> Optional[str]:
    lowered = text.lower()
    for doc_type, keywords in DOC_TYPE_KEYWORDS.items():
        for kw in keywords:
            if kw.lower() in lowered:
                return doc_type
    return None


def _normalize_date(raw_date: str) -> Optional[str]:
    if not raw_date:
        return None
    cleaned = raw_date.strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(cleaned, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return cleaned


def _match_status_in(snippet: str) -> Optional[str]:
    lowered = snippet.lower()
    all_kw = [(status, kw) for status, kws in STATUS_KEYWORDS.items() for kw in kws]
    all_kw.sort(key=lambda x: len(x[1]), reverse=True)
    for status, kw in all_kw:
        if kw.lower() in lowered:
            return status
    return None


def _extract_status(text: str) -> Optional[str]:
    label_match = re.search(r"(?:Status|الحالة)\s*[:\-]\s*(.+)", text, re.IGNORECASE)
    if label_match:
        found = _match_status_in(label_match.group(1))
        if found:
            return found
    return _match_status_in(text)


def extract_fields(text: str, source_file: str = "") -> ExtractedDocument:
    doc_number = _first_match(DOC_NUMBER_PATTERNS, text)
    doc_type = _extract_document_type(text)
    raw_date = _first_match(DATE_PATTERNS, text)
    date = _normalize_date(raw_date) if raw_date else None
    revision = _first_match(REVISION_PATTERNS, text)
    status = _extract_status(text)
    sender = _first_match(SENDER_PATTERNS, text)

    found_count = sum(1 for v in [doc_number, doc_type, date, revision, status, sender] if v)
    if found_count >= 5:
        confidence = "High"
    elif found_count >= 3:
        confidence = "Medium"
    else:
        confidence = "Low"

    return ExtractedDocument(
        source_file=source_file,
        document_number=doc_number,
        document_type=doc_type,
        date=date,
        revision=revision,
        status=status,
        sender=sender,
        confidence=confidence,
    )


# =============================================================================
# 2) قراءة الملفات (PDF / DOCX / TXT)
# =============================================================================

def read_pdf(path: str) -> str:
    import pdfplumber
    text_parts = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text_parts.append(page.extract_text() or "")
    return "\n".join(text_parts)


def read_docx(path: str) -> str:
    import docx
    document = docx.Document(path)
    parts = [p.text for p in document.paragraphs]
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                parts.append(cell.text)
    return "\n".join(parts)


def read_txt(path: str) -> str:
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        return f.read()


READERS = {".pdf": read_pdf, ".docx": read_docx, ".txt": read_txt}


def read_any(path: str) -> str:
    ext = Path(path).suffix.lower()
    reader = READERS.get(ext)
    if reader is None:
        raise ValueError(f"صيغة الملف غير مدعومة: {ext}")
    return reader(path)


# =============================================================================
# 3) تصدير النتائج إلى Excel
# =============================================================================

def export_to_excel(results, output_path: str):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "DocuAI Results"

    headers = [
        "اسم الملف", "رقم المستند", "نوع المستند",
        "التاريخ", "Revision", "Status", "الجهة المرسلة", "دقة الاستخراج",
    ]
    ws.append(headers)

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for r in results:
        ws.append([
            r.source_file, r.document_number, r.document_type,
            r.date, r.revision, r.status, r.sender, r.confidence,
        ])

    for col_idx, header in enumerate(headers, start=1):
        max_len = max(
            [len(str(header))]
            + [len(str(ws.cell(row=row, column=col_idx).value or "")) for row in range(2, ws.max_row + 1)]
        )
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 45)

    ws.freeze_panes = "A2"
    wb.save(output_path)




# =============================================================================
# 4) طبقة تحليل البيانات وذكاء الأعمال
# =============================================================================

def documents_dataframe():
    """Convert processed documents to a pandas DataFrame."""
    if pd is None:
        return None
    rows = [r.to_dict() for r in DOCUMENTS]
    if not rows:
        return pd.DataFrame(columns=[
            "source_file", "document_number", "document_type",
            "date", "revision", "status", "sender", "confidence"
        ])
    df = pd.DataFrame(rows)
    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
    return df


def export_powerbi_dataset(output_path: str):
    """Create a clean CSV dataset ready for Power BI."""
    if pd is None:
        raise RuntimeError("pandas غير مثبتة. أضف pandas إلى requirements.txt")
    df = documents_dataframe()
    df = df.copy()
    if "date" in df.columns:
        df["date"] = df["date"].dt.strftime("%Y-%m-%d")
    df.to_csv(output_path, index=False, encoding="utf-8-sig")


def analyze_tabular_file(path: str):
    """
    Analyze business data files such as CSV/XLSX.
    Returns a compact analytical summary and DataFrame.
    """
    if pd is None:
        raise RuntimeError("pandas غير مثبتة")
    ext = Path(path).suffix.lower()
    if ext == ".csv":
        df = pd.read_csv(path)
    elif ext in {".xlsx", ".xls"}:
        df = pd.read_excel(path)
    else:
        raise ValueError("تحليل البيانات يدعم CSV و Excel حالياً")

    df.columns = [str(c).strip() for c in df.columns]

    numeric = df.select_dtypes(include="number").columns.tolist()
    date_cols = []
    for col in df.columns:
        if df[col].dtype == "object":
            converted = pd.to_datetime(df[col], errors="coerce")
            if converted.notna().mean() >= 0.7:
                df[col] = converted
                date_cols.append(col)
        elif pd.api.types.is_datetime64_any_dtype(df[col]):
            date_cols.append(col)

    summary = {
        "rows": int(len(df)),
        "columns": int(len(df.columns)),
        "numeric_columns": numeric,
        "date_columns": date_cols,
        "missing_cells": int(df.isna().sum().sum()),
    }

    # Common sales/business metrics when recognizable columns exist.
    lowered = {c.lower(): c for c in df.columns}
    sales_col = next(
        (lowered[k] for k in ["sales", "revenue", "amount", "total", "المبيعات", "الإيراد", "المبلغ"] if k in lowered),
        None,
    )
    quantity_col = next(
        (lowered[k] for k in ["quantity", "qty", "units", "الكمية", "الوحدات"] if k in lowered),
        None,
    )
    summary["sales_column"] = sales_col
    summary["quantity_column"] = quantity_col
    summary["total_sales"] = float(df[sales_col].sum()) if sales_col else None
    summary["total_quantity"] = float(df[quantity_col].sum()) if quantity_col else None

    return df, summary


DATASET_INFO = {"name": None, "summary": None, "preview": []}

# =============================================================================
# 4) واجهة الويب الاحترافية (Flask)
# =============================================================================

from flask import Flask, request, render_template_string, send_file, jsonify
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024

UPLOAD_DIR = Path(tempfile.gettempdir()) / "docuai_uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

LAST_RESULTS = []
DOCUMENTS = []


# ---------------------------------------------------------------------------
# UI helpers
# ---------------------------------------------------------------------------

ICON = {
    "dashboard": "▦",
    "documents": "▤",
    "upload": "↑",
    "reports": "▥",
    "settings": "⚙",
    "search": "⌕",
    "file": "▱",
    "check": "✓",
    "clock": "◷",
    "warning": "!",
    "cloud": "☁",
    "download": "↓",
    "menu": "☰",
}


ANALYTICS_TEMPLATE = r"""
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>DocuAI | Analytics & BI</title>
<style>
:root{--navy:#102a43;--blue:#1769aa;--bg:#f5f7fb;--card:#fff;--line:#e5eaf0;--muted:#68778a;--green:#198754}
*{box-sizing:border-box}body{margin:0;background:var(--bg);font-family:"Segoe UI",Tahoma,Arial;color:#172b4d}
.wrap{max-width:1350px;margin:auto;padding:28px}.top{display:flex;justify-content:space-between;align-items:center;margin-bottom:22px}
h1{margin:0;font-size:26px}.sub{color:var(--muted);font-size:13px;margin-top:6px}
.btn{background:var(--blue);color:#fff;text-decoration:none;border:0;border-radius:9px;padding:11px 16px;cursor:pointer;font-weight:600}
.btn.alt{background:#eaf3fb;color:var(--blue)}
.cards{display:grid;grid-template-columns:repeat(4,1fr);gap:15px}.card{background:#fff;border:1px solid var(--line);border-radius:14px;box-shadow:0 6px 22px rgba(16,42,67,.06)}
.stat{padding:19px}.label{font-size:12px;color:var(--muted)}.num{font-size:26px;font-weight:750;margin-top:6px}
.panel{padding:20px;margin-top:18px}.panel h3{margin:0 0 15px}
.upload{border:2px dashed #b9cce0;padding:30px;text-align:center;border-radius:12px;background:#fbfdff}
input[type=file]{margin:14px 0}.grid{display:grid;grid-template-columns:1.3fr .7fr;gap:18px}
.table-wrap{overflow:auto}.table{width:100%;border-collapse:collapse}.table th,.table td{padding:10px;border-bottom:1px solid #edf0f4;text-align:right;font-size:12px;white-space:nowrap}.table th{background:#f8fafc}
.pill{display:inline-block;padding:5px 10px;background:#eaf7f0;color:var(--green);border-radius:20px;font-size:11px;font-weight:700}
.metric{display:flex;justify-content:space-between;padding:12px 0;border-bottom:1px solid #edf0f4}.metric:last-child{border:0}
.note{background:#eef7ff;border-right:4px solid var(--blue);padding:13px;border-radius:8px;font-size:13px;line-height:1.7}
@media(max-width:900px){.cards{grid-template-columns:1fr 1fr}.grid{grid-template-columns:1fr}}
@media(max-width:600px){.wrap{padding:16px}.cards{grid-template-columns:1fr}.top{align-items:flex-start;gap:10px;flex-direction:column}}
</style>
</head>
<body>
<div class="wrap">
  <div class="top">
    <div><h1>📊 مركز تحليل البيانات وذكاء الأعمال</h1>
      <div class="sub">تحليل Excel وCSV وبناء مؤشرات تساعدك على اتخاذ القرار</div></div>
    <div>
      <a class="btn alt" href="/">← DocuAI</a>
      {% if powerbi_ready %}<a class="btn" href="/powerbi-dataset">⬇ Power BI Dataset</a>{% endif %}
    </div>
  </div>

  <div class="card panel">
    <h3>📥 استيراد بيانات المبيعات أو أي بيانات تشغيلية</h3>
    <form method="POST" enctype="multipart/form-data">
      <div class="upload">
        <strong>ارفع ملف Excel أو CSV</strong><br>
        <small style="color:#68778a">مثال: Sales, Revenue, Quantity, Date, Product, Region, Customer</small><br>
        <input type="file" name="data_file" accept=".csv,.xlsx,.xls" required>
        <br><button class="btn" type="submit">🤖 تحليل البيانات</button>
      </div>
    </form>
  </div>

  {% if dataset_name %}
  <div class="note" style="margin-top:18px">
    تم تحليل <strong>{{ dataset_name }}</strong>.
    النظام اكتشف {{ rows }} صف و{{ cols }} عمود و{{ missing }} خلية فارغة.
    <span class="pill">Dataset جاهز للتحليل</span>
  </div>

  <div class="cards" style="margin-top:18px">
    <div class="card stat"><div class="label">عدد الصفوف</div><div class="num">{{ rows }}</div></div>
    <div class="card stat"><div class="label">عدد الأعمدة</div><div class="num">{{ cols }}</div></div>
    <div class="card stat"><div class="label">إجمالي المبيعات</div><div class="num">{{ total_sales if total_sales is not none else '—' }}</div></div>
    <div class="card stat"><div class="label">إجمالي الكمية</div><div class="num">{{ total_quantity if total_quantity is not none else '—' }}</div></div>
  </div>

  <div class="grid">
    <div class="card panel">
      <h3>🔎 معاينة البيانات</h3>
      {% if preview %}
      <div class="table-wrap"><table class="table">
        <thead><tr>{% for c in columns %}<th>{{ c }}</th>{% endfor %}</tr></thead>
        <tbody>{% for row in preview %}<tr>{% for c in columns %}<td>{{ row[c] }}</td>{% endfor %}</tr>{% endfor %}</tbody>
      </table></div>
      {% endif %}
    </div>
    <div class="card panel">
      <h3>🧠 تحليل ذكي</h3>
      <div class="metric"><span>الأعمدة الرقمية</span><strong>{{ columns|length }}</strong></div>
      <div class="metric"><span>مبيعات مكتشفة</span><strong>{{ 'نعم' if total_sales is not none else 'لا' }}</strong></div>
      <div class="metric"><span>كمية مكتشفة</span><strong>{{ 'نعم' if total_quantity is not none else 'لا' }}</strong></div>
      <div class="metric"><span>حالة البيانات</span><span class="pill">جاهزة</span></div>
      <p class="sub">يمكن فتح ملف Power BI Dataset في Power BI Desktop وإنشاء الرسوم والمؤشرات والتقارير التفاعلية.</p>
    </div>
  </div>
  {% endif %}

  <div class="card panel">
    <h3>💼 سيناريوهات التحليل</h3>
    <div class="grid">
      <div>
        <div class="metric"><span>📈 المبيعات والإيرادات</span><strong>Trend / KPI</strong></div>
        <div class="metric"><span>🏆 أفضل المنتجات والعملاء</span><strong>Ranking</strong></div>
        <div class="metric"><span>🌍 المناطق والفروع</span><strong>Comparison</strong></div>
        <div class="metric"><span>📅 الأداء الشهري والسنوي</span><strong>Time Series</strong></div>
      </div>
      <div>
        <div class="note">النسخة الحالية تجهز البيانات والتحليل داخل DocuAI، مع تصدير Dataset نظيف إلى Power BI. يمكن في المرحلة التالية إضافة رسوم تفاعلية داخل الموقع وربط Power BI Service مباشرة.</div>
      </div>
    </div>
  </div>
</div>
</body>
</html>
"""

PAGE_TEMPLATE = r"""
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>DocuAI | Intelligent Document Management</title>
<style>
:root{
  --navy:#102a43; --blue:#1769aa; --blue2:#2f80c9;
  --bg:#f5f7fb; --card:#fff; --text:#172b4d; --muted:#6b778c;
  --line:#e6eaf0; --green:#1f9d68; --orange:#e69b25; --red:#d64545;
  --shadow:0 8px 28px rgba(16,42,67,.08);
}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--text);font-family:"Segoe UI",Tahoma,Arial,sans-serif}
button,input{font:inherit}
.layout{display:flex;min-height:100vh}
.sidebar{
  width:245px;background:linear-gradient(180deg,#0d253d,#123a5b);
  color:#fff;padding:22px 14px;position:fixed;right:0;top:0;bottom:0;
  z-index:20
}
.brand{display:flex;align-items:center;gap:12px;padding:8px 12px 25px}
.logo{
  width:44px;height:44px;border-radius:12px;background:#fff;color:var(--blue);
  display:grid;place-items:center;font-weight:800;font-size:22px;box-shadow:0 5px 18px rgba(0,0,0,.15)
}
.brand strong{font-size:20px;display:block}.brand small{opacity:.7}
.nav-title{font-size:11px;opacity:.55;padding:18px 13px 7px}
.nav a{
  color:#dbe8f4;text-decoration:none;display:flex;align-items:center;gap:12px;
  padding:12px 13px;border-radius:10px;margin:4px 0;font-size:14px
}
.nav a:hover,.nav a.active{background:rgba(255,255,255,.12);color:#fff}
.nav .ico{width:23px;text-align:center;font-size:18px}
.side-bottom{position:absolute;bottom:18px;right:14px;left:14px}
.main{margin-right:245px;width:calc(100% - 245px)}
.topbar{
  height:72px;background:#fff;border-bottom:1px solid var(--line);
  display:flex;align-items:center;justify-content:space-between;padding:0 30px;
  position:sticky;top:0;z-index:10
}
.search{width:min(430px,50%);position:relative}
.search input{
  width:100%;border:1px solid var(--line);background:#f8fafc;border-radius:10px;
  padding:11px 42px 11px 14px;outline:none
}
.search span{position:absolute;right:14px;top:9px;font-size:20px;color:var(--muted)}
.user{display:flex;align-items:center;gap:10px}.avatar{
  width:38px;height:38px;border-radius:50%;background:#e8f2fb;color:var(--blue);
  display:grid;place-items:center;font-weight:700
}
.content{padding:28px 30px 45px;max-width:1400px;margin:auto}
.page-head{display:flex;justify-content:space-between;align-items:center;gap:15px;margin-bottom:25px}
.page-head h1{margin:0;font-size:25px}.page-head p{margin:6px 0 0;color:var(--muted);font-size:13px}
.btn{
  border:0;border-radius:9px;padding:11px 18px;cursor:pointer;font-weight:600;
  background:var(--blue);color:#fff;display:inline-flex;gap:8px;align-items:center
}
.btn:hover{background:#12588e}.btn.secondary{background:#edf4fa;color:var(--blue)}
.cards{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:22px}
.card{background:var(--card);border:1px solid var(--line);border-radius:14px;box-shadow:var(--shadow)}
.stat{padding:20px;display:flex;justify-content:space-between;align-items:center}
.stat .num{font-size:28px;font-weight:750;margin-top:5px}.stat .label{color:var(--muted);font-size:13px}
.stat .ico-box{width:48px;height:48px;border-radius:12px;background:#edf5fc;color:var(--blue);display:grid;place-items:center;font-size:23px}
.grid2{display:grid;grid-template-columns:1.5fr 1fr;gap:18px}
.panel{padding:20px}.panel h3{margin:0 0 17px;font-size:17px}
.upload{
  border:2px dashed #b9cce0;border-radius:13px;padding:32px 20px;text-align:center;
  background:#fbfdff;transition:.2s
}
.upload.drag{border-color:var(--blue);background:#f1f8ff}
.upload .big{font-size:38px;color:var(--blue);margin-bottom:7px}
.upload p{color:var(--muted);margin:6px 0 16px;font-size:13px}
input[type=file]{display:none}
.drop-label{display:inline-flex}
.mini-list{display:flex;flex-direction:column;gap:10px}
.mini-row{display:flex;align-items:center;gap:10px;padding:11px;border-bottom:1px solid #f0f2f5}
.file-ico{width:37px;height:37px;border-radius:9px;background:#eef5fb;color:var(--blue);display:grid;place-items:center}
.mini-row .name{font-size:13px;font-weight:600}.mini-row small{color:var(--muted)}
.badge{padding:4px 9px;border-radius:20px;font-size:11px;font-weight:700;margin-right:auto}
.High{background:#e8f7ef;color:var(--green)}.Medium{background:#fff4dd;color:#a66b00}.Low{background:#fdeaea;color:var(--red)}
.table-wrap{overflow:auto}.table{width:100%;border-collapse:collapse}
.table th,.table td{padding:13px 11px;border-bottom:1px solid #eef1f5;text-align:right;white-space:nowrap;font-size:13px}
.table th{background:#f8fafc;color:#5c6b7c;font-weight:700}
.empty{text-align:center;padding:35px;color:var(--muted)}
.notice{padding:12px 15px;background:#eef7ff;border-right:4px solid var(--blue);border-radius:8px;margin-bottom:18px;font-size:13px}
.footer{text-align:center;color:#98a2b3;font-size:11px;margin-top:35px}
.mobile-menu{display:none;border:0;background:none;font-size:24px;color:var(--text)}
@media(max-width:1000px){
  .sidebar{width:215px}.main{margin-right:215px;width:calc(100% - 215px)}
  .cards{grid-template-columns:repeat(2,1fr)}.grid2{grid-template-columns:1fr}
}
@media(max-width:720px){
  .sidebar{transform:translateX(100%);transition:.25s}.sidebar.open{transform:translateX(0)}
  .main{margin-right:0;width:100%}.mobile-menu{display:block}
  .topbar{padding:0 16px}.search{display:none}.content{padding:20px 15px}
  .cards{grid-template-columns:1fr 1fr}.page-head{align-items:flex-start;flex-direction:column}
}
</style>
</head>
<body>

<aside class="sidebar" id="sidebar">
  <div class="brand">
    <div class="logo">D</div>
    <div><strong>DocuAI</strong><small>Smart Document System</small></div>
  </div>
  <div class="nav-title">MAIN MENU</div>
  <nav class="nav">
    <a href="/" class="active"><span class="ico">▦</span>لوحة التحكم</a>
    <a href="/documents"><span class="ico">▤</span>المستندات</a>
    <a href="/upload"><span class="ico">↑</span>رفع مستندات</a>
    <a href="/reports"><span class="ico">▥</span>التقارير</a>
    <a href="/analytics"><span class="ico">📊</span>تحليل البيانات BI</a>
  </nav>
  <div class="nav-title">SYSTEM</div>
  <nav class="nav"><a href="#"><span class="ico">⚙</span>الإعدادات</a></nav>
  <div class="side-bottom">
    <div style="font-size:11px;opacity:.6;padding:10px">DocuAI v2.0 • 2026</div>
  </div>
</aside>

<main class="main">
  <header class="topbar">
    <div style="display:flex;align-items:center;gap:12px">
      <button class="mobile-menu" onclick="document.getElementById('sidebar').classList.toggle('open')">☰</button>
      <div class="search"><span>⌕</span><input placeholder="ابحث عن مستند..." oninput="filterTable(this.value)"></div>
    </div>
    <div class="user"><div><strong style="font-size:13px">Document Control</strong><br><small style="color:#7b8794">Workspace</small></div><div class="avatar">DC</div></div>
  </header>

  <section class="content">
    <div class="page-head">
      <div><h1>لوحة التحكم</h1><p>مرحباً بك في DocuAI — نظام إدارة وتحليل المستندات الذكي</p></div>
      <a class="btn" href="/upload">＋ رفع مستند</a>
    </div>

    <div class="cards">
      <div class="card stat"><div><div class="label">إجمالي المستندات</div><div class="num">{{ total }}</div></div><div class="ico-box">▤</div></div>
      <div class="card stat"><div><div class="label">تم تحليلها</div><div class="num">{{ analyzed }}</div></div><div class="ico-box">✓</div></div>
      <div class="card stat"><div><div class="label">دقة عالية</div><div class="num">{{ high }}</div></div><div class="ico-box">◉</div></div>
      <div class="card stat"><div><div class="label">ملفات اليوم</div><div class="num">{{ today }}</div></div><div class="ico-box">◷</div></div>
    </div>

    <div class="grid2">
      <div class="card panel">
        <h3>رفع وتحليل المستندات</h3>
        <form method="POST" enctype="multipart/form-data" action="/upload">
          <div class="upload" id="drop">
            <div class="big">↑</div>
            <strong>اسحب الملفات هنا أو اختر من جهازك</strong>
            <p>PDF / DOCX / TXT — حتى 25 MB لكل طلب</p>
            <label class="btn drop-label">اختيار الملفات
              <input type="file" name="files" multiple required onchange="showNames(this)">
            </label>
            <div id="names" style="margin-top:12px;color:#6b778c;font-size:12px"></div>
          </div>
          <button class="btn" style="margin-top:13px;width:100%;justify-content:center" type="submit">🤖 تحليل المستندات</button>
        </form>
      </div>

      <div class="card panel">
        <h3>الأنواع المدعومة</h3>
        <div class="mini-list">
          {% for t in types %}
          <div class="mini-row"><div class="file-ico">▱</div><div><div class="name">{{ t }}</div><small>تصنيف تلقائي</small></div><span style="color:#2f80c9">✓</span></div>
          {% endfor %}
        </div>
      </div>
    </div>

    <div class="card panel" style="margin-top:18px">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:5px">
        <h3 style="margin:0">آخر المستندات</h3>
        <a href="/documents" class="btn secondary">عرض الكل</a>
      </div>
      {% if results %}
      <div class="table-wrap">
      <table class="table" id="docsTable">
        <thead><tr><th>الملف</th><th>رقم المستند</th><th>النوع</th><th>التاريخ</th><th>Revision</th><th>Status</th><th>الثقة</th></tr></thead>
        <tbody>
        {% for r in results %}
        <tr>
          <td>{{ r.source_file }}</td><td>{{ r.document_number or '—' }}</td><td>{{ r.document_type or '—' }}</td>
          <td>{{ r.date or '—' }}</td><td>{{ r.revision or '—' }}</td><td>{{ r.status or '—' }}</td>
          <td><span class="badge {{ r.confidence }}">{{ r.confidence }}</span></td>
        </tr>
        {% endfor %}
        </tbody>
      </table></div>
      {% else %}
        <div class="empty">لا توجد مستندات بعد — ابدأ برفع أول مستند.</div>
      {% endif %}
    </div>
    <div class="footer">DocuAI © 2026 — Intelligent Project Document Management & Business Intelligence</div>
  </section>
</main>

<script>
function showNames(input){
  document.getElementById('names').textContent = [...input.files].map(x=>x.name).join(' • ');
}
function filterTable(q){
  const rows=document.querySelectorAll('#docsTable tbody tr');
  rows.forEach(r=>r.style.display=r.innerText.toLowerCase().includes(q.toLowerCase())?'':'none');
}
const drop=document.getElementById('drop');
if(drop){
 ['dragenter','dragover'].forEach(e=>drop.addEventListener(e,()=>drop.classList.add('drag')));
 ['dragleave','drop'].forEach(e=>drop.addEventListener(e,()=>drop.classList.remove('drag')));
}
</script>
</body>
</html>
"""

UPLOAD_TEMPLATE = PAGE_TEMPLATE.replace(
    '<h1>لوحة التحكم</h1><p>مرحباً بك في DocuAI — نظام إدارة وتحليل المستندات الذكي</p>',
    '<h1>رفع المستندات</h1><p>ارفع ملفات المشروع ليتم استخراج بياناتها وتحليلها تلقائياً</p>'
)

DOCUMENTS_TEMPLATE = PAGE_TEMPLATE.replace(
    '<h1>لوحة التحكم</h1><p>مرحباً بك في DocuAI — نظام إدارة وتحليل المستندات الذكي</p>',
    '<h1>إدارة المستندات</h1><p>البحث واستعراض نتائج تحليل المستندات</p>'
)


def render_page(template=PAGE_TEMPLATE):
    results = DOCUMENTS if False else LAST_RESULTS
    high = sum(1 for r in DOCUMENTS if r.confidence == "High")
    return render_template_string(
        template,
        results=results[-20:][::-1],
        total=len(DOCUMENTS),
        analyzed=len(DOCUMENTS),
        high=high,
        today=len(DOCUMENTS),
        types=[
            "RFI", "Submittal", "Shop Drawing", "Method Statement",
            "Inspection Request", "NCR", "Transmittal", "Correspondence"
        ],
    )


def process_uploaded_files(files):
    global LAST_RESULTS, DOCUMENTS
    new_results = []
    for f in files:
        if not f or not f.filename:
            continue
        safe_name = secure_filename(f.filename)
        if not safe_name:
            continue
        save_path = UPLOAD_DIR / safe_name
        f.save(save_path)
        try:
            text = read_any(str(save_path))
            extracted = extract_fields(text, source_file=f.filename)
            new_results.append(extracted)
        except Exception as e:
            print(f"خطأ في معالجة {f.filename}: {e}")
    if new_results:
        LAST_RESULTS = new_results
        DOCUMENTS.extend(new_results)
    return new_results



@app.route("/analytics", methods=["GET", "POST"])
def analytics_dashboard():
    global DATASET_INFO

    if request.method == "POST":
        f = request.files.get("data_file")
        if f and f.filename:
            safe_name = secure_filename(f.filename)
            path = UPLOAD_DIR / safe_name
            f.save(path)
            try:
                df, summary = analyze_tabular_file(str(path))
                DATASET_INFO = {
                    "name": f.filename,
                    "summary": summary,
                    "preview": df.head(25).fillna("").to_dict("records"),
                    "columns": df.columns.tolist(),
                }
            except Exception as e:
                return f"تعذر تحليل الملف: {e}", 400

    s = DATASET_INFO.get("summary") or {}
    preview = DATASET_INFO.get("preview") or []
    columns = DATASET_INFO.get("columns") or []

    return render_template_string(ANALYTICS_TEMPLATE,
        dataset_name=DATASET_INFO.get("name"),
        rows=s.get("rows", 0),
        cols=s.get("columns", 0),
        missing=s.get("missing_cells", 0),
        total_sales=s.get("total_sales"),
        total_quantity=s.get("total_quantity"),
        columns=columns,
        preview=preview,
        powerbi_ready=bool(DATASET_INFO.get("name"))
    )


@app.route("/powerbi-dataset")
def powerbi_dataset():
    if not DOCUMENTS:
        return "لا توجد بيانات مستندات بعد.", 400
    output_path = UPLOAD_DIR / "docuai_powerbi_dataset.csv"
    export_powerbi_dataset(str(output_path))
    return send_file(output_path, as_attachment=True,
                     download_name="docuai_powerbi_dataset.csv")


@app.route("/", methods=["GET"])
def index():
    return render_page()


@app.route("/upload", methods=["GET", "POST"])
def upload():
    if request.method == "POST":
        process_uploaded_files(request.files.getlist("files"))
        return render_page()
    return render_page(UPLOAD_TEMPLATE)


@app.route("/documents", methods=["GET"])
def documents():
    return render_page(DOCUMENTS_TEMPLATE)


@app.route("/reports", methods=["GET"])
def reports():
    if not DOCUMENTS:
        return render_template_string(
            PAGE_TEMPLATE,
            results=[],
            total=0, analyzed=0, high=0, today=0,
            types=["RFI","Submittal","Shop Drawing","Method Statement",
                   "Inspection Request","NCR","Transmittal","Correspondence"]
        )
    return render_page(DOCUMENTS_TEMPLATE)


@app.route("/download")
def download():
    if not LAST_RESULTS:
        return "لا توجد نتائج بعد. الرجاء رفع مستندات أولاً.", 400
    output_path = UPLOAD_DIR / "docuai_results.xlsx"
    export_to_excel(LAST_RESULTS, str(output_path))
    return send_file(output_path, as_attachment=True,
                     download_name="docuai_results.xlsx")


def find_free_port(preferred=5000):
    for port in [preferred] + list(range(5001, 5011)):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", port)) != 0:
                return port
    return preferred


def open_browser(url):
    threading.Timer(1.2, lambda: webbrowser.open(url)).start()


if __name__ == "__main__":
    env_port = os.environ.get("PORT")
    is_cloud = env_port is not None
    port = int(env_port) if is_cloud else find_free_port(5000)
    host = "0.0.0.0" if is_cloud else "127.0.0.1"
    url = f"http://{host}:{port}"

    print("=" * 60)
    print("  DocuAI Professional UI جاهز للعمل ✅")
    if not is_cloud:
        print(f"  افتح المتصفح على: {url}")
        print("  لإيقاف البرنامج اضغط CTRL+C")
    print("=" * 60)

    if not is_cloud:
        open_browser(url)

    app.run(debug=False, host=host, port=port)
