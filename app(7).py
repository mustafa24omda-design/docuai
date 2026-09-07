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

from flask import Flask, request, render_template_string, send_file, jsonify

# OCR is loaded lazily so the app can still start when OCR dependencies are missing.


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
    """
    Read a PDF. First try normal text extraction. If a page has little/no
    selectable text, OCR the page using Tesseract.

    OCR requirements:
      pip install pytesseract pdf2image pillow
    Also install the Tesseract OCR engine on the operating system.
    """
    import pdfplumber

    text_parts = []
    pages_needing_ocr = []

    # 1) Fast path: extract selectable PDF text.
    with pdfplumber.open(path) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            extracted = (page.extract_text() or "").strip()
            text_parts.append(extracted)
            if len(extracted) < 20:
                pages_needing_ocr.append(page_no)

    # 2) OCR only pages that appear scanned/image-only.
    if pages_needing_ocr:
        try:
            import pytesseract
            from pdf2image import convert_from_path
        except ImportError as exc:
            raise RuntimeError(
                "هذا الـPDF يحتوي صفحات ممسوحة ضوئياً وتحتاج OCR. "
                "ثبّت: pip install pytesseract pdf2image pillow"
            ) from exc

        try:
            images = convert_from_path(
                path,
                dpi=250,
                first_page=min(pages_needing_ocr),
                last_page=max(pages_needing_ocr)
            )
        except Exception as exc:
            raise RuntimeError(
                "تعذر تحويل صفحات PDF إلى صور. تأكد من تثبيت Poppler "
                "وإضافته إلى PATH في Windows."
            ) from exc

        for index, page_no in enumerate(range(min(pages_needing_ocr), max(pages_needing_ocr) + 1)):
            if page_no not in pages_needing_ocr:
                continue

            image = images[index]
            try:
                # Arabic + English OCR. If ara is unavailable, fall back to English.
                langs = "eng+ara"
                try:
                    ocr_text = pytesseract.image_to_string(image, lang=langs)
                except Exception:
                    ocr_text = pytesseract.image_to_string(image, lang="eng")

                text_parts[page_no - 1] = ocr_text.strip()
            except Exception as exc:
                raise RuntimeError(
                    f"فشل OCR في الصفحة {page_no}: {exc}"
                ) from exc

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
# MD DocuAI Professional — Single-file Project Document Management Platform
# V2/V3/V4 foundation: Login, Dashboard, Documents, OCR, Projects, Workflow,
# Transmittals, RFIs, Submittals, BI, Notifications, Audit Trail, Multi-Company,
# Cloud-ready storage and REST API.
# =============================================================================

import json
import secrets
import sqlite3
from functools import wraps
from flask import (
    Flask, request, render_template_string, redirect, url_for,
    session, send_file, jsonify, flash
)
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("DOCUAI_SECRET_KEY", "change-this-secret-key")
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024

BASE_DIR = Path(os.environ.get("DOCUAI_DATA_DIR", "./data"))
STORAGE_DIR = Path(os.environ.get("DOCUAI_STORAGE_DIR", str(BASE_DIR / "storage")))
DB_PATH = Path(os.environ.get("DOCUAI_DB_PATH", str(BASE_DIR / "docuai.db")))
BASE_DIR.mkdir(parents=True, exist_ok=True)
STORAGE_DIR.mkdir(parents=True, exist_ok=True)

ALLOWED_EXTENSIONS = {".pdf", ".docx", ".txt"}

# ---------- Database ----------

def db():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys = ON")
    return con

def init_db():
    con = db()
    con.executescript("""
    CREATE TABLE IF NOT EXISTS companies (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        code TEXT NOT NULL UNIQUE,
        created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER,
        username TEXT NOT NULL UNIQUE,
        password_hash TEXT NOT NULL,
        full_name TEXT NOT NULL,
        role TEXT NOT NULL DEFAULT 'Viewer',
        active INTEGER NOT NULL DEFAULT 1,
        created_at TEXT NOT NULL,
        FOREIGN KEY(company_id) REFERENCES companies(id)
    );

    CREATE TABLE IF NOT EXISTS projects (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL,
        name TEXT NOT NULL,
        code TEXT NOT NULL,
        client TEXT,
        status TEXT NOT NULL DEFAULT 'Active',
        created_at TEXT NOT NULL,
        UNIQUE(company_id, code),
        FOREIGN KEY(company_id) REFERENCES companies(id)
    );

    CREATE TABLE IF NOT EXISTS documents (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL,
        project_id INTEGER,
        uid TEXT NOT NULL UNIQUE,
        filename TEXT NOT NULL,
        stored_name TEXT NOT NULL,
        document_number TEXT,
        document_type TEXT,
        issue_date TEXT,
        revision TEXT,
        status TEXT NOT NULL DEFAULT 'Uploaded',
        sender TEXT,
        confidence TEXT,
        ocr_used INTEGER NOT NULL DEFAULT 0,
        workflow_status TEXT NOT NULL DEFAULT 'Draft',
        file_size INTEGER DEFAULT 0,
        uploaded_by INTEGER,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        FOREIGN KEY(company_id) REFERENCES companies(id),
        FOREIGN KEY(project_id) REFERENCES projects(id),
        FOREIGN KEY(uploaded_by) REFERENCES users(id)
    );

    CREATE TABLE IF NOT EXISTS transmittals (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL,
        project_id INTEGER,
        number TEXT NOT NULL,
        subject TEXT NOT NULL,
        to_company TEXT,
        status TEXT NOT NULL DEFAULT 'Draft',
        created_by INTEGER,
        created_at TEXT NOT NULL,
        UNIQUE(company_id, number),
        FOREIGN KEY(company_id) REFERENCES companies(id),
        FOREIGN KEY(project_id) REFERENCES projects(id)
    );

    CREATE TABLE IF NOT EXISTS rfis (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL,
        project_id INTEGER,
        number TEXT NOT NULL,
        subject TEXT NOT NULL,
        question TEXT,
        status TEXT NOT NULL DEFAULT 'Open',
        priority TEXT NOT NULL DEFAULT 'Normal',
        due_date TEXT,
        created_by INTEGER,
        created_at TEXT NOT NULL,
        UNIQUE(company_id, number),
        FOREIGN KEY(company_id) REFERENCES companies(id),
        FOREIGN KEY(project_id) REFERENCES projects(id)
    );

    CREATE TABLE IF NOT EXISTS submittals (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL,
        project_id INTEGER,
        number TEXT NOT NULL,
        title TEXT NOT NULL,
        type TEXT,
        status TEXT NOT NULL DEFAULT 'Submitted',
        revision TEXT,
        created_by INTEGER,
        created_at TEXT NOT NULL,
        UNIQUE(company_id, number),
        FOREIGN KEY(company_id) REFERENCES companies(id),
        FOREIGN KEY(project_id) REFERENCES projects(id)
    );

    CREATE TABLE IF NOT EXISTS workflow (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        document_id INTEGER NOT NULL,
        action TEXT NOT NULL,
        from_status TEXT,
        to_status TEXT,
        comment TEXT,
        user_id INTEGER,
        created_at TEXT NOT NULL,
        FOREIGN KEY(document_id) REFERENCES documents(id)
    );

    CREATE TABLE IF NOT EXISTS notifications (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER NOT NULL,
        title TEXT NOT NULL,
        message TEXT NOT NULL,
        is_read INTEGER NOT NULL DEFAULT 0,
        created_at TEXT NOT NULL,
        FOREIGN KEY(user_id) REFERENCES users(id)
    );

    CREATE TABLE IF NOT EXISTS audit_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER,
        action TEXT NOT NULL,
        entity_type TEXT,
        entity_id INTEGER,
        details TEXT,
        created_at TEXT NOT NULL,
        FOREIGN KEY(user_id) REFERENCES users(id)
    );

    CREATE INDEX IF NOT EXISTS idx_documents_company ON documents(company_id);
    CREATE INDEX IF NOT EXISTS idx_documents_project ON documents(project_id);
    CREATE INDEX IF NOT EXISTS idx_documents_status ON documents(status);
    CREATE INDEX IF NOT EXISTS idx_documents_type ON documents(document_type);
    CREATE INDEX IF NOT EXISTS idx_documents_number ON documents(document_number);
    """)
    if con.execute("SELECT COUNT(*) FROM companies").fetchone()[0] == 0:
        con.execute(
            "INSERT INTO companies(name,code,created_at) VALUES(?,?,?)",
            ("MD Demo Company", "MD", now())
        )
    if con.execute("SELECT COUNT(*) FROM users").fetchone()[0] == 0:
        company_id = con.execute("SELECT id FROM companies ORDER BY id LIMIT 1").fetchone()["id"]
        con.execute(
            """INSERT INTO users(company_id,username,password_hash,full_name,role,created_at)
               VALUES(?,?,?,?,?,?)""",
            (company_id, "admin", generate_password_hash("Admin@123"),
             "MD Administrator", "Administrator", now())
        )

    # ---------------- Advanced platform tables (V5-V18 foundation) ----------------
    con.executescript("""
    CREATE TABLE IF NOT EXISTS document_revisions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        document_id INTEGER NOT NULL,
        revision TEXT NOT NULL,
        file_uid TEXT NOT NULL,
        status TEXT NOT NULL DEFAULT 'Draft',
        change_summary TEXT,
        created_by INTEGER,
        created_at TEXT NOT NULL,
        FOREIGN KEY(document_id) REFERENCES documents(id)
    );
    CREATE TABLE IF NOT EXISTS document_comments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        document_id INTEGER NOT NULL,
        user_id INTEGER,
        comment TEXT NOT NULL,
        created_at TEXT NOT NULL,
        FOREIGN KEY(document_id) REFERENCES documents(id)
    );
    CREATE TABLE IF NOT EXISTS workflow_tasks (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        document_id INTEGER NOT NULL,
        assigned_to INTEGER,
        stage TEXT NOT NULL,
        due_date TEXT,
        status TEXT NOT NULL DEFAULT 'Pending',
        decision TEXT,
        created_at TEXT NOT NULL,
        completed_at TEXT,
        FOREIGN KEY(document_id) REFERENCES documents(id)
    );
    CREATE TABLE IF NOT EXISTS transmittal_documents (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        transmittal_id INTEGER NOT NULL,
        document_id INTEGER NOT NULL,
        FOREIGN KEY(transmittal_id) REFERENCES transmittals(id),
        FOREIGN KEY(document_id) REFERENCES documents(id)
    );
    CREATE TABLE IF NOT EXISTS rfi_responses (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        rfi_id INTEGER NOT NULL,
        user_id INTEGER,
        response TEXT NOT NULL,
        created_at TEXT NOT NULL,
        FOREIGN KEY(rfi_id) REFERENCES rfis(id)
    );
    CREATE TABLE IF NOT EXISTS submittal_reviews (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        submittal_id INTEGER NOT NULL,
        user_id INTEGER,
        decision TEXT NOT NULL,
        comments TEXT,
        created_at TEXT NOT NULL,
        FOREIGN KEY(submittal_id) REFERENCES submittals(id)
    );
    CREATE TABLE IF NOT EXISTS internal_messages (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL,
        sender_id INTEGER NOT NULL,
        receiver_id INTEGER NOT NULL,
        subject TEXT NOT NULL,
        body TEXT NOT NULL,
        is_read INTEGER NOT NULL DEFAULT 0,
        created_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS signatures (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        document_id INTEGER NOT NULL,
        user_id INTEGER NOT NULL,
        signature_type TEXT NOT NULL DEFAULT 'Electronic Approval',
        signed_at TEXT NOT NULL,
        note TEXT
    );
    CREATE TABLE IF NOT EXISTS saved_searches (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER NOT NULL,
        name TEXT NOT NULL,
        query TEXT NOT NULL,
        created_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS ai_analysis (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        document_id INTEGER NOT NULL,
        summary TEXT,
        key_points TEXT,
        required_actions TEXT,
        risks TEXT,
        keywords TEXT,
        model TEXT,
        created_at TEXT NOT NULL,
        FOREIGN KEY(document_id) REFERENCES documents(id)
    );
    CREATE TABLE IF NOT EXISTS system_settings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL,
        setting_key TEXT NOT NULL,
        setting_value TEXT,
        UNIQUE(company_id, setting_key)
    );
    CREATE INDEX IF NOT EXISTS idx_revisions_document ON document_revisions(document_id);
    CREATE INDEX IF NOT EXISTS idx_comments_document ON document_comments(document_id);
    CREATE INDEX IF NOT EXISTS idx_tasks_document ON workflow_tasks(document_id);
    CREATE INDEX IF NOT EXISTS idx_ai_document ON ai_analysis(document_id);
    """)
    con.commit()
    con.close()

def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def current_user():
    uid = session.get("user_id")
    if not uid:
        return None
    con = db()
    row = con.execute("""
        SELECT u.*, c.name AS company_name, c.code AS company_code
        FROM users u LEFT JOIN companies c ON c.id=u.company_id
        WHERE u.id=? AND u.active=1
    """, (uid,)).fetchone()
    con.close()
    return row

def audit(action, entity_type="", entity_id=None, details=""):
    u = current_user()
    con = db()
    con.execute(
        "INSERT INTO audit_log(user_id,action,entity_type,entity_id,details,created_at) VALUES(?,?,?,?,?,?)",
        (u["id"] if u else None, action, entity_type, entity_id, details, now())
    )
    con.commit()
    con.close()

def notify(user_id, title, message):
    con = db()
    con.execute(
        "INSERT INTO notifications(user_id,title,message,created_at) VALUES(?,?,?,?)",
        (user_id, title, message, now())
    )
    con.commit()
    con.close()

def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not current_user():
            return redirect(url_for("login"))
        return view(*args, **kwargs)
    return wrapped

def role_required(*roles):
    def deco(view):
        @wraps(view)
        def wrapped(*args, **kwargs):
            u = current_user()
            if not u:
                return redirect(url_for("login"))
            if u["role"] not in roles:
                flash("ليس لديك صلاحية لتنفيذ هذا الإجراء.", "error")
                return redirect(url_for("dashboard"))
            return view(*args, **kwargs)
        return wrapped
    return deco

def allowed_file(name):
    return Path(name).suffix.lower() in ALLOWED_EXTENSIONS

# ---------- UI ----------

CSS = """
:root{--primary:#123b63;--secondary:#1e5b8d;--bg:#f3f6fa;--card:#fff;--text:#17212b;--muted:#6b7785;--border:#e1e7ee;--ok:#198754;--warn:#d99000;--bad:#c0392b}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--text);font-family:Segoe UI,Tahoma,Arial,sans-serif}
a{text-decoration:none;color:inherit}.top{height:64px;background:var(--primary);color:white;display:flex;align-items:center;justify-content:space-between;padding:0 24px;position:sticky;top:0;z-index:10}
.brand{font-size:21px;font-weight:800}.brand small{font-size:11px;opacity:.7;margin-right:8px}.top-actions{display:flex;gap:14px;align-items:center}.icon{font-size:20px}
.layout{display:flex;min-height:calc(100vh - 64px)}.side{width:245px;background:#fff;border-left:1px solid var(--border);padding:18px 12px}.side a{display:block;padding:11px 13px;border-radius:9px;margin:4px 0;color:#334455}.side a:hover,.side a.active{background:#eaf2f8;color:var(--primary);font-weight:700}
.main{flex:1;padding:26px;max-width:1500px}.page-title{display:flex;justify-content:space-between;align-items:center;margin-bottom:20px}.page-title h1{margin:0;font-size:27px}.muted{color:var(--muted)}
.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:16px}.card{background:var(--card);border:1px solid var(--border);border-radius:13px;padding:20px;box-shadow:0 3px 12px rgba(20,50,80,.05);margin-bottom:18px}.metric{font-size:29px;font-weight:800;margin-top:8px}.metric-label{color:var(--muted);font-size:13px}
table{width:100%;border-collapse:collapse}th,td{padding:11px;border-bottom:1px solid var(--border);text-align:right;font-size:13px}th{background:#f7f9fb;color:#415468}
.btn{display:inline-block;border:0;border-radius:8px;padding:9px 15px;background:var(--primary);color:#fff;cursor:pointer;font-weight:600}.btn.secondary{background:#e8eef4;color:var(--primary)}.btn.ok{background:var(--ok)}.btn.warn{background:var(--warn)}.btn.bad{background:var(--bad)}
input,select,textarea{width:100%;padding:10px;border:1px solid #ccd6e0;border-radius:8px;background:white;font:inherit}label{font-size:13px;font-weight:700;display:block;margin:10px 0 5px}.form-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:14px}
.badge{display:inline-block;padding:4px 9px;border-radius:20px;background:#e9eef3;font-size:11px}.badge.ok{background:#dff3e8;color:#176a3a}.badge.warn{background:#fff0d2;color:#8b5a00}.badge.bad{background:#fde2df;color:#9e281e}
.flash{padding:11px 14px;border-radius:8px;margin-bottom:12px;background:#e9eef3}.flash.error{background:#fde2df;color:#9e281e}
.chartbar{height:16px;background:#dfe7ef;border-radius:10px;overflow:hidden}.chartbar span{display:block;height:100%;background:var(--secondary)}
.login{min-height:100vh;display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,#123b63,#1f6d9d)}.login-card{width:390px;background:white;border-radius:16px;padding:30px;box-shadow:0 20px 60px rgba(0,0,0,.2)}
@media(max-width:900px){.side{display:none}.grid{grid-template-columns:repeat(2,1fr)}.main{padding:15px}.form-grid{grid-template-columns:1fr}}
@media(max-width:600px){.grid{grid-template-columns:1fr}.top{padding:0 12px}.brand{font-size:17px}}
"""

def shell(title, body, active="dashboard"):
    u = current_user()
    if not u:
        return body
    con = db()
    unread = con.execute("SELECT COUNT(*) FROM notifications WHERE user_id=? AND is_read=0", (u["id"],)).fetchone()[0]
    con.close()
    nav = [
        ("dashboard","🏠 Dashboard","dashboard"),
        ("documents","📁 Documents","documents"),
        ("projects","🏗️ Projects","projects"),
        ("transmittals","📨 Transmittals","transmittals"),
        ("rfis","❓ RFIs","rfis"),
        ("submittals","📋 Submittals","submittals"),
        ("ai_ocr","🤖 AI / OCR","ai_ocr"),
        ("analytics","📊 Analytics / BI","analytics"),
        ("notifications","🔔 Notifications","notifications"),
        ("users","👥 Users","users"),
        ("audit","🕒 Audit Trail","audit"),
        ("storage","💾 Storage","storage"),
        ("search_pro","🔎 Smart Search","smart_search"),
        ("revisions","🧬 Revisions","revisions"),
        ("tasks","⏱️ My Tasks","tasks"),
        ("messages","✉️ Internal Mail","messages"),
        ("settings","⚙️ Settings","settings"),
    ]
    side = "".join(
        f'<a class="{"active" if active==key else ""}" href="{url_for(route)}">{label}</a>'
        for key,label,route in nav
    )
    return f"""<!doctype html><html lang="ar" dir="rtl"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{title} — MD DocuAI</title><style>{CSS}</style></head><body>
<header class="top"><div class="brand">◈ MD DocuAI <small>PROJECT PLATFORM</small></div>
<div class="top-actions"><a class="icon" href="{url_for('notifications')}">🔔 {unread}</a>
<span>👤 {u['full_name']} · {u['role']}</span><a class="btn secondary" href="{url_for('logout')}">Logout</a></div></header>
<div class="layout"><aside class="side">{side}</aside><main class="main">
{"".join(f'<div class="flash {"error" if cat=="error" else ""}">{msg}</div>' for cat,msg in get_flashed_messages(with_categories=True))}
{body}</main></div></body></html>"""

# Flask helper imported through template context manually.
from flask import get_flashed_messages

# ---------- Authentication ----------

LOGIN_HTML = """<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>MD DocuAI Login</title>
<style>{CSS}</style></head><body><div class="login"><div class="login-card">
<div style="text-align:center;font-size:34px">◈</div><h1 style="text-align:center;margin:8px 0">MD DocuAI</h1>
<p class="muted" style="text-align:center">Intelligent Project Document Management</p>
<form method="post">
<label>Username</label><input name="username" required autofocus>
<label>Password</label><div style="position:relative"><input id="pw" name="password" type="password" required style="padding-left:45px">
<button type="button" onclick="togglePw()" style="position:absolute;left:4px;top:4px;border:0;background:transparent;font-size:20px">👁️</button></div>
<button class="btn" style="width:100%;margin-top:18px">🔐 Login</button></form>
<p class="muted" style="font-size:12px;margin-top:18px">Initial administrator: admin / Admin@123 — change it after first login.</p>
</div></div><script>function togglePw(){let x=document.getElementById('pw');x.type=x.type==='password'?'text':'password'}</script></body></html>""".replace("{CSS}", CSS)

@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        con=db()
        u=con.execute("SELECT * FROM users WHERE username=? AND active=1",(request.form.get("username","").strip(),)).fetchone()
        con.close()
        if u and check_password_hash(u["password_hash"], request.form.get("password","")):
            session["user_id"]=u["id"]
            audit("LOGIN","user",u["id"],"Successful login")
            return redirect(url_for("dashboard"))
        flash("بيانات الدخول غير صحيحة.", "error")
    return render_template_string(LOGIN_HTML)

@app.route("/logout")
def logout():
    if current_user(): audit("LOGOUT","user",current_user()["id"],"Logout")
    session.clear()
    return redirect(url_for("login"))

# ---------- Dashboard ----------

@app.route("/")
@app.route("/dashboard")
@login_required
def dashboard():
    u=current_user(); con=db()
    where="WHERE company_id=?"; args=[u["company_id"]]
    counts={
        "documents":con.execute(f"SELECT COUNT(*) FROM documents {where}",args).fetchone()[0],
        "projects":con.execute(f"SELECT COUNT(*) FROM projects {where}",args).fetchone()[0],
        "rfis":con.execute(f"SELECT COUNT(*) FROM rfis {where}",args).fetchone()[0],
        "submittals":con.execute(f"SELECT COUNT(*) FROM submittals {where}",args).fetchone()[0],
        "transmittals":con.execute(f"SELECT COUNT(*) FROM transmittals {where}",args).fetchone()[0],
    }
    recent=con.execute("""SELECT d.*,p.name project_name FROM documents d
        LEFT JOIN projects p ON p.id=d.project_id WHERE d.company_id=? ORDER BY d.id DESC LIMIT 8""",(u["company_id"],)).fetchall()
    con.close()
    rows="".join(f"<tr><td>{r['filename']}</td><td>{r['document_number'] or '-'}</td><td>{r['document_type'] or '-'}</td><td>{r['revision'] or '-'}</td><td><span class='badge'>{r['workflow_status']}</span></td></tr>" for r in recent)
    body=f"""<div class="page-title"><div><h1>Dashboard</h1><div class="muted">Project Control Center — {u['company_name']}</div></div>
<a class="btn" href="{url_for('documents')}">＋ Upload Documents</a></div>
<div class="grid">{"".join(f"<div class='card'><div class='metric-label'>{k.title()}</div><div class='metric'>{v}</div></div>" for k,v in counts.items())}</div>
<div class="card"><h3>Recent Documents</h3><table><tr><th>File</th><th>Document No.</th><th>Type</th><th>Revision</th><th>Workflow</th></tr>{rows or '<tr><td colspan=5>No documents yet</td></tr>'}</table></div>"""
    return render_template_string(shell("Dashboard",body,"dashboard"))

# ---------- Documents + OCR ----------

@app.route("/documents", methods=["GET","POST"])
@login_required
def documents():
    u=current_user()
    if request.method=="POST":
        project_id=request.form.get("project_id") or None
        files=request.files.getlist("files")
        saved=0
        for f in files:
            if not f or not f.filename or not allowed_file(f.filename):
                continue
            safe=secure_filename(f.filename)
            uid=secrets.token_hex(8)
            stored=f"{uid}_{safe}"
            path=STORAGE_DIR/stored
            f.save(path)
            try:
                txt=read_any(str(path))
                result=extract_fields(txt,safe)
                ocr_used=1 if path.suffix.lower()==".pdf" and len(txt.strip())>0 else 0
            except Exception as exc:
                result=ExtractedDocument(safe,None,None,None,None,None,None,"Low")
                ocr_used=0
                flash(f"تعذر تحليل {safe}: {exc}","error")
            con=db()
            con.execute("""INSERT INTO documents
                (company_id,project_id,uid,filename,stored_name,document_number,document_type,issue_date,
                 revision,status,sender,confidence,ocr_used,workflow_status,file_size,uploaded_by,created_at,updated_at)
                 VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (u["company_id"],project_id,uid,safe,stored,result.document_number,result.document_type,
                 result.date,result.revision,result.status or "Uploaded",result.sender,result.confidence,
                 ocr_used,"Draft",path.stat().st_size,u["id"],now(),now()))
            con.commit()
            doc_id=con.execute("SELECT id FROM documents WHERE uid=?",(uid,)).fetchone()["id"]
            con.close()
            audit("UPLOAD","document",doc_id,safe)
            saved+=1
        if saved: flash(f"تم رفع وتحليل {saved} مستند.")
        return redirect(url_for("documents"))
    q=request.args.get("q","").strip()
    con=db()
    projects=con.execute("SELECT * FROM projects WHERE company_id=? ORDER BY name",(u["company_id"],)).fetchall()
    if q:
        docs=con.execute("""SELECT d.*,p.name project_name FROM documents d LEFT JOIN projects p ON p.id=d.project_id
            WHERE d.company_id=? AND (d.filename LIKE ? OR d.document_number LIKE ? OR d.document_type LIKE ? OR d.sender LIKE ?)
            ORDER BY d.id DESC""",(u["company_id"],f"%{q}%",f"%{q}%",f"%{q}%",f"%{q}%")).fetchall()
    else:
        docs=con.execute("""SELECT d.*,p.name project_name FROM documents d LEFT JOIN projects p ON p.id=d.project_id
            WHERE d.company_id=? ORDER BY d.id DESC LIMIT 100""",(u["company_id"],)).fetchall()
    con.close()
    opts="".join(f"<option value='{p['id']}'>{p['code']} — {p['name']}</option>" for p in projects)
    rows="".join(f"""<tr><td>{r['filename']}</td><td>{r['document_number'] or '-'}</td><td>{r['document_type'] or '-'}</td>
<td>{r['issue_date'] or '-'}</td><td>{r['revision'] or '-'}</td><td>{r['status'] or '-'}</td><td>{r['sender'] or '-'}</td>
<td><span class='badge'>{r['workflow_status']}</span></td><td><a class='btn secondary' href='{url_for('download',uid=r['uid'])}'>Download</a></td></tr>""" for r in docs)
    body=f"""<div class="page-title"><div><h1>Documents</h1><div class="muted">Document Register + OCR + Revision + Workflow</div></div></div>
<div class="card"><form method="post" enctype="multipart/form-data"><div class="form-grid">
<div><label>Project</label><select name="project_id"><option value="">General / Unassigned</option>{opts}</select></div>
<div><label>PDF / DOCX / TXT</label><input type="file" name="files" multiple required></div></div>
<button class="btn" style="margin-top:14px">🤖 Upload + AI/OCR Analysis</button></form></div>
<div class="card"><form method="get"><div class="form-grid"><input name="q" value="{q}" placeholder="Search document number, type, sender, filename..."><button class="btn secondary">🔎 Search</button></div></form>
<table><tr><th>File</th><th>No.</th><th>Type</th><th>Date</th><th>Rev.</th><th>Status</th><th>Sender</th><th>Workflow</th><th>Action</th></tr>{rows or '<tr><td colspan=9>No documents</td></tr>'}</table></div>"""
    return render_template_string(shell("Documents",body,"documents"))

@app.route("/download/<uid>")
@login_required
def download(uid):
    u=current_user(); con=db()
    r=con.execute("SELECT * FROM documents WHERE uid=? AND company_id=?",(uid,u["company_id"])).fetchone()
    con.close()
    if not r: return "Not found",404
    path=STORAGE_DIR/r["stored_name"]
    if not path.exists(): return "File missing",404
    audit("DOWNLOAD","document",r["id"],r["filename"])
    return send_file(path,as_attachment=True,download_name=r["filename"])

@app.route("/ai-ocr")
@login_required
def ai_ocr():
    body="""<div class="page-title"><div><h1>AI / OCR</h1><div class="muted">Intelligent document extraction engine</div></div></div>
<div class="card"><h3>Supported extraction</h3><div class="grid">
<div><b>Document Number</b><p class="muted">Regex + project document patterns</p></div>
<div><b>Document Type</b><p class="muted">RFI, Submittal, Drawing, Method Statement, etc.</p></div>
<div><b>Issue Date / Revision</b><p class="muted">Automatic normalization</p></div>
<div><b>Status / Sender</b><p class="muted">Keyword-based extraction</p></div></div>
<p>استخدم صفحة Documents لرفع الملفات. ملفات PDF المصورة تستخدم Tesseract OCR عند الحاجة.</p></div>"""
    return render_template_string(shell("AI / OCR",body,"ai_ocr"))

# ---------- Projects ----------

@app.route("/projects",methods=["GET","POST"])
@login_required
@role_required("Administrator","Manager","Editor")
def projects():
    u=current_user(); con=db()
    if request.method=="POST":
        name=request.form.get("name","").strip(); code=request.form.get("code","").strip().upper()
        client=request.form.get("client","").strip()
        if name and code:
            try:
                con.execute("INSERT INTO projects(company_id,name,code,client,created_at) VALUES(?,?,?,?,?)",(u["company_id"],name,code,client,now()))
                con.commit(); flash("تم إنشاء المشروع."); audit("CREATE","project",None,f"{code} {name}")
            except sqlite3.IntegrityError: flash("رمز المشروع موجود بالفعل.","error")
    rows=con.execute("SELECT * FROM projects WHERE company_id=? ORDER BY id DESC",(u["company_id"],)).fetchall(); con.close()
    tr="".join(f"<tr><td>{r['code']}</td><td>{r['name']}</td><td>{r['client'] or '-'}</td><td>{r['status']}</td><td>{r['created_at']}</td></tr>" for r in rows)
    body=f"""<div class="page-title"><div><h1>Projects</h1></div></div>
<div class="card"><form method="post"><div class="form-grid"><div><label>Project Name</label><input name="name" required></div><div><label>Project Code</label><input name="code" required></div><div><label>Client</label><input name="client"></div></div><button class="btn" style="margin-top:12px">＋ Create Project</button></form></div>
<div class="card"><table><tr><th>Code</th><th>Project</th><th>Client</th><th>Status</th><th>Created</th></tr>{tr}</table></div>"""
    return render_template_string(shell("Projects",body,"projects"))

# ---------- Transmittals / RFIs / Submittals ----------

def module_page(kind):
    u=current_user(); con=db()
    meta={
      "transmittals":("Transmittals","number,subject,to_company","number,subject,to_company","Draft"),
      "rfis":("RFIs","number,subject,question","number,subject,question","Open"),
      "submittals":("Submittals","number,title,type","number,title,type","Submitted")
    }[kind]
    table=kind
    if request.method=="POST":
        project_id=request.form.get("project_id") or None
        if kind=="transmittals":
            con.execute("INSERT INTO transmittals(company_id,project_id,number,subject,to_company,status,created_by,created_at) VALUES(?,?,?,?,?,?,?,?)",
                        (u["company_id"],project_id,request.form["number"],request.form["subject"],request.form.get("to_company"),"Draft",u["id"],now()))
        elif kind=="rfis":
            con.execute("INSERT INTO rfis(company_id,project_id,number,subject,question,status,priority,due_date,created_by,created_at) VALUES(?,?,?,?,?,?,?,?,?,?)",
                        (u["company_id"],project_id,request.form["number"],request.form["subject"],request.form.get("question"),"Open",request.form.get("priority","Normal"),request.form.get("due_date"),u["id"],now()))
        else:
            con.execute("INSERT INTO submittals(company_id,project_id,number,title,type,status,revision,created_by,created_at) VALUES(?,?,?,?,?,?,?,?,?)",
                        (u["company_id"],project_id,request.form["number"],request.form["title"],request.form.get("type"),"Submitted",request.form.get("revision"),u["id"],now()))
        con.commit(); flash(f"تم إنشاء {meta[0]} بنجاح."); audit("CREATE",kind,None,request.form.get("number",""))
    projects=con.execute("SELECT * FROM projects WHERE company_id=? ORDER BY name",(u["company_id"],)).fetchall()
    data=con.execute(f"SELECT * FROM {table} WHERE company_id=? ORDER BY id DESC",(u["company_id"],)).fetchall(); con.close()
    opts="".join(f"<option value='{p['id']}'>{p['code']} — {p['name']}</option>" for p in projects)
    if kind=="transmittals":
        fields=f"""<div><label>Number</label><input name="number" required></div><div><label>Subject</label><input name="subject" required></div><div><label>To Company</label><input name="to_company"></div>"""
        head="<th>No.</th><th>Subject</th><th>To</th><th>Status</th><th>Created</th>"
        tr="".join(f"<tr><td>{r['number']}</td><td>{r['subject']}</td><td>{r['to_company'] or '-'}</td><td>{r['status']}</td><td>{r['created_at']}</td></tr>" for r in data)
    elif kind=="rfis":
        fields=f"""<div><label>Number</label><input name="number" required></div><div><label>Subject</label><input name="subject" required></div><div><label>Question</label><textarea name="question"></textarea></div><div><label>Priority</label><select name="priority"><option>Normal</option><option>High</option><option>Critical</option></select></div><div><label>Due Date</label><input name="due_date" type="date"></div>"""
        head="<th>No.</th><th>Subject</th><th>Priority</th><th>Status</th><th>Due</th>"
        tr="".join(f"<tr><td>{r['number']}</td><td>{r['subject']}</td><td>{r['priority']}</td><td>{r['status']}</td><td>{r['due_date'] or '-'}</td></tr>" for r in data)
    else:
        fields=f"""<div><label>Number</label><input name="number" required></div><div><label>Title</label><input name="title" required></div><div><label>Type</label><input name="type"></div><div><label>Revision</label><input name="revision"></div>"""
        head="<th>No.</th><th>Title</th><th>Type</th><th>Revision</th><th>Status</th>"
        tr="".join(f"<tr><td>{r['number']}</td><td>{r['title']}</td><td>{r['type'] or '-'}</td><td>{r['revision'] or '-'}</td><td>{r['status']}</td></tr>" for r in data)
    body=f"""<div class="page-title"><h1>{meta[0]}</h1></div><div class="card"><form method="post"><div class="form-grid"><div><label>Project</label><select name="project_id"><option value="">General</option>{opts}</select></div>{fields}</div><button class="btn" style="margin-top:12px">＋ Create</button></form></div>
<div class="card"><table><tr>{head}</tr>{tr or '<tr><td colspan=6>No records</td></tr>'}</table></div>"""
    return render_template_string(shell(meta[0],body,kind))

@app.route("/transmittals",methods=["GET","POST"])
@login_required
@role_required("Administrator","Manager","Editor")
def transmittals(): return module_page("transmittals")

@app.route("/rfis",methods=["GET","POST"])
@login_required
@role_required("Administrator","Manager","Editor")
def rfis(): return module_page("rfis")

@app.route("/submittals",methods=["GET","POST"])
@login_required
@role_required("Administrator","Manager","Editor")
def submittals(): return module_page("submittals")

# ---------- Workflow ----------

@app.route("/workflow/<int:doc_id>/<action>",methods=["POST"])
@login_required
@role_required("Administrator","Manager","Editor")
def workflow_action(doc_id,action):
    allowed={"submit":"Submitted","review":"Under Review","approve":"Approved","reject":"Rejected","revise":"Revision Required","close":"Closed"}
    if action not in allowed: return "Invalid action",400
    u=current_user(); con=db()
    d=con.execute("SELECT * FROM documents WHERE id=? AND company_id=?",(doc_id,u["company_id"])).fetchone()
    if not d: con.close(); return "Not found",404
    new=allowed[action]
    con.execute("UPDATE documents SET workflow_status=?,updated_at=? WHERE id=?",(new,now(),doc_id))
    con.execute("INSERT INTO workflow(document_id,action,from_status,to_status,comment,user_id,created_at) VALUES(?,?,?,?,?,?,?)",
                (doc_id,action,d["workflow_status"],new,request.form.get("comment"),u["id"],now()))
    con.commit(); con.close(); audit("WORKFLOW","document",doc_id,f"{d['workflow_status']} -> {new}")
    return redirect(url_for("documents"))

# ---------- Analytics / BI ----------

@app.route("/analytics")
@login_required
def analytics():
    u=current_user(); con=db()
    total=con.execute("SELECT COUNT(*) FROM documents WHERE company_id=?",(u["company_id"],)).fetchone()[0]
    types=con.execute("""SELECT COALESCE(document_type,'Unknown') label,COUNT(*) n FROM documents
                         WHERE company_id=? GROUP BY document_type ORDER BY n DESC LIMIT 10""",(u["company_id"],)).fetchall()
    statuses=con.execute("""SELECT workflow_status label,COUNT(*) n FROM documents
                            WHERE company_id=? GROUP BY workflow_status ORDER BY n DESC""",(u["company_id"],)).fetchall()
    conf=con.execute("""SELECT COALESCE(confidence,'Unknown') label,COUNT(*) n FROM documents
                        WHERE company_id=? GROUP BY confidence""",(u["company_id"],)).fetchall()
    con.close()
    def bars(rows):
        mx=max([r["n"] for r in rows] or [1])
        return "".join(f"<div style='margin:9px 0'><div style='display:flex;justify-content:space-between'><span>{r['label']}</span><b>{r['n']}</b></div><div class='chartbar'><span style='width:{r['n']/mx*100:.1f}%'></span></div></div>" for r in rows)
    body=f"""<div class="page-title"><div><h1>Analytics / BI</h1><div class="muted">Operational project intelligence</div></div></div>
<div class="grid"><div class="card"><div class="metric-label">Total Documents</div><div class="metric">{total}</div></div>
<div class="card"><div class="metric-label">Document Types</div><div class="metric">{len(types)}</div></div>
<div class="card"><div class="metric-label">Workflow States</div><div class="metric">{len(statuses)}</div></div>
<div class="card"><div class="metric-label">Confidence Levels</div><div class="metric">{len(conf)}</div></div></div>
<div class="grid"><div class="card"><h3>Documents by Type</h3>{bars(types) or '<p class="muted">No data</p>'}</div>
<div class="card"><h3>Workflow Status</h3>{bars(statuses) or '<p class="muted">No data</p>'}</div></div>
<div class="card"><h3>Extraction Confidence</h3>{bars(conf) or '<p class="muted">No data</p>'}</div>"""
    return render_template_string(shell("Analytics",body,"analytics"))

# ---------- Notifications ----------

@app.route("/notifications")
@login_required
def notifications():
    u=current_user(); con=db()
    rows=con.execute("SELECT * FROM notifications WHERE user_id=? ORDER BY id DESC LIMIT 100",(u["id"],)).fetchall()
    con.execute("UPDATE notifications SET is_read=1 WHERE user_id=?",(u["id"],)); con.commit(); con.close()
    tr="".join(f"<tr><td>{r['created_at']}</td><td>{r['title']}</td><td>{r['message']}</td></tr>" for r in rows)
    body=f"""<div class="page-title"><h1>Notifications</h1></div><div class="card"><table><tr><th>Date</th><th>Title</th><th>Message</th></tr>{tr or '<tr><td colspan=3>No notifications</td></tr>'}</table></div>"""
    return render_template_string(shell("Notifications",body,"notifications"))

# ---------- Users / Companies ----------

@app.route("/users",methods=["GET","POST"])
@login_required
@role_required("Administrator")
def users():
    u=current_user(); con=db()
    if request.method=="POST":
        username=request.form["username"].strip(); full=request.form["full_name"].strip()
        role=request.form.get("role","Viewer"); pw=request.form["password"]
        try:
            con.execute("""INSERT INTO users(company_id,username,password_hash,full_name,role,created_at)
                           VALUES(?,?,?,?,?,?)""",(u["company_id"],username,generate_password_hash(pw),full,role,now()))
            con.commit(); flash("تم إنشاء المستخدم.")
            audit("CREATE","user",None,username)
        except sqlite3.IntegrityError: flash("اسم المستخدم موجود بالفعل.","error")
    rows=con.execute("SELECT * FROM users WHERE company_id=? ORDER BY id DESC",(u["company_id"],)).fetchall(); con.close()
    tr="".join(f"<tr><td>{r['username']}</td><td>{r['full_name']}</td><td>{r['role']}</td><td>{'Active' if r['active'] else 'Inactive'}</td></tr>" for r in rows)
    body=f"""<div class="page-title"><h1>Users & Permissions</h1></div><div class="card"><form method="post"><div class="form-grid">
<div><label>Username</label><input name="username" required></div><div><label>Full Name</label><input name="full_name" required></div>
<div><label>Password</label><input name="password" type="password" required></div><div><label>Role</label><select name="role"><option>Administrator</option><option>Manager</option><option>Editor</option><option>Viewer</option></select></div>
</div><button class="btn" style="margin-top:12px">＋ Create User</button></form></div>
<div class="card"><table><tr><th>Username</th><th>Name</th><th>Role</th><th>Status</th></tr>{tr}</table></div>"""
    return render_template_string(shell("Users",body,"users"))

# ---------- Audit / Storage ----------

@app.route("/audit")
@login_required
@role_required("Administrator","Manager")
def audit_page():
    u=current_user(); con=db()
    rows=con.execute("""SELECT a.*,u.username FROM audit_log a LEFT JOIN users u ON u.id=a.user_id
                        WHERE u.company_id=? ORDER BY a.id DESC LIMIT 200""",(u["company_id"],)).fetchall(); con.close()
    tr="".join(f"<tr><td>{r['created_at']}</td><td>{r['username'] or '-'}</td><td>{r['action']}</td><td>{r['entity_type']}</td><td>{r['details'] or ''}</td></tr>" for r in rows)
    body=f"""<div class="page-title"><h1>Audit Trail</h1></div><div class="card"><table><tr><th>Date</th><th>User</th><th>Action</th><th>Entity</th><th>Details</th></tr>{tr}</table></div>"""
    return render_template_string(shell("Audit",body,"audit"))

@app.route("/storage")
@login_required
def storage():
    u=current_user(); con=db()
    total=con.execute("SELECT COALESCE(SUM(file_size),0) FROM documents WHERE company_id=?",(u["company_id"],)).fetchone()[0]
    count=con.execute("SELECT COUNT(*) FROM documents WHERE company_id=?",(u["company_id"],)).fetchone()[0]; con.close()
    body=f"""<div class="page-title"><h1>Storage</h1></div><div class="grid">
<div class="card"><div class="metric-label">Stored Files</div><div class="metric">{count}</div></div>
<div class="card"><div class="metric-label">Storage Used</div><div class="metric">{total/1024/1024:.2f} MB</div></div></div>
<div class="card"><h3>Cloud-ready architecture</h3><p class="muted">النسخة الحالية تستخدم Local Storage. يمكن لاحقًا ربط S3/Azure Blob/Google Cloud Storage دون تغيير طبقة المستندات الأساسية.</p></div>"""
    return render_template_string(shell("Storage",body,"storage"))


# =============================================================================
# V5-V18 PROFESSIONAL MODULES
# =============================================================================

@app.route("/smart-search")
@login_required
def smart_search():
    u = current_user()
    q = request.args.get("q","").strip()
    con = db()
    rows = []
    if q:
        like = f"%{q}%"
        rows = con.execute("""
            SELECT d.*, p.name AS project_name
            FROM documents d LEFT JOIN projects p ON p.id=d.project_id
            WHERE d.company_id=? AND (
                d.filename LIKE ? OR d.document_number LIKE ? OR
                d.document_type LIKE ? OR d.status LIKE ? OR
                d.sender LIKE ? OR d.revision LIKE ?
            )
            ORDER BY d.id DESC LIMIT 300
        """,(u["company_id"],like,like,like,like,like,like)).fetchall()
    con.close()
    tr = "".join(
        f"<tr><td>{r['filename']}</td><td>{r['document_number'] or '-'}</td>"
        f"<td>{r['project_name'] or '-'}</td><td>{r['document_type'] or '-'}</td>"
        f"<td>{r['revision'] or '-'}</td><td>{r['status'] or '-'}</td>"
        f"<td>{r['confidence'] or '-'}</td>"
        f"<td><a class='btn secondary' href='{url_for('download',uid=r['uid'])}'>Open</a></td></tr>"
        for r in rows
    )
    body = f"""<div class="hero"><h1>Smart Search</h1>
    <p>Search document number, type, project, revision, status, sender and filename.</p></div>
    <div class="card"><form class="toolbar">
    <div><label>Search</label><input name="q" value="{q}" placeholder="Example: DRG-001 / Approved / Drainage"></div>
    <button class="btn">🔎 Search</button></form></div>
    <div class="card table-wrap"><table><tr><th>File</th><th>Document No.</th><th>Project</th>
    <th>Type</th><th>Revision</th><th>Status</th><th>AI Confidence</th><th>Action</th></tr>
    {tr or '<tr><td colspan="8" class="empty">Enter a search term to begin.</td></tr>'}</table></div>"""
    return render_template_string(shell("Smart Search",body,"search_pro"))

@app.route("/revisions")
@login_required
def revisions():
    u = current_user()
    con = db()
    rows = con.execute("""
        SELECT r.*, d.filename, d.document_number, u.full_name
        FROM document_revisions r
        JOIN documents d ON d.id=r.document_id
        LEFT JOIN users u ON u.id=r.created_by
        WHERE d.company_id=? ORDER BY r.id DESC LIMIT 300
    """,(u["company_id"],)).fetchall()
    con.close()
    tr = "".join(
        f"<tr><td>{r['filename']}</td><td>{r['document_number'] or '-'}</td>"
        f"<td><b>{r['revision']}</b></td><td>{r['status']}</td>"
        f"<td>{r['change_summary'] or '-'}</td><td>{r['full_name'] or '-'}</td>"
        f"<td>{r['created_at']}</td></tr>" for r in rows
    )
    body = f"""<div class="page-title"><div><h1>Revision Control</h1>
    <div class="muted">Complete document version history and change tracking.</div></div></div>
    <div class="card table-wrap"><table><tr><th>Document</th><th>Number</th><th>Revision</th>
    <th>Status</th><th>Change Summary</th><th>Created By</th><th>Date</th></tr>
    {tr or '<tr><td colspan="7" class="empty">No revisions recorded yet.</td></tr>'}</table></div>"""
    return render_template_string(shell("Revisions",body,"revisions"))

@app.route("/tasks")
@login_required
def tasks():
    u = current_user()
    con = db()
    rows = con.execute("""
        SELECT t.*, d.filename, d.document_number
        FROM workflow_tasks t JOIN documents d ON d.id=t.document_id
        WHERE d.company_id=? AND t.assigned_to=?
        ORDER BY CASE WHEN t.status='Pending' THEN 0 ELSE 1 END, t.due_date
    """,(u["company_id"],u["id"])).fetchall()
    con.close()
    tr = "".join(
        f"<tr><td>{r['document_number'] or r['filename']}</td><td>{r['stage']}</td>"
        f"<td>{r['due_date'] or '-'}</td><td><span class='pill'>{r['status']}</span></td>"
        f"<td>{r['decision'] or '-'}</td></tr>" for r in rows
    )
    body = f"""<div class="page-title"><div><h1>My Tasks</h1>
    <div class="muted">Workflow assignments, review deadlines and decisions.</div></div></div>
    <div class="card table-wrap"><table><tr><th>Document</th><th>Stage</th><th>Due Date</th>
    <th>Status</th><th>Decision</th></tr>{tr or '<tr><td colspan="5" class="empty">No assigned tasks.</td></tr>'}</table></div>"""
    return render_template_string(shell("My Tasks",body,"tasks"))

@app.route("/messages", methods=["GET","POST"])
@login_required
def messages():
    u = current_user()
    con = db()
    if request.method == "POST":
        receiver = request.form.get("receiver_id")
        subject = request.form.get("subject","").strip()
        body_txt = request.form.get("body","").strip()
        if receiver and subject and body_txt:
            con.execute("""INSERT INTO internal_messages
                (company_id,sender_id,receiver_id,subject,body,created_at)
                VALUES(?,?,?,?,?,?)""",
                (u["company_id"],u["id"],int(receiver),subject,body_txt,now()))
            con.commit()
            notify(int(receiver),"رسالة جديدة",subject)
            audit("MESSAGE","internal_message",None,subject)
            flash("تم إرسال الرسالة.")
        return redirect(url_for("messages"))
    users = con.execute("""SELECT id,full_name,username FROM users
                           WHERE company_id=? AND id<>? AND active=1 ORDER BY full_name""",
                        (u["company_id"],u["id"])).fetchall()
    inbox = con.execute("""
        SELECT m.*, s.full_name sender_name
        FROM internal_messages m JOIN users s ON s.id=m.sender_id
        WHERE m.receiver_id=? ORDER BY m.id DESC LIMIT 100
    """,(u["id"],)).fetchall()
    con.close()
    opts = "".join(f"<option value='{r['id']}'>{r['full_name']} ({r['username']})</option>" for r in users)
    tr = "".join(
        f"<tr><td>{r['sender_name']}</td><td>{r['subject']}</td><td>{r['body']}</td>"
        f"<td>{r['created_at']}</td></tr>" for r in inbox
    )
    body = f"""<div class="page-title"><div><h1>Internal Mail</h1>
    <div class="muted">Project communication inside your company workspace.</div></div></div>
    <div class="card"><form method="post"><div class="form-grid">
    <div><label>To</label><select name="receiver_id" required>{opts}</select></div>
    <div><label>Subject</label><input name="subject" required></div>
    </div><label>Message</label><textarea name="body" rows="4" required></textarea>
    <button class="btn" style="margin-top:12px">✉️ Send</button></form></div>
    <div class="card table-wrap"><h3>Inbox</h3><table><tr><th>From</th><th>Subject</th>
    <th>Message</th><th>Date</th></tr>{tr or '<tr><td colspan="4" class="empty">Inbox is empty.</td></tr>'}</table></div>"""
    return render_template_string(shell("Internal Mail",body,"messages"))

@app.route("/settings", methods=["GET","POST"])
@login_required
def settings():
    u = current_user()
    con = db()
    if request.method == "POST":
        secret = request.form.get("secret_key","").strip()
        if secret:
            # Store company-level configuration without exposing it in the UI.
            con.execute("""INSERT INTO system_settings(company_id,setting_key,setting_value)
                           VALUES(?,?,?) ON CONFLICT(company_id,setting_key)
                           DO UPDATE SET setting_value=excluded.setting_value""",
                        (u["company_id"],"workspace_secret_hint",secret[-4:]))
            con.commit()
            audit("SETTINGS","company",u["company_id"],"Workspace settings updated")
            flash("تم تحديث الإعدادات.")
        return redirect(url_for("settings"))
    con.close()
    body = """<div class="page-title"><div><h1>Settings</h1>
    <div class="muted">Workspace configuration and deployment-ready controls.</div></div></div>
    <div class="card"><h3>Platform</h3>
    <div class="ai-field"><b>Product</b><span>MD DocuAI Professional</span></div>
    <div class="ai-field"><b>Architecture</b><span>Flask + SQLite + Local/Cloud-ready Storage</span></div>
    <div class="ai-field"><b>AI/OCR</b><span>Regex extraction + optional AI analysis + OCR engine</span></div>
    <div class="ai-field"><b>API</b><span>REST API foundation with API-key authentication</span></div>
    </div>
    <div class="card"><h3>Production checklist</h3>
    <p>Use a strong DOCUAI_SECRET_KEY, configure DOCUAI_API_KEY, move storage to object storage,
    and put the application behind HTTPS before production use.</p></div>"""
    return render_template_string(shell("Settings",body,"settings"))

@app.route("/document/<int:doc_id>")
@login_required
def document_detail(doc_id):
    u = current_user()
    con = db()
    d = con.execute("""SELECT d.*, p.name project_name FROM documents d
                       LEFT JOIN projects p ON p.id=d.project_id
                       WHERE d.id=? AND d.company_id=?""",(doc_id,u["company_id"])).fetchone()
    if not d:
        con.close()
        return "Document not found",404
    comments = con.execute("""SELECT c.*,u.full_name FROM document_comments c
                               LEFT JOIN users u ON u.id=c.user_id
                               WHERE c.document_id=? ORDER BY c.id DESC""",(doc_id,)).fetchall()
    revisions = con.execute("""SELECT * FROM document_revisions
                                WHERE document_id=? ORDER BY id DESC""",(doc_id,)).fetchall()
    tasks_rows = con.execute("""SELECT t.*,u.full_name FROM workflow_tasks t
                                LEFT JOIN users u ON u.id=t.assigned_to
                                WHERE t.document_id=? ORDER BY t.id DESC""",(doc_id,)).fetchall()
    ai = con.execute("""SELECT * FROM ai_analysis WHERE document_id=? ORDER BY id DESC LIMIT 1""",(doc_id,)).fetchone()
    con.close()
    rev_html = "".join(f"<li><b>{r['revision']}</b> — {r['status']} — {r['created_at']} — {r['change_summary'] or ''}</li>" for r in revisions)
    task_html = "".join(f"<li><b>{r['stage']}</b> — {r['status']} — {r['due_date'] or 'No due date'} — {r['full_name'] or 'Unassigned'}</li>" for r in tasks_rows)
    comment_html = "".join(f"<div class='card' style='margin:8px 0'><b>{r['full_name'] or 'User'}</b><br>{r['comment']}<br><small class='muted'>{r['created_at']}</small></div>" for r in comments)
    ai_html = ""
    if ai:
        ai_html = f"""<div class="ai-box"><h3>🤖 AI Intelligence</h3>
        <div class="ai-field"><b>Summary</b><span>{ai['summary'] or '-'}</span></div>
        <div class="ai-field"><b>Key Points</b><span>{ai['key_points'] or '-'}</span></div>
        <div class="ai-field"><b>Required Actions</b><span>{ai['required_actions'] or '-'}</span></div>
        <div class="ai-field"><b>Risks</b><span>{ai['risks'] or '-'}</span></div>
        <div class="ai-field"><b>Keywords</b><span>{ai['keywords'] or '-'}</span></div></div>"""
    body = f"""<div class="page-title"><div><h1>{d['filename']}</h1>
    <div class="muted">{d['document_number'] or 'No document number'} · {d['project_name'] or 'Unassigned project'}</div></div>
    <a class="btn" href="{url_for('download',uid=d['uid'])}">⬇ Download</a></div>
    <div class="card"><div class="form-grid">
    <div><b>Type</b><div>{d['document_type'] or '-'}</div></div><div><b>Revision</b><div>{d['revision'] or '-'}</div></div>
    <div><b>Status</b><div><span class="pill">{d['status'] or '-'}</span></div></div><div><b>Sender</b><div>{d['sender'] or '-'}</div></div>
    <div><b>Issue Date</b><div>{d['issue_date'] or '-'}</div></div><div><b>AI Confidence</b><div>{d['confidence'] or '-'}</div></div>
    </div></div>
    {ai_html}
    <div class="split"><div class="card"><h3>Workflow Timeline</h3><div class="timeline">{task_html or '<div class="empty">No workflow tasks yet.</div>'}</div></div>
    <div class="card"><h3>Revision History</h3><ul>{rev_html or '<li class="empty">No revision history yet.</li>'}</ul></div></div>
    <div class="card"><h3>Comments</h3>{comment_html or '<div class="empty">No comments yet.</div>'}</div>"""
    return render_template_string(shell("Document Detail",body,"documents"))

@app.route("/api/v1/search")
def api_search():
    if not api_auth():
        return jsonify({"error":"API key required"}),401
    q = request.args.get("q","").strip()
    con = db()
    like = f"%{q}%"
    rows = con.execute("""SELECT id,uid,filename,document_number,document_type,revision,status,sender,confidence,created_at
                          FROM documents
                          WHERE filename LIKE ? OR document_number LIKE ? OR document_type LIKE ?
                             OR status LIKE ? OR sender LIKE ? OR revision LIKE ?
                          ORDER BY id DESC LIMIT 200""",(like,like,like,like,like,like)).fetchall()
    con.close()
    return jsonify({"query":q,"results":[dict(r) for r in rows]})

@app.route("/api/v1/dashboard")
def api_dashboard():
    if not api_auth():
        return jsonify({"error":"API key required"}),401
    con = db()
    data = {
        "documents": con.execute("SELECT COUNT(*) FROM documents").fetchone()[0],
        "projects": con.execute("SELECT COUNT(*) FROM projects").fetchone()[0],
        "rfis": con.execute("SELECT COUNT(*) FROM rfis").fetchone()[0],
        "submittals": con.execute("SELECT COUNT(*) FROM submittals").fetchone()[0],
        "transmittals": con.execute("SELECT COUNT(*) FROM transmittals").fetchone()[0],
        "approved": con.execute("SELECT COUNT(*) FROM documents WHERE status='Approved'").fetchone()[0],
        "pending": con.execute("SELECT COUNT(*) FROM documents WHERE status IN ('Pending','For Review','For Approval')").fetchone()[0],
    }
    con.close()
    return jsonify(data)


# ---------- REST API (V4 foundation) ----------

def api_auth():
    token=request.headers.get("X-API-Key") or request.args.get("api_key")
    expected=os.environ.get("DOCUAI_API_KEY")
    return bool(expected and secrets.compare_digest(token or "",expected))

@app.route("/api/v1/health")
def api_health():
    return jsonify({"status":"ok","product":"MD DocuAI","version":"3.0-professional-single-file"})

@app.route("/api/v1/documents")
def api_documents():
    if not api_auth(): return jsonify({"error":"API key required"}),401
    con=db()
    rows=con.execute("""SELECT id,uid,filename,document_number,document_type,issue_date,revision,status,
                        sender,confidence,workflow_status,created_at FROM documents ORDER BY id DESC LIMIT 200""").fetchall()
    con.close()
    return jsonify({"documents":[dict(r) for r in rows]})

# ---------- Error handling / startup ----------

@app.errorhandler(413)
def too_large(e):
    return "File too large. Maximum size is 25 MB.", 413

init_db()

if __name__ == "__main__":
    port=int(os.environ.get("PORT","5000"))
    app.run(host="0.0.0.0",port=port,debug=False)
