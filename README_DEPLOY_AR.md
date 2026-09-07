# DocuAI — ملف التشغيل والنشر على Render

هذا المجلد يحتوي على نسخة جاهزة كبداية لنشر مشروع DocuAI على Render باستخدام Docker.

## 1) الملفات الموجودة

- `app_ocr.py` — التطبيق الرئيسي الحالي.
- `requirements.txt` — مكتبات Python المطلوبة.
- `Dockerfile` — يثبت Python + Tesseract OCR + اللغة العربية + Poppler ثم يشغل Flask بواسطة Gunicorn.
- `.gitignore` — ملفات لا نريد رفعها إلى GitHub.

> لا تحذف `app.py` الموجود في مستودع GitHub إذا كان عندك. هذه الحزمة تعتمد على `app_ocr.py` فقط.

---

# 2) رفع الملفات إلى GitHub — الطريقة الأسهل

افتح مستودع GitHub الخاص بك `docuai`.

يمكنك رفع/استبدال الملفات التالية:

1. `app_ocr.py`
2. `requirements.txt`
3. `Dockerfile`
4. `.gitignore`

اترك `app.py` القديم كما هو حاليًا.

بعد كل تعديل اضغط **Commit changes**.

---

# 3) إذا كنت تستخدم الكمبيوتر وGit — الأوامر

افتح Terminal / PowerShell داخل مجلد المشروع:

```bash
git clone https://github.com/YOUR-USERNAME/docuai.git
cd docuai
```

انسخ الملفات الأربعة إلى المجلد، ثم:

```bash
git add app_ocr.py requirements.txt Dockerfile .gitignore

git commit -m "Prepare DocuAI OCR app for Render Docker deployment"

git push origin main
```

إذا كان اسم الفرع عندك `master` بدل `main` استخدم:

```bash
git push origin master
```

---

# 4) إنشاء الخدمة على Render

1. ادخل إلى Render.
2. اختر **New** ثم **Web Service**.
3. اربط حساب GitHub.
4. اختر المستودع `docuai`.
5. اختر الفرع `main` (أو الفرع الذي تستخدمه).
6. اختر **Docker** كبيئة التشغيل إذا ظهر لك اختيار Runtime.
7. اختر الخطة المناسبة. يمكنك البدء بالخطة المجانية للاختبار إذا كانت متاحة لك.
8. اضغط **Create Web Service / Deploy**.

وجود ملف `Dockerfile` في جذر المستودع يجعل Render يبني التطبيق باستخدام Docker.

---

# 5) أمر التشغيل الموجود داخل Dockerfile

لا تحتاج عادةً إلى كتابة Start Command يدويًا؛ التطبيق يستخدم الأمر التالي داخل Dockerfile:

```bash
gunicorn --bind 0.0.0.0:${PORT:-10000} --workers 1 --timeout 300 app_ocr:app
```

معناه:

- `0.0.0.0` يسمح لـ Render بالوصول إلى التطبيق.
- `${PORT:-10000}` يستخدم منفذ Render أو 10000 افتراضيًا.
- `--workers 1` يقلل استهلاك الذاكرة أثناء OCR.
- `--timeout 300` يعطي ملفات OCR الكبيرة وقتًا أطول.
- `app_ocr:app` يعني ملف `app_ocr.py` وكائن Flask اسمه `app`.

---

# 6) ماذا يفعل Dockerfile؟

الكود الكامل:

```dockerfile
FROM python:3.13-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr \
    tesseract-ocr-ara \
    poppler-utils \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app_ocr.py .

CMD ["sh", "-c", "gunicorn --bind 0.0.0.0:${PORT:-10000} --workers 1 --timeout 300 app_ocr:app"]
```

---

# 7) المكتبات المطلوبة

```text
Flask
pdfplumber
pytesseract
pdf2image
Pillow
python-docx
openpyxl
gunicorn
```

---

# 8) تشغيل المشروع محليًا بدون Docker

إذا كانت Python مثبتة عندك:

```bash
python -m venv .venv
```

Windows PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
```

Windows CMD:

```cmd
.venv\Scripts\activate
```

Linux / macOS:

```bash
source .venv/bin/activate
```

ثم:

```bash
pip install -r requirements.txt
python app_ocr.py
```

> ملاحظة: التشغيل المحلي لـ OCR للـPDF الممسوح ضوئيًا يحتاج أيضًا إلى تثبيت Tesseract وPoppler على جهازك. Docker في Render يتكفل بتثبيتهما داخل الحاوية.

---

# 9) اختبار Gunicorn محليًا

بعد تثبيت المتطلبات:

Linux / macOS:

```bash
PORT=10000 gunicorn --bind 0.0.0.0:10000 --workers 1 --timeout 300 app_ocr:app
```

Windows PowerShell:

```powershell
$env:PORT="10000"
gunicorn --bind 0.0.0.0:10000 --workers 1 --timeout 300 app_ocr:app
```

ثم افتح:

```text
http://127.0.0.1:10000
```

---

# 10) بناء Docker محليًا — اختياري

إذا كان Docker مثبتًا:

```bash
docker build -t docuai .
```

ثم:

```bash
docker run --rm -p 10000:10000 -e PORT=10000 docuai
```

ثم افتح:

```text
http://127.0.0.1:10000
```

---

# 11) فحص سريع قبل النشر

تأكد أن جذر GitHub يحتوي على الأقل على:

```text
app_ocr.py
requirements.txt
Dockerfile
.gitignore
```

ويمكن أن يبقى:

```text
app.py
README.md
```

---

# 12) ماذا يفعل DocuAI الحالي؟

النسخة الحالية ليست نموذج ذكاء اصطناعي توليدي مثل ChatGPT؛ هي نظام استخراج ذكي يعتمد على Regex + OCR + كلمات مفتاحية.

تستطيع حاليًا استخراج حقول مثل:

- Document Number
- Document Type
- Date
- Revision
- Status
- Sender
- Confidence

وتتعامل مع PDF / DOCX / TXT، وتوفر تنزيل النتائج إلى Excel.

في ملفات PDF الممسوحة ضوئيًا يستخدم التطبيق Tesseract OCR، مع دعم اللغة العربية داخل Docker.

---

# 13) إذا فشل Deploy في Render

افتح:

**Render → Service → Logs**

وانسخ آخر 30–50 سطرًا من الخطأ.

أهم الأخطاء الشائعة:

### خطأ: `ModuleNotFoundError`

راجع `requirements.txt`.

### خطأ متعلق بـ Tesseract

تأكد أن Dockerfile يحتوي على:

```dockerfile
tesseract-ocr
tesseract-ocr-ara
```

### خطأ متعلق بـ Poppler / PDF conversion

تأكد أن Dockerfile يحتوي على:

```dockerfile
poppler-utils
```

### خطأ Port / Application failed to respond

تأكد أن التشغيل يستخدم:

```bash
gunicorn --bind 0.0.0.0:${PORT:-10000} app_ocr:app
```

---

# 14) ملاحظات مهمة قبل تحويله إلى نظام إنتاج كامل

هذه النسخة مناسبة جدًا للاختبار الأولي، لكن قبل الاستخدام التجاري/متعدد المستخدمين يفضل لاحقًا:

1. استخدام `secure_filename` لأسماء الملفات.
2. عدم الاعتماد على متغير عالمي واحد للنتائج بين المستخدمين.
3. إضافة قاعدة بيانات لحفظ المستندات والنتائج.
4. إضافة تسجيل دخول وصلاحيات.
5. إضافة تخزين دائم للملفات.
6. إضافة تحليل AI حقيقي للتلخيص واكتشاف المعلومات الناقصة والمخاطر.
7. تحسين معالجة ملفات PDF الكبيرة وOCR على دفعات.

---

# 15) المرحلة التالية بعد نجاح Render

بعد ظهور رابط Render وفتح صفحة DocuAI بنجاح، يمكن تطوير النظام إلى:

```text
Upload PDF
      ↓
OCR / Text Extraction
      ↓
Document Classification
      ↓
AI Analysis
      ↓
Summary
      ↓
Missing Information
      ↓
Key Findings / Risks
      ↓
Document Control Register
      ↓
Excel / PDF Report
```

