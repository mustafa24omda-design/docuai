# DocuAI — Intelligent Project Document Management System

نظام ذكي لتحليل مستندات المشاريع الهندسية/الإنشائية واستخراج بياناتها
الأساسية تلقائيًا:

| الحقل | مثال |
|---|---|
| 📄 رقم المستند | `TR-PMC-00456-A` |
| 📋 نوع المستند | `Transmittal`, `RFI`, `Submittal`, `NCR`... |
| 📅 التاريخ | `2024-05-21` |
| 🔄 Revision | `02`, `A` |
| ✅ Status | `Approved`, `For Review`, `Rejected`... |
| 🏢 الجهة المرسلة | `ABC Engineering Consultants` |

يدعم النظام ملفات **PDF, DOCX, TXT**، ويصدّر النتائج إلى ملف **Excel** جاهز.

---

## 🗂️ هيكل المشروع

```
docuai/
├── docuai/
│   ├── __init__.py
│   ├── extractor.py     # محرك الاستخراج (Regex + قواعد سياقية)
│   ├── readers.py        # قراءة PDF / DOCX / TXT
│   └── cli.py            # واجهة سطر الأوامر + التصدير لـ Excel
├── app.py                 # واجهة ويب بسيطة (Flask)
├── samples/                # ملفات تجريبية للاختبار
├── requirements.txt
└── README.md
```

---

## ⚙️ التثبيت

```bash
pip install -r requirements.txt
```

---

## 🚀 طريقة الاستخدام

### 1) عبر سطر الأوامر (CLI)

معالجة ملف واحد:
```bash
python -m docuai.cli --input path/to/document.pdf --output result.xlsx
```

معالجة مجلد كامل (دفعة مستندات):
```bash
python -m docuai.cli --input path/to/folder --output results.xlsx
```

### 2) عبر واجهة الويب

```bash
python app.py
```
ثم افتح المتصفح على: `http://127.0.0.1:5000`
ارفع ملف أو أكثر، وشاهد النتائج مباشرة، مع إمكانية تنزيلها كملف Excel.

### 3) داخل كودك الخاص (كمكتبة Python)

```python
from docuai import read_any, extract_fields

text = read_any("document.pdf")
result = extract_fields(text, source_file="document.pdf")

print(result.document_number)
print(result.document_type)
print(result.date)
print(result.revision)
print(result.status)
print(result.sender)
```

---

## 🧠 كيف يعمل الاستخراج؟

الإصدار الحالي يعتمد على:
- **Regex Patterns** مبنية على تنسيقات شائعة في مستندات الترانزميتال،
  RFI، Submittal، NCR، وغيرها.
- **قاموس كلمات مفتاحية** لتصنيف نوع المستند والحالة (Status) بدقة،
  مع أولوية للبحث بجانب تسميات صريحة مثل `Status:` لتفادي الالتباس.
- **درجة ثقة تقريبية (Confidence)** بناءً على عدد الحقول التي تم
  العثور عليها بنجاح.

### 🔮 تطوير مستقبلي (اختياري)
يمكن ترقية دقة النظام لاحقًا دون تغيير بنية المشروع عبر:
- استبدال/تعزيز `extractor.py` بنموذج NLP حقيقي (spaCy / LayoutLM) أو
  استدعاء نموذج LLM (مثل Claude API) لفهم سياقي أعمق.
- إضافة OCR (مثل `pytesseract`) لدعم المستندات الممسوحة ضوئيًا (Scanned PDFs).
- ربط النتائج بقاعدة بيانات (PostgreSQL/SQLite) بدلاً من التصدير اليدوي لـ Excel.
- إضافة تصنيف تلقائي بالتعلم الآلي (بدل الكلمات المفتاحية) لدقة أعلى مع
  تنسيقات غير معروفة مسبقًا.

---

## 📌 ملاحظات

- النظام الحالي **قائم على قواعد (Rule-based)**، وهو سريع وشفاف ولا يحتاج
  اتصال إنترنت أو نموذج AI خارجي، لكنه قد يحتاج تعديل الأنماط (patterns)
  إذا اختلف تنسيق مستنداتك بشكل كبير عن الأمثلة الموجودة في `samples/`.
- لأفضل دقة، احرص أن تحتوي مستنداتك على تسميات واضحة مثل
  `Document No.`, `Status:`, `Rev:`, `From:`.
