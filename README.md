# dashboard-project 
# dashboard-project 

# 📊 Excel Data Dashboard using Flask, Pandas, Chart.js, and Azure

## 📝 Overview
هذا المشروع عبارة عن **لوحة معلومات (Dashboard)** تفاعلية تقوم بقراءة وتحليل ملفات **Excel**، واستخراج مؤشرات الأداء (KPIs)، وعرض رسوم بيانية تفاعلية، ومعاينة البيانات—وذلك باستخدام **Flask (Python)** في الخلفية و **HTML/CSS/JavaScript/Jinja** في الواجهة الأمامية.  
تم نشر المشروع بنجاح على **Azure Web App** باستخدام **GitHub Actions CI/CD**.

---

## 🚀 Features
- رفع ملف Excel مباشرة من المتصفح.
- تحليل البيانات باستخدام **Pandas** و **openpyxl**.
- استخراج KPIs:
  - عدد الصفوف  
  - عدد الأعمدة  
  - خلايا فارغة  
  - صفوف مكررة  
- اكتشاف تلقائي لأعمدة:
  - التاريخ Date  
  - القيم الرقمية Numeric  
  - الفئات Category  
- عرض:
  - Bar Chart لأعلى القيم تكرارًا  
  - Line Chart للاتجاه الزمني  
- معاينة أول 20 صفًا من البيانات.
- واجهة مستخدم احترافية باستخدام **Bootstrap 5**.
- رسم المخططات باستخدام **Chart.js**.
- استخدام **Jinja2** لربط الواجهة بالباكند.
- نشر تلقائي على **Azure Web App** عبر **GitHub Actions**.

---

## 🏗 Tech Stack

### 🔧 Backend
- Python  
- Flask  
- Werkzeug  
- Pandas  
- openpyxl  
- Gunicorn (للعمل في Production)  

### 🎨 Frontend
- HTML5  
- CSS3  
- Bootstrap 5  
- JavaScript (ES6)  
- Chart.js  
- Jinja2 Templates  

### ☁ DevOps & Cloud
- GitHub (بحسابين منفصلين Frontend + Backend)  
- GitHub Actions (CI/CD)  
- Azure Web App (Linux)  

---

## 📂 Project Structure
