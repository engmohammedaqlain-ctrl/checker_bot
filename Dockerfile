FROM python:3.12-slim

# LibreOffice (لتحويل ملفات Word إلى PDF) + خطوط عربية و لاتينية للعرض الصحيح
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-writer \
        fonts-noto \
        fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# تثبيت اعتماديات بايثون أولاً (طبقة كاش أسرع)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# نسخ باقي الكود
COPY . .

# البوت يعمل بنظام long polling، والمنفذ للرابط الوسيط (تعليم حالة الإرسال)
EXPOSE 8080
CMD ["python", "main.py"]
