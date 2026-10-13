# ใช้ Python เวอร์ชันที่เป็นทางการและเสถียร
FROM python:3.10-slim

# อัปเดตระบบและติดตั้ง LibreOffice, Poppler (สำหรับแปลง PDF/รูปภาพ) และฟอนต์ภาษาไทย
RUN apt-get update && apt-get install -y \
    libreoffice \
    poppler-utils \
    fonts-thai-tlwg \
    && rm -rf /var/lib/apt/lists/*

# กำหนดโฟลเดอร์ทำงานหลักใน Container
WORKDIR /app

# คัดลอกไฟล์ requirements.txt และติดตั้งไลบรารี Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# คัดลอกไฟล์โค้ดและไฟล์ Excel ทั้งหมดเข้ามาใน Container
COPY . .

# คำสั่งสำหรับเริ่มรันบอท
CMD ["python", "ttt.py"]