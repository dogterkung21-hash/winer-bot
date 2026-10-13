FROM python:3.10-slim

# ติดตั้ง LibreOffice, Poppler และโปรแกรมจัดการฟอนต์ (fontconfig)
RUN apt-get update && apt-get install -y \
    libreoffice \
    poppler-utils \
    fontconfig \
    fonts-thai-tlwg \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# คัดลอกโฟลเดอร์ fonts ของคุณเข้าไปในระบบ Linux และสั่งลงทะเบียนฟอนต์
COPY fonts/ /usr/share/fonts/truetype/custom_fonts/
RUN fc-cache -f -v

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

CMD ["python", "ttt-2.py"]
