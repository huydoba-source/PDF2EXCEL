FROM python:3.11-slim
WORKDIR /app

# Cài đặt các thư viện hệ thống cần thiết cho nhận diện OCR (Tesseract)
RUN apt-get update && \
    apt-get install -y tesseract-ocr libgl1 && \
    rm -rf /var/lib/apt/lists/*

COPY . .
RUN pip install --no-cache-dir -r requirements.txt
EXPOSE 8501
CMD ["streamlit", "run", "src/app.py", "--server.port=8501", "--server.address=0.0.0.0"]
