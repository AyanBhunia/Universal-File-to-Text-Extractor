FROM python:3.9-slim
RUN apt-get update && \
    apt-get install -y tesseract-ocr libtesseract-dev && \
    rm -rf /var/lib/apt/lists/*
WORKDIR /workspace
COPY . .
RUN pip install --no-cache-dir -r requirements.txt
CMD ["vercel", "python", "api/index.py"]
