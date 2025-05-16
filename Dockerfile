FROM python:3.9-slim

# install tesseract + dev libs, then clean up
RUN apt-get update \
 && apt-get install -y tesseract-ocr libtesseract-dev \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /workspace
COPY . .
RUN pip install --no-cache-dir -r requirements.txt

# this tells Vercel to use your Dockerfile for api/index.py
CMD ["vercel", "python", "api/index.py"]
