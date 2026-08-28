FROM python:3.11-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    poppler-utils \
    tesseract-ocr \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY converters.py app.py ./

EXPOSE 8000
CMD gunicorn --bind 0.0.0.0:${PORT:-8000} --timeout 120 --workers 2 app:app
