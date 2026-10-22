FROM python:3.11-slim

# The fonts are metric twins of Calibri, Cambria, Arial, Times New Roman and Courier New:
# without them LibreOffice draws those (the fonts of most Word and PowerPoint files) with a wider
# stand-in, so a PDF made here broke its lines and pages differently from the original.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    poppler-utils \
    tesseract-ocr \
    fonts-crosextra-carlito \
    fonts-crosextra-caladea \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Every Python file in the repository, so adding a module can never break
# the build by being missing from this list.
COPY *.py ./

EXPOSE 8000
# gthread + threads instead of just adding more sync workers: this box is
# 0.5 CPU / 512MB, and each LibreOffice conversion alone can use 150-300MB,
# so spawning more full worker PROCESSES risks the OS OOM-killing the
# container under concurrent load. Threads share one process's memory, so
# 2 workers x 4 threads gives up to 8 requests in flight at roughly the
# same baseline memory as the old 2-process/no-thread setup — a slow
# conversion occupies one thread (subprocess calls release the GIL while
# waiting on the OS), not the whole worker, so quick requests (auth, chat,
# status checks) don't queue behind someone else's file conversion.
# 3.4: 6 threads per worker (12 requests in flight): meetings add many quick requests (joining, the waiting room),
# which shouldn't wait behind Remy's longer answers. LibreOffice itself runs one conversion at a time
# (SOFFICE_AT_ONCE, in converters.py), so more threads don't mean more memory for conversions.
CMD gunicorn --bind 0.0.0.0:${PORT:-8000} --timeout 240 --workers 2 --worker-class gthread --threads 6 app:app
