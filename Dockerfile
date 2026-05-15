# MCQ Shuffler — production container
# Bundles Python + pandoc + LibreOffice so all output formats work in production.

FROM python:3.12-slim

# System deps:
#   pandoc — for KaTeX ↔ Word equation conversion
#   libreoffice-writer — for PDF output (renders docx → pdf)
#   fonts-noto-* — Unicode coverage for Bengali, math symbols, etc.
RUN apt-get update \
 && apt-get install -y --no-install-recommends \
        pandoc \
        libreoffice-writer libreoffice-core \
        fonts-noto-core fonts-noto-cjk fonts-noto-extra \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python deps first so this layer caches well across code edits.
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Now copy the rest of the app.
COPY . .

# DO App Platform sets $PORT; default to 8080 for local docker runs.
ENV PORT=8080
EXPOSE 8080

# PDF generation can take 10-30s per set, so give gunicorn a generous timeout.
CMD gunicorn --workers 2 --threads 2 --timeout 180 \
    --bind 0.0.0.0:${PORT} "app.server:create_app()"
