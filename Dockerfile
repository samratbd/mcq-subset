# MCQ Shuffler — production container
# Bundles Python + pandoc so KaTeX↔OMML math conversion works in production.

FROM python:3.12-slim

# System deps:
#   pandoc — for KaTeX ↔ Word equation conversion
#   libxml2 / libxslt — runtime for lxml (the wheel is usually self-contained,
#     but slim images sometimes need these for older lxml builds)
RUN apt-get update \
 && apt-get install -y --no-install-recommends pandoc \
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

# gunicorn binds to $PORT. 2 workers × 2 threads is plenty for a single-user
# tool; bump these if you have lots of concurrent users.
CMD gunicorn --workers 2 --threads 2 --timeout 120 \
    --bind 0.0.0.0:${PORT} "app.server:create_app()"
