FROM python:3.12-slim

# Chromium needs a few runtime libs; playwright installs the rest with --with-deps
ENV PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PLAYWRIGHT_BROWSERS_PATH=/ms-playwright

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt \
    && playwright install --with-deps chromium

COPY . .

# host injects PORT; default to 8080
ENV PORT=8080
EXPOSE 8080

CMD ["python", "serve.py"]
