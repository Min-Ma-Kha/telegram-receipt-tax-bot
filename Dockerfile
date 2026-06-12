# For cloud deployment (Railway, Render, any Docker host).
FROM python:3.12-slim

RUN apt-get update \
    && apt-get install -y --no-install-recommends tesseract-ocr \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

# Mount a volume at /app/data so receipts.xlsx survives redeploys.
VOLUME ["/app/data"]

# Health check + /backup endpoint for the PC sync.
EXPOSE 8080

CMD ["python", "bot.py"]
