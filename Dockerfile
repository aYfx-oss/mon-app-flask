FROM python:3.11-slim

RUN apt-get update && apt-get install -y \
    libxml2 \
    libxslt1.1 \
    libxml2-dev \
    libxslt-dev \
    gcc \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY backend/ .

RUN pip install --no-cache-dir -r requirements.txt

EXPOSE 8080
CMD ["gunicorn", "app:app", "--bind", "0.0.0.0:8080"]
