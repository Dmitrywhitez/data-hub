FROM python:3.11-slim

# Устанавливаем postgresql-client для работы docker-entrypoint.sh
RUN apt-get update && apt-get install -y postgresql-client && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
RUN mkdir -p /app/uploads

COPY app.py .
COPY docker-entrypoint.sh /docker-entrypoint.sh
RUN chmod +x /docker-entrypoint.sh

CMD ["/bin/bash", "-c", "/docker-entrypoint.sh && gunicorn -w 3 -t 60 -b 0.0.0.0:5000 app:app"]
