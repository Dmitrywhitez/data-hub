FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
RUN mkdir -p /app/uploads
COPY app.py .
COPY init.sql /docker-entrypoint-initdb.d/init.sql
CMD ["python", "app.py"]
