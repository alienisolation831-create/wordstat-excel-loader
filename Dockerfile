FROM python:3.13-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt && useradd --uid 10001 --create-home monitor && mkdir /data && chown monitor:monitor /data
COPY app.py .
COPY static ./static
USER monitor
ENV DATA_DIR=/data PYTHONUNBUFFERED=1
EXPOSE 8000
CMD ["uvicorn", "app:create_app", "--factory", "--host", "0.0.0.0", "--port", "8000", "--workers", "1", "--no-proxy-headers", "--no-access-log", "--limit-concurrency", "32", "--timeout-keep-alive", "5"]
