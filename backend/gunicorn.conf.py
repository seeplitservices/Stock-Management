import os

# Google Cloud Run injects PORT (defaults to 8080)
port = os.environ.get("PORT", "8080")
bind = f"0.0.0.0:{port}"

# Memory-optimized Gunicorn configuration for Render Free (512 MB memory limit)
# Single worker process with multi-threading prevents process memory multiplication.
workers = int(os.environ.get("WEB_CONCURRENCY", "1"))
threads = int(os.environ.get("GUNICORN_THREADS", "4"))
worker_class = "gthread"

timeout = int(os.environ.get("GUNICORN_TIMEOUT", "120"))
graceful_timeout = 30
keepalive = 5

# Memory management & leak prevention
# Automatically recycles worker process after 500 requests to free fragmented heap
max_requests = 500
max_requests_jitter = 50

accesslog = "-"
errorlog = "-"
loglevel = os.environ.get("LOG_LEVEL", "info")
