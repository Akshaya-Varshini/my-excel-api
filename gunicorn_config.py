import os
import multiprocessing

port = int(os.environ.get("PORT", 8000))
bind = f"0.0.0.0:{port}"
workers = min(4, multiprocessing.cpu_count() * 2 + 1)
worker_class = "sync"
timeout = 120
keepalive = 30
max_requests = 1000
max_requests_jitter = 100
preload_app = True
loglevel = "info"
accesslog = "-"
errorlog = "-"
