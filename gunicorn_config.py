import os

# Number of worker processes
# 4 workers × 2 threads = 8 concurrent request handlers
workers = int(os.environ.get("WEB_CONCURRENCY", 4))
threads = 2

# Bind to all interfaces on the port Render provides (default 5000)
bind = f"0.0.0.0:{os.environ.get('PORT', '5000')}"

# Logging
loglevel = "info"
accesslog = "-"
errorlog = "-"

# Timeout (seconds) — increase if PPT generation takes a while
timeout = 120

# Preload the app for faster worker startup
preload_app = True
