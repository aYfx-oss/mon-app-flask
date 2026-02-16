import os

# Timeout de 6 minutes (360s) pour gérer les 3 appels API
timeout = 360
graceful_timeout = 360
keepalive = 5

# Worker configuration
workers = 1
worker_class = "sync"

# Bind
bind = f"0.0.0.0:{os.getenv('PORT', '5000')}"

# Logging
accesslog = "-"
errorlog = "-"
loglevel = "info"
