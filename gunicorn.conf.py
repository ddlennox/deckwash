import os
bind = f"0.0.0.0:{os.environ.get('PORT', '10000')}"
workers = 1  # Single worker so in-memory job dict is shared across all requests
threads = 4
timeout = 300
