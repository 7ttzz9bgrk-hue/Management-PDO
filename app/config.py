import os

FILE_PATHS = [
    "mockData.xlsx",
    # Add more file paths here:
    # "//server/shared/projects.xlsx",
    # "another.xlsx",
]

DEBOUNCE_SECONDS = 3
MAX_RELOAD_RETRIES = 3
RELOAD_RETRY_DELAY = 1.0
READ_RETRY_DELAY = 0.3
READ_RETRY_ATTEMPTS = 3
SSE_POLL_SECONDS = 0.5
SSE_KEEPALIVE_SECONDS = 25

APP_HOST = "127.0.0.1"
APP_PORT = 8889

# Keep extension handling centralized so routes/watcher stay consistent.
ALLOWED_EXCEL_EXTENSIONS = {".xlsx", ".xlsm"}
