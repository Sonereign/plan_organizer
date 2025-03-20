import logging
import os
import datetime
import traceback

# Ensure the logs directory exists
LOG_DIR = "logs"
os.makedirs(LOG_DIR, exist_ok=True)

# Define the log file name (e.g., logs/app.log)
LOG_FILE = os.path.join(LOG_DIR, "app.log")

# Custom formatter to match the desired format
class CustomFormatter(logging.Formatter):
    def format(self, record):
        log_time = datetime.datetime.now().strftime("%d-%m-%y - %H:%M:%S")
        filename = record.filename.replace(".py", "")  # Get filename without .py
        log_message = f"[{log_time}] - [{record.levelname}] - [{filename}:{record.lineno}] - {record.msg}"

        # Include exception traceback if available
        if record.exc_info:
            log_message += f"\n{traceback.format_exc()}"

        return log_message

# Create handlers
file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
console_handler = logging.StreamHandler()

# Set formatter
formatter = CustomFormatter()
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# Get a logger
logger = logging.getLogger("PlanoKratiseon")
logger.setLevel(logging.DEBUG)  # Change to INFO in production
logger.addHandler(file_handler)
logger.addHandler(console_handler)
