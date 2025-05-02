import logging
from datetime import datetime
import os

# Create logs directory if it doesn't exist
os.makedirs("logs", exist_ok=True)

# Configure logging
def setup_logger():
    """Set up and configure the application logger."""
    logger = logging.getLogger("keytrack")
    logger.setLevel(logging.INFO)
    
    # Create handlers
    file_handler = logging.FileHandler(f"logs/keytrack_{datetime.now().strftime('%Y%m%d')}.log")
    console_handler = logging.StreamHandler()
    
    # Create formatters
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    # Add handlers to logger
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger

# Initialize logger
logger = setup_logger()

# Convenience functions
def info(message):
    """Log an info message."""
    logger.info(message)

def error(message):
    """Log an error message."""
    logger.error(message)

def warning(message):
    """Log a warning message."""
    logger.warning(message)

def debug(message):
    """Log a debug message."""
    logger.debug(message)