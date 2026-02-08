"""
Logging configuration for ReadNow generators
Uses rich.logging for beautiful formatted output
"""

import logging

from rich.console import Console
from rich.logging import RichHandler

# Create console for rich output
console = Console()


def setup_logging(level=logging.INFO):
    """
    Set up logging with rich handler for beautiful formatted output.

    Args:
        level: Logging level (default: INFO)
    """
    logging.basicConfig(
        level=level, format="%(message)s", datefmt="[%X]", handlers=[RichHandler(console=console, rich_tracebacks=True)]
    )
    return logging.getLogger("readnow")


# Create default logger
logger = setup_logging()
