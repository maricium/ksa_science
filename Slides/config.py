"""Configuration for slide generator"""

from pathlib import Path

# Paths
TEMPLATE_NAME = "ksa_template.pptx"  # Located in templates/ folder
OUTPUT_DIR = "output"
LESSON_RESOURCES_DIR = "Lesson Resources"
READNOW_DIR = "readnow/readnows"

# Target lesson (set to None to process all lessons)
TARGET_LESSON = "c4.2.14"

# Unit filters (set to None to process all units)
TARGET_UNIT = "C4.2"  # Set to None to process all units

# Slide configuration
MARKSCHEME_SLIDE = 3
OBJECTIVES_SLIDE = 7
EXIT_TICKET_SLIDE = 13
OBJECTIVES_POSITION = {"left": 1.22, "top": 0.3, "width": 11.2, "height": 5.19}
MARKSCHEME_POSITION = {"left": 1.07, "top": 0.34, "width": 11.2, "height": 6.5}
EXIT_TICKET_POSITION = {"left": 1.22, "top": 0.3, "width": 11.2, "height": 6.0}
OBJECTIVES_FONT_SIZE = 28
MARKSCHEME_FONT_SIZE = 24
EXIT_TICKET_FONT_SIZE = 21

# Default sheet name (used as fallback if auto-detection fails)
DEFAULT_SHEET = "Sheet1"

