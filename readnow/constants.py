"""
Constants for ReadNow generators (HPA and LPA)
All shared styling and formatting values
"""

from docx.shared import Inches, Pt, RGBColor

# Font settings
FONT_NAME = "Century Gothic"  # Default font for all documents
FONT_SIZE_TITLE = Pt(14)  # Title font size
FONT_SIZE_HEADING = Pt(16)  # Heading font size
FONT_SIZE_BODY = Pt(12)  # Body text font size
FONT_SIZE_BODY_SMALL = Pt(11)  # Small body text (for HPA)

# Margin settings
MARGIN_TOP_BOTTOM = Inches(0.5)  # Top and bottom margins
MARGIN_LEFT_RIGHT = Inches(0.5)  # Left and right margins (standard)
MARGIN_LEFT_RIGHT_HPA = Inches(0.75)  # Left and right margins for HPA (narrower)

# Image settings
IMAGE_WIDTH = Inches(5.0)  # Image width in documents

# Tab stop settings
TAB_STOP_POSITION = Inches(6)  # Position for date tab stop

# Color settings
HEADING_COLOR = RGBColor(0, 51, 102)  # Dark blue for headings

# Line spacing
LINE_SPACING = 1.5  # Line spacing multiplier

# Date format
DATE_PLACEHOLDER = "Date: _______________"  # Static date placeholder
# Note: For dynamic dates, use DATE field in Word (see unified_generator.py)

# Quick config - change these 8 lines to adjust for different years/reading ages and file locations
STUDENT_YEAR = "Year 9"  # e.g., "Year 7", "Year 8", "Year 9", "Year 10-11"
STUDENT_ATTAINMENT = "LPA"  # "HPA" or "LPA"
READING_AGE = "9 years old"  # e.g., "7 years old", "8 years old", "GCSE level", "14 years old"
WORD_COUNT = "100-120"  # e.g., "80-100", "100-120", "120-150", "140-160"
LANGUAGE_STYLE = "simple language, use short sentences (max 12 words), everyday examples, avoid too many technical terms unless they are in the lesson objectives and defined in the read now content"  # Adjust complexity here

# File location config - change these to point to your PowerPoint files
POWERPOINT_FOLDER = (
    "Lesson Resources"  # Folder containing PowerPoint files (consolidated from mymastery and Lesson Resources)
)
SLIDE_NUMBER = 6  # Which slide to extract objectives from (usually slide 6)
LESSON_CODE_PREFIX = "C5.1."  # Filter lessons by prefix (e.g., "C4.1.", "C3.1.", "C3.2.", "B3.2.", "C5.1.")
EXCEL_FOLDER = "Excel_Files"  # Folder containing all Excel files (consolidated from C3_Excel, C4_Excel, etc.)
