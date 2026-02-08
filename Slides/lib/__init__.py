"""Library modules for slide generation"""

from .processor import create_slide_from_template
from .reader import read_lesson_data, read_markscheme

__all__ = ['create_slide_from_template', 'read_lesson_data', 'read_markscheme']

