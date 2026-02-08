"""
Readers Package
Shared file reading utilities for the entire project.

This package provides readers for:
- Excel files (excel_reader)
- Word documents (word_reader)
- AQA specification PDFs (aqa_reader)

Usage:
    from readers import excel_reader, word_reader, aqa_reader
    
    # Or import specific functions
    from readers.excel_reader import find_excel_file, get_lesson_data
    from readers.word_reader import find_word_document, save_readnow_docx
    from readers.aqa_reader import extract_aqa_specification_content
"""

__all__ = ["excel_reader", "word_reader", "aqa_reader"]
