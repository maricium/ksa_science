"""
AQA Specification Reader
Consolidated utilities for reading and processing AQA specification PDFs

This module provides reusable functions for:
- Finding AQA specification PDF files
- Reading and extracting text from PDFs
- Using AI to identify relevant sections
- Extracting definitions and keywords from specification content

Usage:
    from readers.aqa_reader import extract_aqa_specification_content

    # Extract AQA content for a lesson
    aqa_content, aqa_keywords = extract_aqa_specification_content(
        "Ionic Bonding",
        "Students will understand how ionic bonds form",
        "C4.2"
    )
"""

from __future__ import annotations

import logging
import re
from functools import lru_cache
from pathlib import Path

logger = logging.getLogger("readers.aqa_reader")

# Try to import PyMuPDF
try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None


@lru_cache(maxsize=1)
def find_aqa_pdf() -> Path | None:
    """
    Find the AQA specification PDF file.

    Returns:
        Path to PDF file or None if not found
    """
    aqa_pdf_path = Path("AQA-8464-SP-2016.PDF")
    if not aqa_pdf_path.exists():
        aqa_pdf_path = Path("..") / "AQA-8464-SP-2016.PDF"
    if not aqa_pdf_path.exists():
        aqa_pdf_path = Path("../..") / "AQA-8464-SP-2016.PDF"

    if aqa_pdf_path.exists():
        return aqa_pdf_path

    return None


@lru_cache(maxsize=1)
def _extract_text_from_pdf_cached(pdf_path_str: str, max_pages: int) -> tuple[str, ...]:
    """
    Cached version of extract_text_from_pdf (internal use).
    Returns tuple for hashing.

    Args:
        pdf_path_str: Path to PDF file as string (for hashing)
        max_pages: Maximum number of pages to process

    Returns:
        Tuple of text chunks (converted from list for hashing)
    """
    if not fitz:
        return ()

    try:
        doc = fitz.open(pdf_path_str)

        # Check if PDF is encrypted
        if doc.is_encrypted and not doc.authenticate(""):
            logger.warning("AQA PDF is encrypted and cannot be read")
            doc.close()
            return ()

        logger.info(f"Reading AQA specification PDF ({len(doc)} pages)...")

        # Extract text from PDF in chunks (by page or section)
        specification_chunks = []
        current_chunk = ""
        chunk_size = 0

        for page_num in range(min(len(doc), max_pages)):
            try:
                page = doc[page_num]
                text = page.get_text()
                if not text or len(text.strip()) < 50:
                    continue

                # Clean up text
                text = re.sub(r"\s+", " ", text)  # Normalize whitespace

                # Build chunks (approximately 2000 chars each for AI processing)
                if chunk_size + len(text) > 2000 and current_chunk:
                    specification_chunks.append(current_chunk.strip())
                    current_chunk = text + "\n"
                    chunk_size = len(text)
                else:
                    current_chunk += text + "\n"
                    chunk_size += len(text)
            except Exception:
                continue

        # Add final chunk
        if current_chunk.strip():
            specification_chunks.append(current_chunk.strip())

        doc.close()
        return tuple(specification_chunks)  # Convert to tuple for hashing
    except Exception as e:
        logger.warning(f"Error reading PDF: {e}")
        return ()


def extract_text_from_pdf(pdf_path: Path, max_pages: int = 300) -> list[str]:
    """
    Extract text from PDF and split into chunks for processing.

    Args:
        pdf_path: Path to PDF file
        max_pages: Maximum number of pages to process

    Returns:
        List of text chunks (approximately 2000 chars each)
    """
    # Use cached version and convert tuple back to list
    chunks_tuple = _extract_text_from_pdf_cached(str(pdf_path.absolute()), max_pages)
    return list(chunks_tuple)


def process_chunks_with_ai(
    chunks: list[str], lesson_title: str, first_objective: str, client, batch_size: int = 10
) -> tuple[list[str], list[str]]:
    """
    Use AI to identify relevant chunks from the AQA specification.

    Args:
        chunks: List of text chunks from the PDF
        lesson_title: Lesson title
        first_objective: First learning objective
        client: OpenAI client
        batch_size: Number of chunks to process at once

    Returns:
        Tuple of (relevant_chunks, keywords)
    """
    if not client:
        logger.warning("OpenAI client not available - cannot use AI to find relevant specification sections")
        return [], []

    relevant_chunks = []
    keywords = []

    logger.info(f"Extracted {len(chunks)} chunks from specification")
    logger.info("Using AI to identify relevant sections...")

    # Process chunks in batches to avoid token limits
    for i in range(0, len(chunks), batch_size):
        batch = chunks[i : i + batch_size]
        batch_text = "\n\n---CHUNK SEPARATOR---\n\n".join(
            [f"CHUNK {i + j + 1}:\n{chunk[:1500]}" for j, chunk in enumerate(batch)]
        )

        ai_prompt = f"""You are analyzing the AQA GCSE Science specification (8464) to find content relevant to this lesson:

LESSON: {lesson_title}
OBJECTIVE: {first_objective}

Below are {len(batch)} chunks from the AQA specification. Identify which chunks contain content directly relevant to this lesson.

For each relevant chunk, extract:
1. The exact definitions and key terms
2. The specific content that matches the lesson topic

Return ONLY the relevant chunks with their content. Format:
RELEVANT CHUNK X:
[exact content from that chunk]

If no chunks are relevant, return "NO RELEVANT CONTENT"

SPECIFICATION CHUNKS:
{batch_text}"""

        try:
            response = client.chat.completions.create(
                model="gpt-4o-mini", messages=[{"role": "user", "content": ai_prompt}], temperature=0.3, max_tokens=2000
            )

            ai_response = response.choices[0].message.content

            # Extract relevant chunks from AI response
            if "NO RELEVANT CONTENT" not in ai_response.upper():
                # Find all chunks marked as relevant
                chunk_pattern = r"RELEVANT CHUNK\s+(\d+):\s*(.*?)(?=RELEVANT CHUNK|\Z)"
                matches = re.finditer(chunk_pattern, ai_response, re.DOTALL | re.IGNORECASE)
                for match in matches:
                    chunk_content = match.group(2).strip()
                    if chunk_content and len(chunk_content) > 50:
                        relevant_chunks.append(chunk_content)

                        # Extract keywords from the chunk
                        keyword_matches = re.findall(
                            r"\b([A-Z][a-z]+(?:\s+[a-z]+)?)\s+(?:is|means|refers to|called)",
                            chunk_content,
                            re.IGNORECASE,
                        )
                        for kw in keyword_matches:
                            if kw and kw not in keywords and len(kw) < 30:
                                keywords.append(kw)
        except Exception as e:
            logger.warning(f"Error processing chunk batch: {e}")
            continue

    return relevant_chunks, keywords


def extract_definitions_from_chunks(relevant_chunks: list[str]) -> list[str]:
    """
    Extract definitions from relevant chunks using pattern matching.

    Args:
        relevant_chunks: List of relevant text chunks

    Returns:
        List of definition strings
    """
    definitions = []
    for chunk in relevant_chunks:
        # Find definition patterns
        def_patterns = [
            r"([A-Z][a-zA-Z\s]+?)\s+is\s+([^\.]+)",
            r"([A-Z][a-zA-Z\s]+?)\s+means\s+([^\.]+)",
            r"([A-Z][a-zA-Z\s]+?)\s+refers to\s+([^\.]+)",
            r"([A-Z][a-zA-Z\s]+?)\s+called\s+([^\.]+)",
        ]

        for pattern in def_patterns:
            matches = re.finditer(pattern, chunk, re.IGNORECASE)
            for match in matches:
                term = match.group(1).strip()
                definition = match.group(2).strip()
                if 10 < len(definition) < 200:
                    definitions.append(f"{term} is {definition}")

    return definitions


@lru_cache(maxsize=128)
def _extract_aqa_specification_content_cached(
    lesson_title: str, first_objective: str, unit_code: str | None
) -> tuple[str | None, tuple[str, ...]]:
    """
    Cached version of extract_aqa_specification_content (internal use).
    Caches the full result including AI processing.
    Note: client is not included in cache key since it's not hashable and doesn't change.

    Args:
        lesson_title: Title of the lesson
        first_objective: First learning objective
        unit_code: Optional unit code

    Returns:
        Tuple of (content_string, keywords_tuple) or (None, tuple())
    """
    # Import client here to avoid circular dependency
    try:
        from readnow.main import client as main_client

        client = main_client
    except ImportError:
        logger.warning("OpenAI client not available")
        return None, ()

    if not client:
        logger.warning("OpenAI client not available")
        return None, ()

    # Call the actual implementation
    content, keywords = _extract_aqa_specification_content_impl(lesson_title, first_objective, unit_code, client)

    # Convert keywords list to tuple for hashing
    keywords_tuple = tuple(keywords) if keywords else ()
    return content, keywords_tuple


def _extract_aqa_specification_content_impl(
    lesson_title: str, first_objective: str, unit_code: str | None, client
) -> tuple[str | None, list[str]]:
    """
    Internal implementation of extract_aqa_specification_content (not cached).
    """
    # Check if PyMuPDF is available
    if not fitz:
        logger.warning("PyMuPDF not available - cannot read AQA specification PDF")
        return None, []

    # Find PDF file
    aqa_pdf_path = find_aqa_pdf()
    if not aqa_pdf_path:
        logger.warning("AQA specification PDF not found")
        return None, []

    try:
        # Extract text chunks from PDF (this is already cached)
        specification_chunks = extract_text_from_pdf(aqa_pdf_path)
        if not specification_chunks:
            return None, []

        # Process chunks with AI to find relevant sections
        relevant_chunks, keywords = process_chunks_with_ai(specification_chunks, lesson_title, first_objective, client)

        if relevant_chunks:
            # Combine all relevant chunks
            combined_content = "\n\n".join(relevant_chunks)

            # Extract definitions more systematically
            definitions = extract_definitions_from_chunks(relevant_chunks)

            # Prioritize definitions, then add other relevant content
            final_content = []
            if definitions:
                final_content.extend(definitions[:20])
            final_content.append(combined_content[:2000])  # Add remaining content

            unique_keywords = list(dict.fromkeys(keywords))[:30]  # Remove duplicates, keep order

            logger.info(f"Found {len(relevant_chunks)} relevant sections with {len(definitions)} definitions")

            return "\n\n".join(final_content)[:4000], unique_keywords  # Limit to 4000 chars

        logger.warning("No relevant sections found in AQA specification")
        return None, []

    except Exception as e:
        logger.error(f"Error reading AQA specification: {e}")
        import traceback

        traceback.print_exc()
        return None, []


def extract_aqa_specification_content(
    lesson_title: str, first_objective: str, unit_code: str | None = None, client=None
) -> tuple[str | None, list[str]]:
    """
    Extract relevant content from AQA specification PDF using AI to identify relevant sections.
    Results are cached based on lesson_title, first_objective, and unit_code.

    Args:
        lesson_title: Title of the lesson
        first_objective: First learning objective
        unit_code: Optional unit code (e.g., "C4.2")
        client: OpenAI client (if None, will try to import from main) - not used for cache key

    Returns:
        Tuple of (content_string, keywords_list) or (None, []) if error
    """
    # Use cached version (client is not part of cache key since it doesn't change)
    content, keywords_tuple = _extract_aqa_specification_content_cached(lesson_title, first_objective, unit_code)

    # Convert tuple back to list
    keywords = list(keywords_tuple) if keywords_tuple else []
    return content, keywords
