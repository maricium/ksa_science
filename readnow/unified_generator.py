#!/usr/bin/env python3
"""
Unified ReadNow Generator
Generates ReadNows for any year/reading age - just change constants.py
"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path

# Add project root to path to allow imports from readers module
_project_root = Path(__file__).parent.parent
if str(_project_root) not in sys.path:
    sys.path.insert(0, str(_project_root))

from logging_config import setup_logging
from main import client, download_chemistry_image, extract_slide6_objectives, get_first_objective

# Set up logging
logger = setup_logging()
# Word document imports moved to word_helpers
# Only keep docx imports if still needed elsewhere (they're not)
# PyMuPDF (fitz) import moved to aqa_helpers

from constants import (
    LANGUAGE_STYLE,
    LESSON_CODE_PREFIX,
    POWERPOINT_FOLDER,
    READING_AGE,
    STUDENT_ATTAINMENT,
    STUDENT_YEAR,
    WORD_COUNT,
)
from readers.aqa_reader import extract_aqa_specification_content
from readers.excel_reader import (
    extract_keyword_definition,
    find_excel_file,
    find_keyword_column,
    get_keywords_from_cell,
    get_lesson_data,
    get_lesson_row_data,
    read_excel_with_openpyxl,
)
from readers.word_reader import (
    get_keyword_definition_from_word_doc,
    get_subject_from_code,
    save_readnow_docx,
)

# extract_aqa_specification_content is now imported from readers.aqa_reader


@dataclass
class LessonContext:
    """Container for all lesson-related data needed to generate ReadNows."""

    lesson_code: str
    lesson_title: str
    first_objective: str
    unit_code: str | None = None

    # Keywords and definitions
    keywords: list[str] | None = None
    most_important_keyword: str | None = None
    keyword_definition: str | None = None
    full_know_content: str | None = None

    # AQA specification content
    aqa_content: str | None = None
    aqa_keywords: list[str] | None = None

    # Lesson sequence context
    previous_lesson: dict | None = None
    next_lesson: dict | None = None

    def __post_init__(self):
        """Set unit_code from lesson_code if not provided."""
        if self.unit_code is None:
            self.unit_code = ".".join(self.lesson_code.split(".")[:2])

        # Initialize empty lists if None
        if self.keywords is None:
            self.keywords = []
        if self.aqa_keywords is None:
            self.aqa_keywords = []


@lru_cache(maxsize=256)
def get_most_important_keyword_and_definition(lesson_code, unit=None):
    """
    Extract the most important keyword from the first bullet point of "What do my students need to know by the end of the lesson?"
    Then get its definition from the unit preparation Excel file.
    Returns: (keyword, definition, full_know_content) tuple or (None, None, None)
    """
    try:
        if not unit:
            unit = ".".join(lesson_code.split(".")[:2])

        # Get lesson data from Excel - try the standard function first
        lesson_data = get_lesson_data(lesson_code, unit)

        # If that doesn't work, try reading directly from Excel with header=4
        if not lesson_data or not lesson_data.get("know"):
            # Try to read directly from Excel file using helper
            excel_file = find_excel_file(unit, lesson_code)
            if excel_file:
                try:
                    from readers.excel_reader import read_excel_with_pandas

                    # Read with header=4 (where the column names are)
                    df = read_excel_with_pandas(excel_file, header=4)

                    if df is not None:
                        import pandas as pd

                        # Find the lesson row
                        know_col = None
                        code_col = None
                        for col in df.columns:
                            col_str = str(col).lower()
                            if "lesson" in col_str and "code" in col_str:
                                code_col = col
                            elif "know" in col_str and ("end" in col_str or "lesson" in col_str):
                                know_col = col

                        if code_col and know_col:
                            for _idx, row in df.iterrows():
                                if (
                                    pd.notna(row[code_col])
                                    and str(row[code_col]).strip().upper() == lesson_code.upper()
                                ):
                                    know_content = str(row[know_col]).strip() if pd.notna(row[know_col]) else ""
                                    if know_content and know_content.lower() not in ["none", "nan", ""]:
                                        lesson_data = {"know": know_content}
                                        break
                except Exception as e:
                    logger.warning(f"Error reading Excel directly: {e}")

        if not lesson_data or not lesson_data.get("know"):
            logger.warning(f"No 'know' content found for {lesson_code}")
            return None, None, None

        know_content = lesson_data.get("know", "").strip()
        if not know_content or know_content.lower() in ["none", "nan", ""]:
            return None, None, None

        # Extract first bullet point
        import re

        # Split by bullet points (•, -, *, or numbered)
        bullets = re.split(r"[•\-\*]|\d+\.", know_content)
        first_bullet = bullets[0].strip() if bullets else know_content.split("\n")[0].strip()

        if not first_bullet or len(first_bullet) < 10:
            first_bullet = know_content.split("\n")[0].strip()

        if not first_bullet:
            return None, None

        logger.info(f"First bullet point: {first_bullet[:80]}...")

        # Use AI to identify the most important keyword from the first bullet
        if not client:
            logger.warning("OpenAI client not available for keyword extraction")
            return None, None

        ai_prompt = f"""From this learning objective, identify the SINGLE most important keyword/term that students need to understand:

"{first_bullet}"

Return ONLY the keyword/term (1-3 words maximum). This should be the most critical concept that students must understand.
Examples: "pyramid of biomass", "biodiversity", "electrolysis", "ionic bonding"

Return just the keyword, nothing else."""

        try:
            response = client.chat.completions.create(
                model="gpt-4o-mini", messages=[{"role": "user", "content": ai_prompt}], temperature=0.3, max_tokens=50
            )

            important_keyword = response.choices[0].message.content.strip()
            # Clean up the response (remove quotes, extra text)
            important_keyword = re.sub(r'^["\']|["\']$', "", important_keyword)
            important_keyword = important_keyword.split("\n")[0].strip()

            if not important_keyword or len(important_keyword) < 2:
                return None, None, None

            logger.info(f"Most important keyword: {important_keyword}")

            # First try to get the definition from unit preparation Word document
            definition = get_keyword_definition_from_word_doc(important_keyword, unit)

            # If not found in Word doc, try Excel as fallback
            if not definition:
                definition = get_keyword_definition_from_excel(important_keyword, lesson_code, unit)

            # Also get the full "know" content for HPA - they need the complete definition
            full_know_content = know_content  # The entire "What do my students need to know" content

            if definition:
                logger.info("Found definition from unit preparation")
                return important_keyword, definition, full_know_content
            else:
                logger.warning("No definition found in unit preparation documents")
                return important_keyword, None, full_know_content

        except Exception as e:
            logger.warning(f"Error using AI to extract keyword: {e}")
            return None, None, None

    except Exception as e:
        logger.error(f"Error getting most important keyword: {e}")
        import traceback

        traceback.print_exc()
        return None, None, None


# get_keyword_definition_from_word_doc is now imported from readers.word_reader


@lru_cache(maxsize=128)
def get_keyword_definition_from_excel(keyword, lesson_code, unit=None):
    """Get the definition for a specific keyword from the unit preparation Excel file (Vocabulary and literacy section)"""
    try:
        if not unit:
            unit = ".".join(lesson_code.split(".")[:2])

        # Find Excel file using helper
        excel_file = find_excel_file(unit, lesson_code)
        if not excel_file:
            return None

        # Read Excel file
        wb = read_excel_with_openpyxl(excel_file)
        if not wb:
            return None

        ws = wb.active

        # Find keyword column using helper
        keyword_col_idx, lesson_code_col_idx, header_row = find_keyword_column(ws)
        if keyword_col_idx is None:
            return None

        # Find the lesson row and extract keyword definition
        data_start_row = 7 if header_row == 6 else 6
        row_data = get_lesson_row_data(ws, lesson_code, lesson_code_col_idx, data_start_row)

        if row_data and len(row_data) > keyword_col_idx and row_data[keyword_col_idx]:
            keyword_text = str(row_data[keyword_col_idx]).strip()
            if keyword_text and keyword_text.lower() not in ["none", "nan", ""]:
                # Extract definition using helper
                return extract_keyword_definition(keyword_text, keyword)

        return None
    except Exception as e:
        logger.warning(f"Error extracting keyword definition from Excel: {e}")
        return None


def get_keywords_from_excel(lesson_code, unit=None):
    """Extract keywords from unit preparation Excel file (Vocabulary and literacy section)"""
    try:
        if not unit:
            unit = ".".join(lesson_code.split(".")[:2])

        # Find Excel file using helper
        excel_file = find_excel_file(unit, lesson_code)
        if not excel_file:
            return []

        # Read Excel file
        wb = read_excel_with_openpyxl(excel_file)
        if not wb:
            return []

        ws = wb.active

        # Find keyword column using helper
        keyword_col_idx, lesson_code_col_idx, header_row = find_keyword_column(ws)
        if keyword_col_idx is None:
            return []

        # Find the lesson row and extract keywords
        data_start_row = 7 if header_row == 6 else 6
        row_data = get_lesson_row_data(ws, lesson_code, lesson_code_col_idx, data_start_row)

        if row_data and len(row_data) > keyword_col_idx and row_data[keyword_col_idx]:
            keyword_text = str(row_data[keyword_col_idx]).strip()
            # Extract keywords using helper
            return get_keywords_from_cell(keyword_text)

        return []
    except Exception as e:
        logger.warning(f"Error extracting keywords from Excel: {e}")
        return []


def get_previous_lesson_content(lesson_code, unit=None):
    """Get the previous lesson's content to understand what students already know"""
    try:
        if not unit:
            unit = ".".join(lesson_code.split(".")[:2])

        lesson_num = int(lesson_code.split(".")[-1])
        if lesson_num <= 1:
            return None  # No previous lesson

        prev_code = f"{unit}.{lesson_num - 1}"
        lesson_data = get_lesson_data(prev_code, unit)

        if lesson_data:
            return {
                "code": prev_code,
                "title": lesson_data.get("title", ""),
                "know": lesson_data.get("know", ""),
                "do": lesson_data.get("do", ""),
            }
        return None
    except Exception:
        return None


def get_next_lesson_content(lesson_code, unit=None):
    """Get the next lesson's content to avoid teaching concepts that come later"""
    try:
        if not unit:
            unit = ".".join(lesson_code.split(".")[:2])

        lesson_num = int(lesson_code.split(".")[-1])
        next_code = f"{unit}.{lesson_num + 1}"
        lesson_data = get_lesson_data(next_code, unit)

        if lesson_data:
            return {
                "code": next_code,
                "title": lesson_data.get("title", ""),
                "know": lesson_data.get("know", ""),
                "do": lesson_data.get("do", ""),
            }
        return None
    except Exception:
        return None


def generate_content(context: LessonContext):
    """Generate ReadNow content using OpenAI.

    Args:
        context: LessonContext containing all lesson data
    """
    if not client:
        return "Error: OpenAI API key not set"

    is_hpa = STUDENT_ATTAINMENT.upper() == "HPA"

    if is_hpa:
        questions = """1. Define a key term (1 mark)
2. Recall a fact (1 mark)
3. Recall another fact (1 mark)
4. Recall another fact (1 mark)
5. Apply knowledge/solve a problem (4 marks)"""
        extra = """CRITICAL QUESTION REQUIREMENTS:
- Write SPECIFIC, CLEAR questions that test understanding of the content
- Questions 1-4 MUST be answerable by finding EXACT information in the text above
- Each question should ask about a SPECIFIC fact, definition, or detail mentioned in the text
- Use precise question words: "What is...?", "How does...?", "Why do...?", "Define...", "State...", "Name...", "Describe..."
- AVOID vague questions like "What do you know about...?" or "Tell me about..."
- Question 5 can require application/thinking beyond the text
- Provide a MARK SCHEME with concise, specific answers (1-2 sentences max per answer)
- Each answer should directly quote or paraphrase information from the text"""
    else:
        questions = """1. MCQ (Multiple Choice Question with 4 options A-D)
2. Fill-in-the-gap (1 missing keyword)
3. State question
4. Describe question (1 sentence answer)
5. Explain question scaffolded with WHAT + WHY"""
        extra = """CRITICAL QUESTION REQUIREMENTS:
- Write exactly 5 questions following these formats (do NOT label the question types, just write the questions):

1. Multiple choice question with exactly 4 options labeled A, B, C, D:
   Write: "What is [term/concept]? A) [option1] B) [option2] C) [option3] D) [option4]"

2. Fill-in-the-gap sentence with one missing keyword:
   Write: "The process of _____ occurs when..." (use a single blank/underscore for the missing word)

3. State question:
   Write: "State [something specific from the text]"
   OR "What is [term/concept] called?"

4. Describe question (answerable in 1 sentence):
   Write: "Describe [process/concept/thing]."

5. Explain question scaffolded with WHAT + WHY:
   Write: "Explain [concept/process].
   WHAT: [scaffolding question about what happens]
   WHY: [scaffolding question about why it happens]"

- All questions MUST be answerable from the text above
- Use clear, simple language appropriate for LPA students
- Do NOT write labels like "MCQ:" or "Fill-in-the-gap:" - just write the questions directly
- Provide a MARK SCHEME with concise answers:
  - Multiple choice: Give the correct letter (A, B, C, or D)
  - Fill-in-the-gap: Give the missing keyword
  - State/Describe/Explain: Give 1-2 sentence answers"""

    # Build keywords context from unit preparation
    keywords_context = ""
    all_keywords = []
    if context.keywords:
        all_keywords.extend(context.keywords[:15])
    if context.aqa_keywords:
        all_keywords.extend(context.aqa_keywords[:15])

    if all_keywords:
        keywords_list = ", ".join(list(set(all_keywords))[:20])  # Remove duplicates, limit to 20
        keywords_context = f"\n\nIMPORTANT KEYWORDS TO INCLUDE:\nThe following keywords should be included and defined in the content: {keywords_list}\nMake sure to use these terms and mark them with asterisks (*keyword*) when they appear.\n"

    # MOST IMPORTANT: Add the most important keyword and its definition from unit preparation
    most_important_context = ""
    if context.most_important_keyword:
        # For HPA, include the FULL "What do my students need to know" content
        if is_hpa and context.full_know_content:
            most_important_context = f"\n\n🎯 MOST IMPORTANT KEYWORD AND COMPLETE DEFINITION (FROM 'WHAT STUDENTS NEED TO KNOW'):\nThe most critical keyword for this lesson is: *{context.most_important_keyword}*\n\nCOMPLETE CONTENT FROM UNIT PREPARATION (WHAT STUDENTS NEED TO KNOW):\n{context.full_know_content}\n\nCRITICAL FOR HPA: You MUST start the ReadNow content with the COMPLETE definition and explanation from the unit preparation content above. Include ALL the bullet points and information provided. This is the foundation knowledge students need. Make sure this complete content is prominent and clear at the beginning of the ReadNow.\n"
        elif context.most_important_keyword and context.keyword_definition:
            most_important_context = f"\n\n🎯 MOST IMPORTANT KEYWORD (FROM FIRST BULLET POINT OF 'WHAT STUDENTS NEED TO KNOW'):\nThe most critical keyword for this lesson is: *{context.most_important_keyword}*\n\nDEFINITION FROM UNIT PREPARATION BOOKLET:\n{context.keyword_definition}\n\nCRITICAL: You MUST start the ReadNow content with a clear definition of '{context.most_important_keyword}' using the definition provided above. This is the most important concept students need to understand. Make sure this definition is prominent and clear at the beginning of the content.\n"
        else:
            # If it's a compound term, extract the key component words and ensure they're defined
            keyword_parts = context.most_important_keyword.lower().split()
            key_terms = [word for word in keyword_parts if len(word) > 3]  # Extract significant words

            additional_context = ""
            if len(key_terms) > 1:
                additional_context = f"\n\nIMPORTANT: If '{context.most_important_keyword}' is a compound term (e.g., contains multiple words), you MUST define each key component term. For example, if the keyword is 'pyramid of biomass', you MUST define both 'biomass' and 'pyramid of biomass'. Students cannot understand the compound term without understanding the key components first.\n"

            most_important_context = f"\n\n🎯 MOST IMPORTANT KEYWORD (FROM FIRST BULLET POINT OF 'WHAT STUDENTS NEED TO KNOW'):\nThe most critical keyword for this lesson is: *{context.most_important_keyword}*\n\nCRITICAL: You MUST start the ReadNow content with clear, explicit definitions:\n1. First, define any key component terms (if '{context.most_important_keyword}' contains multiple words, define each important word)\n2. Then, define '{context.most_important_keyword}' itself\n\nThis is the most important concept students need to understand. These definitions MUST be the very first thing students read, before any other content. Use exact definitions from the AQA specification if provided.{additional_context}\n"

    # AQA specification content
    aqa_context = ""
    if context.aqa_content:
        aqa_context = f"\n\nAQA SPECIFICATION CONTENT (EXACT DEFINITIONS FROM SPEC):\n{context.aqa_content[:2000]}\n\nCRITICAL: Use the EXACT definitions from the AQA specification above. These are the definitions that exam markschemes expect. Match the wording and terminology precisely.\n"
    else:
        aqa_context = "\n\nAQA SPECIFICATION REQUIREMENTS:\n- Base all content strictly on the AQA GCSE Science specification (8464)\n- Use only terminology and concepts from the AQA specification\n- Ensure definitions match AQA specification definitions exactly\n- Definitions must be word-for-word as they appear in the AQA specification\n- Do not include content beyond GCSE level\n"

    # Language style - only apply to LPA, not HPA
    language_style_requirement = ""
    if not is_hpa:
        language_style_requirement = f"- {LANGUAGE_STYLE}\n"

    # Build context about previous and next lessons
    teaching_sequence_context = ""
    if context.previous_lesson:
        teaching_sequence_context += f"\n\n📖 PREVIOUS LESSON CONTEXT:\nStudents have already learned: {context.previous_lesson.get('title', '')}\nWhat they know: {context.previous_lesson.get('know', '')[:300]}\n\nYou can reference concepts from this previous lesson if they help students understand the current objective.\n"

    if context.next_lesson:
        teaching_sequence_context += f"\n\n🔮 NEXT LESSON CONTEXT:\nStudents will learn next: {context.next_lesson.get('title', '')}\nWhat they'll learn: {context.next_lesson.get('know', '')[:300]}\n\nDo NOT include concepts, methods, or Tier 3 words from the next lesson. These will be taught later. Only focus on what's needed for THIS lesson's first objective.\n"

    # Cognitive load considerations for HPA
    cognitive_load_note = ""
    if is_hpa:
        cognitive_load_note = "\n\nCRITICAL - COGNITIVE LOAD MANAGEMENT FOR HPA:\n- This ReadNow is used at the BEGINNING of the lesson - students are just starting to learn\n- Act like a professional teacher: Consider the teaching sequence and what's appropriate at this point\n- Avoid cognitive overload: Keep information MINIMAL, focused, and digestible\n- Use ONLY the information from the first learning objective and unit preparation content\n- Do NOT add extra examples, methods, or concepts beyond what's explicitly stated\n- Do NOT introduce new Tier 3 words or concepts students haven't learned yet\n- Keep it SHORT and SIMPLE - students are at the start of the lesson\n- Present ONLY the essential information needed to understand the first learning objective\n- Use clear, straightforward explanations without unnecessary complexity\n- Do NOT include background information, context, or additional details unless explicitly in the objective\n"
    else:
        cognitive_load_note = "\n\nCRITICAL - TEACHING SEQUENCE CONSIDERATION FOR LPA:\n- Act like a professional teacher: Think about what students have learned before and what comes next\n- Consider the common teaching order and what's appropriate at this point in the sequence\n- Do NOT introduce concepts that will be taught in future lessons\n- You can reference what students learned in previous lessons if it helps, but keep it minimal\n- Focus ONLY on what's needed for this lesson's first objective\n"

    prompt = f"""Write for {STUDENT_ATTAINMENT} GCSE Chemistry students ({STUDENT_YEAR}, reading age: {READING_AGE}) about: {context.lesson_title}
Base on: {context.first_objective}
{most_important_context}{aqa_context}{keywords_context}{teaching_sequence_context}
Requirements:
{language_style_requirement}- {WORD_COUNT} words
- Mark important keywords by putting asterisks around them like *keyword* (e.g., *atom*, *molecule*, *electron*)
- Do NOT use **double asterisks** - use single asterisks *keyword* for bold keywords
- Do NOT include the lesson title "{context.lesson_title}" in your response - start directly with the content
{"- Short paragraphs (2-3 sentences each)" if is_hpa else ""}
{cognitive_load_note}

CRITICAL - FOCUS AND TIER 3 VOCABULARY:
- The MAIN GOAL is to help students complete this learning objective: "{context.first_objective}"
- {"MAXIMUM 2 Tier 3 words allowed (absolute maximum)" if not is_hpa else "MAXIMUM 2-3 Tier 3 words allowed"}
- ONLY use Tier 3 words that are EXPLICITLY mentioned in the first learning objective or the "What do my students need to know" content above
- Do NOT introduce ANY Tier 3 words that are NOT explicitly in the learning objective or the unit preparation content
- Do NOT include concepts, methods, or terms that students haven't learned yet (e.g., if the objective is about biodiversity definition, don't include transects, quadrats, sampling methods, or other concepts not mentioned)
- ONLY use what is directly stated in the first learning objective - nothing more
- When you use a Tier 3 word, define it on first use
- Format: Use the word, then immediately add "is..." or "means..." or "refers to..." to define it
- Use EXACT definitions from the AQA specification if provided above
- Stay focused - every sentence should help students understand and achieve the first learning objective ONLY
- If you need to explain concepts, use everyday language instead of adding more Tier 3 words
- DO NOT add extra information, examples, or concepts beyond what's needed for the first learning objective

CRITICAL: When writing the content, make sure to include clear, specific information that will directly answer all 5 questions. The answers to all questions MUST be explicitly stated in the text. Include:
- Clear definitions of key terms (use EXACT definitions from AQA specification if provided above)
- Specific facts and details from the AQA specification
- Step-by-step processes if relevant
- Examples that illustrate concepts

IMPORTANT: If AQA specification content is provided above, you MUST use the EXACT wording and definitions from it. These are the definitions that exam markschemes expect. Do not paraphrase or change the wording - use the exact definitions as they appear in the AQA specification.

Then 5 questions (NUMBER THEM 1-5):
{questions}

{extra}

BAD QUESTION EXAMPLES (DO NOT WRITE LIKE THIS):
- "What do you know about ionic bonding?" (too vague)
- "Can you tell me about atoms?" (too vague)
- "What is interesting about this topic?" (not specific)

GOOD QUESTION EXAMPLES (WRITE LIKE THIS):
- "What is an ion?" (specific, answerable from text)
- "How do metal atoms form positive ions?" (specific, answerable from text)
- "Define the term 'ionic bond'." (specific, answerable from text)
- "State one property of ionic compounds." (specific, answerable from text)

Format: content → "Questions" → "MARK SCHEME"
"""

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini", messages=[{"role": "user", "content": prompt}], temperature=0.7
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Error generating content: {e}"


def save_docx(content, lesson_code, lesson_title, output_dir="readnows"):
    """Save ReadNow to Word document organized by year and subject.

    This is a wrapper function that calls save_readnow_docx from readers.word_reader.
    """
    import constants

    return save_readnow_docx(
        content=content,
        lesson_code=lesson_code,
        lesson_title=lesson_title,
        output_dir=output_dir,
        constants_module=constants,
        get_subject_from_code_func=get_subject_from_code,
        download_chemistry_image=download_chemistry_image,
    )


def main():
    """Generate ReadNows from PowerPoint files using settings from constants.py"""
    # TARGET_LESSON: Set to a specific lesson code (e.g., "C4.2.4") to generate only that one, or None to generate all
    TARGET_LESSON = "C3.2.6"  # Uncertainty (Year 9) - or None to generate all

    logger.info(f"Starting {STUDENT_ATTAINMENT} ReadNow generation for {STUDENT_YEAR}")
    logger.info(f"Reading age: {READING_AGE}")
    logger.info(f"Word count: {WORD_COUNT}")
    if TARGET_LESSON:
        logger.info(f"Target lesson: {TARGET_LESSON} (only generating this one)")

    # Resolve PowerPoint folder path (check current dir and parent dir)
    ppt_folder = Path(POWERPOINT_FOLDER)
    if not ppt_folder.exists():
        ppt_folder = Path("..") / POWERPOINT_FOLDER
    if not ppt_folder.exists():
        logger.error(f"PowerPoint folder not found: {POWERPOINT_FOLDER}")
        logger.error(f"Tried: {Path(POWERPOINT_FOLDER).absolute()}")
        logger.error(f"Tried: {(Path('..') / POWERPOINT_FOLDER).absolute()}")
        return

    logger.info(f"Using folder: {ppt_folder.absolute()}")

    # Look for PowerPoint files in subfolders (e.g., Lesson Resources/C3.1.3 Electronic Configuration/)
    pptx_files = list(ppt_folder.rglob("*.pptx"))
    if not pptx_files:
        logger.error(f"No PowerPoint files found in {POWERPOINT_FOLDER}")
        return

    logger.info(f"Found {len(pptx_files)} PowerPoint files")

    successful = failed = 0
    found_lessons = []

    for pptx_file in pptx_files:
        # Skip temp files
        if pptx_file.name.startswith("~"):
            continue

        # Try to extract lesson code from folder name or filename
        # Folder structure: Lesson Resources/C3.1.3 Electronic Configuration/file.pptx
        folder_name = pptx_file.parent.name
        if folder_name and any(c.isdigit() for c in folder_name.split()[0] if folder_name.split()):
            # Extract from folder name: "C3.1.3 Electronic Configuration" -> "C3.1.3", "Electronic Configuration"
            parts = folder_name.split(" ", 1)
            lesson_code = parts[0]
            lesson_title = parts[1] if len(parts) > 1 else "Unknown"
        else:
            # Fallback: try filename
            parts = pptx_file.stem.split(" ", 1)
            lesson_code = parts[0] if parts else pptx_file.stem[:7]
            lesson_title = parts[1].strip().replace(" (1)", "") if len(parts) > 1 else "Unknown"

        # Normalize lesson codes for comparison (uppercase)
        lesson_code_normalized = lesson_code.upper()
        target_normalized = TARGET_LESSON.upper() if TARGET_LESSON else None
        prefix_normalized = LESSON_CODE_PREFIX.upper() if LESSON_CODE_PREFIX else None

        found_lessons.append(lesson_code)

        # If TARGET_LESSON is set, skip LESSON_CODE_PREFIX filter and match by lesson code
        if TARGET_LESSON:
            if lesson_code_normalized != target_normalized:
                continue
        else:
            # Only filter by prefix if no specific target lesson
            if prefix_normalized and not lesson_code_normalized.startswith(prefix_normalized):
                continue

        logger.info(f"{lesson_code}: {lesson_title}")

        # Pass the full path to the PowerPoint file - searches slides 4-7 automatically
        objectives = extract_slide6_objectives(lesson_code, str(pptx_file.parent), slide_num=None)
        if not objectives:
            logger.warning("No objectives found")
            failed += 1
            continue

        first_objective = get_first_objective(objectives)
        logger.info(f"First objective: {first_objective[:60]}...")

        unit = ".".join(lesson_code.split(".")[:2])

        # STEP 1: Get the most important keyword from first bullet point of "What do my students need to know"
        logger.info("Step 1: Extracting most important keyword from first bullet point...")
        most_important_keyword, keyword_definition, full_know_content = get_most_important_keyword_and_definition(
            lesson_code, unit
        )

        # STEP 2: Get previous lesson to understand what students already know
        logger.info("Step 2: Checking previous lesson to understand prior knowledge...")
        previous_lesson = get_previous_lesson_content(lesson_code, unit)
        if previous_lesson:
            logger.info(
                f"Found previous lesson: {previous_lesson.get('code', '')} - {previous_lesson.get('title', '')[:50]}"
            )

        # STEP 3: Get next lesson to understand what comes next (to avoid teaching ahead)
        logger.info("Step 3: Checking next lesson to avoid teaching ahead...")
        next_lesson = get_next_lesson_content(lesson_code, unit)
        if next_lesson:
            logger.info(f"Found next lesson: {next_lesson.get('code', '')} - {next_lesson.get('title', '')[:50]}")

        # STEP 4: Extract keywords from unit preparation Excel file
        keywords = get_keywords_from_excel(lesson_code, unit)
        if keywords:
            logger.info(f"Found {len(keywords)} keywords from unit preparation")
        else:
            logger.warning("No keywords found in unit preparation file")

        # STEP 5: Extract AQA specification content using AI to find relevant sections
        logger.info("Step 5: Extracting AQA specification content...")
        aqa_content, aqa_keywords = extract_aqa_specification_content(
            lesson_title, first_objective, unit
        )
        if aqa_content:
            logger.info(f"Found AQA specification content ({len(aqa_content)} chars)")
        if aqa_keywords:
            logger.info(f"Found {len(aqa_keywords)} keywords from AQA specification")

        # STEP 6: Generate content with the most important keyword and definition prioritized
        # Create LessonContext to pass all data
        context = LessonContext(
            lesson_code=lesson_code,
            lesson_title=lesson_title,
            first_objective=first_objective,
            unit_code=unit,
            keywords=keywords,
            aqa_content=aqa_content,
            aqa_keywords=aqa_keywords,
            most_important_keyword=most_important_keyword,
            keyword_definition=keyword_definition,
            full_know_content=full_know_content,
            previous_lesson=previous_lesson,
            next_lesson=next_lesson,
        )

        content = generate_content(context)

        if save_docx(content, lesson_code, lesson_title):
            logger.info("Saved")
            successful += 1
        else:
            failed += 1

    logger.info("=" * 60)
    if TARGET_LESSON and successful == 0 and failed == 0:
        logger.warning(f"No lessons found matching '{TARGET_LESSON}'")
        logger.info("Found lessons in PowerPoint files:")
        unique_lessons = sorted(set(found_lessons))[:20]  # Show first 20
        for lesson in unique_lessons:
            logger.info(f"  - {lesson}")
        if len(set(found_lessons)) > 20:
            logger.info(f"  ... and {len(set(found_lessons)) - 20} more")
        logger.info("Tip: Make sure the lesson code matches exactly (case-insensitive)")
    else:
        logger.info(f"Success: {successful} | Failed: {failed}")
    year_folder = STUDENT_YEAR.replace(" ", "_")
    logger.info(f"Files saved to: readnows/{year_folder}/[Subject]/")


if __name__ == "__main__":
    main()
