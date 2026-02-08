#!/usr/bin/env python3
"""
Simple Streamlit interface for Education Resource Generator
"""

import streamlit as st
import subprocess
import re
from pathlib import Path
import os
import glob

st.set_page_config(page_title="Education Resource Generator", page_icon="🎓", layout="wide")

st.title("🎓 Education Resource Generator")

# Simple input section
st.markdown("### Enter lesson details:")

col1, col2 = st.columns([2, 1])

with col1:
    lesson_code = st.text_input(
        "Lesson Code",
        placeholder="e.g., C4.3.2 (or leave blank for all lessons)",
        key="lesson_input"
    )

with col2:
    year_group = st.selectbox(
        "Year Group",
        ["Year 9", "Year 10", "Year 11"],
        key="year_select"
    )

# ReadNow Configuration Section
st.markdown("---")
st.markdown("### 📖 ReadNow Configuration:")

config_col1, config_col2 = st.columns(2)

with config_col1:
    attainment = st.selectbox(
        "Student Attainment",
        ["LPA", "HPA"],
        key="attainment_select",
        help="LPA = Lower Prior Attainment, HPA = Higher Prior Attainment"
    )
    
    reading_age_input = st.number_input(
        "Reading Age (years)",
        min_value=7,
        max_value=16,
        value=14,
        step=1,
        key="reading_age_input",
        help="Target reading age in years (e.g., 11, 12, 14)"
    )

with config_col2:
    word_count_range = st.selectbox(
        "Word Count Range",
        ["80-100", "100-120", "120-150", "140-160", "150-180"],
        index=1,
        key="word_count_select"
    )
    
    reading_age_display = f"{reading_age_input} years old"

# Question Types Information
st.markdown("#### Question Types:")
question_info_col1, question_info_col2 = st.columns(2)

with question_info_col1:
    st.markdown(f"**Selected: {attainment}**")
    if attainment == "LPA":
        st.markdown("""
        **LPA Question Format:**
        - ✓ Multiple Choice (MCQ) - 4 options A-D
        - ✓ Fill-in-the-gap (1 missing keyword)
        - ✓ State question
        - ✓ Describe question (1 sentence)
        - ✓ Explain question (scaffolded with WHAT + WHY)
        """)
    else:
        st.markdown("""
        **HPA Question Format:**
        - ✓ Define a key term (1 mark)
        - ✓ Recall facts (3 questions, 1 mark each)
        - ✓ Apply knowledge/solve a problem (4 marks)
        """)

with question_info_col2:
    st.markdown(f"**Reading Age:** {reading_age_display}")
    st.markdown(f"**Word Count:** {word_count_range} words")
    st.markdown(f"**Year Group:** {year_group}")

# Store these for later use
reading_age = reading_age_input
word_count = word_count_range
language_style = "Accessible"
custom_title = ""
custom_objectives = ""

def update_config_file(file_path, target_lesson):
    """Update TARGET_LESSON in a config file"""
    try:
        with open(file_path, 'r') as f:
            content = f.read()
        
        # Replace TARGET_LESSON line
        if target_lesson:
            new_line = f'TARGET_LESSON = "{target_lesson}"'
        else:
            new_line = 'TARGET_LESSON = None'
        
        content = re.sub(r'TARGET_LESSON\s*=\s*.*', new_line, content)
        
        with open(file_path, 'w') as f:
            f.write(content)
        return True
    except Exception as e:
        st.error(f"Error updating config: {e}")
        return False

def update_readnow_target(script_path, target_lesson):
    """Update TARGET_LESSON in readnow unified_generator.py"""
    try:
        with open(script_path, 'r') as f:
            content = f.read()
        
        # Find and replace TARGET_LESSON line (handle any format)
        pattern = r'(TARGET_LESSON\s*=\s*)(None|"[^"]*"|\'[^\']*\')'
        if target_lesson:
            # Normalize to uppercase for consistency
            target_lesson_clean = target_lesson.strip().upper()
            replacement = f'\\1"{target_lesson_clean}"'
        else:
            replacement = r'\1None'
        
        content = re.sub(pattern, replacement, content)
        
        with open(script_path, 'w') as f:
            f.write(content)
        return True
    except Exception as e:
        st.error(f"Error updating ReadNow script: {e}")
        return False

def update_readnow_constants(constants_path, year, attainment, reading_age, word_count):
    """Update constants in readnow/constants.py"""
    try:
        with open(constants_path, 'r') as f:
            content = f.read()
        
        # Update STUDENT_YEAR
        content = re.sub(
            r'STUDENT_YEAR\s*=\s*"[^"]*"',
            f'STUDENT_YEAR = "{year}"',
            content
        )
        
        # Update STUDENT_ATTAINMENT
        content = re.sub(
            r'STUDENT_ATTAINMENT\s*=\s*"[^"]*"',
            f'STUDENT_ATTAINMENT = "{attainment}"',
            content
        )
        
        # Update READING_AGE
        content = re.sub(
            r'READING_AGE\s*=\s*"[^"]*"',
            f'READING_AGE = "{reading_age} years old"',
            content
        )
        
        # Update WORD_COUNT
        content = re.sub(
            r'WORD_COUNT\s*=\s*"[^"]*"',
            f'WORD_COUNT = "{word_count}"',
            content
        )
        
        with open(constants_path, 'w') as f:
            f.write(content)
        return True
    except Exception as e:
        st.error(f"Error updating ReadNow constants: {e}")
        return False

def run_command(command, label):
    """Run a command and return output"""
    try:
        result = subprocess.run(
            command,
            shell=True,
            capture_output=True,
            text=True,
            cwd=Path(__file__).parent
        )
        output = result.stdout + result.stderr
        return result.returncode == 0, output
    except Exception as e:
        return False, f"Error: {str(e)}"

base_path = Path(__file__).parent
slides_config = base_path / "Slides" / "config.py"
readnow_script = base_path / "readnow" / "unified_generator.py"
readnow_constants = base_path / "readnow" / "constants.py"

# Simple generation buttons
st.markdown("---")
st.markdown("### Generate Resources:")

gen_col1, gen_col2, gen_col3 = st.columns(3, gap="large")

with gen_col1:
    if st.button("📊 Generate Slides", use_container_width=True, key="slides_btn"):
        if update_config_file(slides_config, lesson_code if lesson_code else None):
            with st.spinner("Generating slides..."):
                success, output = run_command("cd Slides && uv run python generate.py", "slides")
                if success:
                    st.success("✅ Slides ready!")
                    # Show where files are saved
                    output_dir = base_path / "Slides" / "output"
                    if lesson_code:
                        # Try to find the generated file
                        import glob
                        pattern = str(output_dir / f"*{lesson_code.upper()}*.pptx")
                        files = glob.glob(pattern)
                        if files:
                            st.info(f"📁 Saved to: `{files[0]}`")
                        else:
                            st.info(f"📁 Output directory: `{output_dir}`")
                    else:
                        st.info(f"📁 Output directory: `{output_dir}`")
                else:
                    st.error("❌ Error")
                    st.code(output, language="text")

with gen_col2:
    if st.button("📖 Generate ReadNows", use_container_width=True, key="readnow_btn"):
        # Update all ReadNow settings before generating
        if (update_readnow_target(readnow_script, lesson_code if lesson_code else None) and
            update_readnow_constants(readnow_constants, year_group, attainment, reading_age, word_count)):
            with st.spinner("Generating ReadNows..."):
                success, output = run_command("PYTHONPATH=. uv run python readnow/unified_generator.py", "readnow")
                if success:
                    st.success("✅ ReadNows ready!")
                else:
                    st.error("❌ Error")
                    st.code(output, language="text")
        else:
            st.error("❌ Failed to update ReadNow configuration")

with gen_col3:
    if st.button("📝 Generate Worksheets", use_container_width=True, key="worksheet_btn"):
        with st.spinner("Generating worksheets..."):
            success, output = run_command("cd worksheet && uv run python main.py", "worksheet")
            if success:
                st.success("✅ Worksheets ready!")
            else:
                st.error("❌ Error")
                st.code(output, language="text")

st.markdown("---")
st.caption("💡 Enter a lesson code (e.g., C4.3.2) or leave blank to process all lessons")
