#!/usr/bin/env python3
"""
Streamlit front end for the Daily Review / Core Knowledge generator.
Flow: (1) Which half term? (2) Which units? (3) Assign lessons to each week → Generate.
Optional: set DAILY_REVIEW_PIN and APP_TITLE via environment.
"""

import os
import streamlit as st
from pathlib import Path

from daily_review_words import (
    _lesson_sort_key,
    find_unit_resources,
    list_available_units,
    list_half_terms,
    load_term_data,
    read_lessons_with_spec,
    run_daily_review_generator,
    run_daily_review_generator_multi,
)

APP_TITLE = os.environ.get("APP_TITLE", "KSA Science").strip() or "KSA Science"

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="📋",
    layout="wide",
)

# Optional PIN
REQUIRED_PIN = os.environ.get("DAILY_REVIEW_PIN", "").strip()
if REQUIRED_PIN and not st.session_state.get("pin_ok"):
    st.title(f"📋 {APP_TITLE}")
    st.markdown("Enter the PIN to access the app.")
    pin = st.text_input("PIN", type="password", placeholder="Enter PIN", key="pin_input")
    if st.button("Submit"):
        if pin == REQUIRED_PIN:
            st.session_state["pin_ok"] = True
            st.rerun()
        else:
            st.error("Incorrect PIN.")
    st.stop()

st.title(f"📋 {APP_TITLE}")
st.markdown("Generate Core Knowledge word lists. Choose the half term, which units, then assign lessons to each week.")

script_dir = Path(__file__).resolve().parent
lesson_resources = script_dir / "Lesson Resources"

if not lesson_resources.is_dir():
    st.error(f"Lesson Resources folder not found: {lesson_resources}")
    st.stop()

units = list_available_units(lesson_resources)
if not units:
    st.warning("No units found in Lesson Resources.")
    st.stop()

# ——— 1. Which half term is this for? ———
st.markdown("### 1. Which half term is this for?")
term_data = load_term_data(script_dir)
half_term_options = list_half_terms(term_data)
teaching_weeks_from_term = None
options_for_select = ["— Manual (I'll set number of weeks below) —"]
if half_term_options:
    options_for_select += [f"{y} {t} – {ht} ({w} weeks)" for y, t, ht, w in half_term_options]

half_term_choice = st.selectbox(
    "Half term",
    options=options_for_select,
    key="half_term",
    help="This sets the default number of weeks. You can change it in step 3.",
)
if half_term_choice and half_term_choice != options_for_select[0] and half_term_options:
    idx = options_for_select.index(half_term_choice) - 1
    _, _, _, teaching_weeks_from_term = half_term_options[idx]
    st.caption(f"Default: **{teaching_weeks_from_term} weeks** for this half term.")

# ——— 2. Which units do you want to make words for? ———
st.markdown("### 2. Which units do you want to make words for?")
selected_units = st.multiselect(
    "Select one or more units",
    options=units,
    default=[],
    key="units",
    help="You can pick a single unit (e.g. C4.3) or several (e.g. B3.2, C4.2, P3.2) to combine words from different units in the same week.",
)
if not selected_units:
    st.warning("Select at least one unit.")
    st.stop()

@st.cache_data(ttl=60)
def load_lessons(uc: str):
    try:
        found = find_unit_resources(uc, lesson_resources)
        lessons = read_lessons_with_spec(found["excel_path"])
        return found, lessons
    except Exception:
        return None, None

# Build combined lesson list
combined_lessons = []
for uc in selected_units:
    found, lessons = load_lessons(uc)
    if not found or not lessons:
        continue
    for lc in sorted(
        [c for c in lessons if "Feedback" not in lessons[c].get("title", "")],
        key=lambda x: _lesson_sort_key(x.split("/")[0] if "/" in x else x),
    ):
        title = lessons[lc].get("title", "")[:40]
        combined_lessons.append((lc, f"{lc} – {title} ({uc})"))

if not combined_lessons:
    st.error("No lessons found in the selected units.")
    st.stop()

st.success(f"**{len(combined_lessons)} lessons** from {', '.join(selected_units)}.")

# ——— 3. Assign lessons to each week ———
st.markdown("### 3. Assign lessons to each week")
default_weeks = teaching_weeks_from_term if teaching_weeks_from_term else 4
n_weeks = st.number_input(
    "Number of weeks",
    min_value=1,
    max_value=20,
    value=default_weeks,
    step=1,
    key="n_weeks",
    help="Match your half term or set how many weeks you want in the document.",
)
lesson_options = [label for _, label in combined_lessons]
code_by_label = {label: code for code, label in combined_lessons}

assignments = []
for w in range(1, int(n_weeks) + 1):
    selected_labels = st.multiselect(
        f"Week {w}",
        options=lesson_options,
        default=[],
        key=f"week_{w}",
        help="Select which lessons belong in this week. You can mix units in the same week.",
    )
    assignments.append([code_by_label[l] for l in selected_labels])

if st.button("Preview week summary", key="preview"):
    for i, a in enumerate(assignments, 1):
        st.caption(f"Week {i}: {' & '.join(a) if a else '(none)'}")

# ——— Generate ———
st.markdown("---")
if st.button("Generate Core Knowledge document", type="primary", key="generate"):
    week_assignments_clean = [a for a in assignments if a]
    if not week_assignments_clean:
        st.warning("Add at least one lesson to at least one week.")
    else:
        with st.spinner("Generating document…"):
            if len(selected_units) == 1:
                path, err = run_daily_review_generator(
                    selected_units[0],
                    lesson_resources,
                    week_assignments=week_assignments_clean,
                    output_dir=script_dir,
                    max_weeks=None,
                )
            else:
                path, err = run_daily_review_generator_multi(
                    lesson_resources,
                    week_assignments_clean,
                    output_dir=script_dir,
                )
        if err:
            st.error(err)
        else:
            st.success("Document generated.")
            with open(path, "rb") as f:
                st.download_button(
                    "Download Core Knowledge (.docx)",
                    data=f.read(),
                    file_name=Path(path).name,
                    mime="application/octet-stream",
                    key="download",
                )
