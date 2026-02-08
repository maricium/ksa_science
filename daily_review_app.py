#!/usr/bin/env python3
"""
Streamlit front end for the Daily Review / Core Knowledge generator.
Lets teachers choose a unit and either use auto weeks or assign lessons to each week
(e.g. Week 1: C4.3.2, C4.3.3 only; Week 2: C4.3.4, C4.3.5, C4.3.6).

Optional: set environment variable DAILY_REVIEW_PIN to require a PIN before using the app.
Optional: set APP_TITLE to your site name (default: "KSA Science") for the browser tab and headings.
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

# Optional PIN: if DAILY_REVIEW_PIN is set, user must enter it to see the app
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
st.markdown("Generate AQA-style Core Knowledge documents. Choose a unit (or multiple units) and how lessons are grouped into weeks.")

script_dir = Path(__file__).resolve().parent
lesson_resources = script_dir / "Lesson Resources"

if not lesson_resources.is_dir():
    st.error(f"Lesson Resources folder not found: {lesson_resources}")
    st.stop()

units = list_available_units(lesson_resources)
if not units:
    st.warning("No units found in Lesson Resources (expected folders like 'C4.3 Quantitative Chemistry').")
    st.stop()

# Single unit vs multiple units (e.g. B3.2, C4.2, P3.2 in the same week)
scope = st.radio(
    "Who is this for?",
    options=["single", "multi"],
    format_func=lambda x: "Single unit (one unit, e.g. C4.3 only)" if x == "single" else "Multiple units (I teach B3.2, C4.2, P3.2 etc. in the same week – combine words from several units per week)",
    horizontal=True,
    key="scope",
)

@st.cache_data(ttl=60)
def load_lessons(uc: str):
    try:
        found = find_unit_resources(uc, lesson_resources)
        lessons = read_lessons_with_spec(found["excel_path"])
        return found, lessons
    except Exception as e:
        return None, None

if scope == "multi":
    # Multiple units: select which units, then assign lessons to each week (any mix of units per week)
    st.markdown("### Multiple units – one word list per week")
    st.caption("Pick the units you teach, then assign lessons to each week. You can put B3.2.4, C4.2.2 and P3.2.2 in the same week to get one combined list of words for that week.")
    selected_units = st.multiselect("Which units do you teach?", options=units, default=units[:1] if units else [], key="multi_units")
    if not selected_units:
        st.warning("Select at least one unit.")
        st.stop()

    # Build combined lesson list: (lesson_code, label) for display, sorted by unit then lesson
    combined_lessons = []  # (code, label)
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

    term_data = load_term_data(script_dir)
    half_term_options = list_half_terms(term_data)
    teaching_weeks_from_term = None
    if half_term_options:
        options_for_select = ["— Manual —"] + [f"{y} {t} – {ht} ({w} weeks)" for y, t, ht, w in half_term_options]
        choice = st.selectbox("Half term (sets default number of weeks)", options=options_for_select, key="multi_half_term")
        if choice and choice != "— Manual —":
            idx = options_for_select.index(choice) - 1
            _, _, _, teaching_weeks_from_term = half_term_options[idx]

    default_weeks = teaching_weeks_from_term if teaching_weeks_from_term else 4
    n_weeks = st.number_input("Number of weeks", min_value=1, max_value=20, value=default_weeks, step=1, key="multi_n_weeks")
    lesson_options = [label for _, label in combined_lessons]
    code_by_label = {label: code for code, label in combined_lessons}

    assignments = []
    for w in range(1, int(n_weeks) + 1):
        selected_labels = st.multiselect(
            f"Week {w}",
            options=lesson_options,
            default=[],
            key=f"multi_week_{w}",
            help="Select which lessons (from any unit) belong in this week.",
        )
        assignments.append([code_by_label[l] for l in selected_labels])

    if st.button("Preview week summary", key="multi_preview"):
        for i, a in enumerate(assignments, 1):
            st.caption(f"Week {i}: {' & '.join(a) if a else '(none)'}")

    st.markdown("---")
    if st.button("Generate Core Knowledge document (multiple units)", type="primary", key="multi_gen"):
        week_assignments_clean = [a for a in assignments if a]
        if not week_assignments_clean:
            st.warning("Add at least one lesson to at least one week.")
        else:
            with st.spinner("Generating document (questions from all units, validation)…"):
                path, err = run_daily_review_generator_multi(lesson_resources, week_assignments_clean, output_dir=script_dir)
            if err:
                st.error(err)
            else:
                st.success("Document generated.")
                with open(path, "rb") as f:
                    st.download_button(
                        "Download Core Knowledge (.docx)",
                        data=f.read(),
                        file_name=Path(path).name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="multi_download",
                    )
    st.stop()

# Single-unit flow below
unit_code = st.selectbox(
    "Unit",
    options=units,
    format_func=lambda x: x,
    key="unit",
)

found, lessons = load_lessons(unit_code)
if not found or not lessons:
    st.error("Could not load unit or lessons.")
    st.stop()

lesson_list = sorted(
    [lc for lc in lessons if "Feedback" not in lessons[lc]["title"]],
    key=lambda lc: _lesson_sort_key(lc.split("/")[0] if "/" in lc else lc),
)

st.success(f"**{found['unit_name']}** – {len(lesson_list)} lessons loaded.")

term_data = load_term_data(script_dir)
half_term_options = list_half_terms(term_data)
teaching_weeks_from_term = None
if half_term_options:
    st.markdown("### Half term")
    st.caption("Choose the half term to match how many teaching weeks to generate (from term.json).")
    options_for_select = ["— Manual / don't use half term —"] + [
        f"{y} {t} – {ht} ({w} weeks)" for y, t, ht, w in half_term_options
    ]
    choice = st.selectbox("Half term", options=options_for_select, key="half_term")
    if choice and choice != "— Manual / don't use half term —":
        idx = options_for_select.index(choice) - 1
        _, _, _, teaching_weeks_from_term = half_term_options[idx]
        st.info(f"Will generate **{teaching_weeks_from_term} weeks** of content for this half term.")

mode = st.radio(
    "How should weeks be defined?",
    options=["auto", "custom"],
    format_func=lambda x: "Auto (combine lessons by word count)" if x == "auto" else "Custom (I choose which lessons go in each week)",
    horizontal=True,
)

week_assignments = None

if mode == "custom":
    st.markdown("---")
    st.markdown("**Assign lessons to each week.** One teacher might have 2 lessons per week, another 4.")
    default_weeks = teaching_weeks_from_term if teaching_weeks_from_term else 4
    n_weeks = st.number_input("Number of weeks", min_value=1, max_value=20, value=default_weeks, step=1)
    assignments = []
    for w in range(1, int(n_weeks) + 1):
        selected = st.multiselect(
            f"Week {w}",
            options=lesson_list,
            default=[],
            key=f"week_{w}",
            help="Select which lessons belong in this week.",
        )
        assignments.append(selected)
    if st.button("Preview week summary"):
        for i, a in enumerate(assignments, 1):
            st.caption(f"Week {i}: {' & '.join(a) if a else '(none)'}")
    week_assignments = assignments

st.markdown("---")
if st.button("Generate Core Knowledge document", type="primary"):
    if mode == "custom" and week_assignments:
        week_assignments_clean = [a for a in week_assignments if a]
        if not week_assignments_clean:
            st.warning("Add at least one lesson to at least one week.")
        else:
            with st.spinner("Generating document (questions, state questions, validation)…"):
                path, err = run_daily_review_generator(
                    unit_code,
                    lesson_resources,
                    week_assignments=week_assignments_clean,
                    output_dir=script_dir,
                    max_weeks=None,
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
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="download",
                    )
    else:
        with st.spinner("Generating document (auto weeks)…"):
            path, err = run_daily_review_generator(
                unit_code,
                lesson_resources,
                week_assignments=None,
                output_dir=script_dir,
                max_weeks=teaching_weeks_from_term,
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
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="download",
                )
