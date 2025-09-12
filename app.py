import os
import json
import requests
import streamlit as st
from dotenv import load_dotenv
from io import BytesIO
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
import pandas as pd
import altair as alt
from report_generator import build_pdf_report  # new modular generator

load_dotenv()

API_BASE = os.getenv("API_BASE", "http://localhost:8000")
TITLE = "Excel Mock Interviewer"

st.set_page_config(page_title=TITLE, page_icon="📊")

# Session state keys
STATE_KEYS = {
    "stage": "stage",  # intro, asking, summary
    "questions": "questions",
    "current_index": "current_index",
    "qa": "qa",  # list of dicts: {question, answer, score, feedback}
    "pdf_bytes": "pdf_bytes",
    "pdf_name": "pdf_name",
    "candidate": "candidate",  # dict: {name, email, role, experience}
    "followup_pending": "followup_pending",
    "followup_question": "followup_question",
    "followup_for_index": "followup_for_index",
}


def init_state():
    if STATE_KEYS["stage"] not in st.session_state:
        st.session_state[STATE_KEYS["stage"]] = "intro"
    if STATE_KEYS["questions"] not in st.session_state:
        try:
            resp = requests.get(f"{API_BASE}/questions", timeout=10)
            questions = resp.json() if resp.ok else []
        except Exception:
            questions = []
        st.session_state[STATE_KEYS["questions"]] = questions
    if STATE_KEYS["current_index"] not in st.session_state:
        st.session_state[STATE_KEYS["current_index"]] = 0
    if STATE_KEYS["qa"] not in st.session_state:
        st.session_state[STATE_KEYS["qa"]] = []
    if STATE_KEYS["pdf_bytes"] not in st.session_state:
        st.session_state[STATE_KEYS["pdf_bytes"]] = None
    if STATE_KEYS["pdf_name"] not in st.session_state:
        st.session_state[STATE_KEYS["pdf_name"]] = None
    if STATE_KEYS["candidate"] not in st.session_state:
        st.session_state[STATE_KEYS["candidate"]] = {"name": "", "email": "", "role": "Excel Analyst", "experience": "0"}
    if STATE_KEYS["followup_pending"] not in st.session_state:
        st.session_state[STATE_KEYS["followup_pending"]] = False
    if STATE_KEYS["followup_question"] not in st.session_state:
        st.session_state[STATE_KEYS["followup_question"]] = ""
    if STATE_KEYS["followup_for_index"] not in st.session_state:
        st.session_state[STATE_KEYS["followup_for_index"]] = -1


def render_header():
    st.title(TITLE)
    st.caption("Practice basic → advanced Excel skills. Get graded feedback after each answer.")


def introduction():
    st.info("Hi, I’m your Excel Mock Interviewer. Please provide your details to begin.")
    with st.form("candidate_form"):
        name = st.text_input("Full name", value=st.session_state[STATE_KEYS["candidate"]].get("name", ""))
        email = st.text_input("Email (optional)", value=st.session_state[STATE_KEYS["candidate"]].get("email", ""))
        role = st.text_input("Target role", value=st.session_state[STATE_KEYS["candidate"]].get("role", "Excel Analyst"))
        experience = st.text_input("Years of experience", value=st.session_state[STATE_KEYS["candidate"]].get("experience", "0"))
        submitted = st.form_submit_button("Start Interview", type="primary")
    if submitted:
        st.session_state[STATE_KEYS["candidate"]] = {
            "name": name.strip(),
            "email": email.strip(),
            "role": role.strip() or "Excel Analyst",
            "experience": experience.strip() or "0",
        }
        if name.strip():
            st.session_state[STATE_KEYS["stage"]] = "asking"
            st.rerun()
        else:
            st.warning("Please enter your full name to start.")


def grade_answer(question_text: str, answer_text: str):
    try:
        resp = requests.post(
            f"{API_BASE}/grade",
            json={"question": question_text, "answer": answer_text},
            timeout=30,
        )
        if resp.ok:
            data = resp.json()
            return int(data.get("score", 0)), str(data.get("feedback", ""))
    except Exception as e:
        return 0, f"Error contacting grader: {e}"
    return 0, "Grading failed."


# Interactive visuals per question index
VISUALS = {
    2: {  # PivotTable question (index 2 in our default list after intro)
        "table": pd.DataFrame({
            "Month": ["Jan", "Jan", "Feb", "Feb", "Mar", "Mar"],
            "Product": ["A", "B", "A", "B", "A", "B"],
            "Sales": [1200, 800, 1500, 950, 1700, 1100],
        }),
        "chart": pd.DataFrame({
            "Month": ["Jan", "Feb", "Mar"],
            "A": [1200, 1500, 1700],
            "B": [800, 950, 1100],
        })
    }
}


def render_visuals_for_index(idx: int):
    v = VISUALS.get(idx)
    if not v:
        return
    st.markdown("Interactive reference data:")
    st.dataframe(v["table"], use_container_width=True)
    chart_df = v["chart"].melt("Month", var_name="Product", value_name="Sales")
    chart = alt.Chart(chart_df).mark_bar().encode(
        x="Month:N",
        y="Sales:Q",
        color="Product:N",
        tooltip=["Month", "Product", "Sales"]
    ).properties(height=220)
    st.altair_chart(chart, use_container_width=True)


def get_followup(question_text: str, score: int) -> str:
    qt_lower = (question_text or "").lower()
    if "vlookup" in qt_lower or "xlookup" in qt_lower or "index-match" in qt_lower:
        return (
            "Great. As a follow-up, when would you prefer INDEX/MATCH or XLOOKUP over VLOOKUP, "
            "and how do you handle approximate vs. exact matches?" if score >= 4 else
            "As a follow-up, clarify the role of the last argument (range_lookup/exact match) "
            "in VLOOKUP and why it matters."
        )
    if "pivot" in qt_lower:
        return (
            "Nice. Follow-up: how would you add a slicer and sort top 5 products by sales?" if score >= 4 else
            "Follow-up: explain how to change aggregation from SUM to AVERAGE and add a value filter."
        )
    if "conditional formatting" in qt_lower:
        return (
            "Good. Follow-up: how would you use formula-based rules to highlight rows based on multiple conditions?" if score >= 4 else
            "Follow-up: explain duplicate vs. unique highlighting and managing rule precedence."
        )
    if "macro" in qt_lower or "vba" in qt_lower:
        return (
            "Good. Follow-up: outline steps to create a parameterized macro and assign to a button." if score >= 4 else
            "Follow-up: describe how to record a macro and edit its VBA to fix an absolute reference."
        )
    if "power query" in qt_lower:
        return (
            "Follow-up: how would you merge two queries with mismatched keys and handle nulls?" if score >= 3 else
            "Follow-up: what is the difference between merge vs. append in Power Query?"
        )
    return "Follow-up: provide a concrete example and common pitfalls."


def ask_questions():
    questions = st.session_state[STATE_KEYS["questions"]]
    idx = st.session_state[STATE_KEYS["current_index"]]

    if not questions:
        st.error("No questions available. Please ensure the backend is running.")
        return

    if idx >= len(questions):
        st.session_state[STATE_KEYS["stage"]] = "summary"
        st.rerun()
        return

    q = questions[idx]
    st.subheader(f"Question {idx + 1} of {len(questions)} ({q.get('level', '').title()})")
    st.write(q.get("text", ""))

    # Interactive visuals for certain questions
    render_visuals_for_index(idx)

    # If there is a pending follow-up for this index
    if st.session_state[STATE_KEYS["followup_pending"]] and st.session_state[STATE_KEYS["followup_for_index"]] == idx:
        st.markdown("**Follow-up**: " + st.session_state[STATE_KEYS["followup_question"]])
        with st.form(key=f"followup_form_{idx}", clear_on_submit=True):
            answer = st.text_area("Your follow-up answer", height=140)
            c1, c2 = st.columns(2)
            with c1:
                submitted = st.form_submit_button("Submit Follow-up")
            with c2:
                skip = st.form_submit_button("Skip follow-up / Next question")
        if submitted:
            score, feedback = grade_answer(st.session_state[STATE_KEYS["followup_question"]], answer)
            st.session_state[STATE_KEYS["qa"]].append(
                {"question": st.session_state[STATE_KEYS["followup_question"]], "answer": answer, "score": score, "feedback": feedback}
            )
            st.session_state[STATE_KEYS["followup_pending"]] = False
            st.session_state[STATE_KEYS["followup_question"]] = ""
            st.session_state[STATE_KEYS["followup_for_index"]] = -1
            st.session_state[STATE_KEYS["current_index"]] += 1
            st.rerun()
        elif skip:
            st.session_state[STATE_KEYS["followup_pending"]] = False
            st.session_state[STATE_KEYS["followup_question"]] = ""
            st.session_state[STATE_KEYS["followup_for_index"]] = -1
            st.session_state[STATE_KEYS["current_index"]] += 1
            st.rerun()
        return

    # Primary question form
    with st.form(key=f"answer_form_{idx}", clear_on_submit=True):
        answer = st.text_area("Your answer", height=140)
        submitted = st.form_submit_button("Submit Answer")
    if submitted:
        score, feedback = grade_answer(q.get("text", ""), answer)
        st.session_state[STATE_KEYS["qa"]].append(
            {"question": q.get("text", ""), "answer": answer, "score": score, "feedback": feedback}
        )
        # Schedule a follow-up once per main question
        st.session_state[STATE_KEYS["followup_pending"]] = True
        st.session_state[STATE_KEYS["followup_question"]] = get_followup(q.get("text", ""), score)
        st.session_state[STATE_KEYS["followup_for_index"]] = idx
        st.rerun()

    # Show running results
    if st.session_state[STATE_KEYS["qa"]]:
        st.divider()
        st.markdown("**Running Scores**")
        for i, item in enumerate(st.session_state[STATE_KEYS["qa"]], start=1):
            st.write(f"Q{i}: Score {item['score']} — {item['feedback']}")


def generate_summary_report(qa_list):
    if not qa_list:
        return {
            "strengths": ["No answers provided"],
            "weaknesses": ["Provide answers to receive feedback"],
            "learning_path": ["Start with basic Excel tutorials and practice exercises."],
            "average_score": 0,
        }

    avg = sum(item.get("score", 0) for item in qa_list) / len(qa_list)

    strengths = []
    weaknesses = []
    for item in qa_list:
        feedback = (item.get("feedback") or "").lower()
        if any(k in feedback for k in ["good", "correct", "nice", "well", "clear", "complete"]):
            strengths.append(item.get("question", ""))
        if any(k in feedback for k in ["missing", "incorrect", "improve", "lack", "confus"]):
            weaknesses.append(item.get("question", ""))

    learning_path = [
        "Master lookup functions: VLOOKUP, INDEX/MATCH, XLOOKUP with approximate vs exact.",
        "Build PivotTables with slicers, top-N filters, and calculated fields.",
        "Use formula-based Conditional Formatting and manage rule precedence.",
        "Automate workflows with recorded macros and simple VBA refactors.",
        "Use Power Query to clean, merge, and append tables robustly."
    ]

    return {
        "strengths": strengths or ["Demonstrated understanding in some areas."],
        "weaknesses": weaknesses or ["Needs deeper explanations in several topics."],
        "learning_path": learning_path,
        "average_score": round(avg, 2),
    }


def overall_rating(avg_score: float) -> str:
    if avg_score >= 4.5:
        return "Outstanding"
    if avg_score >= 3.8:
        return "Strong"
    if avg_score >= 3.0:
        return "Competent"
    if avg_score >= 2.0:
        return "Developing"
    return "Needs Improvement"


def summary_view():
    st.success("Interview complete. Here's your summary report:")
    report = generate_summary_report(st.session_state[STATE_KEYS["qa"]])
    st.metric("Average Score", report["average_score"])

    st.markdown("**Strengths**")
    for s in report["strengths"]:
        st.write(f"- {s}")

    st.markdown("**Weaknesses**")
    for w in report["weaknesses"]:
        st.write(f"- {w}")

    st.markdown("**Suggested Learning Path**")
    for l in report["learning_path"]:
        st.write(f"- {l}")

    candidate = st.session_state[STATE_KEYS["candidate"]]
    candidate_name = candidate.get("name") or "Candidate"

    desired_pdf_name = f"excel_mock_interview_report_{candidate_name.replace(' ', '_')}.pdf"
    # Automatically build/rebuild the PDF
    pdf_buf = build_pdf_report(candidate, st.session_state[STATE_KEYS["qa"]], report)
    st.session_state[STATE_KEYS["pdf_bytes"]] = pdf_buf.getvalue()
    st.session_state[STATE_KEYS["pdf_name"]] = desired_pdf_name

    st.download_button(
        label="Download Report PDF",
        data=st.session_state[STATE_KEYS["pdf_bytes"]],
        file_name=st.session_state[STATE_KEYS["pdf_name"]] or desired_pdf_name,
        mime="application/pdf",
    )

    if st.button("Restart Interview"):
        for k in STATE_KEYS.values():
            if k in st.session_state:
                del st.session_state[k]
        st.rerun()


def main():
    init_state()
    render_header()

    stage = st.session_state[STATE_KEYS["stage"]]
    if stage == "intro":
        introduction()
    elif stage == "asking":
        ask_questions()
    elif stage == "summary":
        summary_view()
    else:
        st.write("Unknown stage.")


if __name__ == "__main__":
    main()
