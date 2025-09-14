import os
import json
import requests
import streamlit as st
from dotenv import load_dotenv
from io import BytesIO
from datetime import datetime, timedelta
import time
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
import pandas as pd
import numpy as np
import altair as alt
from report_generator import build_pdf_report  # new modular generator

load_dotenv()

API_BASE = os.getenv("API_BASE", "http://localhost:8000")
TITLE = "Excel Mock Interviewer"

st.set_page_config(page_title=TITLE, page_icon="📊", layout="wide")

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
    "timer_start": "timer_start",
    "time_remaining": "time_remaining",
    "answer_submitted": "answer_submitted",
    "question_type": "question_type",
    "selected_mcq": "selected_mcq",
}

# Enhanced question bank with different types
QUESTION_BANK = [
    {
        "text": "What is the primary difference between VLOOKUP and XLOOKUP functions?",
        "type": "mcq",
        "options": [
            "XLOOKUP can only search left to right",
            "XLOOKUP can search in both directions and has better error handling",
            "VLOOKUP is faster than XLOOKUP",
            "There is no difference"
        ],
        "correct": 1,
        "level": "intermediate"
    },
    {
        "text": "Complete the formula: To sum values in cells A1 to A10 where corresponding B1 to B10 equals 'Sales', use =______(B1:B10,'Sales',A1:A10)",
        "type": "fill_blank",
        "answer": "SUMIF",
        "level": "basic"
    },
    {
        "text": "Analyze the sales data chart below. What is the trend and which month had the highest growth rate?",
        "type": "visual_analysis",
        "level": "advanced",
        "visual_data": "sales_trend"
    },
    {
        "text": "Which Excel function would you use to find the most frequently occurring value in a dataset?",
        "type": "mcq",
        "options": [
            "AVERAGE",
            "MEDIAN",
            "MODE.SNGL",
            "COUNT"
        ],
        "correct": 2,
        "level": "intermediate"
    },
    {
        "text": "In PivotTables, the ______ area is used to aggregate numerical data.",
        "type": "fill_blank",
        "answer": "VALUES",
        "level": "basic"
    },
    {
        "text": "Look at the data table below. Create a PivotTable strategy to analyze product performance by region and quarter.",
        "type": "data_analysis",
        "level": "advanced"
    },
    {
        "text": "What keyboard shortcut opens the 'Go To Special' dialog in Excel?",
        "type": "mcq",
        "options": [
            "Ctrl + G then Alt + S",
            "Ctrl + Shift + G",
            "F5 then Alt + S",
            "Both A and C"
        ],
        "correct": 3,
        "level": "intermediate"
    },
    {
        "text": "The Excel function =IF(A1>100,A1*0.1,A1*0.05) calculates a ______ based on the value in A1.",
        "type": "fill_blank",
        "answer": "commission",
        "alternative_answers": ["discount", "percentage", "rate"],
        "level": "intermediate"
    },
    {
        "text": "Examine the scatter plot below showing sales vs. advertising spend. What type of relationship exists and what would be your recommendation?",
        "type": "visual_analysis",
        "level": "advanced",
        "visual_data": "scatter_analysis"
    },
    {
        "text": "Which of these is NOT a valid Excel chart type?",
        "type": "mcq",
        "options": [
            "Waterfall Chart",
            "Sunburst Chart",
            "Treemap Chart",
            "Network Chart"
        ],
        "correct": 3,
        "level": "basic"
    }
]

def init_state():
    if STATE_KEYS["stage"] not in st.session_state:
        st.session_state[STATE_KEYS["stage"]] = "intro"
    if STATE_KEYS["questions"] not in st.session_state:
        # Use enhanced question bank
        st.session_state[STATE_KEYS["questions"]] = QUESTION_BANK
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
    if STATE_KEYS["timer_start"] not in st.session_state:
        st.session_state[STATE_KEYS["timer_start"]] = None
    if STATE_KEYS["time_remaining"] not in st.session_state:
        st.session_state[STATE_KEYS["time_remaining"]] = 15
    if STATE_KEYS["answer_submitted"] not in st.session_state:
        st.session_state[STATE_KEYS["answer_submitted"]] = False
    if STATE_KEYS["selected_mcq"] not in st.session_state:
        st.session_state[STATE_KEYS["selected_mcq"]] = None


def render_header():
    col1, col2, col3 = st.columns([3, 1, 1])
    with col1:
        st.title(TITLE)
    with col2:
        if st.session_state[STATE_KEYS["stage"]] == "asking":
            # Display timer
            if st.session_state[STATE_KEYS["timer_start"]] is not None:
                elapsed = (datetime.now() - st.session_state[STATE_KEYS["timer_start"]]).total_seconds()
                remaining = max(0, 15 - elapsed)
                st.session_state[STATE_KEYS["time_remaining"]] = remaining
                
                if remaining > 10:
                    st.success(f"⏱️ Time: {int(remaining)}s")
                elif remaining > 5:
                    st.warning(f"⏱️ Time: {int(remaining)}s")
                else:
                    st.error(f"⏱️ Time: {int(remaining)}s")
                
                if remaining <= 0 and not st.session_state[STATE_KEYS["answer_submitted"]]:
                    st.error("⏰ Time's up!")
    with col3:
        if st.session_state[STATE_KEYS["stage"]] == "asking":
            # Fix: Ensure progress value stays within [0.0, 1.0] range
            current_idx = st.session_state[STATE_KEYS["current_index"]]
            total_questions = len(st.session_state[STATE_KEYS["questions"]])
            
            # Clamp the progress value to prevent exceeding 1.0
            progress = min(1.0, (current_idx + 1) / total_questions) if total_questions > 0 else 0.0
            
            st.progress(progress)
            
            # Also fix the caption to show correct question numbers
            display_current = min(current_idx + 1, total_questions)
            st.caption(f"Q {display_current}/{total_questions}")

def introduction():
    st.info("👋 Hi, I'm your Excel Mock Interviewer. You'll have 15 seconds per question. Questions include MCQs, fill-in-the-blanks, and data analysis.")
    
    with st.form("candidate_form"):
        col1, col2 = st.columns(2)
        with col1:
            name = st.text_input("Full name*", value=st.session_state[STATE_KEYS["candidate"]].get("name", ""))
            email = st.text_input("Email (optional)", value=st.session_state[STATE_KEYS["candidate"]].get("email", ""))
        with col2:
            role = st.text_input("Target role", value=st.session_state[STATE_KEYS["candidate"]].get("role", "Excel Analyst"))
            experience = st.selectbox("Years of experience", ["0-1", "2-3", "4-5", "6-10", "10+"], 
                                     index=0)
        submitted = st.form_submit_button("🚀 Start Interview", type="primary", use_container_width=True)
    
    if submitted:
        st.session_state[STATE_KEYS["candidate"]] = {
            "name": name.strip(),
            "email": email.strip(),
            "role": role.strip() or "Excel Analyst",
            "experience": experience,
        }
        if name.strip():
            st.session_state[STATE_KEYS["stage"]] = "asking"
            st.session_state[STATE_KEYS["timer_start"]] = datetime.now()
            st.rerun()
        else:
            st.warning("Please enter your full name to start.")

def generate_visual_data(visual_type):
    """Generate interactive visual data for questions"""
    if visual_type == "sales_trend":
        # Generate sales trend data
        months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
        sales = [45000, 48000, 52000, 49000, 58000, 65000]
        growth = [0] + [(sales[i] - sales[i-1])/sales[i-1]*100 for i in range(1, len(sales))]
        
        df = pd.DataFrame({
            "Month": months,
            "Sales": sales,
            "Growth %": growth
        })
        
        # Create interactive chart
        base = alt.Chart(df).encode(x="Month:O")
        
        bars = base.mark_bar(color="steelblue").encode(
            y=alt.Y("Sales:Q", title="Sales ($)"),
            tooltip=["Month", "Sales", alt.Tooltip("Growth %:Q", format=".1f")]
        )
        
        line = base.mark_line(color="red", strokeWidth=2).encode(
            y=alt.Y("Growth %:Q", title="Growth Rate (%)", axis=alt.Axis(orient="right")),
            tooltip=alt.Tooltip("Growth %:Q", format=".1f")
        )
        
        points = line.mark_point(color="red", size=100)
        
        chart = alt.layer(bars, line + points).resolve_scale(y="independent").properties(
            width=600, height=300, title="Monthly Sales Performance Analysis"
        )
        
        return df, chart
    
    elif visual_type == "scatter_analysis":
        # Generate scatter plot data
        np.random.seed(42)
        ad_spend = np.random.uniform(1000, 10000, 50)
        sales = ad_spend * 3.5 + np.random.normal(0, 1000, 50) + 5000
        
        df = pd.DataFrame({
            "Ad Spend ($)": ad_spend,
            "Sales ($)": sales,
            "ROI": sales / ad_spend
        })
        
        scatter = alt.Chart(df).mark_circle(size=100).encode(
            x=alt.X("Ad Spend ($):Q", scale=alt.Scale(zero=False)),
            y=alt.Y("Sales ($):Q", scale=alt.Scale(zero=False)),
            color=alt.Color("ROI:Q", scale=alt.Scale(scheme="viridis")),
            tooltip=[alt.Tooltip("Ad Spend ($):Q", format=",.0f"),
                    alt.Tooltip("Sales ($):Q", format=",.0f"),
                    alt.Tooltip("ROI:Q", format=".2f")]
        ).properties(
            width=600, height=400,
            title="Sales vs. Advertising Spend Analysis"
        )
        
        # Add trend line
        scatter = scatter + scatter.transform_regression(
            "Ad Spend ($)", "Sales ($)"
        ).mark_line(color="red", strokeDash=[5, 5])
        
        return df, scatter
    
    elif visual_type == "pivot_data":
        # Generate sample data for PivotTable question
        products = ["Laptop", "Phone", "Tablet", "Monitor", "Keyboard"]
        regions = ["North", "South", "East", "West"]
        quarters = ["Q1", "Q2", "Q3", "Q4"]
        
        data = []
        for _ in range(100):
            data.append({
                "Product": np.random.choice(products),
                "Region": np.random.choice(regions),
                "Quarter": np.random.choice(quarters),
                "Sales": np.random.randint(10000, 100000),
                "Units": np.random.randint(10, 500)
            })
        
        df = pd.DataFrame(data)
        return df, None
    
    return None, None

def grade_answer(question, answer, question_type="text"):
    """Grade different types of answers"""
    if question_type == "mcq":
        # For MCQ, check if the selected option is correct
        correct_idx = question.get("correct", -1)
        selected_idx = answer  # answer is the index for MCQ
        if selected_idx == correct_idx:
            return 5, "Correct! Well done."
        else:
            return 2, f"Incorrect. The correct answer was: {question['options'][correct_idx]}"
    
    elif question_type == "fill_blank":
        # For fill-in-the-blank, check exact or alternative answers
        correct = question.get("answer", "").lower().strip()
        alternatives = question.get("alternative_answers", [])
        user_answer = answer.lower().strip()
        
        if user_answer == correct or user_answer in [a.lower().strip() for a in alternatives]:
            return 5, "Correct! Perfect answer."
        elif correct in user_answer or any(alt in user_answer for alt in alternatives):
            return 3, "Partially correct. Close, but not exact."
        else:
            return 1, f"Incorrect. The answer was: {question.get('answer')}"
    
    else:
        # For other types, use API or default grading
        try:
            resp = requests.post(
                f"{API_BASE}/grade",
                json={"question": question.get("text", ""), "answer": answer},
                timeout=30,
            )
            if resp.ok:
                data = resp.json()
                return int(data.get("score", 0)), str(data.get("feedback", ""))
        except Exception:
            # Fallback grading based on answer length and keywords
            if len(answer) > 50:
                return 3, "Good attempt. Consider being more specific."
            else:
                return 2, "Answer needs more detail and explanation."
    
    return 2, "Answer evaluated."

def ask_questions():
    questions = st.session_state[STATE_KEYS["questions"]]
    idx = st.session_state[STATE_KEYS["current_index"]]
    
    if not questions:
        st.error("No questions available.")
        return
    
    if idx >= len(questions):
        st.session_state[STATE_KEYS["stage"]] = "summary"
        st.rerun()
        return
    
    ##Start timer for new question
    if st.session_state[STATE_KEYS["timer_start"]] is None:
        st.session_state[STATE_KEYS["timer_start"]] = datetime.now()
        st.session_state[STATE_KEYS["answer_submitted"]] = False
    
    q = questions[idx]
    question_type = q.get("type", "text")
    
    # Question header
    col1, col2 = st.columns([3, 1])
    with col1:
        st.subheader(f"Question {idx + 1}: {q.get('level', '').title()} Level")
    with col2:
        badge_color = {"basic": "🟢", "intermediate": "🟡", "advanced": "🔴"}.get(q.get("level", ""), "⚪")
        st.markdown(f"### {badge_color}")
    
    # Display question based on type
    if question_type == "mcq":
        st.markdown(f"**{q.get('text', '')}**")
        
        # MCQ Options
        selected = st.radio(
            "Select your answer:",
            options=range(len(q.get("options", []))),
            format_func=lambda x: q["options"][x],
            key=f"mcq_{idx}",
            disabled=st.session_state[STATE_KEYS["answer_submitted"]]
        )
        
        if st.button("Submit Answer", disabled=st.session_state[STATE_KEYS["answer_submitted"]] or 
                     st.session_state[STATE_KEYS["time_remaining"]] <= 0):
            score, feedback = grade_answer(q, selected, "mcq")
            st.session_state[STATE_KEYS["qa"]].append({
                "question": q.get("text", ""),
                "answer": q["options"][selected],
                "score": score,
                "feedback": feedback,
                "time_taken": 15 - st.session_state[STATE_KEYS["time_remaining"]]
            })
            st.session_state[STATE_KEYS["answer_submitted"]] = True
            time.sleep(1)
            st.session_state[STATE_KEYS["current_index"]] += 1
            st.session_state[STATE_KEYS["timer_start"]] = None
            st.rerun()
    
    elif question_type == "fill_blank":
        st.markdown(f"**{q.get('text', '')}**")
        
        answer = st.text_input(
            "Fill in the blank:",
            key=f"fill_{idx}",
            disabled=st.session_state[STATE_KEYS["answer_submitted"]]
        )
        
        if st.button("Submit Answer", disabled=st.session_state[STATE_KEYS["answer_submitted"]] or 
                     st.session_state[STATE_KEYS["time_remaining"]] <= 0):
            score, feedback = grade_answer(q, answer, "fill_blank")
            st.session_state[STATE_KEYS["qa"]].append({
                "question": q.get("text", ""),
                "answer": answer,
                "score": score,
                "feedback": feedback,
                "time_taken": 15 - st.session_state[STATE_KEYS["time_remaining"]]
            })
            st.session_state[STATE_KEYS["answer_submitted"]] = True
            time.sleep(1)
            st.session_state[STATE_KEYS["current_index"]] += 1
            st.session_state[STATE_KEYS["timer_start"]] = None
            st.rerun()
    
    elif question_type in ["visual_analysis", "data_analysis"]:
        st.markdown(f"**{q.get('text', '')}**")
        
        # Generate and display visual data
        visual_data = q.get("visual_data", "pivot_data")
        df, chart = generate_visual_data(visual_data)
        
        if df is not None:
            col1, col2 = st.columns([1, 1])
            with col1:
                st.markdown("📊 **Data Table:**")
                st.dataframe(df.head(10), use_container_width=True)
            with col2:
                if chart is not None:
                    st.markdown("📈 **Visualization:**")
                    st.altair_chart(chart, use_container_width=True)
        
        answer = st.text_area(
            "Your analysis:",
            height=100,
            key=f"analysis_{idx}",
            disabled=st.session_state[STATE_KEYS["answer_submitted"]]
        )
        
        if st.button("Submit Answer", disabled=st.session_state[STATE_KEYS["answer_submitted"]] or 
                     st.session_state[STATE_KEYS["time_remaining"]] <= 0):
            score, feedback = grade_answer(q, answer, "analysis")
            st.session_state[STATE_KEYS["qa"]].append({
                "question": q.get("text", ""),
                "answer": answer,
                "score": score,
                "feedback": feedback,
                "time_taken": 15 - st.session_state[STATE_KEYS["time_remaining"]]
            })
            st.session_state[STATE_KEYS["answer_submitted"]] = True
            time.sleep(1)
            st.session_state[STATE_KEYS["current_index"]] += 1
            st.session_state[STATE_KEYS["timer_start"]] = None
            st.rerun()
    
    else:  # Default text question
        st.markdown(f"**{q.get('text', '')}**")
        
        answer = st.text_area(
            "Your answer:",
            height=120,
            key=f"text_{idx}",
            disabled=st.session_state[STATE_KEYS["answer_submitted"]]
        )
        
        if st.button("Submit Answer", disabled=st.session_state[STATE_KEYS["answer_submitted"]] or 
                     st.session_state[STATE_KEYS["time_remaining"]] <= 0):
            score, feedback = grade_answer(q, answer, "text")
            st.session_state[STATE_KEYS["qa"]].append({
                "question": q.get("text", ""),
                "answer": answer,
                "score": score,
                "feedback": feedback,
                "time_taken": 15 - st.session_state[STATE_KEYS["time_remaining"]]
            })
            st.session_state[STATE_KEYS["answer_submitted"]] = True
            time.sleep(1)
            st.session_state[STATE_KEYS["current_index"]] += 1
            st.session_state[STATE_KEYS["timer_start"]] = None
            st.rerun()
    
    # Auto-submit when timer reaches 0
    if st.session_state[STATE_KEYS["time_remaining"]] <= 0 and not st.session_state[STATE_KEYS["answer_submitted"]]:
        st.session_state[STATE_KEYS["qa"]].append({
            "question": q.get("text", ""),
            "answer": "Time expired - No answer provided",
            "score": 0,
            "feedback": "Time ran out before answer was submitted.",
            "time_taken": 15
        })
        st.session_state[STATE_KEYS["answer_submitted"]] = True
        time.sleep(1)
        st.session_state[STATE_KEYS["current_index"]] += 1
        st.session_state[STATE_KEYS["timer_start"]] = None
        st.rerun()
    
    # Show progress
    if st.session_state[STATE_KEYS["qa"]]:
        st.divider()
        st.markdown("### 📊 Performance Summary")
        cols = st.columns(4)
        scores = [item["score"] for item in st.session_state[STATE_KEYS["qa"]]]
        with cols[0]:
            st.metric("Questions Answered", len(scores))
        with cols[1]:
            st.metric("Average Score", f"{np.mean(scores):.1f}/5" if scores else "0/5")
        with cols[2]:
            st.metric("Correct Answers", sum(1 for s in scores if s >= 4))
        with cols[3]:
            avg_time = np.mean([item.get("time_taken", 15) for item in st.session_state[STATE_KEYS["qa"]]])
            st.metric("Avg. Time", f"{avg_time:.1f}s")
    
    # Refresh timer
    if st.session_state[STATE_KEYS["time_remaining"]] > 0 and not st.session_state[STATE_KEYS["answer_submitted"]]:
        time.sleep(0.1)
        st.rerun()

def generate_summary_report(qa_list):
    if not qa_list:
        return {
            "strengths": ["No answers provided"],
            "weaknesses": ["Provide answers to receive feedback"],
            "learning_path": ["Start with basic Excel tutorials and practice exercises."],
            "average_score": 0,
            "time_efficiency": 0
        }
    
    avg = sum(item.get("score", 0) for item in qa_list) / len(qa_list)
    avg_time = np.mean([item.get("time_taken", 15) for item in qa_list])
    
    strengths = []
    weaknesses = []
    
    for item in qa_list:
        if item.get("score", 0) >= 4:
            strengths.append(f"{item.get('question', '')[:50]}...")
        elif item.get("score", 0) <= 2:
            weaknesses.append(f"{item.get('question', '')[:50]}...")
    
    learning_path = [
        "Master lookup functions: VLOOKUP, INDEX/MATCH, XLOOKUP",
        "Build complex PivotTables with calculated fields",
        "Learn Power Query for data transformation",
        "Practice data visualization and chart selection",
        "Develop VBA skills for automation"
    ]
    
    return {
        "strengths": strengths or ["Good overall understanding"],
        "weaknesses": weaknesses or ["Room for improvement in advanced topics"],
        "learning_path": learning_path,
        "average_score": round(avg, 2),
        "time_efficiency": round((15 - avg_time) / 15 * 100, 1)
    }

def summary_view():
    st.balloons()
    st.success("🎉 Interview Complete! Here's your comprehensive report:")
    
    report = generate_summary_report(st.session_state[STATE_KEYS["qa"]])
    
    # Metrics row
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Overall Score", f"{report['average_score']}/5.0", 
                 delta=f"{(report['average_score']/5*100):.0f}%")
    with col2:
        total_q = len(st.session_state[STATE_KEYS["qa"]])
        correct = sum(1 for item in st.session_state[STATE_KEYS["qa"]] if item.get("score", 0) >= 4)
        st.metric("Accuracy", f"{correct}/{total_q}", 
                 delta=f"{(correct/total_q*100):.0f}%")
    with col3:
        st.metric("Time Efficiency", f"{report['time_efficiency']:.0f}%")
    with col4:
        rating = "⭐" * min(5, max(1, int(report['average_score'])))
        st.metric("Rating", rating)
    
    # Detailed breakdown
    st.divider()
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### ✅ Strengths")
        for s in report["strengths"][:5]:
            st.write(f"• {s}")
    
    with col2:
        st.markdown("### 📚 Areas for Improvement")
        for w in report["weaknesses"][:5]:
            st.write(f"• {w}")
    
    st.divider()
    
    # Question-by-question breakdown
    st.markdown("### 📋 Detailed Results")
    
    results_df = pd.DataFrame([
        {
            "Question": item["question"][:60] + "...",
            "Score": f"{item['score']}/5",
            "Time": f"{item.get('time_taken', 15):.1f}s",
            "Result": "✅" if item["score"] >= 4 else "❌"
        }
        for item in st.session_state[STATE_KEYS["qa"]]
    ])
    
    st.dataframe(results_df, use_container_width=True, hide_index=True)
    
    # Learning recommendations
    st.markdown("### 🎯 Personalized Learning Path")
    for i, path in enumerate(report["learning_path"], 1):
        st.write(f"{i}. {path}")
    
    # Download and restart buttons
    col1, col2 = st.columns(2)
    
    with col1:
        candidate = st.session_state[STATE_KEYS["candidate"]]
        candidate_name = candidate.get("name") or "Candidate"
        pdf_name = f"excel_interview_{candidate_name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.pdf"
        
        # FIXED: Generate PDF using the report_generator
        if st.button("📄 Generate PDF Report", type="primary", use_container_width=True):
            try:
                with st.spinner("Generating comprehensive PDF report..."):
                    # Prepare enhanced report data for PDF generation
                    enhanced_report = {
                        **report,
                        'qa_list': st.session_state[STATE_KEYS["qa"]],
                        'category_breakdown': analyze_performance_by_category(st.session_state[STATE_KEYS["qa"]]),
                        'difficulty_analysis': analyze_performance_by_difficulty(st.session_state[STATE_KEYS["qa"]]),
                        'time_analysis': analyze_time_performance(st.session_state[STATE_KEYS["qa"]]),
                        'proctoring_notes': generate_proctoring_notes(st.session_state[STATE_KEYS["qa"]])
                    }
                    
                    # Generate PDF using your report generator
                    pdf_buffer = build_pdf_report(
                        candidate=candidate,
                        qa_list=st.session_state[STATE_KEYS["qa"]],
                        report=enhanced_report
                    )
                    
                    # Store in session state
                    st.session_state[STATE_KEYS["pdf_bytes"]] = pdf_buffer.getvalue()
                    st.session_state[STATE_KEYS["pdf_name"]] = pdf_name
                    
                    st.success("✅ PDF report generated successfully!")
                    
            except Exception as e:
                st.error(f"Error generating PDF: {str(e)}")
                st.error("Please try again or contact support if the issue persists.")
        
        # Download button (appears after PDF is generated)
        if st.session_state[STATE_KEYS["pdf_bytes"]] is not None:
            st.download_button(
                label="⬇️ Download PDF Report",
                data=st.session_state[STATE_KEYS["pdf_bytes"]],
                file_name=st.session_state[STATE_KEYS["pdf_name"]],
                mime="application/pdf",
                use_container_width=True
            )
    
    with col2:
        if st.button("🔄 Start New Interview", use_container_width=True):
            for k in STATE_KEYS.values():
                if k in st.session_state:
                    del st.session_state[k]
            st.rerun()


# Helper functions for enhanced report data
def analyze_performance_by_category(qa_list):
    """Analyze performance by question categories"""
    if not qa_list:
        return {}
    
    categories = {}
    for qa in qa_list:
        # Extract category from question or use default
        category = qa.get('category', 'general')
        if not category or category == 'general':
            # Try to infer category from question text
            question_text = qa.get('question', '').lower()
            if any(word in question_text for word in ['vlookup', 'xlookup', 'index', 'match']):
                category = 'lookup_functions'
            elif any(word in question_text for word in ['pivot', 'pivottable']):
                category = 'pivot_tables'
            elif any(word in question_text for word in ['chart', 'graph', 'visualization']):
                category = 'visualization'
            elif any(word in question_text for word in ['formula', 'function']):
                category = 'functions'
            else:
                category = 'general'
        
        if category not in categories:
            categories[category] = {'scores': [], 'count': 0}
        
        categories[category]['scores'].append(qa.get('score', 0))
        categories[category]['count'] += 1
    
    # Calculate averages
    for category in categories:
        scores = categories[category]['scores']
        categories[category]['average'] = sum(scores) / len(scores) if scores else 0
    
    return categories


def analyze_performance_by_difficulty(qa_list):
    """Analyze performance by difficulty levels"""
    if not qa_list:
        return {}
    
    difficulties = {}
    for qa in qa_list:
        # Get difficulty from question data or infer from score patterns
        difficulty = qa.get('level', 'intermediate')
        
        if difficulty not in difficulties:
            difficulties[difficulty] = {'scores': [], 'count': 0}
        
        difficulties[difficulty]['scores'].append(qa.get('score', 0))
        difficulties[difficulty]['count'] += 1
    
    # Calculate averages
    for difficulty in difficulties:
        scores = difficulties[difficulty]['scores']
        difficulties[difficulty]['average'] = sum(scores) / len(scores) if scores else 0
    
    return difficulties


def analyze_time_performance(qa_list):
    """Analyze time-related performance metrics"""
    if not qa_list:
        return {}
    
    times = [qa.get('time_taken', 15) for qa in qa_list if qa.get('time_taken')]
    total_time = sum(times) if times else 0
    avg_time = total_time / len(times) if times else 0
    
    # Determine efficiency rating
    if avg_time <= 10:
        efficiency = "Excellent"
    elif avg_time <= 12:
        efficiency = "Good" 
    elif avg_time <= 14:
        efficiency = "Average"
    else:
        efficiency = "Needs Improvement"
    
    return {
        'total_time': total_time,
        'average_time_per_question': avg_time,
        'time_efficiency': efficiency,
        'fastest_response': min(times) if times else 0,
        'slowest_response': max(times) if times else 0
    }


def generate_proctoring_notes(qa_list):
    """Generate proctoring and integrity notes"""
    if not qa_list:
        return []
    
    notes = []
    
    # Check for consistent timing patterns
    times = [qa.get('time_taken', 15) for qa in qa_list if qa.get('time_taken')]
    if times:
        avg_time = sum(times) / len(times)
        if avg_time < 5:
            notes.append("Unusually fast response times detected - may indicate prior knowledge or assistance")
        elif avg_time > 14.5:
            notes.append("Consistent use of full time allocation - thorough approach to questions")
        
        # Check for time pattern consistency
        if len(times) > 3:
            import statistics
            std_dev = statistics.stdev(times)
            if std_dev < 2:
                notes.append("Very consistent response timing observed")
            elif std_dev > 8:
                notes.append("Highly variable response timing - adapted approach per question type")
    
    # Check score patterns
    scores = [qa.get('score', 0) for qa in qa_list]
    if scores:
        if all(score >= 4 for score in scores):
            notes.append("Exceptional performance across all questions")
        elif all(score <= 1 for score in scores):
            notes.append("Consistent low performance - may need skill development")
    
    # Check for incomplete responses
    incomplete = sum(1 for qa in qa_list if not qa.get('answer') or qa.get('answer', '').strip() == '')
    if incomplete > 0:
        notes.append(f"{incomplete} questions had incomplete or missing responses")
    
    return notes if notes else ["Standard assessment completion with no integrity concerns noted"]

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