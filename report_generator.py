from io import BytesIO
from datetime import datetime
from typing import List, Dict, Any

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors


def overall_rating(avg_score: float) -> str:
    # Per product requirements
    if avg_score <= 1.5:
        return "Needs Improvement"
    if avg_score <= 2.9:
        return "Developing"
    if avg_score <= 3.9:
        return "Competent"
    return "Strong"


def sanitize_feedback(text: str) -> str:
    if not text or "Error" in text or "error" in text:
        return "Evaluation not available due to a system error. Please retry."
    lowered = text.lower()
    replacements = {
        "completely inadequate": "The response does not demonstrate an understanding of the concept.",
        "terrible": "The response does not sufficiently address the question.",
        "bad": "The response is incomplete or contains inaccuracies.",
        "awful": "The explanation lacks clarity or correctness.",
    }
    for k, v in replacements.items():
        if k in lowered:
            return v
    return text


def _styles():
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Small", fontSize=9, leading=11))
    styles.add(ParagraphStyle(name="HeadingTight", parent=styles['Heading2'], spaceAfter=6))
    styles.add(ParagraphStyle(name="SectionTitle", parent=styles['Heading2'], spaceBefore=12, spaceAfter=6))
    return styles


def format_candidate_info(story, styles, candidate: Dict[str, Any]):
    story.append(Paragraph("Candidate Information", styles['SectionTitle']))
    meta = [
        ["Name", candidate.get('name') or 'N/A'],
        ["Email", candidate.get('email') or 'N/A'],
        ["Target Role", candidate.get('role') or 'Excel Analyst'],
        ["Experience", f"{candidate.get('experience') or '0'} years"],
        ["Date", datetime.now().strftime('%Y-%m-%d %H:%M')],
    ]
    meta_table = Table(meta, hAlign='LEFT', colWidths=[110, 380])
    meta_table.setStyle(TableStyle([
        ('FONT', (0,0), (-1,-1), 'Helvetica', 10),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('BOX', (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#f7f7f7')),
    ]))
    story.append(meta_table)


def format_summary(story, styles, average_score: float):
    story.append(Paragraph("Summary", styles['SectionTitle']))
    rating = overall_rating(average_score)
    rating_color = colors.HexColor('#2e7d32') if average_score >= 4.0 else (
        colors.HexColor('#1976d2') if average_score >= 3.0 else (
            colors.HexColor('#ef6c00') if average_score >= 1.6 else colors.HexColor('#c62828')
        )
    )
    summary_table = Table([
        ["Average Score (0–5)", f"{round(average_score, 2)}", "Overall Rating", rating]
    ], colWidths=[170, 80, 140, 120])
    summary_table.setStyle(TableStyle([
        ('FONT', (0,0), (-1,-1), 'Helvetica-Bold', 11),
        ('TEXTCOLOR', (3,0), (3,0), rating_color),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#eef2f7')),
        ('BOX', (0,0), (-1,-1), 0.5, colors.grey),
        ('INNERGRID', (0,0), (-1,-1), 0.25, colors.grey),
        ('ALIGN', (1,0), (1,0), 'LEFT'),
        ('ALIGN', (3,0), (3,0), 'LEFT'),
    ]))
    story.append(summary_table)


def format_question_evaluations(story, styles, qa_list: List[Dict[str, Any]]):
    story.append(Paragraph("Per-Question Evaluation", styles['SectionTitle']))
    data = [["Question", "Score", "Feedback"]]
    for item in qa_list:
        score = int(item.get("score", 0))
        feedback = sanitize_feedback(str(item.get("feedback", "")))
        data.append([item.get("question", ""), str(score), feedback])

    table = Table(data, colWidths=[280, 40, 220])
    tbl_style = [
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#f0f2f6')),
        ('GRID', (0,0), (-1,-1), 0.25, colors.grey),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONT', (0,1), (-1,-1), 'Helvetica', 9),
    ]
    # Highlight low/high scores
    for row_idx in range(1, len(data)):
        try:
            score_val = int(data[row_idx][1])
        except Exception:
            score_val = 0
        if score_val <= 1:
            tbl_style.append(('TEXTCOLOR', (1, row_idx), (1, row_idx), colors.HexColor('#c62828')))
        if score_val >= 4:
            tbl_style.append(('TEXTCOLOR', (1, row_idx), (1, row_idx), colors.HexColor('#2e7d32')))
    table.setStyle(TableStyle(tbl_style))
    story.append(table)


def format_feedback_sections(story, styles, strengths: List[str], weaknesses: List[str], learning_path: List[str]):
    story.append(Paragraph("Strengths", styles['SectionTitle']))
    for s in strengths:
        story.append(Paragraph(f"• {s}", styles['Small']))
    story.append(Spacer(1, 6))

    story.append(Paragraph("Areas to Improve", styles['SectionTitle']))
    for w in weaknesses:
        clean = sanitize_feedback(w)
        story.append(Paragraph(f"• {clean}", styles['Small']))
    story.append(Spacer(1, 6))

    story.append(Paragraph("Suggested Learning Path", styles['SectionTitle']))
    for l in learning_path:
        story.append(Paragraph(f"• {l}", styles['Small']))


def build_pdf_report(candidate: Dict[str, Any], qa_list: List[Dict[str, Any]], report: Dict[str, Any]) -> BytesIO:
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    styles = _styles()
    story = []

    # Title
    title = f"Excel Mock Interview Report — {candidate.get('name') or 'Candidate'}"
    story.append(Paragraph(title, styles['Title']))
    story.append(Spacer(1, 8))

    # Candidate Info
    format_candidate_info(story, styles, candidate)
    story.append(Spacer(1, 10))

    # Summary
    avg = float(report.get('average_score', 0))
    format_summary(story, styles, avg)
    story.append(Spacer(1, 10))

    # Per-Question evaluations
    format_question_evaluations(story, styles, qa_list)
    story.append(Spacer(1, 10))

    # Feedback sections
    strengths = report.get('strengths', [])
    weaknesses = report.get('weaknesses', [])
    learning_path = report.get('learning_path', [])
    format_feedback_sections(story, styles, strengths, weaknesses, learning_path)

    doc.build(story)
    buffer.seek(0)
    return buffer
