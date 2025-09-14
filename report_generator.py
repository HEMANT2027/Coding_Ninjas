from io import BytesIO
from datetime import datetime
from typing import List, Dict, Any
import statistics

from reportlab.lib.pagesizes import A4, letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.piecharts import Pie
from reportlab.lib.units import inch


def overall_rating(avg_score: float) -> str:
    """Enhanced rating system with more granular categories"""
    if avg_score >= 4.5:
        return "Outstanding (Top 5%)"
    elif avg_score >= 4.0:
        return "Excellent (Top 15%)"
    elif avg_score >= 3.5:
        return "Strong (Above Average)"
    elif avg_score >= 3.0:
        return "Competent (Average)"
    elif avg_score >= 2.0:
        return "Developing (Below Average)"
    else:
        return "Needs Improvement (Bottom 20%)"


def get_performance_color(score: float) -> colors.Color:
    """Return color based on performance score"""
    if score >= 4.5:
        return colors.HexColor('#1B5E20')  # Dark Green
    elif score >= 4.0:
        return colors.HexColor('#2E7D32')  # Green
    elif score >= 3.5:
        return colors.HexColor('#388E3C')  # Light Green
    elif score >= 3.0:
        return colors.HexColor('#FFA000')  # Amber
    elif score >= 2.0:
        return colors.HexColor('#F57C00')  # Orange
    else:
        return colors.HexColor('#D32F2F')  # Red


def truncate_long_text(text: str, max_length: int = 400) -> str:
    """Truncate long text to prevent PDF overflow"""
    if not text:
        return ""
    
    if len(text) <= max_length:
        return text
    
    # Find the last complete sentence within the limit
    truncated = text[:max_length]
    last_period = truncated.rfind('.')
    last_question = truncated.rfind('?')
    last_exclamation = truncated.rfind('!')
    
    last_sentence_end = max(last_period, last_question, last_exclamation)
    
    if last_sentence_end > max_length * 0.7:  # If we can keep most of the text
        return text[:last_sentence_end + 1] + "... [Response truncated for report brevity]"
    else:
        return truncated + "... [Response truncated for report brevity]"


def sanitize_feedback(text: str) -> str:
    """Enhanced feedback sanitization with more professional language"""
    if not text or "Error" in text or "error" in text:
        return "Technical evaluation unavailable. Response noted for manual review."
    
    # More comprehensive replacements for professional report
    replacements = {
        "completely inadequate": "The response demonstrates limited understanding of the core concepts.",
        "terrible": "The response requires significant improvement to meet professional standards.",
        "bad": "The response contains gaps in understanding or application.",
        "awful": "The explanation lacks the necessary depth and accuracy.",
        "wrong": "The approach described may lead to incorrect results.",
        "horrible": "The response needs substantial enhancement in both content and clarity.",
        "stupid": "The response shows misunderstanding of fundamental principles.",
        "useless": "The response does not adequately address the question requirements."
    }
    
    text_lower = text.lower()
    for negative, professional in replacements.items():
        if negative in text_lower:
            return professional
    
    # Enhance positive feedback too
    positive_enhancements = {
        "good": "demonstrates solid understanding",
        "nice": "shows good practical knowledge",
        "great": "exhibits excellent comprehension",
        "perfect": "demonstrates mastery-level understanding"
    }
    
    for casual, professional in positive_enhancements.items():
        if casual in text_lower and len(text) < 50:  # Only for short feedback
            return text.replace(casual, professional)
    
    return text


def _enhanced_styles():
    """Create enhanced styles for professional report"""
    styles = getSampleStyleSheet()
    
    # Add custom styles
    styles.add(ParagraphStyle(
        name='CustomTitle',
        parent=styles['Title'],
        fontSize=24,
        spaceAfter=30,
        textColor=colors.HexColor('#1565C0'),
        alignment=TA_CENTER
    ))
    
    styles.add(ParagraphStyle(
        name='SectionHeader',
        parent=styles['Heading2'],
        fontSize=16,
        spaceBefore=20,
        spaceAfter=10,
        textColor=colors.HexColor('#1976D2'),
        borderWidth=1,
        borderColor=colors.HexColor('#E3F2FD'),
        borderPadding=5,
        backColor=colors.HexColor('#F8FBFF')
    ))
    
    styles.add(ParagraphStyle(
        name='SubHeader',
        parent=styles['Heading3'],
        fontSize=14,
        spaceBefore=15,
        spaceAfter=8,
        textColor=colors.HexColor('#424242')
    ))
    
    styles.add(ParagraphStyle(
        name='CustomBodyText',
        parent=styles['Normal'],
        fontSize=10,
        leading=16,
        spaceAfter=8,
        alignment=TA_JUSTIFY,
        wordWrap='LTR'
    ))
    
    styles.add(ParagraphStyle(
        name='BulletPoint',
        parent=styles['Normal'],
        fontSize=10,
        leading=15,
        leftIndent=20,
        bulletIndent=10,
        spaceAfter=6,
        wordWrap='LTR'
    ))
    
    styles.add(ParagraphStyle(
        name='SmallText',
        fontSize=8,
        leading=12,
        textColor=colors.grey,
        wordWrap='LTR'
    ))
    
    styles.add(ParagraphStyle(
        name='HighlightBox',
        parent=styles['Normal'],
        fontSize=11,
        leading=16,
        backColor=colors.HexColor('#E8F5E8'),
        borderColor=colors.HexColor('#4CAF50'),
        borderWidth=1,
        borderPadding=10,
        spaceAfter=12,
        wordWrap='LTR'
    ))
    
    return styles


def create_performance_chart(category_data: Dict[str, Dict]) -> Drawing:
    """Create a performance visualization chart"""
    drawing = Drawing(400, 200)
    
    # Create bar chart
    chart = VerticalBarChart()
    chart.x = 50
    chart.y = 50
    chart.height = 125
    chart.width = 300
    
    categories = list(category_data.keys())[:6]  # Limit to 6 categories for readability
    scores = [category_data[cat]["average"] for cat in categories]
    
    chart.data = [scores]
    chart.categoryAxis.categoryNames = [cat.replace('_', ' ').title()[:10] for cat in categories]
    chart.valueAxis.valueMin = 0
    chart.valueAxis.valueMax = 5
    chart.valueAxis.valueStep = 1
    
    # Styling
    chart.bars[0].fillColor = colors.HexColor('#2196F3')
    chart.categoryAxis.labels.angle = 45
    chart.categoryAxis.labels.fontSize = 8
    chart.valueAxis.labels.fontSize = 8
    
    drawing.add(chart)
    return drawing


def format_executive_summary(story, styles, candidate: Dict[str, Any], report: Dict[str, Any]):
    """Create executive summary section"""
    story.append(Paragraph("Executive Summary", styles['SectionHeader']))
    
    avg_score = report.get('average_score', 0)
    rating = overall_rating(avg_score)
    
    # Performance overview
    candidate_name = candidate.get('name', 'Candidate')
    role = candidate.get('role', 'Excel Analyst')
    experience = candidate.get('experience', '0')
    
    readiness_text = ('<b>strong readiness</b> for the target role' if avg_score >= 3.5 
                     else '<b>potential for growth</b> with additional development' if avg_score >= 2.5 
                     else '<b>significant skill development needs</b> before role readiness')
    
    summary_text = f"""
    <b>{candidate_name}</b> completed a comprehensive Excel proficiency assessment 
    targeting the <b>{role}</b> position. With <b>{experience} years</b> 
    of stated experience, the candidate achieved an overall score of <b>{avg_score:.1f}/5.0</b>, 
    placing them in the <b>"{rating}"</b> performance category.
    
    The assessment covered multiple competency areas including data analysis, advanced functions, 
    visualization, automation, and business intelligence tools. Based on the comprehensive evaluation, 
    the candidate demonstrates {readiness_text}.
    """
    
    story.append(Paragraph(summary_text, styles['CustomBodyText']))
    story.append(Spacer(1, 15))
    
    # Key metrics table - FIXED VERSION
    metrics_data = [
        ["Assessment Metric", "Result", "Benchmark"],
        ["Overall Score", f"{avg_score:.1f}/5.0", "3.5+ (Role Ready)"],
        ["Questions Completed", str(len(report.get('qa_list', []))), f"{len(report.get('qa_list', []))} Total"],
        ["Performance Rating", rating, "Strong+ Recommended"],
        ["Assessment Duration", f"{report.get('time_analysis', {}).get('total_time', 0)//60:.0f} minutes", "45-90 minutes"],
    ]
    
    # Fixed column widths - made wider for better text fit
    metrics_table = Table(metrics_data, colWidths=[140, 160, 180])  # Increased column widths
    metrics_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1976D2')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold', 10),
        ('FONT', (0, 1), (-1, -1), 'Helvetica', 9),  # Reduced font size for better fit
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),  # Left align for better readability
        ('PADDING', (0, 0), (-1, -1), 8),  # Increased padding
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F5F5F5')]),
        ('WORDWRAP', (0, 0), (-1, -1), True)  # Enable word wrapping
    ]))
    
    story.append(metrics_table)
    story.append(Spacer(1, 20))


def format_candidate_profile(story, styles, candidate: Dict[str, Any]):
    """Enhanced candidate information with professional formatting"""
    story.append(Paragraph("Candidate Profile", styles['SectionHeader']))
    
    profile_data = [
        ["Full Name:", candidate.get('name', 'N/A')],
        ["Email Address:", candidate.get('email', 'N/A')],
        ["Target Position:", candidate.get('role', 'Excel Analyst')],
        ["Stated Experience:", f"{candidate.get('experience', '0')} years in Excel"],
        ["Assessment Date:", datetime.now().strftime('%B %d, %Y at %I:%M %p')],
        ["Assessment Type:", "AI-Powered Comprehensive Excel Evaluation"]
    ]
    
    # Fixed column widths for candidate profile
    profile_table = Table(profile_data, colWidths=[140, 340])  # Adjusted widths
    profile_table.setStyle(TableStyle([
        ('FONT', (0, 0), (0, -1), 'Helvetica-Bold', 10),
        ('FONT', (1, 0), (1, -1), 'Helvetica', 10),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('ALIGN', (1, 0), (1, -1), 'LEFT'),
        ('PADDING', (0, 0), (-1, -1), 8),
        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#F0F4F8')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E0E0E0')),
        ('WORDWRAP', (0, 0), (-1, -1), True)
    ]))
    
    story.append(profile_table)
    story.append(Spacer(1, 20))


def format_detailed_performance_analysis(story, styles, report: Dict[str, Any]):
    """Comprehensive performance breakdown with analytics"""
    story.append(Paragraph("Detailed Performance Analysis", styles['SectionHeader']))
    
    # Category performance breakdown
    if report.get('category_breakdown'):
        story.append(Paragraph("Performance by Excel Competency Area", styles['SubHeader']))
        
        category_data = []
        category_data.append(["Competency Area", "Questions", "Average Score", "Performance Level", "Industry Benchmark"])
        
        for category, data in report['category_breakdown'].items():
            category_name = category.replace('_', ' ').title()
            avg_score = data['average']
            performance_level = "Excellent" if avg_score >= 4.0 else "Good" if avg_score >= 3.0 else "Developing"
            benchmark = "Above Average" if avg_score >= 3.5 else "Average" if avg_score >= 2.5 else "Below Average"
            
            category_data.append([
                category_name,
                str(data['count']),
                f"{avg_score:.1f}/5.0",
                performance_level,
                benchmark
            ])
        
        # Fixed column widths for category table
        category_table = Table(category_data, colWidths=[120, 60, 80, 90, 100])  # Adjusted widths
        
        # Enhanced table styling with performance-based colors
        table_style = [
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1565C0')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold', 8),  # Smaller header font
            ('FONT', (0, 1), (-1, -1), 'Helvetica', 8),  # Smaller body font
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('PADDING', (0, 0), (-1, -1), 4),  # Reduced padding
            ('WORDWRAP', (0, 0), (-1, -1), True)
        ]
        
        # Color-code performance levels
        for i, (_, data) in enumerate(report['category_breakdown'].items(), 1):
            score_color = get_performance_color(data['average'])
            table_style.append(('TEXTCOLOR', (2, i), (2, i), score_color))
            table_style.append(('BACKGROUND', (3, i), (3, i), colors.HexColor('#FAFAFA')))
        
        category_table.setStyle(TableStyle(table_style))
        story.append(category_table)
        story.append(Spacer(1, 15))
    
    # Difficulty progression analysis
    if report.get('difficulty_analysis'):
        story.append(Paragraph("Skill Progression Analysis", styles['SubHeader']))
        
        difficulty_text = "Assessment results across difficulty levels indicate: "
        diff_analysis = report['difficulty_analysis']
        
        if 'beginner' in diff_analysis:
            beginner_score = diff_analysis['beginner']['average']
            difficulty_text += f"<b>Foundational skills</b> (Beginner): {beginner_score:.1f}/5.0"
        
        if 'intermediate' in diff_analysis:
            intermediate_score = diff_analysis['intermediate']['average']
            difficulty_text += f" | <b>Applied skills</b> (Intermediate): {intermediate_score:.1f}/5.0"
        
        if 'advanced' in diff_analysis:
            advanced_score = diff_analysis['advanced']['average']
            difficulty_text += f" | <b>Expert skills</b> (Advanced): {advanced_score:.1f}/5.0"
        
        # Add skill progression insight
        scores = [data['average'] for data in diff_analysis.values()]
        if len(scores) > 1:
            if max(scores) - min(scores) < 0.5:
                difficulty_text += "<br/><br/><b>Insight:</b> Consistent performance across difficulty levels suggests well-rounded Excel expertise."
            elif scores[0] > scores[-1] if len(scores) > 1 else False:
                difficulty_text += "<br/><br/><b>Insight:</b> Stronger foundation than advanced skills - excellent base for further development."
            else:
                difficulty_text += "<br/><br/><b>Insight:</b> Progressive skill development demonstrated across complexity levels."
        
        story.append(Paragraph(difficulty_text, styles['CustomBodyText']))
        story.append(Spacer(1, 15))


def format_comprehensive_feedback(story, styles, report: Dict[str, Any]):
    """Enhanced feedback sections with actionable insights"""
    
    # Strengths section
    story.append(Paragraph("Key Strengths & Competencies", styles['SectionHeader']))
    
    strengths = report.get('strengths', [])
    if strengths:
        for i, strength in enumerate(strengths, 1):
            strength_text = f"<b>{i}.</b> {sanitize_feedback(strength)}"
            strength_text = truncate_long_text(strength_text, 350)
            story.append(Paragraph(strength_text, styles['BulletPoint']))
        story.append(Spacer(1, 8))
    else:
        story.append(Paragraph("Areas of strength will be identified as assessment data is analyzed.", styles['CustomBodyText']))
    
    story.append(Spacer(1, 20))
    
    # Areas for improvement
    story.append(Paragraph("Development Opportunities", styles['SectionHeader']))
    
    weaknesses = report.get('weaknesses', [])
    if weaknesses:
        for i, weakness in enumerate(weaknesses, 1):
            weakness_text = f"<b>{i}.</b> {sanitize_feedback(weakness)}"
            weakness_text = truncate_long_text(weakness_text, 350)
            story.append(Paragraph(weakness_text, styles['BulletPoint']))
        story.append(Spacer(1, 8))
    else:
        story.append(Paragraph("Continue building expertise across all Excel competency areas.", styles['CustomBodyText']))
    
    story.append(Spacer(1, 20))
    
    # Strategic learning path
    story.append(Paragraph("Strategic Learning & Development Plan", styles['SectionHeader']))
    
    learning_path = report.get('learning_path', [])
    if learning_path:
        story.append(Paragraph("Based on your assessment results, we recommend the following development sequence:", styles['CustomBodyText']))
        story.append(Spacer(1, 8))
        
        for i, recommendation in enumerate(learning_path, 1):
            rec_text = f"<b>Phase {i}:</b> {recommendation}"
            rec_text = truncate_long_text(rec_text, 350)
            story.append(Paragraph(rec_text, styles['BulletPoint']))
            story.append(Spacer(1, 6))
    
    # Additional professional development recommendations
    story.append(Spacer(1, 15))
    story.append(Paragraph("Professional Development Resources", styles['SubHeader']))
    
    avg_score = report.get('average_score', 0)
    
    if avg_score >= 4.0:
        resources_text = """
        <b>Advanced Certification Track:</b> Consider pursuing Microsoft Excel Expert (Microsoft 365 Apps) 
        certification to formally validate your advanced skills. Explore Power BI integration and 
        advanced analytics to extend your expertise into business intelligence.
        """
    elif avg_score >= 3.0:
        resources_text = """
        <b>Skill Consolidation Track:</b> Focus on building a portfolio of Excel projects that 
        demonstrate your capabilities. Consider intermediate certifications and participate in 
        Excel user communities for continuous learning.
        """
    else:
        resources_text = """
        <b>Foundation Building Track:</b> Establish a structured learning plan with Microsoft Learn's 
        Excel fundamentals courses. Practice daily with real-world datasets and consider 
        mentorship or formal training programs.
        """
    
    story.append(Paragraph(resources_text, styles['CustomBodyText']))


def format_question_analysis(story, styles, qa_list: List[Dict[str, Any]]):
    """Detailed question-by-question analysis"""
    story.append(PageBreak())
    story.append(Paragraph("Detailed Question Analysis", styles['SectionHeader']))
    
    story.append(Paragraph(
        "This section provides specific feedback on each question, including scoring rationale and improvement suggestions.",
        styles['CustomBodyText']
    ))
    story.append(Spacer(1, 15))
    
    # Summary statistics
    if qa_list:
        scores = [qa.get('score', 0) for qa in qa_list]
        avg_score = sum(scores) / len(scores)
        max_score = max(scores)
        min_score = min(scores)
        
        stats_data = [
            ["Statistical Summary", "Value", "Interpretation"],
            ["Questions Answered", str(len(qa_list)), "Total assessment coverage"],
            ["Average Score", f"{avg_score:.2f}/5.0", f"{'Above average' if avg_score >= 3.0 else 'Below average'} performance"],
            ["Highest Score", f"{max_score}/5.0", "Peak performance demonstrated"],
            ["Lowest Score", f"{min_score}/5.0", "Area needing most attention"],
            ["Standard Deviation", f"{statistics.stdev(scores):.2f}" if len(scores) > 1 else "N/A", 
             "Consistency of performance"]
        ]
        
        # Fixed statistics table
        stats_table = Table(stats_data, colWidths=[140, 120, 220])
        stats_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#37474F')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold', 9),
            ('FONT', (0, 1), (-1, -1), 'Helvetica', 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('PADDING', (0, 0), (-1, -1), 6),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8F9FA')]),
            ('WORDWRAP', (0, 0), (-1, -1), True)
        ]))
        
        story.append(stats_table)
        story.append(Spacer(1, 20))
    
    # Individual question analysis
    for i, qa in enumerate(qa_list, 1):
        # Question header
        q_header = f"Question {i}: {qa.get('category', 'General').replace('_', ' ').title()}"
        story.append(Paragraph(q_header, styles['SubHeader']))
        
        # Question details table
        question_data = [
            ["Difficulty Level:", qa.get('level', 'intermediate').title()],
            ["Category:", qa.get('category', 'general').replace('_', ' ').title()],
            ["Score Achieved:", f"{qa.get('score', 0)}/5.0"],
            ["Response Time:", f"{qa.get('answer_time', 0):.0f} seconds" if qa.get('answer_time') else "Not recorded"]
        ]
        
        # Fixed question details table
        q_table = Table(question_data, colWidths=[100, 380])
        q_table.setStyle(TableStyle([
            ('FONT', (0, 0), (0, -1), 'Helvetica-Bold', 9),
            ('FONT', (1, 0), (1, -1), 'Helvetica', 9),
            ('PADDING', (0, 0), (-1, -1), 4),
            ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#F5F5F5')),
            ('GRID', (0, 0), (-1, -1), 0.25, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('WORDWRAP', (0, 0), (-1, -1), True)
        ]))
        
        story.append(q_table)
        story.append(Spacer(1, 8))
        
        # Question text
        story.append(Paragraph("<b>Question:</b>", styles['CustomBodyText']))
        question_text = truncate_long_text(qa.get('question', 'Question text not available'), 300)
        story.append(Paragraph(question_text, styles['CustomBodyText']))
        story.append(Spacer(1, 10))
        
        # Candidate answer
        story.append(Paragraph("<b>Your Response:</b>", styles['CustomBodyText']))
        answer_text = truncate_long_text(qa.get('answer', 'No response provided'), 350)
        story.append(Paragraph(answer_text, styles['CustomBodyText']))
        story.append(Spacer(1, 10))
        
        # AI feedback
        story.append(Paragraph("<b>AI Evaluation & Feedback:</b>", styles['CustomBodyText']))
        feedback = sanitize_feedback(qa.get('feedback', 'Feedback not available'))
        feedback = truncate_long_text(feedback, 400)
        story.append(Paragraph(feedback, styles['CustomBodyText']))
        
        # Performance indicator
        score = qa.get('score', 0)
        if score >= 4:
            performance_note = "🎯 <b>Excellent Response:</b> Demonstrates strong understanding and practical knowledge."
        elif score >= 3:
            performance_note = "✓ <b>Good Response:</b> Shows solid grasp of concepts with room for enhancement."
        elif score >= 2:
            performance_note = "⚠ <b>Developing Response:</b> Basic understanding present, needs deeper exploration."
        else:
            performance_note = "📚 <b>Learning Opportunity:</b> Significant improvement needed in this area."
        
        story.append(Spacer(1, 10))
        story.append(Paragraph(performance_note, styles['HighlightBox']))
        story.append(Spacer(1, 20))
        
        # Add page break after every 3 questions for readability
        if i % 3 == 0 and i < len(qa_list):
            story.append(PageBreak())


def format_assessment_integrity(story, styles, report: Dict[str, Any]):
    """Assessment integrity and proctoring analysis"""
    story.append(PageBreak())
    story.append(Paragraph("Assessment Integrity & Analytics", styles['SectionHeader']))
    
    story.append(Paragraph(
        "This section provides information about assessment conditions and integrity measures to ensure fair evaluation.",
        styles['CustomBodyText']
    ))
    story.append(Spacer(1, 15))
    
    # Proctoring analysis
    proctoring_notes = report.get('proctoring_notes', [])
    if proctoring_notes:
        story.append(Paragraph("Proctoring Analysis", styles['SubHeader']))
        
        for note in proctoring_notes:
            # Clean up the note for professional presentation
            clean_note = note.replace('⚠️', '').replace('✅', '').replace('⏰', '').replace('⚡', '').strip()
            story.append(Paragraph(f"• {clean_note}", styles['BulletPoint']))
        
        story.append(Spacer(1, 15))
    
    # Time analysis
    time_analysis = report.get('time_analysis', {})
    if time_analysis:
        story.append(Paragraph("Response Time Analysis", styles['SubHeader']))
        
        avg_time = time_analysis.get('average_time_per_question', 0)
        total_time = time_analysis.get('total_time', 0)
        efficiency = time_analysis.get('time_efficiency', 'Unknown')
        
        time_data = [
            ["Time Metric", "Value", "Assessment"],
            ["Total Assessment Time", f"{total_time//60:.0f} minutes {total_time%60:.0f} seconds", "Complete duration"],
            ["Average Time per Question", f"{avg_time:.0f} seconds", f"{efficiency} pacing"],
            ["Response Consistency", 
             "Consistent" if 60 <= avg_time <= 240 else "Variable", 
             "Time management evaluation"]
        ]
        
        # Fixed time analysis table
        time_table = Table(time_data, colWidths=[150, 150, 180])
        time_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#455A64')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold', 9),
            ('FONT', (0, 1), (-1, -1), 'Helvetica', 9),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('PADDING', (0, 0), (-1, -1), 6),
            ('WORDWRAP', (0, 0), (-1, -1), True)
        ]))
        
        story.append(time_table)
        story.append(Spacer(1, 20))
    
    # Assessment methodology
    story.append(Paragraph("Assessment Methodology", styles['SubHeader']))
    methodology_text = """
    This Excel proficiency assessment utilizes AI-powered evaluation combined with industry-standard 
    competency frameworks. Questions are dynamically selected based on candidate experience level 
    and target role requirements. Scoring follows a 5-point scale where:
    
    • <b>5.0:</b> Exceptional - Expert-level understanding with practical insights
    • <b>4.0:</b> Excellent - Strong competency with minor gaps
    • <b>3.0:</b> Competent - Adequate for role requirements
    • <b>2.0:</b> Developing - Basic understanding, needs development
    • <b>1.0:</b> Insufficient - Significant learning required
    
    The assessment covers key Excel competencies including data analysis, advanced functions, 
    automation, visualization, and business intelligence tools. Follow-up questions probe deeper 
    understanding and practical application abilities.
    """
    
    story.append(Paragraph(methodology_text, styles['CustomBodyText']))


def format_recommendations_summary(story, styles, report: Dict[str, Any], candidate: Dict[str, Any]):
    """Final recommendations and next steps"""
    story.append(PageBreak())
    story.append(Paragraph("Hiring Recommendation & Next Steps", styles['SectionHeader']))
    
    avg_score = report.get('average_score', 0)
    candidate_name = candidate.get('name', 'Candidate')
    target_role = candidate.get('role', 'Excel Analyst')
    
    # Overall recommendation
    if avg_score >= 4.0:
        recommendation = f"""
        <b>RECOMMENDATION: STRONGLY RECOMMENDED</b>
        
        {candidate_name} demonstrates exceptional Excel proficiency well-suited for the {target_role} position. 
        The candidate shows mastery across multiple competency areas and would likely excel in the role 
        from day one. Consider for advanced responsibilities or mentoring opportunities.
        """
        recommendation_color = colors.HexColor('#1B5E20')
        
    elif avg_score >= 3.5:
        recommendation = f"""
        <b>RECOMMENDATION: RECOMMENDED</b>
        
        {candidate_name} shows strong Excel competency appropriate for the {target_role} position. 
        With solid foundational skills and good performance across assessment areas, the candidate 
        would be a valuable addition to the team with minimal onboarding requirements.
        """
        recommendation_color = colors.HexColor('#2E7D32')
        
    elif avg_score >= 3.0:
        recommendation = f"""
        <b>RECOMMENDATION: CONDITIONAL RECOMMENDATION</b>
        
        {candidate_name} demonstrates adequate Excel skills for the {target_role} position with some development needs. 
        Consider hire with structured onboarding plan focusing on identified improvement areas. 
        Strong potential for growth with proper support.
        """
        recommendation_color = colors.HexColor('#F57C00')
        
    elif avg_score >= 2.0:
        recommendation = f"""
        <b>RECOMMENDATION: ADDITIONAL DEVELOPMENT NEEDED</b>
        
        {candidate_name} shows basic Excel understanding but requires significant skill development 
        before being fully effective in the {target_role} position. Consider for junior role or 
        with extensive training program. Monitor progress closely.
        """
        recommendation_color = colors.HexColor('#FF5722')
        
    else:
        recommendation = f"""
        <b>RECOMMENDATION: NOT RECOMMENDED AT THIS TIME</b>
        
        {candidate_name} needs substantial Excel skill development before being suitable for the {target_role} position. 
        Recommend focused training program and reassessment after 3-6 months of dedicated learning.
        """
        recommendation_color = colors.HexColor('#D32F2F')
    
    # Create recommendation box
    rec_style = ParagraphStyle(
        'RecommendationBox',
        parent=styles['CustomBodyText'],
        fontSize=11,
        leading=16,
        backColor=colors.HexColor('#F8F9FA'),
        borderColor=recommendation_color,
        borderWidth=2,
        borderPadding=15,
        spaceAfter=20
    )
    
    story.append(Paragraph(recommendation, rec_style))
    
    # Next steps
    story.append(Paragraph("Suggested Next Steps", styles['SubHeader']))
    
    if avg_score >= 3.5:
        next_steps = [
            "Schedule final interview focusing on cultural fit and role-specific scenarios",
            "Discuss advanced project opportunities and potential career growth path",
            "Consider for immediate start with standard onboarding process"
        ]
    elif avg_score >= 3.0:
        next_steps = [
            "Conduct practical Excel exercise or case study interview",
            "Develop targeted training plan for identified skill gaps",
            "Set 30-60-90 day performance milestones"
        ]
    else:
        next_steps = [
            "Provide detailed feedback on assessment results",
            "Recommend specific learning resources and timeline",
            "Schedule reassessment after skill development period"
        ]
    
    for i, step in enumerate(next_steps, 1):
        story.append(Paragraph(f"<b>{i}.</b> {step}", styles['BulletPoint']))
    
    story.append(Spacer(1, 20))
    
    # Contact information
    story.append(Paragraph("Assessment Contact Information", styles['SubHeader']))
    contact_text = """
    This assessment was generated by the AI-Powered Excel Mock Interviewer system. 
    For questions about results, methodology, or to schedule reassessment, please contact 
    your hiring team or system administrator.
    
    <i>Report generated on """ + datetime.now().strftime('%B %d, %Y at %I:%M %p') + """</i>
    """
    
    story.append(Paragraph(contact_text, styles['SmallText']))


def build_pdf_report(candidate: Dict[str, Any], qa_list: List[Dict[str, Any]], report: Dict[str, Any]) -> BytesIO:
    """Build comprehensive PDF report with enhanced formatting and analytics"""
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, 
        pagesize=A4, 
        leftMargin=0.75*inch, 
        rightMargin=0.75*inch, 
        topMargin=1*inch, 
        bottomMargin=1*inch
    )
    
    styles = _enhanced_styles()
    story = []
    
    # Title page
    title = f"Excel Proficiency Assessment Report"
    subtitle = f"{candidate.get('name', 'Candidate')} • {datetime.now().strftime('%B %Y')}"
    
    story.append(Paragraph(title, styles['CustomTitle']))
    story.append(Spacer(1, 20))
    story.append(Paragraph(subtitle, styles['SubHeader']))
    story.append(Spacer(1, 30))
    
    # Executive summary
    format_executive_summary(story, styles, candidate, report)
    
    # Candidate profile
    format_candidate_profile(story, styles, candidate)
    
    # Performance analysis
    format_detailed_performance_analysis(story, styles, report)
    
    # Comprehensive feedback
    format_comprehensive_feedback(story, styles, report)
    
    # Question analysis
    format_question_analysis(story, styles, qa_list)
    
    # Assessment integrity
    format_assessment_integrity(story, styles, report)
    
    # Final recommendations
    format_recommendations_summary(story, styles, report, candidate)
    
    # Build the PDF
    doc.build(story)
    buffer.seek(0)
    return buffer