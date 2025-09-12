from report_generator import build_pdf_report

mock_candidate = {
    "name": "Alex Johnson",
    "email": "alex.j@example.com",
    "role": "Senior Excel Analyst",
    "experience": "5"
}

mock_qa = [
    {"question": "Explain VLOOKUP and give an example.", "answer": "...", "score": 4, "feedback": "Clear explanation with correct example."},
    {"question": "Difference between VLOOKUP and INDEX-MATCH?", "answer": "...", "score": 3, "feedback": "Good distinction; could mention left lookup advantage."},
    {"question": "Create a PivotTable to summarize monthly sales by product.", "answer": "...", "score": 2, "feedback": "Understands basics but missed slicers and filtering."},
    {"question": "What are Excel Macros and when to use them?", "answer": "...", "score": 5, "feedback": "Excellent use cases and clarity about recording vs. editing."},
]

mock_report = {
    "average_score": 3.5,
    "strengths": ["Good understanding of lookups", "Clear macro use cases"],
    "weaknesses": ["PivotTable filtering and slicer usage need practice"],
    "learning_path": [
        "Practice INDEX/MATCH and XLOOKUP with approximate matches.",
        "Build PivotTables with top-N filters and slicers.",
        "Automate tasks with recorded macros; review generated VBA.",
    ]
}

if __name__ == "__main__":
    buf = build_pdf_report(mock_candidate, mock_qa, mock_report)
    with open("sample_report.pdf", "wb") as f:
        f.write(buf.getvalue())
    print("Wrote sample_report.pdf")
