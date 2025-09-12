import os
from typing import List, Optional, Dict, Any
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from dotenv import load_dotenv
import google.generativeai as genai

load_dotenv()

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
if not GEMINI_API_KEY:
    # Allow Render/Vercel to pass at runtime; avoid crash at import time locally
    GEMINI_API_KEY = ""

genai.configure(api_key=GEMINI_API_KEY) if GEMINI_API_KEY else None

app = FastAPI(title="Excel Mock Interviewer API")


class Question(BaseModel):
    id: int
    level: str
    text: str


class GradeRequest(BaseModel):
    question: str
    answer: str


class GradeResponse(BaseModel):
    score: int
    feedback: str


QUESTION_BANK: List[Question] = [
    Question(id=1, level="basic", text="Can you explain how VLOOKUP works and give a simple example?"),
    Question(id=2, level="basic", text="How can you use Conditional Formatting to highlight duplicate values?"),
    Question(id=3, level="intermediate", text="What is the difference between VLOOKUP and INDEX-MATCH?"),
    Question(id=4, level="intermediate", text="How would you create a PivotTable to summarize monthly sales by product?"),
    Question(id=5, level="advanced", text="What are Excel Macros, and when would you use them?"),
    Question(id=6, level="advanced", text="Describe how you would use Power Query to clean and merge data from multiple sources."),
    Question(id=7, level="advanced", text="Explain how to use XLOOKUP and its advantages over VLOOKUP.")
]


SYSTEM_GRADER_PROMPT = (
    "You are an expert Excel interviewer. Grade the candidate's answer on a 0-5 scale. "
    "Be strict but fair. Consider correctness, clarity, and practical understanding. "
    "Return ONLY a compact JSON object with fields 'score' (integer 0-5) and 'feedback' (short sentence). "
    "No markdown, no extra text."
)


def grade_with_gemini(question: str, answer: str) -> Dict[str, Any]:
    if not GEMINI_API_KEY:
        # Fallback for local dev without key
        return {"score": 0, "feedback": "No API key provided; cannot grade. Provide GEMINI_API_KEY."}

    try:
        model = genai.GenerativeModel("gemini-1.5-flash")
        prompt = (
            f"{SYSTEM_GRADER_PROMPT}\n\n"
            f"Question: {question}\n"
            f"Answer: {answer}\n"
            f"Respond with JSON only."
        )
        resp = model.generate_content(prompt)
        text = resp.text.strip()
        # Best effort JSON extraction
        import json, re
        match = re.search(r"\{[\s\S]*\}", text)
        if match:
            parsed = json.loads(match.group(0))
            score = int(parsed.get("score", 0))
            feedback = str(parsed.get("feedback", ""))
            score = max(0, min(5, score))
            return {"score": score, "feedback": feedback}
        return {"score": 0, "feedback": "Grader returned unexpected format."}
    except Exception as e:
        return {"score": 0, "feedback": f"Error grading: {e}"}


@app.get("/health")
async def health() -> Dict[str, str]:
    return {"status": "ok"}


@app.get("/questions", response_model=List[Question])
async def get_questions(limit: int = 6) -> List[Question]:
    # Return 5-7 questions in difficulty order
    ordered = sorted(QUESTION_BANK, key=lambda q: ["basic", "intermediate", "advanced"].index(q.level))
    return ordered[: max(5, min(7, limit))]


@app.post("/grade", response_model=GradeResponse)
async def grade(req: GradeRequest) -> GradeResponse:
    result = grade_with_gemini(req.question, req.answer)
    try:
        score = int(result.get("score", 0))
    except Exception:
        score = 0
    feedback = str(result.get("feedback", ""))
    return GradeResponse(score=max(0, min(5, score)), feedback=feedback)
