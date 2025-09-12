# Excel Mock Interviewer (Gemini + FastAPI + Streamlit)

An AI-powered mock interviewer focused on Microsoft Excel skills. It asks 5–7 questions (basic → advanced), grades each answer using Gemini, and produces a final strengths/weaknesses report with a suggested learning path.

## Tech Stack
- Backend: FastAPI
- Frontend: Streamlit
- LLM: Gemini (google-generativeai)
- State/Flow: Streamlit session state
- Deployment: Render (recommended) or Vercel

## Project Structure
```
api/server.py        # FastAPI backend with grading endpoint
app.py               # Streamlit UI
report_generator.py  # Modular PDF report builder
sample_report.py     # Script to generate a sample polished PDF
requirements.txt     # Python dependencies
render.yaml          # Render deployment config
vercel.json          # Vercel deployment config (alternative)
```

## Setup
1. Python 3.10+
2. Create and activate a virtual environment
```bash
python -m venv .venv
. .venv/Scripts/activate  # Windows PowerShell: .venv\Scripts\Activate.ps1
```
3. Install dependencies
```bash
pip install -r requirements.txt
```
4. Set environment variable for Gemini
```bash
$env:GEMINI_API_KEY="YOUR_KEY_HERE"   # PowerShell
```

## Local Run
Open two terminals (or run API first, then Streamlit).

- Start API:
```bash
uvicorn api.server:app --reload --port 8000
```
- Start UI:
```bash
streamlit run app.py
```
The UI expects the API at http://localhost:8000. To use a different base, set `API_BASE` before starting Streamlit:
```bash
$env:API_BASE="https://your-api.onrender.com"
streamlit run app.py
```

## Generate a Sample Polished PDF
```bash
python sample_report.py
# Outputs: sample_report.pdf
```

## Deployment (Render)
1. Push this repo to GitHub.
2. In Render, create a new Web Service for the API using this repo:
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `uvicorn api.server:app --host 0.0.0.0 --port $PORT`
   - Add environment variable `GEMINI_API_KEY` with your key.
3. Create another Web Service for the UI:
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`
   - Add environment variable `API_BASE` to the API URL.
4. Visit the UI URL to use the app.

## Deployment (Vercel)
Vercel is not ideal for Streamlit. The provided `vercel.json` can route to Python entries, but running Streamlit on Vercel is unsupported in many setups. Prefer Render for simplicity.

## Clone and Use from GitHub
```bash
# Clone
git clone https://github.com/HEMANT2027/Coding_Ninjas.git
cd Coding_Ninjas

# Setup
python -m venv .venv
. .venv/Scripts/activate   # Windows PowerShell
pip install -r requirements.txt

# Configure environment
$env:GEMINI_API_KEY="YOUR_KEY_HERE"

# Run API (terminal 1)
uvicorn api.server:app --reload --port 8000

# Run UI (terminal 2)
. .venv/Scripts/activate
$env:API_BASE="http://localhost:8000"
streamlit run app.py
```

## Notes
- The grading endpoint returns JSON like `{ "score": 3, "feedback": "Good use of VLOOKUP but missing explanation of range_lookup" }`.
- If no `GEMINI_API_KEY` is set, backend will return a fallback message and score 0.
- You can customize questions in `api/server.py` (`QUESTION_BANK`).
