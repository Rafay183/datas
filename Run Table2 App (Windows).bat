@echo off
cd /d "%~dp0"
if not exist .venv (
  echo First run: setting up Python environment...
  python -m venv .venv
  call .venv\Scripts\activate
  pip install -r requirements.txt
) else (
  call .venv\Scripts\activate
)
streamlit run app.py
pause
