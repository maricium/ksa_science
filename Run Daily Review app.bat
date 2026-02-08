@echo off
cd /d "%~dp0"
echo Starting Daily Review app...
echo Share the "Network URL" (e.g. http://192.168.x.x:8501) with colleagues.
echo.
uv run streamlit run daily_review_app.py --server.address 0.0.0.0
pause
