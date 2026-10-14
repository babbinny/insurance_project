@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo 브라우저에서 http://127.0.0.1:8000 을 여세요. (끄려면 Ctrl+C)
python -m uvicorn app:app --port 8000
pause
