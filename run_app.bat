@echo off
setlocal
cd /d "%~dp0"
python -m pip install -r requirements.txt
if errorlevel 1 goto fail
python app.py
if errorlevel 1 goto fail
exit /b 0

:fail
echo.
echo Program failed. Please check the message above.
pause
exit /b 1
