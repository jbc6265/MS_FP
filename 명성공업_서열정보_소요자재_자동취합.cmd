@echo off
setlocal
cd /d "%~dp0"

title 명성공업 서열정보^&소요자재 자동 취합 프로그램

echo ================================================
echo  명성공업 서열정보^&소요자재 자동 취합 프로그램
echo ================================================
echo.
echo [1/2] 실행에 필요한 Python 패키지를 확인합니다.
python -m pip install -r requirements.txt
if errorlevel 1 goto fail

echo.
echo [2/2] 프로그램을 실행합니다.
python app.py
if errorlevel 1 goto fail

exit /b 0

:fail
echo.
echo 프로그램 실행 중 오류가 발생했습니다.
echo 위 메시지를 확인해 주세요.
pause
exit /b 1
