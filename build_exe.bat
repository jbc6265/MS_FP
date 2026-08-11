@echo off
setlocal
cd /d "%~dp0"

echo [1/3] Installing dependencies...
python -m pip install -r requirements.txt
if errorlevel 1 goto fail

echo [2/3] Building executable...
python -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --onefile ^
  --name "Myeongseong_Sequence_Material_Merge" ^
  app.py
if errorlevel 1 goto fail

echo [3/3] Done.
echo.
echo EXE location:
echo %cd%\dist\Myeongseong_Sequence_Material_Merge.exe
pause
exit /b 0

:fail
echo.
echo Build failed. Please check the message above.
pause
exit /b 1
