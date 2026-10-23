@echo off
chcp 65001 >nul
echo ============================================
echo  Geoseong Hotel Night Closing Tool - EXE build
echo ============================================
where py >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Python is not installed. Install from https://www.python.org/downloads/
  echo         and check "Add python.exe to PATH" during setup.
  pause
  exit /b 1
)
py -m pip install --upgrade pip
py -m pip install -r requirements.txt pyinstaller
py -m PyInstaller --noconfirm --onefile --windowed --name GeoseongClosingTool --add-data "config.json;." app.py
if errorlevel 1 (
  echo [ERROR] Build failed.
  pause
  exit /b 1
)
copy /Y config.json dist\config.json >nul
if exist templates xcopy /E /I /Y templates dist\templates >nul
copy /Y docs\사용법.txt dist\ >nul
echo.
echo Build complete: dist\GeoseongClosingTool.exe
echo Copy BOTH files in dist\ (exe + config.json) to the hotel PC.
pause
