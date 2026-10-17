@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Python Launcher not found. Install Python 3.11 or 3.12.
  pause
  exit /b 1
)
py -m pip install -r requirements.txt
if errorlevel 1 goto fail
py -m unittest discover -s tests -v
if errorlevel 1 goto fail
rem Always build the portable folder; dependencies stay outside the EXE.
py -m PyInstaller --noconfirm --clean --onedir --windowed --name WeComActivityNotifier --collect-all openpyxl --collect-all PIL --collect-all certifi --hidden-import xlrd main.py
if errorlevel 1 goto fail
py -c "import hashlib,pathlib; p=pathlib.Path('dist/WeComActivityNotifier/WeComActivityNotifier.exe'); h=hashlib.file_digest(p.open('rb'),'sha256').hexdigest(); pathlib.Path('dist/WeComActivityNotifier/SHA256.txt').write_text(h+'  WeComActivityNotifier.exe\n'); print('SHA256:',h)"
if errorlevel 1 goto fail
echo.
echo [OK] dist\WeComActivityNotifier\WeComActivityNotifier.exe
echo [INFO] Keep the entire dist\WeComActivityNotifier folder; it is portable.
pause
exit /b 0
:fail
echo [ERROR] Build failed.
pause
exit /b 1
