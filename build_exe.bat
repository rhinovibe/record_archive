@echo off
setlocal

set "NO_RUN=0"
if /I "%~1"=="--no-run" set "NO_RUN=1"

python -m pip install -r requirements.txt
if errorlevel 1 goto :error

pyinstaller --noconfirm --onefile --windowed --name record_archive --hidden-import docx --collect-all docx app.py
if errorlevel 1 goto :error

echo Build complete. EXE is in dist\record_archive.exe

if "%NO_RUN%"=="1" goto :eof

echo Running app...
start "" "%~dp0dist\record_archive.exe"

goto :eof

:error
echo Build failed. Please check the error messages above.
exit /b 1
