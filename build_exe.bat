@echo off
setlocal

python -m pip install -r requirements.txt
if errorlevel 1 goto :error

pyinstaller --noconfirm --onefile --windowed --name record_archive app.py
if errorlevel 1 goto :error

echo Build complete. EXE is in dist\record_archive.exe

echo Running app...
start "" "%~dp0dist\record_archive.exe"

goto :eof

:error
echo Build failed. Please check the error messages above.
exit /b 1
