@echo off
setlocal

python -m pip install -r requirements.txt
pyinstaller --noconfirm --onefile --windowed --name record_archive app.py

echo Build complete. EXE is in dist\record_archive.exe
endlocal
