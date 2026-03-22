@echo off
setlocal

if exist "%~dp0dist\record_archive.exe" (
  start "" "%~dp0dist\record_archive.exe"
) else (
  echo dist\record_archive.exe not found. Build first with build_exe.bat
  exit /b 1
)
