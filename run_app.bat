@echo off
setlocal

call "%~dp0build_exe.bat" --no-run
if errorlevel 1 (
  echo Build failed in build_exe.bat
  exit /b 1
)

if exist "%~dp0dist\record_archive.exe" (
  start "" "%~dp0dist\record_archive.exe"
) else (
  echo dist\record_archive.exe not found even after build
  exit /b 1
)
