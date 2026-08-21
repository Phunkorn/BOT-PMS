@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0BUILD_EXE.ps1" %*
set "BUILD_EXIT=%ERRORLEVEL%"
if not "%BUILD_EXIT%"=="0" (
    echo.
    echo Build failed with exit code %BUILD_EXIT%.
)
exit /b %BUILD_EXIT%
