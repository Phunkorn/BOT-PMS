@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\pythonw.exe" goto gui_runtime_error
for %%I in (".venv\Scripts\pythonw.exe") do if %%~zI LEQ 0 goto gui_runtime_error
start "" ".venv\Scripts\pythonw.exe" "src\tvc_control.py"
exit /b 0

:gui_runtime_error
echo ERROR: Python environment for GUI is missing or damaged.
echo Please install or repair .venv before starting T.V.C JOB BOT.
pause
exit /b 1
