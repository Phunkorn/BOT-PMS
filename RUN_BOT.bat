@echo off
title TVC Desktop Bot v0.5
if not exist .venv\Scripts\python.exe (
  echo ERROR: .venv not found
  echo Copy .venv from your old bot folder into this folder first.
  pause
  exit /b 1
)
.venv\Scripts\python.exe src\bot.py
pause
