@echo off
REM Excel/CSV Automation Tool - Direct Launcher
REM This batch file launches the Python application directly

cd /d "%~dp0"
start python.exe app.py %*
