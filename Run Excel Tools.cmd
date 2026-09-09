@echo off
setlocal
cd /d "%~dp0"
start "" pyw.exe -3 "%~dp0app.py"
