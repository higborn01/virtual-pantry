@echo off
cd /d "%~dp0"
rem Use Anaconda's Python if it's there; otherwise whatever "python" is on the PATH.
set "PY=%USERPROFILE%\anaconda3\python.exe"
if not exist "%PY%" set "PY=python"
"%PY%" pantry_app.py
if errorlevel 1 pause
