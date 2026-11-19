@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_rp.ps1" %*
pause
