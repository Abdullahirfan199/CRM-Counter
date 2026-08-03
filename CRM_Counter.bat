@echo off
title CRM Counter
mode con: cols=70 lines=35
reg add "HKCU\Console" /v ScreenBufferSize /t REG_DWORD /d 0xBB800046 /f >nul
cd /d "%~dp0"
python crm_counter.py
if errorlevel 1 (
    echo.
    echo CRM Counter exited with an error. Check logs\crm_counter.log
    pause
)
