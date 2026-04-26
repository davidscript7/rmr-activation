@echo off
title RMR Activation - STEP 1
color 0A
cd /d "%~dp0"
echo.
echo  ============================================================
echo    RMR ACTIVATION - STEP 1: MONITORING
echo  ============================================================
echo.
python STEP1.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  ERROR: Script finished with errors. Code: %ERRORLEVEL%
    pause
)
