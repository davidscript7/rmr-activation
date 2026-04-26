@echo off
title RMR Activation - STEP 2: IW75
color 0B
cd /d "%~dp0"
echo.
echo  ============================================================
echo    RMR ACTIVATION — STEP 2: IW75
echo  ============================================================
echo.
python STEP2.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  ERROR: Script finished with errors. Code: %ERRORLEVEL%
    pause
)
