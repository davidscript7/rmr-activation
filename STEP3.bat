@echo off
title RMR Activation - STEP 3: CONTRACT DETAIL
color 0E
cd /d "%~dp0"
echo.
echo  ============================================================
echo    RMR ACTIVATION — STEP 3: CONTRACT DETAIL
echo  ============================================================
echo.
python STEP3.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  ERROR: Script finished with errors. Code: %ERRORLEVEL%
    pause
)
