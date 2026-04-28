@echo off
title RMR Activation - STEP 3: CONTRACT DETAIL
color 0E
cd /d "C:\Users\arenahe\OneDrive - Securitas\RMR_ACTIVATION"
echo.
echo  ============================================================
echo    RMR ACTIVATION — STEP 3: CONTRACT DETAIL
echo  ============================================================
echo.
python STEP3.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  ERROR: El script termino con errores. Codigo: %ERRORLEVEL%
    pause
)
