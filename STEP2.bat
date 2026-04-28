@echo off
title RMR Activation - STEP 2: IW75
color 0B
cd /d "C:\Users\arenahe\OneDrive - Securitas\RMR_ACTIVATION"
echo.
echo  ============================================================
echo    RMR ACTIVATION — STEP 2: IW75
echo  ============================================================
echo.
python STEP2.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  ERROR: El script termino con errores. Codigo: %ERRORLEVEL%
    pause
)
