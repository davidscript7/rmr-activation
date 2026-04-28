@echo off
title RMR Activation - STEP 1: MONITORING
color 0A
cd /d "C:\Users\arenahe\OneDrive - Securitas\RMR_ACTIVATION"
echo.
echo  ============================================================
echo    RMR ACTIVATION — STEP 1: MONITORING
echo  ============================================================
echo.
python STEP1.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  ERROR: El script termino con errores. Codigo: %ERRORLEVEL%
    pause
)
