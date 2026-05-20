@echo off
title RMR Activation - STEP 4: Sync to SharePoint
color 0D
cd /d "C:\Users\arenahe\OneDrive - Securitas\RMR_ACTIVATION"
echo.
echo  ============================================================
echo    RMR ACTIVATION — STEP 4: Sync Excel to SharePoint
echo  ============================================================
echo.
python STEP4.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  ERROR: Script failed with code %ERRORLEVEL%
    pause
)