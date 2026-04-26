@echo off
title RMR - Install Dependencies
color 0F
echo.
echo  ============================================================
echo    RMR ACTIVATION - INSTALL DEPENDENCIES
echo    Run this only ONCE before using the scripts
echo  ============================================================
echo.
echo  Installing required packages...
echo.
pip install pandas --quiet
echo    [OK] pandas
pip install openpyxl --quiet
echo    [OK] openpyxl
pip install msal --quiet
echo    [OK] msal
pip install requests --quiet
echo    [OK] requests
pip install pyperclip --quiet
echo    [OK] pyperclip
echo.
echo  ============================================================
echo    Installation complete!
echo    You can now use STEP1.bat, STEP2.bat and STEP3.bat
echo  ============================================================
echo.
pause
