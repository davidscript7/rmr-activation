@echo off
title RMR - Install Dependencies
color 0F
echo.
echo  ============================================================
echo    RMR ACTIVATION - INSTALL DEPENDENCIES
echo    Power Query + SharePoint Lists version
echo  ============================================================
echo.
echo  Installing required packages...
echo.
pip install pandas --quiet
echo    [OK] pandas
pip install openpyxl --quiet
echo    [OK] openpyxl
pip install pyperclip --quiet
echo    [OK] pyperclip
pip install pywin32 --quiet
echo    [OK] pywin32
pip install xlwings --quiet
echo    [OK] xlwings
pip install Office365-REST-Python-Client --quiet
echo    [OK] Office365-REST-Python-Client
echo.
echo  ============================================================
echo    Installation complete!
echo.
echo    Installed packages:
echo      • pandas         - Data processing
echo      • openpyxl       - Excel reading
echo      • pyperclip      - Clipboard integration
echo      • pywin32        - Power Query auto-refresh
echo      • xlwings        - Excel COM automation
echo      • Office365-REST - SharePoint API (STEP4)
echo.
echo    You can now use STEP1.bat through STEP4.bat
echo  ============================================================
echo.
pause