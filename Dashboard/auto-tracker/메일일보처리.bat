@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo.
echo  ===============================================
echo    [ Daily Report (Email) Auto Process ]
echo  ===============================================
echo.

python email_fetcher.py
set RC=%ERRORLEVEL%

if %RC% NEQ 0 (
    echo.
    echo  [ FAIL ] email_fetcher exit code = %RC%
    exit /b %RC%
)

echo.
echo  ===============================================
echo    [ ALL DONE ]
echo  ===============================================
echo.
exit /b 0
