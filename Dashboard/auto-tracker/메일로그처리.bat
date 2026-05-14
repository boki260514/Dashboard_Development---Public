@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo.
echo  ===============================================
echo    [ Mail Log Auto Process ]
echo  ===============================================
echo.

python mail_log_fetcher.py
set RC=%ERRORLEVEL%

if %RC% NEQ 0 (
    echo.
    echo  [ FAIL ] mail_log_fetcher exit code = %RC%
    exit /b %RC%
)

echo.
echo  ===============================================
echo    [ ALL DONE ]
echo  ===============================================
echo.
exit /b 0
