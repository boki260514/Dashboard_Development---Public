@echo off
chcp 65001 > nul
set TASK_NAME=HCNS_메일일보처리_매일오전9시

echo.
echo  ========================================================
echo    현재 등록 상태
echo  ========================================================
echo.

schtasks /query /tn "%TASK_NAME%" /v /fo list 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo  [INFO] 등록된 자동화가 없습니다.
    echo         "자동화_등록.bat"으로 등록할 수 있어요.
)

echo.
pause
