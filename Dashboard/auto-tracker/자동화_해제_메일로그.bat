@echo off
chcp 65001 > nul

set TASK_NEW=HCNS_메일로그처리_매일오전
set TASK_OLD=HCNS_메일로그처리_매일오전9시

echo.
echo  메일 로그 자동 실행 해제
echo.

set FOUND=0

schtasks /query /tn "%TASK_NEW%" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    schtasks /delete /tn "%TASK_NEW%" /f
    set FOUND=1
)

schtasks /query /tn "%TASK_OLD%" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    schtasks /delete /tn "%TASK_OLD%" /f
    set FOUND=1
)

if "%FOUND%"=="0" (
    echo  [INFO] 등록된 작업이 없습니다.
) else (
    echo  [OK] 해제 완료.
)

pause
exit /b 0
