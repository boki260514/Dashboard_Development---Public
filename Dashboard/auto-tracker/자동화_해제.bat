@echo off
chcp 65001 > nul
REM ============================================================
REM  작업일보 자동 처리 스케줄 해제 (현재 이름 + 옛 이름 모두)
REM ============================================================

set TASK_NEW=HCNS_메일일보처리_매일오전
set TASK_OLD=HCNS_메일일보처리_매일오전9시

echo.
echo  ========================================================
echo    자동 실행 해제
echo  ========================================================
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
    echo  [INFO] 등록된 작업이 없어요. (이미 해제되어 있음)
) else (
    echo  [OK] 해제 완료. 이제 자동 실행되지 않습니다.
    echo       다시 자동화하려면 "자동화_등록.bat" 더블클릭.
)
echo.
pause
exit /b 0
