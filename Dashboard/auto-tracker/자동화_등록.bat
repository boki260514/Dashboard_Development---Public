@echo off
chcp 65001 > nul
REM ============================================================
REM  매일 오전 8시~10시 30분 간격 작업일보 자동 처리 등록
REM ============================================================
REM  이 파일을 한 번만 더블클릭하면 매일 오전 8:00, 8:30, 9:00,
REM  9:30, 10:00 (총 5회) 자동으로 메일 → 엑셀 → 대시보드 갱신이
REM  실행됩니다. 실패 시 본인 메일로 알림이 전송됩니다.
REM ============================================================

setlocal

cd /d "%~dp0"

set TASK_NAME=HCNS_메일일보처리_매일오전
set TASK_TARGET=%~dp0메일일보처리.bat

echo.
echo  ========================================================
echo    매일 오전 8~10시 30분 간격 자동 실행 등록
echo  ========================================================
echo.
echo   작업 이름: %TASK_NAME%
echo   실행 파일: %TASK_TARGET%
echo   실행 시각: 매일 08:00 시작, 30분마다, 2시간 동안
echo              (08:00, 08:30, 09:00, 09:30, 10:00 — 총 5회)
echo.

REM 기존 동일 이름 작업이 있으면 삭제 (덮어쓰기)
schtasks /query /tn "%TASK_NAME%" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo  [기존 등록 감지됨 - 새로 덮어씁니다]
    schtasks /delete /tn "%TASK_NAME%" /f >nul 2>&1
)

REM 이전 이름(09시 단일)도 있으면 청소
schtasks /query /tn "HCNS_메일일보처리_매일오전9시" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo  [이전 이름 작업 청소]
    schtasks /delete /tn "HCNS_메일일보처리_매일오전9시" /f >nul 2>&1
)

REM 등록: 매일 08:00 시작, 30분 간격, 2시간 동안 반복
schtasks /create /tn "%TASK_NAME%" /tr "\"%TASK_TARGET%\"" /sc daily /st 08:00 /ri 30 /du 02:00 /f

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  [ FAIL ] 등록 실패. 위 메시지를 확인하세요.
    echo.
    pause
    exit /b 1
)

echo.
echo  ========================================================
echo    [OK] 등록 완료!
echo  ========================================================
echo.
echo  - 매일 오전 8시부터 10시까지 30분 간격으로 자동 실행됩니다.
echo  - 실패 시 yoon01@hcnsservice.co.kr 로 알림 메일 발송.
echo  - 작업 스케줄러에서 확인하려면:
echo      Win+R → taskschd.msc → 작업 스케줄러 라이브러리
echo  - 해제하려면 "자동화_해제.bat" 더블클릭
echo.
pause
exit /b 0
