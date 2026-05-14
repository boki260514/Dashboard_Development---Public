@echo off
chcp 65001 > nul
REM ============================================================
REM  매일 오전 8~10시 30분 간격 일반 업무 메일 로그 자동 처리 등록
REM ============================================================
REM  이 파일을 한 번만 더블클릭하면 매일 08:05, 08:35, 09:05,
REM  09:35, 10:05 (총 5회) 자동으로 INBOX의 업무 메일이
REM  Mail_Log.xlsx + mail_dashboard.html 에 반영됩니다.
REM  (작업일보 자동화 5분 뒤에 어긋나게 실행 — 충돌 회피)
REM ============================================================

setlocal

cd /d "%~dp0"

set TASK_NAME=HCNS_메일로그처리_매일오전
set TASK_TARGET=%~dp0메일로그처리.bat

echo.
echo  ========================================================
echo    매일 오전 8~10시 30분 간격 메일 로그 자동 실행 등록
echo  ========================================================
echo.
echo   작업 이름: %TASK_NAME%
echo   실행 파일: %TASK_TARGET%
echo   실행 시각: 매일 08:05 시작, 30분마다, 2시간 동안
echo              (08:05, 08:35, 09:05, 09:35, 10:05 — 총 5회)
echo.

REM 기존 동일 이름 작업 삭제
schtasks /query /tn "%TASK_NAME%" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo  [기존 등록 감지됨 - 새로 덮어씁니다]
    schtasks /delete /tn "%TASK_NAME%" /f >nul 2>&1
)

REM 이전 이름(09:05 단일)도 청소
schtasks /query /tn "HCNS_메일로그처리_매일오전9시" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo  [이전 이름 작업 청소]
    schtasks /delete /tn "HCNS_메일로그처리_매일오전9시" /f >nul 2>&1
)

schtasks /create /tn "%TASK_NAME%" /tr "\"%TASK_TARGET%\"" /sc daily /st 08:05 /ri 30 /du 02:00 /f

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
echo  - 매일 오전 8시 5분부터 10시 5분까지 30분 간격으로 실행됩니다.
echo  - 결과는 Mail_Log.xlsx 와 mail_dashboard.html 에 기록됩니다.
echo  - 작업 스케줄러에서 확인하려면:
echo      Win+R → taskschd.msc → 작업 스케줄러 라이브러리
echo  - 해제하려면 "자동화_해제_메일로그.bat" 더블클릭
echo.
pause
exit /b 0
