@echo off
REM 엑셀 감시 프로세스 종료 (watch_excel.py를 실행 중인 python 프로세스만 정확히 종료)

REM watch_excel.py가 명령줄에 포함된 python.exe 프로세스를 찾아 종료
for /f "tokens=2" %%a in ('wmic process where "name='python.exe' and commandline like '%%watch_excel.py%%'" get processid /value ^| find "ProcessId="') do (
    taskkill /pid %%a /f
)

echo.
echo 대시보드 감시 프로세스가 종료되었습니다.
pause
