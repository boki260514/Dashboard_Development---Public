@echo off
chcp 65001 > nul
cd /d "C:\auto-tracker"
echo.
echo ===========================================
echo  대시보드 자동 갱신
echo ===========================================
echo.
python make_dashboard.py
echo.
echo ===========================================
echo  완료. 아무 키나 누르면 닫힙니다.
echo ===========================================
pause > nul
