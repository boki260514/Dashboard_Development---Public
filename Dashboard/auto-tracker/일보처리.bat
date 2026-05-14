@echo off
cd /d "C:\auto-tracker"

cls
echo.
echo  ===============================================
echo    [ Daily Report Auto Process ]
echo  ===============================================
echo.

echo  [ 1 / 2 ] Update Excel Tracking Sheet
echo  -----------------------------------------------
echo.
python update_tracker.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  [ FAIL ] Tracker update problem
    pause
    exit /b 1
)

echo.
echo  [ 2 / 2 ] Refresh Dashboard HTML
echo  -----------------------------------------------
echo.
python make_dashboard.py --quiet
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  [ FAIL ] Dashboard problem
    pause
    exit /b 1
)
echo  [ OK ] Dashboard refreshed

echo.
echo  ===============================================
echo    [ ALL DONE ]
echo  ===============================================
echo.
echo    Open dashboard.html and press F5 to refresh
echo.
pause