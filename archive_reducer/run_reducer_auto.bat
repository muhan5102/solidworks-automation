@echo off
chcp 65001 >nul
cd /d "%~dp0"
set RC=1
echo [%time%] start > "%~dp0run_reducer_status.txt"
echo ============================================
echo  一级减速器 - SolidWorks Auto Build
echo  (scheduled task, no keypress required)
echo ============================================
"D:\Python\python.exe" "%~dp0build_reducer.py" > "%~dp0run_reducer_console.log" 2>&1
set RC=%errorlevel%
echo [%time%] exit=%RC% >> "%~dp0run_reducer_status.txt"
echo Build finished. exit code = %RC%. See reducer_build_log.txt
timeout /t 5 /nobreak >nul
exit /b %RC%
