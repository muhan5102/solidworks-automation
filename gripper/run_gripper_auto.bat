@echo off
chcp 65001 >nul
cd /d "%~dp0"
set RC=1
echo [%time%] start > "%~dp0run_status.txt"
echo ============================================
echo  UMI Fin-Ray Gripper - Auto Build
echo  (scheduled task, no keypress required)
echo ============================================
"D:\Python\python.exe" "%~dp0build_umi_fin_gripper.py" > "%~dp0run_console.log" 2>&1
set RC=%errorlevel%
echo [%time%] exit=%RC% >> "%~dp0run_status.txt"
echo Build finished. exit code = %RC%. See build_log.txt / run_console.log
timeout /t 5 /nobreak >nul
exit /b %RC%
