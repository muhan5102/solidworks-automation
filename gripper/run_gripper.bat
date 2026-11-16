@echo off
chcp 65001 >nul
echo ============================================
echo  UMI Fin-Ray Gripper - SolidWorks 2024 build
echo ============================================
"D:\Python\python.exe" "%~dp0build_umi_fin_gripper.py"
echo.
echo Build script finished. You can close this window.
pause
