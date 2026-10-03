@echo off
chcp 65001 >nul
echo ==================================================
echo  Python env check: D:\Python (pywin32/numpy/Pillow/yaml)
echo ==================================================
"D:\Python\python.exe" -m pip install pyyaml setuptools wheel
echo.
"D:\Python\python.exe" -c "import sys, numpy, win32com.client, PIL, yaml; print('PY', sys.version.split()[0]); print('numpy', numpy.__version__); print('pywin32 OK'); print('Pillow', PIL.__version__); print('yaml OK')"
echo.
echo If no red ERROR lines above, the environment is ready for run_gripper.bat.
pause
