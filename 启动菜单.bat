@echo off
setlocal
cd /d "%~dp0"

:menu
cls
echo ChromeAutoRun Project Menu
echo ==========================
echo 1. Start development version
echo 2. Build Windows EXE
echo 3. Build Linux DEB through WSL Ubuntu
echo 0. Exit
echo.
choice /c 1230 /n /m "Select an option: "
if errorlevel 4 goto end
if errorlevel 3 goto build_linux
if errorlevel 2 goto build_windows
if errorlevel 1 goto start_dev

:start_dev
set "DEV_PYTHON=%USERPROFILE%\miniconda3\envs\HZPBC\python.exe"
if not exist "%DEV_PYTHON%" set "DEV_PYTHON=python"
call "%DEV_PYTHON%" -c "import selenium, openpyxl" >nul 2>nul
if errorlevel 1 (
  echo.
  echo ERROR: Python with Selenium and openpyxl was not found.
  echo Install the HZPBC environment or install these packages for Python on PATH.
  goto wait
)
echo Starting development version with %DEV_PYTHON% ...
call "%DEV_PYTHON%" main.py
goto wait

:build_windows
call build_quick.bat
if errorlevel 1 (
  echo.
  echo ERROR: Windows EXE build did not complete. Review the build output above.
)
goto wait

:build_linux
set "WSL_DISTRO=Ubuntu"
set "WSL_PROJECT_DIR="
set "WSL_WINDOWS_PATH=%~dp0"
set "WSL_WINDOWS_PATH=%WSL_WINDOWS_PATH:\=/%"
for /f "delims=" %%I in ('wsl.exe -d %WSL_DISTRO% -- wslpath -a "%WSL_WINDOWS_PATH%" 2^>nul') do set "WSL_PROJECT_DIR=%%I"
if not defined WSL_PROJECT_DIR (
  echo.
  echo ERROR: Could not map this project path into WSL distribution "%WSL_DISTRO%".
  echo Check that Ubuntu is installed and this project is under a WSL-accessible drive.
  goto wait
)
echo Building Linux DEB in WSL. Output directory:
echo   %WSL_PROJECT_DIR%/dist/linux
wsl.exe -d %WSL_DISTRO% -- bash -lc "cd '%WSL_PROJECT_DIR%' && sh packaging/linux/build-deb.sh"
if errorlevel 1 (
  echo.
  echo ERROR: Linux DEB build failed. See the WSL build output above.
)
goto wait

:wait
echo.
pause
goto menu

:end
endlocal
exit /b 0
