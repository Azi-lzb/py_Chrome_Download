@echo off
setlocal

cd /d "%~dp0"

set "PYTHON_CMD=%USERPROFILE%\miniconda3\envs\HZPBC\python.exe"
if not exist "%PYTHON_CMD%" (
  echo ERROR: Required HZPBC Python environment was not found:
  echo   %PYTHON_CMD%
  echo Install or restore the HZPBC environment, then try again.
  exit /b 1
)

echo [1/4] Checking HZPBC Python and dependencies...
call "%PYTHON_CMD%" --version
if errorlevel 1 exit /b 1
call "%PYTHON_CMD%" -c "import selenium, openpyxl; print('Selenium', selenium.__version__)"
if errorlevel 1 (
  echo ERROR: Selenium or openpyxl is missing from the HZPBC environment.
  exit /b 1
)

echo [2/4] Checking PyInstaller...
call "%PYTHON_CMD%" -m PyInstaller --version
if errorlevel 1 (
  echo ERROR: PyInstaller is not installed in the HZPBC environment.
  echo Install it with:
  echo   "%PYTHON_CMD%" -m pip install pyinstaller
  exit /b 1
)

echo [3/4] Building ChromeAutoRun.exe...
call "%PYTHON_CMD%" -m PyInstaller --clean --noconfirm main.spec
if errorlevel 1 (
  echo ERROR: PyInstaller build failed.
  exit /b 1
)

if not exist "dist\ChromeAutoRun.exe" (
  echo ERROR: dist\ChromeAutoRun.exe was not generated.
  exit /b 1
)

echo [4/4] Copying deliverables to dist_hzpbc...
if not exist "dist_hzpbc" mkdir "dist_hzpbc"
for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "STAMP=%%I"
if not defined STAMP (
  echo ERROR: Could not generate a timestamp for the archive copy.
  exit /b 1
)
if exist "dist_hzpbc\ChromeAutoRun_%STAMP%.exe" (
  echo ERROR: Timestamped deliverable already exists; refusing to overwrite it.
  exit /b 1
)
copy /Y "dist\ChromeAutoRun.exe" "dist_hzpbc\ChromeAutoRun.exe" >nul
if errorlevel 1 (
  echo ERROR: Failed to copy the current executable to dist_hzpbc.
  exit /b 1
)
copy "dist\ChromeAutoRun.exe" "dist_hzpbc\ChromeAutoRun_%STAMP%.exe" >nul
if errorlevel 1 (
  echo ERROR: Failed to create the timestamped rollback copy.
  exit /b 1
)

echo Build completed:
echo   %CD%\dist\ChromeAutoRun.exe
echo   %CD%\dist_hzpbc\ChromeAutoRun.exe
echo   %CD%\dist_hzpbc\ChromeAutoRun_%STAMP%.exe
exit /b 0
