@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo.
echo ============================================================
echo  Attendance Excel to JSON Converter v5
echo ============================================================
echo.
echo Folder: %CD%
echo.

if "%~1"=="" goto ASK_PATH
set "XLSX=%~1"
goto CHECK_FILE

:ASK_PATH
echo Drag and drop the XLSX file onto this BAT file, or type/paste the full XLSX path.
echo Example: C:\Users\User\Desktop\attendance.xlsx
echo.
set /p "XLSX=XLSX path: "
goto CHECK_FILE

:CHECK_FILE
set "XLSX=%XLSX:"=%"
if "%XLSX%"=="" goto NO_FILE
if not exist "%XLSX%" goto NO_FILE

if not exist "%~dp0data" mkdir "%~dp0data"
set "LOG=%~dp0conversion_log.txt"
if exist "%LOG%" del "%LOG%" >nul 2>nul

echo Input file:
echo %XLSX%
echo.

where powershell >nul 2>nul
if errorlevel 1 goto TRY_PYTHON

echo [1/2] Trying PowerShell + Microsoft Excel conversion...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0excel_to_json_excelcom.ps1" -XlsxPath "%XLSX%" > "%LOG%" 2>&1
if not errorlevel 1 goto SUCCESS

echo.
echo PowerShell + Excel conversion failed. Log:
type "%LOG%"
echo.
echo [2/2] Trying Python conversion...
goto TRY_PYTHON

:TRY_PYTHON
where py >nul 2>nul
if not errorlevel 1 (
  py -3 "%~dp0excel_to_json.py" "%XLSX%" >> "%LOG%" 2>&1
  if not errorlevel 1 goto SUCCESS
)

where python >nul 2>nul
if not errorlevel 1 (
  python "%~dp0excel_to_json.py" "%XLSX%" >> "%LOG%" 2>&1
  if not errorlevel 1 goto SUCCESS
)

goto FAILED

:SUCCESS
echo.
echo SUCCESS: data\current.json has been updated.
echo You can now open index.html and choose data\current.json.
echo.
if exist "%LOG%" (
  echo Last log:
  type "%LOG%"
  echo.
)
pause
exit /b 0

:NO_FILE
echo.
echo ERROR: XLSX file was not found.
echo Value: %XLSX%
echo.
echo Tips:
echo - Put the XLSX file on Desktop and try again.
echo - Close the XLSX file before conversion.
echo - Avoid special characters in folder names if possible.
echo.
pause
exit /b 1

:FAILED
echo.
echo ERROR: Conversion failed.
echo.
echo Check these items:
echo 1. Microsoft Excel is installed and the XLSX file is closed.
echo 2. The workbook must have the required attendance sheet.
echo 3. If Python is used, install openpyxl: pip install openpyxl
echo 4. If company security blocks BAT/PS1, ask IT to allow this local script.
echo.
if exist "%LOG%" (
  echo Log:
  type "%LOG%"
)
echo.
pause
exit /b 1
