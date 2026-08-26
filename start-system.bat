@echo off
setlocal
cd /d "%~dp0"
if not exist "data\main data 2.xlsx" goto missing_data
set "GARAGE_DATA_FILE=%CD%\data\main data 2.xlsx"

if exist "garage-server.pid" (
  set /p GARAGE_OLD_PID=<"garage-server.pid"
  taskkill /PID %GARAGE_OLD_PID% /T /F >nul 2>nul
  del /q "garage-server.pid" >nul 2>nul
)

rem عند تحديث البرنامج قد تكون النسخة السابقة لا تملك ملف PID؛ أوقف فقط العملية التي تستخدم منفذ النظام.
for /f "tokens=5" %%P in ('netstat -ano ^| findstr /R /C:":3210 .*LISTENING"') do taskkill /PID %%P /T /F >nul 2>nul

if exist "runtime\node.exe" (
  set "GARAGE_NODE=%CD%\runtime\node.exe"
) else (
  where node.exe >nul 2>nul
  if errorlevel 1 goto missing_node
  set "GARAGE_NODE=node.exe"
)

if not exist "node_modules\xlsx" (
  where npm.cmd >nul 2>nul
  if errorlevel 1 goto missing_modules
  echo Preparing the system for first use...
  call npm.cmd install --omit=dev --cache .npm-cache
  if errorlevel 1 goto install_failed
)

start "" /b "%GARAGE_NODE%" server.js french-center-garage
ping 127.0.0.1 -n 3 >nul
start "" http://localhost:3210
exit /b 0

:missing_data
echo The database file data\main data 2.xlsx was not found.
goto failed

:missing_node
echo Node.js was not found. Please use FrenchCenter-Setup.exe.
goto failed

:missing_modules
echo The system libraries are missing. Please reinstall using FrenchCenter-Setup.exe.
goto failed

:install_failed
echo The system libraries could not be installed. Check the internet connection and try again.

:failed
pause
exit /b 1

