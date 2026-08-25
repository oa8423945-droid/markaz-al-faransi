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

start "" http://localhost:3210
"%GARAGE_NODE%" server.js french-center-garage
exit /b %errorlevel%

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

