@echo off
cd /d "%~dp0"
if not exist "data\main data 2.xlsx" goto missing
set "GARAGE_DATA_FILE=%CD%\data\main data 2.xlsx"
if exist "garage-server.pid" (
  set /p OLD_GARAGE_PID=<"garage-server.pid"
  taskkill /PID %OLD_GARAGE_PID% /T /F >nul 2>nul
  del /q "garage-server.pid" >nul 2>nul
)
if not exist "node_modules" (
  echo Preparing the system for first use...
  call npm.cmd install --cache .npm-cache
)
start "" http://localhost:3210
node server.js french-center-garage
exit /b 0
:missing
echo data\main data 2.xlsx was not found.
pause
exit /b 1
