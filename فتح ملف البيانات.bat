@echo off
cd /d "%~dp0"
if not exist "data\main data 2.xlsx" goto missing
start "" "%CD%\data\main data 2.xlsx"
exit /b 0
:missing
echo data\main data 2.xlsx was not found.
pause
exit /b 1
