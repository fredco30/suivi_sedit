@echo off
setlocal EnableExtensions

REM Usage: drag a .py file ONTO this .cmd
if "%~1"=="" (
  echo Drag and drop your AE_Gestion .py onto this file to run it with logging.
  pause
  goto :EOF
)

set "HERE=%~dp0"
pushd "%HERE%"

set "SCRIPT=%~1"
echo [CHECK] Script: "%SCRIPT%"
if not exist "%SCRIPT%" (
  echo [ERROR] File not found: "%SCRIPT%"
  pause
  goto :EOF
)

set "PYCMD="
for %%P in (py.exe,py,python.exe,python,python3.exe,python3) do (
  where %%P >nul 2>&1 && (set "PYCMD=%%P" & goto :FOUND)
)
:FOUND
if not defined PYCMD (
  set "PYCMD=%LocalAppData%\Programs\Python\Python313\python.exe"
)

set "LOGDIR=run_logs"
if not exist "%LOGDIR%" mkdir "%LOGDIR%" >nul 2>&1
for /f "tokens=1-3 delims=/ " %%a in ("%date%") do set "D=%%c%%b%%a"
for /f "tokens=1-3 delims=:." %%a in ("%time%") do set "T=%%a%%b%%c"
set "STAMP=%D%_%T%"
set "STAMP=%STAMP: =0%"
set "LOG=%LOGDIR%\AE_RUN_%STAMP%.log"

echo [RUN] "%PYCMD%" -X dev -u "%SCRIPT%"
"%PYCMD%" -X dev -u "%SCRIPT%" >> "%LOG%" 2>&1
set "RC=%ERRORLEVEL%"
echo [DONE] RC=%RC%
start "" notepad.exe "%LOG%"
pause
