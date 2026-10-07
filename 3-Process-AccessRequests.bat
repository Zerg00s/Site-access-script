@echo off
setlocal
pushd "%~dp0"
title Access Request Processor
set "PS=powershell.exe"
where pwsh.exe >nul 2>&1 && set "PS=pwsh.exe"
%PS% -NoProfile -ExecutionPolicy Bypass -File "%~dp03-Process-AccessRequests.ps1"
set RC=%ERRORLEVEL%
popd
echo ExitCode=%RC%
pause
exit /b %RC%
