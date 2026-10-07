@echo off
setlocal
pushd "%~dp0"
set "PS=powershell.exe"
where pwsh.exe >nul 2>&1 && set "PS=pwsh.exe"
for /f %%a in ('%PS% -NoProfile -Command "Get-Date -Format yyyy-MM-dd_HHmmss"') do set STAMP=%%a
if not exist "%~dp0BatLogs" mkdir "%~dp0BatLogs"
set "LOG=%~dp0BatLogs\2-Create-AccessRequestLists_%STAMP%.log"
%PS% -NoProfile -ExecutionPolicy Bypass -Command "& { & '%~dp02-Create-AccessRequestLists.ps1' *>&1 | Tee-Object -FilePath '%LOG%' -Append; exit $LASTEXITCODE }"
set RC=%ERRORLEVEL%
popd
echo.
echo ExitCode=%RC%   Log: %LOG%
pause
exit /b %RC%
