@echo off
REM ---------------------------------------------------------------------------
REM  RUN_<job>.cmd  -  template launcher for pcat-win.ps1 (fill in the
REM  PS_ARGS line). Lives in <drive>:\_PhotoJob\<job>\ next to pcat-win.ps1.
REM  On success: deletes the working folder and itself. On failure: keeps
REM  everything so the run can be inspected and retried.
REM ---------------------------------------------------------------------------
setlocal
set "JOBDIR=%~dp0"
set "PS_ARGS=-Mode VerifyTree -A "C:\Media" -B "D:\Media" -OutDir "C:\Media\Photo & Video Library\_Catalog\verify_D_YYYY-MM-DD""

powershell -NoProfile -ExecutionPolicy Bypass -File "%JOBDIR%pcat-win.ps1" %PS_ARGS%
set "RC=%ERRORLEVEL%"
echo.
if not "%RC%"=="0" (
  echo *** The run reported problems ^(exit code %RC%^). Nothing was cleaned up.
  echo *** Check the log in the -OutDir folder.
  pause
  exit /b %RC%
)
echo Done. Results are in the -OutDir folder. Removing the working folder...
pause
cd /d "%TEMP%"
del /q "%JOBDIR%pcat-win.ps1" "%JOBDIR%*.csv" 2>nul
(goto) 2>nul & del "%~f0" & rmdir "%JOBDIR%" 2>nul
