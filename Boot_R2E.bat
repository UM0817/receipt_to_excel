@echo off
setlocal
powershell -NoExit -ExecutionPolicy Bypass -File "C:\Users\masat\project\receipt_to_excel\Shutdown_R2E.ps1" %*
if errorlevel 1 (
  echo.
  echo The launcher ended with an error.
)
pause
