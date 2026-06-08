@echo off
setlocal
pushd "%~dp0"
set LOG=%TEMP%\MonthlySalesPlanCompareExe.log
echo [%DATE% %TIME%] Starting MonthlySalesPlanCompareExe from %CD% > "%LOG%"

start "MonthlySalesPlanCompareExe" "%CD%\MonthlySalesPlanCompareExe.exe" >> "%LOG%" 2>&1

echo Waiting for http://127.0.0.1:5002/ ...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ok=$false; for($i=0;$i -lt 30;$i++){ try { $r=Invoke-WebRequest -UseBasicParsing -Uri 'http://127.0.0.1:5002/' -TimeoutSec 1; if($r.StatusCode -eq 200){ $ok=$true; break } } catch { Start-Sleep -Seconds 1 } }; if($ok){ exit 0 } else { exit 1 }"

if errorlevel 1 (
  echo [%DATE% %TIME%] Server did not respond on port 5002. >> "%LOG%"
  echo.
  echo MonthlySalesPlanCompare did not start correctly.
  echo Please check this log file:
  echo %LOG%
  echo.
  echo Possible causes:
  echo - Port 5002 is already in use
  echo - Windows security blocked the exe
  echo - Network folder execution is blocked
  echo - You do not have access to the shared outputs folder
  pause
) else (
  echo [%DATE% %TIME%] Server is ready. >> "%LOG%"
  echo Opening browser: http://127.0.0.1:5002/
  powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process 'http://127.0.0.1:5002/'"
  echo.
  echo If the browser did not open, please open this URL manually:
  echo http://127.0.0.1:5002/
)
popd
exit /b 0
