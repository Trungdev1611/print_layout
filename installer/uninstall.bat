@echo off
REM Remove PrintLayoutAddin from AutoCAD ApplicationPlugins
setlocal
set TARGET=%APPDATA%\Autodesk\ApplicationPlugins\PrintLayoutAddin.bundle

if not exist "%TARGET%" (
  echo Nothing to remove. %TARGET% does not exist.
  pause
  exit /b 0
)

echo.
echo Please close all AutoCAD windows before continuing.
echo.
pause

rmdir /s /q "%TARGET%"
if exist "%TARGET%" (
  echo [ERROR] Removal failed. AutoCAD may still be running.
  pause
  exit /b 1
)

echo PrintLayoutAddin has been removed. Restart AutoCAD to finish.
pause
endlocal
