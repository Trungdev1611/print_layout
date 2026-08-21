@echo off
REM Install PrintLayoutAddin into AutoCAD ApplicationPlugins (auto-loads on every AutoCAD start)
setlocal

set TARGET=%APPDATA%\Autodesk\ApplicationPlugins\PrintLayoutAddin.bundle
REM Bundle sits next to install.bat (distributed zip) or one level up (repo: installer\ -> root)
set SRC=%~dp0PrintLayoutAddin.bundle
if not exist "%SRC%" set SRC=%~dp0..\PrintLayoutAddin.bundle

if not exist "%SRC%" (
  echo [ERROR] Folder not found: %SRC%
  echo Make sure install.bat sits next to the PrintLayoutAddin.bundle folder.
  pause
  exit /b 1
)

echo.
echo ====================================================
echo   PrintLayoutAddin - Installer
echo ====================================================
echo.
echo Please close all AutoCAD windows before continuing.
echo.
pause

echo Copying bundle to: %TARGET%
if exist "%TARGET%" rmdir /s /q "%TARGET%"
xcopy /e /i /y "%SRC%" "%TARGET%" >nul

if errorlevel 1 (
  echo [ERROR] Copy failed. AutoCAD might still be running and locking the DLL.
  pause
  exit /b 1
)

echo.
echo Installation successful.
echo Open AutoCAD to see the "Print Layout" tab on the ribbon.
echo Commands available from the command line: PLSTT, PLAYOUT, PLPRINT, PLVP, PLKEYS.
echo.
pause
endlocal
