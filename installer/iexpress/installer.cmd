@echo off
REM Runs inside the IExpress temp extraction folder.
REM Sibling files: PrintLayoutAddin.bundle.zip
setlocal enabledelayedexpansion

set TARGET_PARENT=%APPDATA%\Autodesk\ApplicationPlugins
set TARGET=%TARGET_PARENT%\PrintLayoutAddin.bundle
set ZIP=%~dp0PrintLayoutAddin.bundle.zip

echo.
echo ====================================================
echo   PrintLayoutAddin - Setup
echo ====================================================
echo.

if not exist "%ZIP%" (
  echo [ERROR] Bundle archive missing: %ZIP%
  echo The installer .exe is corrupt or was extracted incorrectly.
  pause
  exit /b 1
)

REM Warn if AutoCAD is running (DLL is locked while AutoCAD owns it).
tasklist /FI "IMAGENAME eq acad.exe" 2>nul | find /I "acad.exe" >nul
if not errorlevel 1 (
  echo [WARNING] AutoCAD is running. Close all AutoCAD windows before continuing,
  echo           otherwise the install will fail because the plugin DLL is locked.
  echo.
  pause
)

if not exist "%TARGET_PARENT%" mkdir "%TARGET_PARENT%"

if exist "%TARGET%" (
  echo Removing previous installation at:
  echo   %TARGET%
  rmdir /s /q "%TARGET%"
  if exist "%TARGET%" (
    echo [ERROR] Could not remove previous installation. Is AutoCAD still running?
    pause
    exit /b 1
  )
)

echo Extracting bundle to:
echo   %TARGET_PARENT%
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { Expand-Archive -LiteralPath '%ZIP%' -DestinationPath '%TARGET_PARENT%' -Force; exit 0 } catch { Write-Host $_; exit 1 }"
if errorlevel 1 (
  echo [ERROR] Extraction failed.
  pause
  exit /b 1
)

if not exist "%TARGET%\PackageContents.xml" (
  echo [ERROR] Bundle did not extract correctly.
  echo         Expected: %TARGET%\PackageContents.xml
  pause
  exit /b 1
)

echo.
echo Installation successful.
echo Open AutoCAD to see the "Print Layout" tab on the ribbon.
echo Commands available from the command line: PLSTT, PLAYOUT, PLVP, PLAUTO, PLCLEAN.
echo.
pause
endlocal
