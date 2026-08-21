powershell -NoProfile -ExecutionPolicy Bypass -File installer\build_portable_zip.ps1 -AutoCADPath "C:\Program Files\Autodesk\AutoCAD 2024"

powershell -NoProfile -ExecutionPolicy Bypass -File installer\build_exe.ps1 -AutoCADPath "C:\Program Files\Autodesk\AutoCAD 2024"
