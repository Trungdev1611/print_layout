# build_portable_zip.ps1
# Package PrintLayoutAddin as a portable .zip for distribution to other machines.
# The recipient simply extracts the zip and double-clicks install.bat
# (no code-signed .exe, so Windows SmartScreen does not silently block it).
#
# Output: installer\PrintLayoutAddin-Setup.zip
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File installer\build_portable_zip.ps1
#   powershell -NoProfile -ExecutionPolicy Bypass -File installer\build_portable_zip.ps1 -SkipBuild
#   powershell -NoProfile -ExecutionPolicy Bypass -File installer\build_portable_zip.ps1 -AutoCADPath "C:\Program Files\Autodesk\AutoCAD 2024"

[CmdletBinding()]
param(
    [string]$AutoCADPath = "",
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"

$RepoRoot       = Split-Path -Parent $PSScriptRoot
$Project        = Join-Path $RepoRoot "PrintLayoutAddin\PrintLayoutAddin.csproj"
$BinRelease     = Join-Path $RepoRoot "PrintLayoutAddin\bin\Release"
$Bundle         = Join-Path $RepoRoot "PrintLayoutAddin.bundle"
$BundleContents = Join-Path $Bundle "Contents"
$InstallBat     = Join-Path $PSScriptRoot "install.bat"
$InstallerDir   = $PSScriptRoot
$StageDir       = Join-Path $InstallerDir "stage"
$PayloadDir     = Join-Path $StageDir "PrintLayoutAddin"
$OutputZip      = Join-Path $InstallerDir "PrintLayoutAddin-Setup.zip"

function Step($msg) { Write-Host "==> $msg" -ForegroundColor Cyan }
function Info($msg) { Write-Host "    $msg" -ForegroundColor DarkGray }

# --- 1. Build the project ----------------------------------------------------
if ($SkipBuild) {
    Step "Skipping build (-SkipBuild)"
} else {
    Step "Building PrintLayoutAddin (Release)"
    $buildArgs = @("build", $Project, "-c", "Release", "--nologo", "-v", "minimal")
    if ($AutoCADPath) { $buildArgs += "/p:AutoCADPath=$AutoCADPath" }
    & dotnet @buildArgs
    if ($LASTEXITCODE -ne 0) { throw "dotnet build failed (exit $LASTEXITCODE)" }
}

# --- 2. Sync built artefacts into the bundle ---------------------------------
# Two builds: net48 (AutoCAD 2018-2024) and net8.0-windows (AutoCAD 2025+).
# Each goes into its own Contents subfolder, matching PackageContents.xml.
Step "Syncing built DLLs into bundle/Contents (net48 + net8)"
$tfmMap = [ordered]@{ "net48" = "net48"; "net8.0-windows" = "net8" }
foreach ($tfm in $tfmMap.Keys) {
    $srcDir  = Join-Path $BinRelease $tfm
    $destDir = Join-Path $BundleContents $tfmMap[$tfm]
    if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir | Out-Null }
    foreach ($f in @("PrintLayoutAddin.v201.dll", "config.json")) {
        $src = Join-Path $srcDir $f
        if (-not (Test-Path $src)) { throw "Missing build output: $src" }
        Copy-Item -Path $src -Destination (Join-Path $destDir $f) -Force
        Info "copied $tfm/$f"
    }
}

# --- 3. Stage payload --------------------------------------------------------
Step "Preparing staging folder"
if (Test-Path $StageDir) { Remove-Item -LiteralPath $StageDir -Recurse -Force }
New-Item -ItemType Directory -Path $PayloadDir | Out-Null

Copy-Item -Path $Bundle -Destination $PayloadDir -Recurse -Force
Copy-Item -Path $InstallBat -Destination (Join-Path $PayloadDir "install.bat") -Force

# --- 4. Write a Vietnamese quick-start readme --------------------------------
$readme = @"
PrintLayoutAddin - Huong dan cai dat
====================================

1. Dong tat ca cua so AutoCAD truoc khi cai.
2. Nhap chuot phai vao file install.bat -> Run as administrator
   (hoac double-click install.bat). Neu Windows hien bang xanh
   "Windows protected your PC" thi bam "More info" -> "Run anyway".
3. Lam theo huong dan tren man hinh (nhan phim bat ky khi gap "pause").
4. Mo lai AutoCAD -> tab "Print Layout" se xuat hien tren thanh ribbon.

Cac lenh: PLSTT, PLAYOUT, PLPRINT, PLVP, PLAUTO, PLCLEAN, PLLICENSE.

Cai dat duoc chep vao:
  %APPDATA%\Autodesk\ApplicationPlugins\PrintLayoutAddin.bundle
(khong can quyen admin, AutoCAD tu nap addin o lan khoi dong ke tiep).

Go cai dat: xoa thu muc o duong dan tren, hoac chay uninstall neu co.
"@
$readmeCrlf = ($readme -replace "`r`n", "`n") -replace "`n", "`r`n"
[IO.File]::WriteAllText((Join-Path $PayloadDir "HUONG_DAN_CAI_DAT.txt"), $readmeCrlf, [Text.Encoding]::UTF8)

# --- 5. Zip the payload ------------------------------------------------------
Step "Compressing to .zip"
if (Test-Path $OutputZip) { Remove-Item -LiteralPath $OutputZip -Force }
Compress-Archive -Path (Join-Path $StageDir "PrintLayoutAddin") -DestinationPath $OutputZip -CompressionLevel Optimal -Force

Remove-Item -LiteralPath $StageDir -Recurse -Force

# --- 6. Done -----------------------------------------------------------------
$sizeKB = [math]::Round((Get-Item $OutputZip).Length / 1KB, 1)
Write-Host ""
Write-Host "Portable installer built:" -ForegroundColor Green
Write-Host "  $OutputZip ($sizeKB KB)" -ForegroundColor Green
Write-Host ""
Write-Host "Send that single .zip. Recipient extracts it and runs install.bat." -ForegroundColor Green
