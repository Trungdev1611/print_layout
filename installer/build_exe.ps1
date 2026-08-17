# build_exe.ps1
# Package PrintLayoutAddin as a self-extracting .exe installer using Windows IExpress.
# The recipient just double-clicks the .exe; it extracts the bundle to
# %APPDATA%\Autodesk\ApplicationPlugins and AutoCAD auto-loads it next start.
#
# Output: installer\PrintLayoutAddinSetup.exe
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File installer\build_exe.ps1
#   powershell -NoProfile -ExecutionPolicy Bypass -File installer\build_exe.ps1 -SkipBuild
#   powershell -NoProfile -ExecutionPolicy Bypass -File installer\build_exe.ps1 -AutoCADPath "C:\Program Files\Autodesk\AutoCAD 2024"
#
# How it works: iexpress/PrintLayoutAddinSetup.SED bundles two files from iexpress/ —
#   installer.cmd (the runtime extractor) and PrintLayoutAddin.bundle.zip (the payload).
# This script regenerates that .zip from the freshly built bundle, then runs IExpress.

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
$IExpressDir    = Join-Path $PSScriptRoot "iexpress"
$Sed            = Join-Path $IExpressDir "PrintLayoutAddinSetup.SED"
$BundleZip      = Join-Path $IExpressDir "PrintLayoutAddin.bundle.zip"
$OutputExe      = Join-Path $PSScriptRoot "PrintLayoutAddinSetup.exe"

function Step($msg) { Write-Host "==> $msg" -ForegroundColor Cyan }
function Info($msg) { Write-Host "    $msg" -ForegroundColor DarkGray }

# --- 1. Build both targets (net48 + net8.0-windows) --------------------------
if ($SkipBuild) {
    Step "Skipping build (-SkipBuild)"
} else {
    Step "Building PrintLayoutAddin (Release, both targets)"
    $buildArgs = @("build", $Project, "-c", "Release", "--nologo", "-v", "minimal")
    if ($AutoCADPath) { $buildArgs += "/p:AutoCADPath=$AutoCADPath" }
    & dotnet @buildArgs
    if ($LASTEXITCODE -ne 0) { throw "dotnet build failed (exit $LASTEXITCODE)" }
}

# --- 2. Sync built DLLs into the bundle subfolders ---------------------------
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

# --- 3. Zip the bundle folder (root entry = PrintLayoutAddin.bundle\) ---------
# installer.cmd extracts this straight into %APPDATA%\Autodesk\ApplicationPlugins,
# so the archive must contain the PrintLayoutAddin.bundle folder as its root.
Step "Compressing bundle -> PrintLayoutAddin.bundle.zip"
if (Test-Path $BundleZip) { Remove-Item -LiteralPath $BundleZip -Force }
Compress-Archive -Path $Bundle -DestinationPath $BundleZip -CompressionLevel Optimal -Force

# --- 4. Patch SED paths for this machine, then run IExpress ------------------
Step "Patching SED TargetName / SourceFiles0 for this repo"
if (-not (Test-Path $Sed)) { throw "Missing SED: $Sed" }
$sedText = Get-Content -LiteralPath $Sed -Raw
$sedText = [regex]::Replace($sedText, '(?m)^TargetName=.*$', "TargetName=$OutputExe")
$sedText = [regex]::Replace($sedText, '(?m)^SourceFiles0=.*$', ("SourceFiles0=" + $IExpressDir.TrimEnd('\') + '\'))
Set-Content -LiteralPath $Sed -Value $sedText -NoNewline -Encoding ASCII

Step "Running IExpress"
$iexpress = Join-Path $env:WINDIR "System32\iexpress.exe"
if (-not (Test-Path $iexpress)) { $iexpress = Join-Path $env:WINDIR "SysWOW64\iexpress.exe" }
if (-not (Test-Path $iexpress)) { throw "iexpress.exe not found on this machine" }
if (Test-Path $OutputExe) { Remove-Item -LiteralPath $OutputExe -Force }

& $iexpress /N /Q $Sed | Out-Null
if (-not (Test-Path $OutputExe)) {
    throw "IExpress did not produce the .exe. Check $Sed (TargetName / SourceFiles paths)."
}

# --- 5. Done -----------------------------------------------------------------
$sizeKB = [math]::Round((Get-Item $OutputExe).Length / 1KB, 1)
Write-Host ""
Write-Host "Self-extracting installer built:" -ForegroundColor Green
Write-Host "  $OutputExe ($sizeKB KB)" -ForegroundColor Green
Write-Host ""
Write-Host "Send that single .exe. Recipient double-clicks it; AutoCAD loads the addin next start." -ForegroundColor Green
