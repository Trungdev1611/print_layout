# build_exe.ps1
# Package the addin as a self-extracting .exe installer using Windows IExpress.
# The recipient just double-clicks the .exe; it extracts the bundle to
# %APPDATA%\Autodesk\ApplicationPlugins and AutoCAD auto-loads it next start.
#
# Generic on purpose: everything addin-specific lives in addin.config.json and
# templates\*.tmpl. To reuse this whole installer\ folder for another addin,
# copy it and edit addin.config.json only -- this script should not need changes.
# The one file NOT generic is iexpress\<AddinName>Setup.SED, which IExpress itself
# requires as a static file; its per-addin fields (TargetName, FriendlyName,
# SourceFiles0) are patched below from the config on every run.
#
# Output: installer\<AddinName>Setup.exe
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File installer\build_exe.ps1
#   powershell -NoProfile -ExecutionPolicy Bypass -File installer\build_exe.ps1 -SkipBuild
#   powershell -NoProfile -ExecutionPolicy Bypass -File installer\build_exe.ps1 -AutoCADPath "C:\Program Files\Autodesk\AutoCAD 2024"
#
# How it works: iexpress/<AddinName>Setup.SED bundles two files from iexpress/ —
# installer.cmd (the runtime extractor, generated from templates\installer.cmd.tmpl)
# and <BundleFolderName>.zip (the payload). This script regenerates both, then runs
# IExpress.

[CmdletBinding()]
param(
    [string]$AutoCADPath = "",
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "_common.ps1")

$Config         = Get-AddinConfig -InstallerDir $PSScriptRoot
$RepoRoot       = Split-Path -Parent $PSScriptRoot
$Project        = Join-Path $RepoRoot $Config.ProjectPath
$ProjectDir     = Join-Path $RepoRoot (Split-Path $Config.ProjectPath -Parent)
$BinRelease     = Join-Path $ProjectDir "bin\Release"
$Bundle         = Join-Path $RepoRoot $Config.BundleFolderName
$BundleContents = Join-Path $Bundle "Contents"
$TemplatesDir   = Join-Path $PSScriptRoot "templates"
$IExpressDir    = Join-Path $PSScriptRoot "iexpress"
$Sed            = Join-Path $IExpressDir "$($Config.AddinName)Setup.SED"
$BundleZip      = Join-Path $IExpressDir "$($Config.BundleFolderName).zip"
$InstallerCmd   = Join-Path $IExpressDir "installer.cmd"
$OutputExe      = Join-Path $PSScriptRoot "$($Config.AddinName)Setup.exe"

$TemplateValues = @{
    ADDIN_NAME    = $Config.AddinName
    BUNDLE_FOLDER = $Config.BundleFolderName
    RIBBON_TAB    = $Config.RibbonTabName
    COMMANDS      = ($Config.Commands -join ", ")
}

function Step($msg) { Write-Host "==> $msg" -ForegroundColor Cyan }

# --- 1. Build both targets (net48 + net8.0-windows) --------------------------
if ($SkipBuild) {
    Step "Skipping build (-SkipBuild)"
} else {
    Step "Building $($Config.AddinName) (Release, both targets)"
    $buildArgs = @("build", $Project, "-c", "Release", "--nologo", "-v", "minimal")
    if ($AutoCADPath) { $buildArgs += "/p:AutoCADPath=$AutoCADPath" }
    & dotnet @buildArgs
    if ($LASTEXITCODE -ne 0) { throw "dotnet build failed (exit $LASTEXITCODE)" }
}

# --- 2. Sync built DLLs into the bundle + write version.log -----------------
Step "Syncing built DLLs into bundle/Contents (net48 + net8)"
Sync-BundleContents -BinRelease $BinRelease -BundleContents $BundleContents -DllFileName $Config.DllFileName

# --- 2b. Validate the bundle before packaging --------------------------------
Step "Validating bundle"
Test-BundleValid -Bundle $Bundle

# --- 2c. Regenerate installer.cmd from template -------------------------------
Step "Generating iexpress/installer.cmd from template"
Expand-Template -TemplatePath (Join-Path $TemplatesDir "installer.cmd.tmpl") -OutputPath $InstallerCmd -Values $TemplateValues

# --- 3. Zip the bundle folder (root entry = <BundleFolderName>\) -------------
# installer.cmd extracts this straight into %APPDATA%\Autodesk\ApplicationPlugins,
# so the archive must contain the bundle folder as its root.
Step "Compressing bundle -> $($Config.BundleFolderName).zip"
if (Test-Path $BundleZip) { Remove-Item -LiteralPath $BundleZip -Force }
Compress-Archive -Path $Bundle -DestinationPath $BundleZip -CompressionLevel Optimal -Force

# --- 4. Patch SED fields for this machine/addin, then run IExpress -----------
Step "Patching SED TargetName / FriendlyName / SourceFiles0"
if (-not (Test-Path $Sed)) { throw "Missing SED: $Sed (expected iexpress\$($Config.AddinName)Setup.SED)" }
$sedText = Get-Content -LiteralPath $Sed -Raw
$sedText = [regex]::Replace($sedText, '(?m)^TargetName=.*$', "TargetName=$OutputExe")
$sedText = [regex]::Replace($sedText, '(?m)^FriendlyName=.*$', "FriendlyName=$($Config.AddinName) Setup")
$sedText = [regex]::Replace($sedText, '(?m)^SourceFiles0=.*$', ("SourceFiles0=" + $IExpressDir.TrimEnd('\') + '\'))
$sedText = [regex]::Replace($sedText, '(?m)^FILE1=.*$', "FILE1=$($Config.BundleFolderName).zip")
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
