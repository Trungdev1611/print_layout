# build_portable_zip.ps1
# Package the addin as a portable .zip for distribution to other machines.
# The recipient simply extracts the zip and double-clicks install.bat
# (no code-signed .exe, so Windows SmartScreen does not silently block it).
#
# Generic on purpose: everything addin-specific lives in addin.config.json and
# templates\*.tmpl. To reuse this whole installer\ folder for another addin,
# copy it and edit addin.config.json only -- this script should not need changes.
#
# Output: installer\<AddinName>-Setup.zip
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
. (Join-Path $PSScriptRoot "_common.ps1")

$Config         = Get-AddinConfig -InstallerDir $PSScriptRoot
$RepoRoot       = Split-Path -Parent $PSScriptRoot
$Project        = Join-Path $RepoRoot $Config.ProjectPath
$ProjectDir     = Join-Path $RepoRoot (Split-Path $Config.ProjectPath -Parent)
$BinRelease     = Join-Path $ProjectDir "bin\Release"
$Bundle         = Join-Path $RepoRoot $Config.BundleFolderName
$BundleContents = Join-Path $Bundle "Contents"
$TemplatesDir   = Join-Path $PSScriptRoot "templates"
$InstallerDir   = $PSScriptRoot
$StageDir       = Join-Path $InstallerDir "stage"
$PayloadDir     = Join-Path $StageDir $Config.AddinName
$OutputZip      = Join-Path $InstallerDir "$($Config.AddinName)-Setup.zip"

$TemplateValues = @{
    ADDIN_NAME    = $Config.AddinName
    BUNDLE_FOLDER = $Config.BundleFolderName
    RIBBON_TAB    = $Config.RibbonTabName
    COMMANDS      = ($Config.Commands -join ", ")
}

function Step($msg) { Write-Host "==> $msg" -ForegroundColor Cyan }

# --- 1. Build the project ----------------------------------------------------
if ($SkipBuild) {
    Step "Skipping build (-SkipBuild)"
} else {
    Step "Building $($Config.AddinName) (Release)"
    $buildArgs = @("build", $Project, "-c", "Release", "--nologo", "-v", "minimal")
    if ($AutoCADPath) { $buildArgs += "/p:AutoCADPath=$AutoCADPath" }
    & dotnet @buildArgs
    if ($LASTEXITCODE -ne 0) { throw "dotnet build failed (exit $LASTEXITCODE)" }
}

# --- 2. Sync built artefacts into the bundle + write version.log ------------
# Two builds: net48 (AutoCAD 2018-2024) and net8.0-windows (AutoCAD 2025+).
# Each goes into its own Contents subfolder, matching PackageContents.xml.
Step "Syncing built DLLs into bundle/Contents (net48 + net8)"
Sync-BundleContents -BinRelease $BinRelease -BundleContents $BundleContents -DllFileName $Config.DllFileName

# --- 2b. Validate the bundle before packaging --------------------------------
Step "Validating bundle"
Test-BundleValid -Bundle $Bundle

# --- 2c. Regenerate install.bat / uninstall.bat from templates ---------------
# Regenerated every run so they can never drift from addin.config.json.
Step "Generating install.bat / uninstall.bat from templates"
$InstallBat   = Join-Path $InstallerDir "install.bat"
$UninstallBat = Join-Path $InstallerDir "uninstall.bat"
Expand-Template -TemplatePath (Join-Path $TemplatesDir "install.bat.tmpl") -OutputPath $InstallBat -Values $TemplateValues
Expand-Template -TemplatePath (Join-Path $TemplatesDir "uninstall.bat.tmpl") -OutputPath $UninstallBat -Values $TemplateValues

# --- 3. Stage payload --------------------------------------------------------
Step "Preparing staging folder"
if (Test-Path $StageDir) { Remove-Item -LiteralPath $StageDir -Recurse -Force }
New-Item -ItemType Directory -Path $PayloadDir | Out-Null

Copy-Item -Path $Bundle -Destination $PayloadDir -Recurse -Force
Copy-Item -Path $InstallBat -Destination (Join-Path $PayloadDir "install.bat") -Force

# --- 4. Write a Vietnamese quick-start readme --------------------------------
Step "Writing quick-start readme"
Expand-Template -TemplatePath (Join-Path $TemplatesDir "readme.tmpl") -OutputPath (Join-Path $PayloadDir "HUONG_DAN_CAI_DAT.txt") -Values $TemplateValues

# --- 5. Zip the payload ------------------------------------------------------
Step "Compressing to .zip"
if (Test-Path $OutputZip) { Remove-Item -LiteralPath $OutputZip -Force }
Compress-Archive -Path $PayloadDir -DestinationPath $OutputZip -CompressionLevel Optimal -Force

Remove-Item -LiteralPath $StageDir -Recurse -Force

# --- 6. Done -----------------------------------------------------------------
$sizeKB = [math]::Round((Get-Item $OutputZip).Length / 1KB, 1)
Write-Host ""
Write-Host "Portable installer built:" -ForegroundColor Green
Write-Host "  $OutputZip ($sizeKB KB)" -ForegroundColor Green
Write-Host ""
Write-Host "Send that single .zip. Recipient extracts it and runs install.bat." -ForegroundColor Green
