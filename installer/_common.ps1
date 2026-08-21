# _common.ps1
# Shared helpers for build_portable_zip.ps1 / build_exe.ps1. Generic on purpose --
# everything addin-specific comes from addin.config.json, not from this file.
# Dot-source it: . (Join-Path $PSScriptRoot "_common.ps1")

function Get-AddinConfig {
    param([Parameter(Mandatory)] [string]$InstallerDir)
    $path = Join-Path $InstallerDir "addin.config.json"
    if (-not (Test-Path $path)) { throw "Missing addin.config.json at $path" }
    Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
}

# Replaces {{TOKEN}} placeholders in a template file and writes the result with
# CRLF line endings (these are all plain-text files read on Windows/cmd.exe).
function Expand-Template {
    param(
        [Parameter(Mandatory)] [string]$TemplatePath,
        [Parameter(Mandatory)] [string]$OutputPath,
        [Parameter(Mandatory)] [hashtable]$Values
    )
    $text = Get-Content -LiteralPath $TemplatePath -Raw
    foreach ($key in $Values.Keys) {
        # Replacement side of -replace treats $ specially; our config values never
        # contain one, so a literal string substitute (not a regex) is used here.
        $text = $text.Replace("{{$key}}", [string]$Values[$key])
    }
    $crlf = ($text -replace "`r`n", "`n") -replace "`n", "`r`n"
    # No BOM: these are read by cmd.exe (batch files), which doesn't reliably
    # tolerate a UTF-8 BOM before "@echo off".
    $noBomUtf8 = New-Object System.Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($OutputPath, $crlf, $noBomUtf8)
}

# Copies the WHOLE build output folder (not a hardcoded file list) into
# bundle/Contents/<net48|net8>, for both target frameworks, then writes a
# version.log snapshot next to the DLL. A fixed file list silently drops runtime
# dependencies (e.g. a license DLL) -- exactly the bug that broke a ribbon-loaded
# install while Add-in Manager, pointed straight at bin\Debug, worked fine.
function Sync-BundleContents {
    param(
        [Parameter(Mandatory)] [string]$BinRelease,
        [Parameter(Mandatory)] [string]$BundleContents,
        [Parameter(Mandatory)] [string]$DllFileName
    )
    $tfmMap = [ordered]@{ "net48" = "net48"; "net8.0-windows" = "net8" }
    foreach ($tfm in $tfmMap.Keys) {
        $srcDir  = Join-Path $BinRelease $tfm
        $destDir = Join-Path $BundleContents $tfmMap[$tfm]
        if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir | Out-Null }
        if (-not (Test-Path $srcDir)) { throw "Missing build output folder: $srcDir" }
        Get-ChildItem -Path $srcDir -File | Where-Object { $_.Extension -ne ".pdb" } | ForEach-Object {
            Copy-Item -Path $_.FullName -Destination (Join-Path $destDir $_.Name) -Force
            Write-Host "    copied $tfm/$($_.Name)" -ForegroundColor DarkGray
        }

        $dllPath = Join-Path $destDir $DllFileName
        $fileVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($dllPath).FileVersion
        Set-Content -LiteralPath (Join-Path $destDir "version.log") -Encoding UTF8 -Value @(
            "$DllFileName v$fileVersion"
            "Packaged: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        )
        Write-Host "    wrote $tfm/version.log ($fileVersion)" -ForegroundColor DarkGray
    }
}

# PackageContents.xml is a hand-written, git-tracked source file -- no build step
# generates it. Without it AutoCAD IGNORES the whole .bundle folder silently: no
# ribbon, no commands, no error message, and NETLOAD still works (which makes it
# look like a code bug for hours). Fail loudly here instead of shipping another
# silently-broken installer.
function Test-BundleValid {
    param([Parameter(Mandatory)] [string]$Bundle)
    $pkgXmlPath = Join-Path $Bundle "PackageContents.xml"
    if (-not (Test-Path $pkgXmlPath)) {
        throw "MISSING $pkgXmlPath -- AutoCAD would silently ignore the bundle. Restore it (git restore) before packaging."
    }

    [xml]$pkgXml = Get-Content -LiteralPath $pkgXmlPath
    $moduleNames = $pkgXml.ApplicationPackage.Components.ComponentEntry |
        ForEach-Object { $_.ModuleName } | Where-Object { $_ }
    if (-not $moduleNames) { throw "No ComponentEntry/ModuleName found in $pkgXmlPath" }
    foreach ($m in $moduleNames) {
        $rel = $m -replace '^\./', '' -replace '/', '\'
        $abs = Join-Path $Bundle $rel
        if (-not (Test-Path $abs)) {
            throw "PackageContents.xml points at a missing file: $m (expected $abs)"
        }
        Write-Host "    verified $m" -ForegroundColor DarkGray
    }
}
