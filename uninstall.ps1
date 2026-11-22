param([switch]$PurgeSettings)

$ErrorActionPreference = "Stop"

$TargetDir = [System.IO.Path]::GetFullPath($env:APPDATA)
foreach ($directoryName in @("Microsoft", "AddIns", "PptMathFontAddin")) {
    $TargetDir = Join-Path $TargetDir $directoryName
    if (Test-Path -LiteralPath $TargetDir -PathType Container) {
        $directory = Get-Item -LiteralPath $TargetDir -Force
        if ($directory.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
            throw "The uninstall path contains a directory link: $TargetDir"
        }
    }
}
$PpamPath = Join-Path $TargetDir "PptMathFontAddin.ppam"

function Find-InstalledAddInIndex($AddIns) {
    for ($i = 1; $i -le $AddIns.Count; $i++) {
        $addin = $AddIns.Item($i)
        $fullName = [System.IO.Path]::GetFullPath($addin.FullName)
        if ([string]::Equals($fullName, $PpamPath, [StringComparison]::OrdinalIgnoreCase)) {
            return $i
        }
    }
    return 0
}

$running = Get-Process POWERPNT -ErrorAction SilentlyContinue
if ($running) {
    Write-Host "Close every PowerPoint window, then run uninstall.ps1 again."
    Write-Host "Press Enter to close."
    [void](Read-Host)
    exit 1
}

$ppt = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $addIns = $ppt.AddIns
    $index = Find-InstalledAddInIndex $addIns

    if ($index -gt 0) {
        $addin = $addIns.Item($index)
        $addin.AutoLoad = 0
        $addin.Loaded = 0
        $addIns.Remove($index)
    }

    if ((Find-InstalledAddInIndex $addIns) -gt 0) {
        throw "PowerPoint still lists the add-in at $PpamPath."
    }
}
catch {
    throw "Could not unregister the add-in. No files or settings were deleted. $($_.Exception.Message)"
}
finally {
    if ($null -ne $ppt) {
        $ppt.Quit()
    }
}

foreach ($fileName in @("PptMathFontAddin.ppam", "MathFontAddin.bas", "CAppEvents.cls", "customUI.xml")) {
    $filePath = Join-Path $TargetDir $fileName
    if (Test-Path -LiteralPath $filePath -PathType Leaf) {
        Remove-Item -LiteralPath $filePath -Force
    }
}

if ((Test-Path -LiteralPath $TargetDir -PathType Container) -and
    -not (Get-ChildItem -LiteralPath $TargetDir -Force | Select-Object -First 1)) {
    [System.IO.Directory]::Delete($TargetDir)
}

if ($PurgeSettings) {
    $settingsPath = "HKCU:\Software\VB and VBA Program Settings\PptMathFontAddin\Settings"
    if (Test-Path -LiteralPath $settingsPath) {
        $valueNames = (Get-Item -LiteralPath $settingsPath).GetValueNames()
        foreach ($valueName in @("TargetFont", "AutoEnabled")) {
            if ($valueNames -contains $valueName) {
                Remove-ItemProperty -LiteralPath $settingsPath -Name $valueName
            }
        }
    }
}

Write-Host ""
Write-Host "Uninstalled."
Write-Host "Press Enter to close."
[void](Read-Host)
