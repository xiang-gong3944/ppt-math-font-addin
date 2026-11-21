$ErrorActionPreference = "Stop"

$TargetDir = Join-Path $env:APPDATA "Microsoft\AddIns\PptMathFontAddin"
$PpamPath = Join-Path $TargetDir "PptMathFontAddin.ppam"

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

    for ($i = $ppt.AddIns.Count; $i -ge 1; $i--) {
        $addin = $ppt.AddIns.Item($i)
        $full = ""
        try { $full = $addin.FullName } catch {}

        if ($full -eq $PpamPath -or $addin.Name -eq "PptMathFontAddin.ppam") {
            try { $addin.AutoLoad = 0 } catch {}
            try { $addin.Loaded = 0 } catch {}
            try { $ppt.AddIns.Remove($i) } catch {}
        }
    }

    $ppt.Quit()
    $ppt = $null
}
catch {
    if ($ppt -ne $null) { try { $ppt.Quit() } catch {} }
    Write-Host "PowerPoint add-in registration cleanup returned:"
    Write-Host $_.Exception.Message
}

if (Test-Path $TargetDir) {
    Remove-Item $TargetDir -Recurse -Force
}

Remove-Item "HKCU:\Software\VB and VBA Program Settings\PptMathFontAddin" -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "Uninstalled."
Write-Host "Press Enter to close."
[void](Read-Host)
