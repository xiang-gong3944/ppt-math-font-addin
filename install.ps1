param([switch]$BuildOnly)

$ErrorActionPreference = "Stop"

$ScriptDir = $PSScriptRoot
$BasPath = Join-Path $ScriptDir "MathFontAddin.bas"
$ClsPath = Join-Path $ScriptDir "CAppEvents.cls"
$RibbonPath = Join-Path $ScriptDir "customUI.xml"
$PpamName = "PptMathFontAddin.ppam"

function Stop-WithMessage([string]$Message) {
    if ($BuildOnly) { throw $Message }
    Write-Host ""
    Write-Host $Message
    Write-Host ""
    Write-Host "Press Enter to close."
    [void](Read-Host)
    exit 1
}

if ($env:OS -ne "Windows_NT") {
    throw "This script requires desktop PowerPoint on Windows. On Mac, add PptMathFontAddin.ppam from Tools > PowerPoint Add-ins."
}

if ($BuildOnly) {
    $TargetDir = $ScriptDir
}
else {
    $TargetDir = Join-Path $env:APPDATA "Microsoft\AddIns\PptMathFontAddin"
}
$PpamPath = Join-Path $TargetDir $PpamName

if (-not (Test-Path -LiteralPath $BasPath) -or -not (Test-Path -LiteralPath $ClsPath) -or -not (Test-Path -LiteralPath $RibbonPath)) {
    Stop-WithMessage "Required source files are missing. Extract the whole ZIP before running install.ps1."
}

$running = Get-Process POWERPNT -ErrorAction SilentlyContinue
if ($running) {
    Stop-WithMessage "PowerPoint is running. Close every PowerPoint window, then run install.ps1 again."
}

function New-PpamPackage([string]$BuildPath) {
    $buildApp = $null
    $pres = $null
    try {
        Write-Host "Building PowerPoint add-in..."
        $buildApp = New-Object -ComObject PowerPoint.Application
        $buildApp.Visible = -1
        $pres = $buildApp.Presentations.Add()
        [void]$pres.Slides.Add(1, 12)

        try {
            [void]$pres.VBProject.VBComponents.Import($BasPath)
            [void]$pres.VBProject.VBComponents.Import($ClsPath)
        }
        catch {
            throw @"
PowerPoint blocked programmatic access to the VBA project.

Temporarily enable:
PowerPoint > File > Options > Trust Center > Trust Center Settings
> Macro Settings > Trust access to the VBA project object model

Then close PowerPoint and run install.ps1 again.
You can turn that checkbox off again after installation.
"@
        }

        # ppSaveAsOpenXMLAddin = 30
        $pres.SaveAs($BuildPath, 30)
    }
    finally {
        if ($null -ne $pres) {
            try { $pres.Saved = -1; $pres.Close() } catch {}
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($pres)
        }
        if ($null -ne $buildApp) {
            try { $buildApp.Quit() } catch {}
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($buildApp)
        }
    }

    Write-Host "Adding the Ribbon UI..."
    Add-RibbonToPpam $BuildPath $RibbonPath
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Save-XmlToZipEntry($Entry, [xml]$XmlDoc) {
    $stream = $Entry.Open()
    try {
        $stream.SetLength(0)
        $settings = New-Object System.Xml.XmlWriterSettings
        $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
        $settings.Indent = $false
        $writer = [System.Xml.XmlWriter]::Create($stream, $settings)
        try { $XmlDoc.Save($writer) }
        finally { $writer.Dispose() }
    }
    finally { $stream.Dispose() }
}

function Add-RibbonToPpam([string]$PackagePath, [string]$UiPath) {
    $zip = [System.IO.Compression.ZipFile]::Open($PackagePath,[System.IO.Compression.ZipArchiveMode]::Update)
    try {
        # .NET Framework's ZIP update can retain deflate bytes in an empty part
        # while marking it as uncompressed. Recreate empty parts to keep PPAM valid.
        $emptyEntries = @($zip.Entries | Where-Object { $_.Length -eq 0 })
        foreach ($emptyEntry in $emptyEntries) {
            $emptyName = $emptyEntry.FullName
            $emptyEntry.Delete()
            [void]$zip.CreateEntry($emptyName)
        }

        $old = $zip.GetEntry("customUI/customUI.xml")
        if ($old -ne $null) { $old.Delete() }

        $entry = $zip.CreateEntry("customUI/customUI.xml")
        $out = $entry.Open()
        try {
            $bytes = [System.IO.File]::ReadAllBytes($UiPath)
            $out.Write($bytes, 0, $bytes.Length)
        }
        finally { $out.Dispose() }

        $relsEntry = $zip.GetEntry("_rels/.rels")
        if ($relsEntry -eq $null) { throw "The generated PPAM has no _rels/.rels file." }

        $relsStream = $relsEntry.Open()
        try {
            $relsDoc = New-Object System.Xml.XmlDocument
            $relsDoc.PreserveWhitespace = $true
            $relsDoc.Load($relsStream)
        }
        finally { $relsStream.Dispose() }

        $relsNs = $relsDoc.DocumentElement.NamespaceURI
        $existing = $null
        foreach ($node in $relsDoc.DocumentElement.ChildNodes) {
            if ($node.LocalName -eq "Relationship" -and $node.GetAttribute("Target") -eq "customUI/customUI.xml") {
                $existing = $node
                break
            }
        }

        if ($existing -eq $null) {
            $rel = $relsDoc.CreateElement("Relationship", $relsNs)
            $rel.SetAttribute("Id", "rIdPptMathFontRibbon")
            $rel.SetAttribute("Type", "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility")
            $rel.SetAttribute("Target", "customUI/customUI.xml")
            [void]$relsDoc.DocumentElement.AppendChild($rel)
        }

        Save-XmlToZipEntry $relsEntry $relsDoc

        $ctEntry = $zip.GetEntry("[Content_Types].xml")
        if ($ctEntry -eq $null) { throw "The generated PPAM has no [Content_Types].xml file." }

        $ctStream = $ctEntry.Open()
        try {
            $ctDoc = New-Object System.Xml.XmlDocument
            $ctDoc.PreserveWhitespace = $true
            $ctDoc.Load($ctStream)
        }
        finally { $ctStream.Dispose() }

        $ctNs = $ctDoc.DocumentElement.NamespaceURI
        $hasOverride = $false
        foreach ($node in $ctDoc.DocumentElement.ChildNodes) {
            if ($node.LocalName -eq "Override" -and $node.GetAttribute("PartName") -eq "/customUI/customUI.xml") {
                $node.SetAttribute("ContentType", "application/xml")
                $hasOverride = $true
                break
            }
        }

        if (-not $hasOverride) {
            $ov = $ctDoc.CreateElement("Override", $ctNs)
            $ov.SetAttribute("PartName", "/customUI/customUI.xml")
            $ov.SetAttribute("ContentType", "application/xml")
            [void]$ctDoc.DocumentElement.AppendChild($ov)
        }

        Save-XmlToZipEntry $ctEntry $ctDoc
    }
    finally { $zip.Dispose() }
}

$buildDir = Join-Path ([System.IO.Path]::GetTempPath()) ("ppt-math-build-" + [guid]::NewGuid().ToString("N"))
# PowerPoint can reject SaveAs if an add-in with that name is already loaded.
$buildPath = Join-Path $buildDir "PptMathFontBuild.ppam"
$buildError = $null
try {
    [void](New-Item -ItemType Directory -Path $buildDir)
    New-PpamPackage $buildPath
    if (-not $BuildOnly) {
        # The previous add-in can remain locked until the build's PowerPoint process exits.
        $shutdownDeadline = [DateTime]::UtcNow.AddSeconds(10)
        while (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
            if ([DateTime]::UtcNow -ge $shutdownDeadline) {
                throw "PowerPoint is still running. Close it before replacing the installed add-in."
            }
            Start-Sleep -Milliseconds 100
        }
    }
    [void](New-Item -ItemType Directory -Force -Path $TargetDir)
    Copy-Item -LiteralPath $buildPath -Destination $PpamPath -Force
}
catch {
    $buildError = $_.Exception.Message
}
finally {
    Remove-Item -LiteralPath $buildPath -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $buildDir -ErrorAction SilentlyContinue
}
if ($buildError) { Stop-WithMessage $buildError }

if ($BuildOnly) {
    Write-Host "Built add-in: $PpamPath"
    return
}

Write-Host "Registering the add-in..."
$ppt = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $addin = $ppt.AddIns.Add($PpamPath)
    $addin.Loaded = -1
    if ($addin.Loaded -ne -1) {
        throw "PowerPoint did not load the add-in. Open PowerPoint > File > Options > Add-ins > Manage: PowerPoint Add-ins > Go, then load the generated PPAM and check the message shown there."
    }
    $addin.AutoLoad = -1
    $ppt.Quit()
    $ppt = $null
}
catch {
    if ($ppt -ne $null) { try { $ppt.Quit() } catch {} }
    Stop-WithMessage ("The PPAM was built, but registration failed." + [Environment]::NewLine + "File: " + $PpamPath + [Environment]::NewLine + [Environment]::NewLine + $_.Exception.Message)
}

Copy-Item $BasPath (Join-Path $TargetDir "MathFontAddin.bas") -Force
Copy-Item $ClsPath (Join-Path $TargetDir "CAppEvents.cls") -Force
Copy-Item $RibbonPath (Join-Path $TargetDir "customUI.xml") -Force

Write-Host ""
Write-Host "Installed successfully."
Write-Host "Add-in: $PpamPath"
Write-Host ""
Write-Host "Open PowerPoint and look for the 'Math Font' Ribbon tab."
Write-Host "Default target font: Latin Modern Math"
Write-Host ""
Write-Host "You may now turn off 'Trust access to the VBA project object model' again."
Write-Host ""
Write-Host "Press Enter to close."
[void](Read-Host)
