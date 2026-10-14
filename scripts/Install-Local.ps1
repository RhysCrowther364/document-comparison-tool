param([switch]$Uninstall)
$ErrorActionPreference = 'Stop'
$sourceRoot = Split-Path -Parent $PSScriptRoot
$appRoot = Join-Path $env:LOCALAPPDATA 'LocalDocumentCompare'
$desktopPath = [Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktopPath 'Document Compare.lnk'
$managedNames = @('DocumentCompare.exe', 'DocumentCompare.exe.config', 'DocumentCompare-icon-1.0.8.ico', 'README.md', 'VALIDATION.md', 'Install.cmd', 'Uninstall.cmd', 'DocumentCompare-QuickGuide.docx')
$managedScripts = @('Convert-PdfLocal.ps1', 'Install-Local.ps1')
$managedTests = @('AcceptanceChecklist.md', 'PilotTestPlan.md', 'PilotTestPlan.docx')
$managedFixtures = @('README.txt', 'WF-four-archive.docx', 'WF-four-identical.docx', 'WF-five-added.docx',
    'WF-three-deleted.docx', 'WF-three-archive.docx', 'WF-three-unrelated.docx', 'WF-four-formatting.docx',
    'WF-move-archive.docx', 'WF-move-new.docx', 'EP-prepared-archive.docx', 'EP-prepared-identical.docx',
    'EP-uncropped-archive.docx', 'EP-description-changed.docx', 'EP-claims-changed.docx', 'EP-both-changed.docx')

try {
    $resolvedApp = [IO.Path]::GetFullPath($appRoot)
    $expectedApp = [IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'LocalDocumentCompare'))
    if ($resolvedApp -ne $expectedApp) { throw 'The installation path is invalid.' }
    if ($Uninstall) {
        if (Test-Path -LiteralPath $shortcutPath) {
            $shell = New-Object -ComObject WScript.Shell
            $shortcut = $shell.CreateShortcut($shortcutPath)
            if ($shortcut.TargetPath -eq (Join-Path $appRoot 'DocumentCompare.exe')) {
                Remove-Item -LiteralPath $shortcutPath -Force
            }
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($shell)
        }
        # Remove only files installed by this package. Results and other files stay.
        foreach ($name in $managedNames) {
            $filePath = Join-Path $appRoot $name
            if (Test-Path -LiteralPath $filePath) { Remove-Item -LiteralPath $filePath -Force }
        }
        foreach ($name in $managedScripts) {
            $filePath = Join-Path (Join-Path $appRoot 'scripts') $name
            if (Test-Path -LiteralPath $filePath) { Remove-Item -LiteralPath $filePath -Force }
        }
        foreach ($name in $managedTests) {
            $filePath = Join-Path (Join-Path $appRoot 'tests') $name
            if (Test-Path -LiteralPath $filePath) { Remove-Item -LiteralPath $filePath -Force }
        }
        foreach ($name in $managedFixtures) {
            $filePath = Join-Path (Join-Path $appRoot 'tests\PilotFixtures') $name
            if (Test-Path -LiteralPath $filePath) { Remove-Item -LiteralPath $filePath -Force }
        }
        Write-Host 'Document Compare removed. Saved results have been kept.'
    } else {
        $executable = Join-Path $sourceRoot 'DocumentCompare.exe'
        if (-not (Test-Path -LiteralPath $executable)) { throw 'Extract the complete package before installing.' }
        New-Item -ItemType Directory -Path $appRoot -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $appRoot 'scripts') -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $appRoot 'tests') -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $appRoot 'tests\PilotFixtures') -Force | Out-Null
        if ([IO.Path]::GetFullPath($sourceRoot) -ne $resolvedApp) {
            foreach ($name in $managedNames) { Copy-Item -LiteralPath (Join-Path $sourceRoot $name) -Destination (Join-Path $appRoot $name) -Force }
            foreach ($name in $managedScripts) { Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination (Join-Path (Join-Path $appRoot 'scripts') $name) -Force }
            foreach ($name in $managedTests) {
                Copy-Item -LiteralPath (Join-Path (Join-Path $sourceRoot 'tests') $name) -Destination (Join-Path $appRoot 'tests') -Force
            }
            # The editable log may contain a user's pilot results: preserve it on
            # updates and uninstall. The release ZIP always includes a fresh template.
            $pilotLogPath = Join-Path $appRoot 'tests\PilotResults.csv'
            if (-not (Test-Path -LiteralPath $pilotLogPath)) {
                Copy-Item -LiteralPath (Join-Path $sourceRoot 'tests\PilotResults.csv') -Destination $pilotLogPath
            }
            foreach ($name in $managedFixtures) {
                Copy-Item -LiteralPath (Join-Path (Join-Path $sourceRoot 'tests\PilotFixtures') $name) -Destination (Join-Path $appRoot 'tests\PilotFixtures') -Force
            }
        }
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = Join-Path $appRoot 'DocumentCompare.exe'
        $shortcut.IconLocation = (Join-Path $appRoot 'DocumentCompare-icon-1.0.8.ico') + ',0'
        $shortcut.WorkingDirectory = $appRoot
        $shortcut.Description = 'RWS | Compare archive and new documents locally with Microsoft Word'
        $shortcut.Save()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($shell)
        Write-Host 'Installed. Open Document Compare using the desktop shortcut.'
    }
} catch {
    Write-Host ('Setup could not finish: ' + $_.Exception.Message)
    exit 1
}
