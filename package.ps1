$ErrorActionPreference = 'Stop'
& (Join-Path $PSScriptRoot 'build.ps1')
$release = Join-Path $PSScriptRoot 'dist\LocalDocumentCompare'
$zipPath = Join-Path $PSScriptRoot 'dist\DocumentCompare-Team.zip'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$stagedZip = $zipPath + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
$files = @(Get-ChildItem -LiteralPath $release -File -Recurse)
$zip = [IO.Compression.ZipFile]::Open($stagedZip, [IO.Compression.ZipArchiveMode]::Create)
try {
    foreach ($file in $files) {
        $relative = $file.FullName.Substring($release.Length).TrimStart('\').Replace('\', '/')
        $entry = $zip.CreateEntry(('LocalDocumentCompare/' + $relative))
        # Editors/scanners may hold write handles; do not deny those existing handles.
        $inputStream = [IO.File]::Open($file.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        try {
            $outputStream = $entry.Open()
            try { $inputStream.CopyTo($outputStream) } finally { $outputStream.Dispose() }
        } finally { $inputStream.Dispose() }
    }
} finally { $zip.Dispose() }
$checkZip = [IO.Compression.ZipFile]::OpenRead($stagedZip)
try { if ($checkZip.Entries.Count -ne $files.Count) { throw 'Package file count mismatch.' } }
finally { $checkZip.Dispose() }
Move-Item -LiteralPath $stagedZip -Destination $zipPath -Force
$hash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash
[IO.File]::WriteAllText(($zipPath + '.sha256'), ($hash + '  DocumentCompare-Team.zip' + [Environment]::NewLine))
Write-Output ('Team package: ' + $zipPath)
Write-Output ('SHA256: ' + $hash)

# A separate copy of the pilot material is useful for reviewers who already
# have the application installed. It contains only instructions and synthetic files.
$pilotZipPath = Join-Path $PSScriptRoot 'dist\PilotTestMaterials.zip'
$pilotStage = $pilotZipPath + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
$pilotFiles = @('AcceptanceChecklist.md', 'PilotTestPlan.md', 'PilotTestPlan.docx', 'PilotResults.csv') |
    ForEach-Object { Get-Item -LiteralPath (Join-Path (Join-Path $PSScriptRoot 'tests') $_) }
$pilotFiles += @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'tests\PilotFixtures') -File)
$pilotZip = [IO.Compression.ZipFile]::Open($pilotStage, [IO.Compression.ZipArchiveMode]::Create)
try {
    foreach ($file in $pilotFiles) {
        $testRoot = Join-Path $PSScriptRoot 'tests'
        $relative = $file.FullName.Substring($testRoot.Length).TrimStart('\').Replace('\', '/')
        $entry = $pilotZip.CreateEntry(('PilotTestMaterials/' + $relative))
        $inputStream = [IO.File]::Open($file.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        try {
            $outputStream = $entry.Open()
            try { $inputStream.CopyTo($outputStream) } finally { $outputStream.Dispose() }
        } finally { $inputStream.Dispose() }
    }
} finally { $pilotZip.Dispose() }
$pilotCheck = [IO.Compression.ZipFile]::OpenRead($pilotStage)
try { if ($pilotCheck.Entries.Count -ne $pilotFiles.Count) { throw 'Pilot package file count mismatch.' } }
finally { $pilotCheck.Dispose() }
Move-Item -LiteralPath $pilotStage -Destination $pilotZipPath -Force
Write-Output ('Pilot materials: ' + $pilotZipPath)
