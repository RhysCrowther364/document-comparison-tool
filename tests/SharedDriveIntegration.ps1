param(
    [Parameter(Mandatory = $true)][string]$ShareRoot,
    [string]$AlternateShareRoot,
    [switch]$SkipBuild,
    [switch]$KeepArtifacts
)
$ErrorActionPreference = 'Stop'
$project = Split-Path -Parent $PSScriptRoot
$fullRoot = [IO.Path]::GetFullPath($ShareRoot)
if (-not [IO.Path]::IsPathRooted($ShareRoot) -or -not (Test-Path -LiteralPath $fullRoot -PathType Container)) {
    throw 'ShareRoot must be an existing, approved, writable File Explorer folder. This script does not create shares or drive mappings.'
}
if (-not $SkipBuild) { & (Join-Path $project 'build.ps1') }
$release = Join-Path $project 'dist\LocalDocumentCompare\DocumentCompare.exe'
if (-not (Test-Path -LiteralPath $release -PathType Leaf)) { throw 'Build DocumentCompare.exe first.' }
$id = [guid]::NewGuid().ToString('N')
$owned = Join-Path $fullRoot ('Compare-test-' + $id.Substring(0, 10))
$localOwned = Join-Path ([IO.Path]::GetTempPath()) ('LocalDocumentCompare\shared-test-' + $id)
$artifactRoot = Join-Path $PSScriptRoot ('artifacts\shared-' + $id)
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path -LiteralPath $compiler)) { $compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe' }
New-Item -ItemType Directory -Path $artifactRoot -Force | Out-Null
$helper = Join-Path $artifactRoot 'SharedDriveIntegration.exe'
$compilerArguments = @('/nologo', '/target:exe', '/platform:anycpu', '/utf8output', ('/out:' + $helper))
foreach ($ref in @('System.dll', 'System.Core.dll', 'System.Drawing.dll', 'System.Windows.Forms.dll', 'System.Web.Extensions.dll', 'System.Xml.dll', 'System.Xml.Linq.dll', 'System.IO.Compression.dll', 'System.IO.Compression.FileSystem.dll', $release)) {
    $compilerArguments += '/reference:' + $ref
}
& $compiler @compilerArguments (Join-Path $PSScriptRoot 'SharedDriveIntegration.cs')
if ($LASTEXITCODE -ne 0) { throw 'Shared-drive helper compilation failed.' }
if ((Test-Path -LiteralPath $owned) -or (Test-Path -LiteralPath $localOwned)) { throw 'Unique test folders already exist; nothing was changed.' }
New-Item -ItemType Directory -Path $owned | Out-Null
New-Item -ItemType Directory -Path $localOwned | Out-Null
$marker = Join-Path $owned '.synthetic-test-owner'
[IO.File]::WriteAllText($marker, $id)
$alternateOwned = ''
if (-not [String]::IsNullOrWhiteSpace($AlternateShareRoot)) {
    $alternateRoot = [IO.Path]::GetFullPath($AlternateShareRoot)
    if (-not (Test-Path -LiteralPath $alternateRoot -PathType Container)) { throw 'AlternateShareRoot must be an existing alias of the approved ShareRoot.' }
    $alternateOwned = Join-Path $alternateRoot ([IO.Path]::GetFileName($owned))
    if (-not (Test-Path -LiteralPath (Join-Path $alternateOwned '.synthetic-test-owner') -PathType Leaf) -or
        [IO.File]::ReadAllText((Join-Path $alternateOwned '.synthetic-test-owner')) -ne $id) {
        throw 'AlternateShareRoot does not resolve to the same newly created owned test folder.'
    }
}
$passed = $false
try {
    # The helper runs directly on this PC, with a hidden form and private Word
    # workers. It writes only synthetic documents below the new owned folder.
    $log = Join-Path $artifactRoot 'result.log'
    $errorLog = Join-Path $artifactRoot 'error.log'
    $arguments = '"' + $release + '" "' + $owned + '" "' + $localOwned + '"'
    if ($alternateOwned) { $arguments += ' "' + $alternateOwned + '"' }
    $running = Start-Process -FilePath $helper -ArgumentList $arguments -WindowStyle Hidden -PassThru -RedirectStandardOutput $log -RedirectStandardError $errorLog
    $identityTicks = $running.StartTime.ToUniversalTime().Ticks
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $linesShown = 0
    while (-not $running.WaitForExit(1000)) {
        if ($timer.Elapsed.TotalMinutes -gt 20) {
            $current = Get-Process -Id $running.Id -ErrorAction SilentlyContinue
            if ($current -and $current.ProcessName -eq 'SharedDriveIntegration' -and $current.StartTime.ToUniversalTime().Ticks -eq $identityTicks) {
                Stop-Process -InputObject $current -Force
            }
            throw ('Synthetic helper exceeded 20 minutes; exact owned helper stopped; inspect its recorded job processes. Artifacts: ' + $artifactRoot)
        }
        $lines = @(Get-Content -LiteralPath $log)
        while ($linesShown -lt $lines.Count) { Write-Output $lines[$linesShown]; $linesShown++ }
    }
    $lines = @(Get-Content -LiteralPath $log)
    while ($linesShown -lt $lines.Count) { Write-Output $lines[$linesShown]; $linesShown++ }
    if ((Get-Item -LiteralPath $errorLog).Length -gt 0) { Get-Content -LiteralPath $errorLog }
    if ($running.ExitCode -ne 0) { throw ('Shared-drive integration failed. Synthetic test folder: ' + $owned + '; log: ' + $artifactRoot) }
    $passed = $true
    Write-Output ('PASS shared-location test completed for ' + $fullRoot + '; log: ' + $artifactRoot)
} finally {
    if ($passed -and -not $KeepArtifacts) {
        # Validate both absolute ownership boundaries before any recursive removal.
        $shareBoundary = $fullRoot.TrimEnd('\') + '\'
        $ownedResolved = [IO.Path]::GetFullPath($owned)
        $tempBoundary = [IO.Path]::GetFullPath((Join-Path ([IO.Path]::GetTempPath()) 'LocalDocumentCompare')).TrimEnd('\') + '\'
        $localResolved = [IO.Path]::GetFullPath($localOwned)
        if (-not $ownedResolved.StartsWith($shareBoundary, [StringComparison]::OrdinalIgnoreCase) -or
            -not $localResolved.StartsWith($tempBoundary, [StringComparison]::OrdinalIgnoreCase) -or
            [IO.File]::ReadAllText($marker) -ne $id) { throw 'Owned test cleanup boundary check failed; artifacts retained.' }
        Remove-Item -LiteralPath $ownedResolved -Recurse -Force
        Remove-Item -LiteralPath $localResolved -Recurse -Force
        Write-Output 'Removed only the unique owned synthetic shared folder and its private local test work. Local log retained.'
    } elseif (-not $passed) {
        Write-Output ('Failed test artifacts retained for diagnosis: ' + $owned + ' and ' + $localOwned)
    }
}
