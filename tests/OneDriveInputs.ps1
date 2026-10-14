$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Integration.ps1') -DefineOnly
$oneDriveRoot = Join-Path $fixtures 'OneDrive - Test Company'
$outputRoot = Join-Path $fixtures 'Local Results'
$workRoot = Join-Path $fixtures 'Local Work'
New-Item -ItemType Directory -Path $oneDriveRoot,$outputRoot,$workRoot -Force | Out-Null
$archivePath = Join-Path $oneDriveRoot 'Archive.docx'
$newPath = Join-Path $oneDriveRoot 'New.docx'
New-Docx $archivePath @('Alpha beta gamma.')
New-Docx $newPath @('Alpha bright beta gamma.')
$before = (Get-FileHash -LiteralPath $archivePath).Hash
$assembly = [Reflection.Assembly]::LoadFrom($exe)
$engine = $assembly.GetType('LocalDocumentCompare.ComparisonEngine', $true)
$form = $assembly.GetType('LocalDocumentCompare.MainForm', $true)
$requestType = $assembly.GetType('LocalDocumentCompare.ComparisonRequest', $true)
$flags = [Reflection.BindingFlags]'Static,NonPublic'
$request = [Activator]::CreateInstance($requestType)
$request.ArchivePath = $archivePath
$request.NewPath = $newPath
$request.OutputPath = Join-Path $outputRoot 'New vs Archive.docx'
$request.WorkDirectory = $workRoot
$oldRoots = @{}
try {
    foreach ($name in @('OneDrive','OneDriveCommercial','OneDriveConsumer')) {
        $oldRoots[$name] = [Environment]::GetEnvironmentVariable($name)
        [Environment]::SetEnvironmentVariable($name, $oneDriveRoot, 'Process')
    }
    $null = $form.GetMethod('ValidateInput', $flags).Invoke($null, @([string]$archivePath, 'archive'))
    $null = $form.GetMethod('ValidateInput', $flags).Invoke($null, @([string]$newPath, 'new'))
    $null = $engine.GetMethod('Validate', $flags).Invoke($null, @($request.PSObject.BaseObject))
    $null = $engine.GetMethod('CopyInput', $flags).Invoke($null, @([string]$archivePath, 'archive', [string]$workRoot))
    Assert-True ((Get-FileHash -LiteralPath (Join-Path $workRoot 'archive-source.docx')).Hash -eq $before) 'Working copy differs from source'
    Assert-True ((Get-FileHash -LiteralPath $archivePath).Hash -eq $before) 'OneDrive source was changed'
    $request.OutputPath = Join-Path $oneDriveRoot 'Result.docx'
    $null = $engine.GetMethod('Validate', $flags).Invoke($null, @($request.PSObject.BaseObject))
    $request.WorkDirectory = $oneDriveRoot
    $rejected = $false
    try { $null = $engine.GetMethod('Validate', $flags).Invoke($null, @($request.PSObject.BaseObject)) }
    catch { $rejected = $_.Exception.ToString().Contains('outside OneDrive') }
    Assert-True $rejected 'Private work must stay outside OneDrive'
    $null = $engine.GetMethod('RequireSourcePath', $flags).Invoke($null, @('\\server\share\archive.docx'))
    $null = $engine.GetMethod('RequireOutputPath', $flags).Invoke($null, @('\\server\share\Result.docx'))
    Write-Host 'PASS OneDrive inputs and outputs accepted; local copy preserves source; private work remains outside OneDrive; UNC storage paths accepted.'
    Write-Host 'Uses a simulated OneDrive root; no files were uploaded and live cloud downloading was not tested.'
} finally {
    foreach ($name in $oldRoots.Keys) { [Environment]::SetEnvironmentVariable($name, $oldRoots[$name], 'Process') }
}
