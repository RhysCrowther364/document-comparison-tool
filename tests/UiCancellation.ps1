$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Integration.ps1') -DefineOnly
Add-Type -AssemblyName System.Windows.Forms
$assembly = [Reflection.Assembly]::LoadFrom($exe)
$type = $assembly.GetType('LocalDocumentCompare.MainForm', $true)
$flags = [Reflection.BindingFlags]'Instance,NonPublic'
$form = [Activator]::CreateInstance($type)
$output = Join-Path $fixtures 'cancel-result'
New-Item -ItemType Directory -Path $output -Force | Out-Null
$archivePath = Join-Path $fixtures 'Archive.docx'
$newPath = Join-Path $fixtures 'New.docx'
New-Docx $archivePath @('Alpha beta gamma.', 'A stable paragraph to compare.')
New-Docx $newPath @('Alpha bright beta gamma.', 'A stable paragraph to compare.')
try {
    $null = $form.Handle
    $type.GetField('archive', $flags).GetValue($form).Text = $archivePath
    $type.GetField('revised', $flags).GetValue($form).Text = $newPath
    $type.GetField('destination', $flags).GetValue($form).Text = $output
    $null = $type.GetMethod('StartComparison', $flags).Invoke($form, @($null, [EventArgs]::Empty))
    $job = $type.GetField('jobDirectory', $flags).GetValue($form)
    Assert-True (-not [string]::IsNullOrEmpty($job)) 'UI did not start a job'
    $watch = [Diagnostics.Stopwatch]::StartNew()
    while (-not (Test-Path -LiteralPath (Join-Path $job 'word.process')) -and $watch.Elapsed.TotalSeconds -lt 40) {
        [Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 50
    }
    Assert-True (Test-Path -LiteralPath (Join-Path $job 'word.process')) 'Word ownership record missing'
    $record = Get-Content -LiteralPath (Join-Path $job 'word.process') -Raw | ConvertFrom-Json
    $stopped = $type.GetMethod('CancelComparison', $flags).Invoke($form, @())
    Assert-True $stopped 'UI could not stop worker'
    Assert-True ($null -eq $type.GetField('worker', $flags).GetValue($form)) 'UI worker still present'
    Assert-True (-not (Test-Path -LiteralPath $job)) 'Job directory was not cleaned'
    Assert-True ($null -eq (Get-Process -Id $record.Id -ErrorAction SilentlyContinue)) 'Owned Word process survived cancellation'
    Assert-True $type.GetField('compare', $flags).GetValue($form).Enabled 'Compare button not restored'
    Write-Host 'PASS UI cancellation stops its worker and owned Word, cleans temporary files, and restores controls.'
} finally { $form.Dispose() }
