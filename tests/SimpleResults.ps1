$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Integration.ps1') -DefineOnly
$output = Run-Case 'simple-results' @('Alpha beta gamma.', 'Keep this stable paragraph.', 'Old ending.') @('Alpha bright beta gamma.', 'Keep this stable paragraph.', 'Fresh ending.') 80 10 2
$summary = @(Get-Content -LiteralPath ([IO.Path]::ChangeExtension($output, '.txt')) -Encoding UTF8)
Assert-True ($summary.Count -ge 3) 'The comparison must contain the three summary figures'
Assert-True ($summary[0] -eq 'Crossover: 80.0%') 'Summary crossover is incorrect'
Assert-True ($summary[1] -eq 'Total word count of new file: 10') 'Summary new-file count is incorrect'
Assert-True ($summary[2] -eq 'Highlighted insertion word count: 2') 'Summary insertion count is incorrect'
Assert-True ($summary -match 'Source language was confirmed manually') 'Short-fixture manual language check must be recorded'
Write-Host 'PASS summary contains the three requested figures and the manual language review note.'
