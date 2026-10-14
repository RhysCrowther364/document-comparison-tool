$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Integration.ps1') -DefineOnly
$retained = (@('retained') * 500) -join ' '
$inserted = (@('addition') * 625) -join ' '
$output = Run-Case 'insertion-pages' @($retained) @($retained, $inserted) (100.0 * 500 / 1125) 1125 625
$summary = Get-Content -LiteralPath ([IO.Path]::ChangeExtension($output, '.txt')) -Raw -Encoding UTF8
Assert-True ($summary -match '(?m)^Pages of new content \(highlighted insertion words / 500\): 1\.3\r?$') 'Summary must round 625 insertion words to 1.3 pages'
Assert-True ($summary -notmatch 'Whole-page allowance') 'Obsolete rounded-up allowance remains'
$output = Run-Case 'no-insertion-pages' @($retained) @($retained) 100 500 0
$summary = Get-Content -LiteralPath ([IO.Path]::ChangeExtension($output, '.txt')) -Raw -Encoding UTF8
Assert-True ($summary -match '(?m)^Pages of new content \(highlighted insertion words / 500\): 0\.0\r?$') 'No insertions must display 0.0 pages'
Write-Host 'PASS insertion-only page estimate and one-decimal summary formatting.'
