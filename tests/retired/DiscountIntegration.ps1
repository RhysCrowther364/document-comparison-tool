$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Integration.ps1') -DefineOnly
$output = Run-Case 'discount' @('Alpha beta gamma.', 'Keep this stable paragraph.', 'Old ending.') @('Alpha bright beta gamma.', 'Keep this stable paragraph.', 'Fresh ending.') 80 10 2
$summary = Get-Content -LiteralPath ([IO.Path]::ChangeExtension($output, '.txt')) -Raw -Encoding UTF8
Assert-True ($summary -match '(?m)^Discount: 25%\r?$') 'Summary should display 25% for 2 inserted words out of 10'
$resultPath = Join-Path (Split-Path -Parent $output) 'job\result.json'
$result = Get-Content -LiteralPath $resultPath -Raw -Encoding UTF8 | ConvertFrom-Json
Assert-True ($result.DiscountPercent -eq 25) 'Worker must send 25% discount to the UI'
Write-Host 'PASS actual Word comparison reports 25% discount in worker result and saved summary.'
