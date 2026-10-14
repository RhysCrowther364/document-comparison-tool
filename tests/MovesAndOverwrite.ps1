$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Integration.ps1') -DefineOnly
$opening = 'The opening section is unchanged and provides a clear starting point for this comparison.'
$moved = 'This entire paragraph describes the earlier archive process and will be moved to a different position without changing any of its original words.'
$anchor = 'The next section contains detailed information about project delivery schedules, review responsibilities, expected turnaround times, and the local document storage arrangements used by the team.'
$closing = 'The final fixed section explains the approval process in detail. Each reviewer checks the document carefully before the final version is released for use by the wider team.'
$newParagraphs = @($opening, $anchor, $closing, $moved)
$wordCount = (($newParagraphs -join ' ') -split '\s+').Count
$output = Run-Case 'moved-paragraph' @($opening, $moved, $anchor, $closing) $newParagraphs 100 $wordCount 0
$zip = [IO.Compression.ZipFile]::OpenRead($output)
try {
    $reader = New-Object IO.StreamReader($zip.GetEntry('word/document.xml').Open())
    try { $xml = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
    $ns = New-Object Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    Assert-True ($xml.SelectNodes('//w:moveFrom', $ns).Count -gt 0) 'Native move origin missing'
    Assert-True ($xml.SelectNodes('//w:moveTo', $ns).Count -gt 0) 'Native move destination missing'
} finally { $zip.Dispose() }
Write-Host 'PASS native Word move markup preserved and counted as retained text.'
$before = (Get-FileHash -LiteralPath $output).Hash
$caseRoot = Split-Path -Parent $output
$requestPath = Join-Path $caseRoot 'job\request.json'
$secondResult = Join-Path $caseRoot 'job\second-result.json'
$process = Start-Process -FilePath $exe -ArgumentList ('--worker "' + $requestPath + '" "' + $secondResult + '"') -WindowStyle Hidden -PassThru
Assert-True ($process.WaitForExit(15000)) 'No-overwrite check timed out'
$result = Get-Content -LiteralPath $secondResult -Raw -Encoding UTF8 | ConvertFrom-Json
Assert-True (-not $result.Success -and $result.Error -match 'already exists') 'Existing output was not refused'
Assert-True ((Get-FileHash -LiteralPath $output).Hash -eq $before) 'Existing output changed'
Write-Host 'PASS existing output refused without modification.'
