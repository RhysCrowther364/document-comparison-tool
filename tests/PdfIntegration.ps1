$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Integration.ps1') -DefineOnly
$pdfRoot = Join-Path $fixtures 'pdf'
New-Item -ItemType Directory -Path $pdfRoot -Force | Out-Null

function New-TextPdf([string]$Path, [string]$Text) {
    $ascii = [Text.Encoding]::ASCII
    $stream = New-Object IO.MemoryStream
    $offsets = New-Object 'System.Collections.Generic.List[long]'
    function Write-Pdf([string]$Value) { $bytes = $ascii.GetBytes($Value); $stream.Write($bytes, 0, $bytes.Length) }
    $content = 'BT /F1 12 Tf 50 740 Td (' + $Text + ') Tj ET'
    $objects = @(
        '<< /Type /Catalog /Pages 2 0 R >>',
        '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
        '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>',
        '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
        ('<< /Length ' + $content.Length + " >>`nstream`n" + $content + "`nendstream")
    )
    try {
        Write-Pdf "%PDF-1.4`n"
        for ($i = 0; $i -lt $objects.Count; $i++) {
            $offsets.Add($stream.Position)
            Write-Pdf (($i + 1).ToString() + " 0 obj`n" + $objects[$i] + "`nendobj`n")
        }
        $xref = $stream.Position
        Write-Pdf "xref`n0 6`n0000000000 65535 f `n"
        foreach ($offset in $offsets) { Write-Pdf ($offset.ToString('D10') + " 00000 n `n") }
        Write-Pdf ("trailer`n<< /Size 6 /Root 1 0 R >>`nstartxref`n" + $xref + "`n%%EOF`n")
        [IO.File]::WriteAllBytes($Path, $stream.ToArray())
    } finally { $stream.Dispose() }
}

function Invoke-Pair([string]$Name, [string]$ArchivePath, [string]$NewPath, [bool]$OcrArchive, [bool]$OcrNew, [bool]$ExpectSuccess) {
    $job = Join-Path $pdfRoot $Name
    New-Item -ItemType Directory -Path $job -Force | Out-Null
    $requestPath = Join-Path $job 'request.json'
    $resultPath = Join-Path $job 'result.json'
    $request = @{ ArchivePath=$ArchivePath; NewPath=$NewPath; OutputPath=(Join-Path $job 'New vs Archive.docx'); WorkDirectory=$job; OcrArchive=$OcrArchive; OcrNew=$OcrNew; MatchingLanguageConfirmed=$true }
    [IO.File]::WriteAllText($requestPath, ($request | ConvertTo-Json), $utf8)
    $process = Start-Process -FilePath $exe -ArgumentList ('--worker "' + $requestPath + '" "' + $resultPath + '"') -WindowStyle Hidden -PassThru
    if (-not $process.WaitForExit(180000)) { throw ('PDF integration timeout: ' + $job) }
    $result = Get-Content -LiteralPath $resultPath -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-True ($result.Success -eq $ExpectSuccess) ($Name + ': ' + $result.Error)
    Write-Host ('PASS ' + $Name + ': success=' + $result.Success + ', crossover=' + $result.CrossoverPercent)
    return $result
}

$archivePdf = Join-Path $pdfRoot 'Archive.pdf'
$newPdf = Join-Path $pdfRoot 'New.pdf'
New-TextPdf $archivePdf 'Alpha beta gamma. Keep this stable paragraph. Old ending.'
New-TextPdf $newPdf 'Alpha bright beta gamma. Keep this stable paragraph. Fresh ending.'
$result = Invoke-Pair 'text-pdfs' $archivePdf $newPdf $false $false $true
Assert-True ($result.NewWords -eq 10 -and $result.CrossoverPercent -eq 80) 'Text PDF metrics incorrect'

& (Join-Path $PSScriptRoot 'OcrSmoke.ps1') -KeepFiles
$ocrFixture = Get-ChildItem -LiteralPath $PSScriptRoot -Directory -Filter '.ocr-smoke-*' | Sort-Object LastWriteTime -Descending | Select-Object -First 1
$mixedPdf = Join-Path $ocrFixture.FullName 'mixed.pdf'
$result = Invoke-Pair 'mixed-pdf-ocr' $mixedPdf $mixedPdf $true $true $true
Assert-True ($result.CrossoverPercent -eq 100 -and $result.NewWords -gt 0) 'Identical OCR inputs should match fully'
$result = Invoke-Pair 'scan-without-ocr' $mixedPdf $newPdf $false $false $false
Assert-True ($result.Error -match 'OCR') 'Scan omission should instruct OCR'
$result = Invoke-Pair 'unreadable-ocr-page' (Join-Path $ocrFixture.FullName 'text-and-blank.pdf') $newPdf $true $false $false
Assert-True ($result.Error -match 'without recognised words') 'Blank OCR page must prevent a misleading score'
Write-Host ('PDF integration artifacts: ' + $pdfRoot)
