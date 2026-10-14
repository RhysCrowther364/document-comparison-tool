#requires -Version 5.1
[CmdletBinding()]
param([string]$Executable)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing,System.IO.Compression,System.IO.Compression.FileSystem
$project = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
if ([string]::IsNullOrWhiteSpace($Executable)) { $Executable = Join-Path $project 'dist\LocalDocumentCompare\DocumentCompare.exe' }
$fixtureRoot = Join-Path $PSScriptRoot ('artifacts\ep-pdf-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $fixtureRoot -Force | Out-Null
$ascii = [Text.Encoding]::ASCII
$utf8 = New-Object Text.UTF8Encoding($false)

function Assert-True([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}

function Write-PdfAscii([IO.Stream]$Stream, [string]$Value) {
    $bytes = $ascii.GetBytes($Value)
    $Stream.Write($bytes, 0, $bytes.Length)
}

function New-SyntheticDiagram {
    $bitmap = New-Object Drawing.Bitmap(600, 240)
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    $font = New-Object Drawing.Font('Arial', 24)
    $jpeg = New-Object IO.MemoryStream
    try {
        $graphics.Clear([Drawing.Color]::White)
        $graphics.DrawRectangle([Drawing.Pens]::Black, 30, 30, 220, 100)
        $graphics.DrawLine([Drawing.Pens]::Black, 250, 80, 400, 80)
        $graphics.DrawString('Synthetic diagram', $font, [Drawing.Brushes]::Black, [single]30, [single]160)
        $bitmap.Save($jpeg, [Drawing.Imaging.ImageFormat]::Jpeg)
        return ,$jpeg.ToArray()
    } finally {
        $jpeg.Dispose()
        $font.Dispose()
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

function New-EpPdf([string]$Path, [bool]$Revised, [bool]$ImageInDescription) {
    # Build a native text PDF directly, with widely separated text blocks so
    # Word can retain standalone section headings. All fixture text is synthetic.
    $description = if ($Revised) { 'Alpha bright beta gamma delta.' } else { 'Alpha beta gamma delta.' }
    $claim = if ($Revised) { '1. A modified device comprising an element.' } else { '1. A device comprising an element.' }
    $streams = @(
        "BT /F1 24 Tf 50 700 Td (Synthetic cover sheet) Tj /F1 12 Tf 0 -80 Td (Cover content must be excluded.) Tj ET`n",
        ('BT /F1 24 Tf 50 700 Td (Description) Tj /F1 12 Tf 0 -80 Td (' + $description + ") Tj 0 -80 Td (Keep this stable paragraph.) Tj ET`n"),
        ('BT /F1 24 Tf 50 700 Td (Claims) Tj /F1 12 Tf 0 -80 Td (' + $claim + ") Tj ET`n"),
        "BT /F1 24 Tf 50 700 Td (Drawings) Tj /F1 12 Tf 0 -80 Td (Excluded drawing caption.) Tj ET`n"
    )
    $image = New-SyntheticDiagram
    $imagePage = if ($ImageInDescription) { 1 } else { 3 }
    $streams[$imagePage] += "q 480 0 0 192 50 280 cm /Diagram Do Q`n"
    $pdf = New-Object IO.MemoryStream
    $offsets = New-Object 'Collections.Generic.List[long]'
    try {
        Write-PdfAscii $pdf "%PDF-1.4`n"
        $objects = @(
            '<< /Type /Catalog /Pages 2 0 R >>',
            '<< /Type /Pages /Kids [3 0 R 5 0 R 7 0 R 9 0 R] /Count 4 >>'
        )
        for ($page = 0; $page -lt 4; $page++) {
            $contentNumber = 4 + 2 * $page
            $resources = '<< /Font << /F1 11 0 R >>'
            if ($page -eq $imagePage) { $resources += ' /XObject << /Diagram 12 0 R >>' }
            $resources += ' >>'
            $objects += ('<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources ' + $resources + ' /Contents ' + $contentNumber + ' 0 R >>')
            $objects += ('<< /Length ' + $ascii.GetByteCount($streams[$page]) + " >>`nstream`n" + $streams[$page] + 'endstream')
        }
        $objects += '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'
        for ($i = 0; $i -lt $objects.Count; $i++) {
            $offsets.Add($pdf.Position)
            Write-PdfAscii $pdf (($i + 1).ToString() + " 0 obj`n" + $objects[$i] + "`nendobj`n")
        }
        $offsets.Add($pdf.Position)
        Write-PdfAscii $pdf ('12 0 obj << /Type /XObject /Subtype /Image /Width 600 /Height 240 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ' + $image.Length + " >>`nstream`n")
        $pdf.Write($image, 0, $image.Length)
        Write-PdfAscii $pdf "`nendstream`nendobj`n"
        $xref = $pdf.Position
        Write-PdfAscii $pdf "xref`n0 13`n0000000000 65535 f `n"
        foreach ($offset in $offsets) { Write-PdfAscii $pdf ($offset.ToString('D10') + " 00000 n `n") }
        Write-PdfAscii $pdf ("trailer`n<< /Size 13 /Root 1 0 R >>`nstartxref`n" + $xref + "`n%%EOF`n")
        [IO.File]::WriteAllBytes($Path, $pdf.ToArray())
    } finally { $pdf.Dispose() }
}

function Assert-OwnedWordStopped([string]$Job) {
    $recordPath = Join-Path $Job 'word.process'
    Assert-True (Test-Path -LiteralPath $recordPath) 'No private Word process record was written'
    $record = [IO.File]::ReadAllText($recordPath) | ConvertFrom-Json
    $process = Get-Process -Id $record.Id -ErrorAction SilentlyContinue
    if ($null -ne $process) {
        try {
            $sameProcess = $process.ProcessName -eq 'WINWORD' -and $process.StartTime.ToUniversalTime().Ticks -eq $record.StartTimeUtcTicks
            Assert-True (-not $sameProcess) ('Private Word process remains after the test: ' + $record.Id)
        } finally { $process.Dispose() }
    }
}

function Invoke-EpPair([string]$Name, [string]$ArchivePath, [string]$NewPath, [bool]$ExpectSuccess) {
    $caseRoot = Join-Path $fixtureRoot $Name
    $job = Join-Path $caseRoot 'job'
    New-Item -ItemType Directory -Path $job -Force | Out-Null
    $requestPath = Join-Path $job 'request.json'
    $resultPath = Join-Path $job 'result.json'
    $outputPath = Join-Path $caseRoot 'New vs Archive.docx'
    $archiveHash = (Get-FileHash -LiteralPath $ArchivePath -Algorithm SHA256).Hash
    $newHash = (Get-FileHash -LiteralPath $NewPath -Algorithm SHA256).Hash
    $request = @{ ArchivePath=$ArchivePath; NewPath=$NewPath; OutputPath=$outputPath; WorkDirectory=$job; OcrArchive=$false; OcrNew=$false; JobType='EP'; MatchingLanguageConfirmed=$true }
    [IO.File]::WriteAllText($requestPath, ($request | ConvertTo-Json), $utf8)
    $process = Start-Process -FilePath $Executable -ArgumentList ('--worker "' + $requestPath + '" "' + $resultPath + '"') -WindowStyle Hidden -PassThru
    try {
        $started = [DateTime]::UtcNow
        while (-not $process.WaitForExit(15000)) {
            if (([DateTime]::UtcNow - $started).TotalMinutes -ge 3) { throw ('EP PDF comparison timed out. Process ownership records are in ' + $job) }
        }
        Assert-True (Test-Path -LiteralPath $resultPath) ($Name + ': worker returned no result')
        $result = [IO.File]::ReadAllText($resultPath) | ConvertFrom-Json
        Assert-True ($result.Success -eq $ExpectSuccess) ($Name + ': ' + $result.Error)
        Assert-True ((Get-FileHash -LiteralPath $ArchivePath -Algorithm SHA256).Hash -eq $archiveHash) ($Name + ': archive PDF changed')
        Assert-True ((Get-FileHash -LiteralPath $NewPath -Algorithm SHA256).Hash -eq $newHash) ($Name + ': new PDF changed')
        Assert-OwnedWordStopped $job
        if (-not $ExpectSuccess) {
            Assert-True ($result.Error -match 'OCR') ($Name + ': retained images must instruct local OCR')
            Assert-True (-not (Test-Path -LiteralPath $outputPath)) ($Name + ': failed comparison left an output')
            Assert-True (-not (Test-Path -LiteralPath ([IO.Path]::ChangeExtension($outputPath, '.txt')))) ($Name + ': failed comparison left a summary')
        }
        Write-Host ('PASS ' + $Name + ': success=' + $result.Success)
        return $result
    } finally { $process.Dispose() }
}

$archivePdf = Join-Path $fixtureRoot 'Archive.pdf'
$newPdf = Join-Path $fixtureRoot 'New.pdf'
New-EpPdf $archivePdf $false $false
New-EpPdf $newPdf $true $false
$result = Invoke-EpPair 'excluded-drawing-image' $archivePdf $newPdf $true
Assert-True ($result.JobType -eq 'EP' -and $result.Sections.Count -eq 3) 'EP result must have three scope reports'
$expectedScopes = @('Description', 'Full specification', 'Claims')
$expectedWords = @(9, 16, 7)
$expectedInsertions = @(1, 2, 1)
for ($i = 0; $i -lt 3; $i++) {
    $section = $result.Sections[$i]
    Assert-True ($section.Scope -eq $expectedScopes[$i]) 'EP scope order is incorrect'
    Assert-True ($section.NewWords -eq $expectedWords[$i]) ($section.Scope + ': incorrect PDF word count: ' + $section.NewWords)
    Assert-True ($section.InsertedWords -eq $expectedInsertions[$i]) ($section.Scope + ': incorrect highlighted insertion count: ' + $section.InsertedWords)
    $expectedCrossover = 100.0 * ($expectedWords[$i] - $expectedInsertions[$i]) / $expectedWords[$i]
    Assert-True ($null -ne $section.CrossoverPercent -and [Math]::Abs($section.CrossoverPercent - $expectedCrossover) -lt 0.01) ($section.Scope + ': incorrect crossover')
}
$summary = [IO.File]::ReadAllText($result.SummaryPath)
foreach ($scope in $expectedScopes) { Assert-True ($summary -match ('(?m)^' + [regex]::Escape($scope) + '\r?$')) ('Missing summary scope: ' + $scope) }
Assert-True ([regex]::Matches($summary, '(?m)^Crossover: \d+\.\d%\r?$').Count -eq 3) 'Summary must report all three crossover percentages to one decimal place'
Assert-True ([regex]::Matches($summary, '(?m)^Total word count of new file: ').Count -eq 3) 'Summary must report all three new word counts'
Assert-True ([regex]::Matches($summary, '(?m)^Highlighted insertion word count: ').Count -eq 3) 'Summary must report all three insertion word counts'
$zip = [IO.Compression.ZipFile]::OpenRead($result.OutputPath)
try {
    $reader = New-Object IO.StreamReader($zip.GetEntry('word/document.xml').Open())
    try { $xml = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
    $ns = New-Object Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    $outputText = ($xml.SelectNodes('//w:t | //w:delText', $ns) | ForEach-Object InnerText) -join ' '
    Assert-True ($outputText -notmatch 'Synthetic cover|Cover content|Drawings|Excluded drawing caption|Synthetic diagram') 'Excluded PDF content leaked into the EP comparison'
    Assert-True ($xml.SelectNodes('//w:drawing | //w:pict', $ns).Count -eq 0) 'Excluded drawing image remains in the EP comparison body'
    Assert-True ($xml.SelectNodes('//w:ins//w:r[w:t]/w:rPr/w:highlight[@w:val="yellow"]', $ns).Count -ge 2) 'PDF insertions are not highlighted yellow'
} finally { $zip.Dispose() }

$scannedDescriptionPdf = Join-Path $fixtureRoot 'Description-image.pdf'
New-EpPdf $scannedDescriptionPdf $false $true
$null = Invoke-EpPair 'retained-description-image' $scannedDescriptionPdf $newPdf $false
Write-Host ('PASS EP PDF scoping, retained-image OCR refusal, unchanged sources and private Word cleanup. Artifacts: ' + $fixtureRoot)
