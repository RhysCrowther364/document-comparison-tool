param([switch]$SkipPdf)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Integration.ps1') -DefineOnly

# Synthetic, non-confidential prose. Repeated complete passages provide enough
# independent text windows for the conservative native Windows detector.
$english = 'The present invention relates to a method for controlling a motor. The apparatus comprises a housing, a controller and a sensor. The controller receives the signal from the sensor and adjusts the speed of the motor. This arrangement improves the operation of the device. The following description explains how the components work together and how the operator can use the apparatus under different conditions. Each component performs a specific function within the system, and the controller records the result before the next operation begins.'
$german = 'Die vorliegende Erfindung betrifft ein Verfahren zur Steuerung eines Motors. Die Vorrichtung besteht aus einer Steuerung und einem Sensor. Die Steuerung verarbeitet das Signal des Sensors und stellt die Drehzahl des Motors ein. Durch diese Anordnung wird der Betrieb der Vorrichtung verbessert. Das Verfahren wird mit einer weiteren Einheit verbunden, damit der Motor unter verschiedenen Bedingungen sicher arbeiten kann. Die Beschreibung zeigt, wie die einzelnen Bauteile zusammenwirken und welche Funktionen bei der Anwendung des Verfahrens erforderlich sind.'
$french = 'La presente invention concerne un dispositif de commande pour un moteur. Le dispositif comprend un boitier, un controleur et un capteur. Le controleur recoit le signal du capteur et ajuste la vitesse du moteur. Cet agencement facilite le fonctionnement du dispositif. La description suivante explique comment les composants travaillent ensemble et comment un utilisateur peut appliquer le dispositif dans des conditions variables. Chaque composant remplit une fonction dans le systeme, et le controleur enregistre le resultat avant le commencement de la prochaine operation.'
$englishLong = (($english + ' ') * 8).Trim()
$germanLong = (($german + ' ') * 8).Trim()
$frenchLong = (($french + ' ') * 8).Trim()
$inputRoot = Join-Path $fixtures 'language-inputs'
New-Item -ItemType Directory -Path $inputRoot | Out-Null

function New-LanguageDocx([string]$Name, [string[]]$Paragraphs) {
    $path = Join-Path $inputRoot ($Name + '.docx')
    New-Docx $path $Paragraphs
    return $path
}

function Set-EnglishProofing([string]$Path) {
    $zip = [IO.Compression.ZipFile]::Open($Path, [IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry = $zip.GetEntry('word/document.xml')
        $reader = New-Object IO.StreamReader($entry.Open())
        try { $xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $xml = $xml.Replace('<w:r>', '<w:r><w:rPr><w:lang w:val="en-GB"/></w:rPr>')
        $entry.Delete()
        Add-ZipText $zip 'word/document.xml' $xml
    } finally { $zip.Dispose() }
}

function Get-OwnedLanguageProcess([string]$Job, [string]$Prefix) {
    $recordPath = Join-Path $Job ($Prefix + '.process')
    if (-not (Test-Path -LiteralPath $recordPath)) { return $null }
    $record = Get-Content -LiteralPath $recordPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $candidate = Get-Process -Id $record.Id -ErrorAction SilentlyContinue
    if ($null -ne $candidate -and $candidate.StartTime.ToUniversalTime().Ticks -eq $record.StartTimeUtcTicks) { return $candidate }
    return $null
}

function Invoke-LanguageCase {
    param(
        [string]$Name, [string]$ArchivePath, [string]$NewPath,
        [ValidateSet('Mismatch','Review','Success')][string]$Expected,
        [bool]$Confirmed = $false, [string]$JobType = 'WF', [bool]$ExpectWord = $false,
        [string[]]$Languages = @(), [string]$ExpectedErrorScope = '',
        [int]$ExpectedWords = -1, [int]$ExpectedInserted = -1,
        [double]$ExpectedCrossover = -1
    )
    $case = Join-Path $fixtures $Name
    $job = Join-Path $case 'job'
    New-Item -ItemType Directory -Path $job -Force | Out-Null
    $output = Join-Path $case 'New vs Archive.docx'
    $summary = [IO.Path]::ChangeExtension($output, '.txt')
    $requestPath = Join-Path $job 'request.json'
    $resultPath = Join-Path $job 'result.json'
    $archiveHash = (Get-FileHash -LiteralPath $ArchivePath).Hash
    $newHash = (Get-FileHash -LiteralPath $NewPath).Hash
    $request = @{
        ArchivePath = $ArchivePath; NewPath = $NewPath; OutputPath = $output
        WorkDirectory = $job; JobType = $JobType
        OcrArchive = $false; OcrNew = $false; MatchingLanguageConfirmed = $Confirmed
    }
    [IO.File]::WriteAllText($requestPath, ($request | ConvertTo-Json), $utf8)
    $worker = $null
    try {
        $worker = Start-Process -FilePath $exe -ArgumentList ('--worker "' + $requestPath + '" "' + $resultPath + '"') -WindowStyle Hidden -PassThru
        if (-not $worker.WaitForExit(120000)) { throw ($Name + ': worker timed out') }
        Assert-True (Test-Path -LiteralPath $resultPath) ($Name + ': worker did not return a result')
        $result = Get-Content -LiteralPath $resultPath -Raw -Encoding UTF8 | ConvertFrom-Json
        Assert-True ((Get-FileHash -LiteralPath $ArchivePath).Hash -eq $archiveHash) ($Name + ': archive changed')
        Assert-True ((Get-FileHash -LiteralPath $NewPath).Hash -eq $newHash) ($Name + ': new source changed')
        $hasWordRecord = Test-Path -LiteralPath (Join-Path $job 'word.process')
        Assert-True ($hasWordRecord -eq $ExpectWord) ($Name + ': unexpected Word startup/preflight behavior')
        foreach ($prefix in @('word','ocr')) {
            Assert-True ($null -eq (Get-OwnedLanguageProcess $job $prefix)) ($Name + ': owned helper did not close')
        }
        if ($Expected -eq 'Success') {
            Assert-True $result.Success ($Name + ': ' + $result.Error)
            Assert-True (-not $result.LanguageReviewRequired) ($Name + ': successful comparison still asks for language review')
            Assert-True ((Test-Path -LiteralPath $output) -and (Test-Path -LiteralPath $summary)) ($Name + ': successful outputs missing')
            if ($ExpectedWords -ge 0) { Assert-True ($result.NewWords -eq $ExpectedWords) ($Name + ': word count changed') }
            else { Assert-True ($result.NewWords -gt 200) ($Name + ': long fixture has too little measured text') }
            if ($ExpectedInserted -ge 0) { Assert-True ($result.InsertedWords -eq $ExpectedInserted) ($Name + ': insertion count changed') }
            if ($ExpectedCrossover -ge 0) {
                Assert-True ($null -ne $result.CrossoverPercent -and [Math]::Abs($result.CrossoverPercent - $ExpectedCrossover) -lt 0.01) ($Name + ': crossover differs')
            }
            if ($JobType -eq 'EP') {
                Assert-True ($result.Sections.Count -eq 3) ($Name + ': EP scopes missing')
                foreach ($section in $result.Sections) {
                    Assert-True ($section.InsertedWords -eq 0 -and $null -ne $section.CrossoverPercent -and [Math]::Abs($section.CrossoverPercent - 100) -lt 0.01) ($Name + ': translated excluded claims affected selected content')
                }
            }
        }
        else {
            Assert-True (-not $result.Success) ($Name + ': language guard did not stop comparison')
            Assert-True (-not (Test-Path -LiteralPath $output) -and -not (Test-Path -LiteralPath $summary)) ($Name + ': stopped comparison published outputs')
            Assert-True (-not (Test-Path -LiteralPath (Join-Path $job 'comparison.docx'))) ($Name + ': comparison was saved despite language stop')
            Assert-True ($null -eq $result.OutputPath -and $null -eq $result.SummaryPath) ($Name + ': stopped result advertises an output')
            Assert-True ($result.LanguageReviewRequired -eq ($Expected -eq 'Review')) ($Name + ': incorrect manual-review state')
            Assert-True ($result.Error -match '(?i)language') ($Name + ': language failure explanation missing')
            foreach ($language in $Languages) {
                Assert-True ($result.Error -match [regex]::Escape($language)) ($Name + ': detected language missing from mismatch error')
            }
            if ($ExpectedErrorScope) { Assert-True ($result.Error -match [regex]::Escape($ExpectedErrorScope)) ($Name + ': mismatch scope missing from error') }
            if (-not $ExpectWord) {
                Assert-True (-not (Test-Path -LiteralPath (Join-Path $job 'archive-final.docx'))) ($Name + ': preflight stop converted documents unnecessarily')
            }
            else {
                Assert-True ((Test-Path -LiteralPath (Join-Path $job 'archive-final.docx')) -and (Test-Path -LiteralPath (Join-Path $job 'new-final.docx'))) ($Name + ': canonical guard did not inspect converted text')
            }
        }
        Write-Host ('PASS language guard ' + $Name + ': ' + $Expected + ', Word started=' + $hasWordRecord)
        return $result
    }
    finally {
        if ($null -ne $worker) {
            if (-not $worker.HasExited) { $worker.Kill(); $null = $worker.WaitForExit(5000) }
            $worker.Dispose()
        }
        # Cleanup only processes whose ID and creation time match this test job.
        foreach ($prefix in @('word','ocr')) {
            $owned = Get-OwnedLanguageProcess $job $prefix
            if ($null -ne $owned) {
                try { $owned.Kill(); $null = $owned.WaitForExit(5000) } finally { $owned.Dispose() }
            }
        }
    }
}

$en = New-LanguageDocx 'English' @($englishLong)
$de = New-LanguageDocx 'German' @($germanLong)
$fr = New-LanguageDocx 'French' @($frenchLong)
$null = Invoke-LanguageCase 'docx-en-de' $en $de 'Mismatch' -Languages @('English','German')
$null = Invoke-LanguageCase 'docx-en-fr' $en $fr 'Mismatch' -Languages @('English','French')
$null = Invoke-LanguageCase 'confirmed-known-mismatch' $en $de 'Mismatch' -Confirmed $true -Languages @('English','German')
$falseMetadata = New-LanguageDocx 'German-with-English-proofing' @($germanLong)
Set-EnglishProofing $falseMetadata
$null = Invoke-LanguageCase 'proofing-does-not-decide' $en $falseMetadata 'Mismatch' -Languages @('English','German')
$null = Invoke-LanguageCase 'same-long-english' $en $en 'Success' -ExpectWord $true -ExpectedInserted 0 -ExpectedCrossover 100

$shortArchive = New-LanguageDocx 'Short-archive' @('Alpha beta gamma.', 'Keep this stable paragraph.', 'Old ending.')
$shortNew = New-LanguageDocx 'Short-new' @('Alpha bright beta gamma.', 'Keep this stable paragraph.', 'Fresh ending.')
$null = Invoke-LanguageCase 'short-needs-review' $shortArchive $shortNew 'Review'
$null = Invoke-LanguageCase 'confirmed-short' $shortArchive $shortNew 'Success' -Confirmed $true -ExpectWord $true -ExpectedWords 10 -ExpectedInserted 2 -ExpectedCrossover 80
$mixed = New-LanguageDocx 'Mixed' @($englishLong,$germanLong)
$null = Invoke-LanguageCase 'mixed-needs-review' $mixed $mixed 'Review'
$null = Invoke-LanguageCase 'confirmed-mixed' $mixed $mixed 'Success' -Confirmed $true -ExpectWord $true -ExpectedInserted 0 -ExpectedCrossover 100

$germanClaimsHeading = 'Anspr' + [char]0x00fc + 'che'
$epArchive = New-LanguageDocx 'EP-English-with-French-translation' @('Description',$englishLong,'Claims',('1. ' + $englishLong),'Revendications',('1. ' + $frenchLong))
$epNew = New-LanguageDocx 'EP-English-with-German-translation' @('Description',$englishLong,'Claims',('1. ' + $englishLong),$germanClaimsHeading,('1. ' + $germanLong))
$null = Invoke-LanguageCase 'ep-translated-claims-excluded' $epArchive $epNew 'Success' -JobType 'EP' -ExpectWord $true -ExpectedInserted 0 -ExpectedCrossover 100
$epDifferentClaims = New-LanguageDocx 'EP-English-description-French-original-claims' @('Description',$englishLong,'Revendications',('1. ' + $frenchLong))
$null = Invoke-LanguageCase 'ep-original-claims-mismatch' $epArchive $epDifferentClaims 'Mismatch' -JobType 'EP' -Languages @('English','French') -ExpectedErrorScope 'Claims'

function New-LanguageTextPdf([string]$Path, [string]$Text) {
    $lines = New-Object 'Collections.Generic.List[string]'
    $line = ''
    foreach ($word in ($Text -split '\s+')) {
        if ($line.Length -gt 0 -and $line.Length + $word.Length + 1 -gt 85) { $lines.Add($line); $line = '' }
        if ($line.Length -gt 0) { $line += ' ' }
        $line += $word
    }
    if ($line) { $lines.Add($line) }
    $pages = [int][Math]::Ceiling($lines.Count / 50.0)
    $objects = New-Object 'Collections.Generic.List[string]'
    $objects.Add('<< /Type /Catalog /Pages 2 0 R >>')
    $kids = ((0..($pages-1)) | ForEach-Object { (4 + 2 * $_).ToString() + ' 0 R' }) -join ' '
    $objects.Add('<< /Type /Pages /Kids [' + $kids + '] /Count ' + $pages + ' >>')
    $objects.Add('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>')
    for ($page = 0; $page -lt $pages; $page++) {
        $content = "BT /F1 10 Tf 45 750 Td 13 TL`n"
        for ($index = $page * 50; $index -lt [Math]::Min(($page+1)*50,$lines.Count); $index++) {
            $escaped = $lines[$index].Replace('\','\\').Replace('(','\(').Replace(')','\)')
            $content += '(' + $escaped + ") Tj T*`n"
        }
        $content += "ET`n"
        $objects.Add('<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 3 0 R >> >> /Contents ' + (5 + 2 * $page) + ' 0 R >>')
        $objects.Add('<< /Length ' + [Text.Encoding]::ASCII.GetByteCount($content) + " >>`nstream`n" + $content + 'endstream')
    }
    $bytes = New-Object IO.MemoryStream
    $offsets = New-Object 'Collections.Generic.List[long]'
    function Write-LanguagePdfPart([string]$Value) { $part = [Text.Encoding]::ASCII.GetBytes($Value); $bytes.Write($part,0,$part.Length) }
    try {
        Write-LanguagePdfPart "%PDF-1.4`n"
        for ($index = 0; $index -lt $objects.Count; $index++) {
            $offsets.Add($bytes.Position)
            Write-LanguagePdfPart (($index+1).ToString() + " 0 obj`n" + $objects[$index] + "`nendobj`n")
        }
        $xref = $bytes.Position
        Write-LanguagePdfPart ('xref' + "`n0 " + ($objects.Count+1) + "`n0000000000 65535 f `n")
        foreach ($offset in $offsets) { Write-LanguagePdfPart ($offset.ToString('D10') + " 00000 n `n") }
        Write-LanguagePdfPart ('trailer' + "`n<< /Size " + ($objects.Count+1) + ' /Root 1 0 R >>' + "`nstartxref`n" + $xref + "`n%%EOF`n")
        [IO.File]::WriteAllBytes($Path,$bytes.ToArray())
    } finally { $bytes.Dispose() }
}

if (-not $SkipPdf) {
    $englishPdf = Join-Path $inputRoot 'English.pdf'
    New-LanguageTextPdf $englishPdf $englishLong
    $null = Invoke-LanguageCase 'pdf-canonical-en-de' $englishPdf $de 'Mismatch' -ExpectWord $true -Languages @('English','German')
    $null = Invoke-LanguageCase 'pdf-confirmed-known-mismatch' $englishPdf $de 'Mismatch' -Confirmed $true -ExpectWord $true -Languages @('English','German')
}
Write-Host ('Language guard integration artifacts: ' + $fixtures)
