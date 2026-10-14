param([switch]$IncludeSamples, [string]$SampleArchivePath, [string]$SamplePreparedPath)
$ErrorActionPreference = 'Stop'
if ($IncludeSamples -and (-not $SampleArchivePath -or -not $SamplePreparedPath)) {
    throw 'With IncludeSamples, provide SampleArchivePath and SamplePreparedPath explicitly.'
}
. (Join-Path $PSScriptRoot 'Integration.ps1') -DefineOnly

function Invoke-EpCase([string]$Name, [string]$ArchivePath, [string]$NewPath, [bool]$ExpectSuccess = $true) {
    $case = Join-Path $fixtures $Name
    New-Item -ItemType Directory -Path $case -Force | Out-Null
    $job = Join-Path $case 'job'
    New-Item -ItemType Directory -Path $job -Force | Out-Null
    $output = Join-Path $case 'New vs Archive.docx'
    $archiveHash = (Get-FileHash -LiteralPath $ArchivePath).Hash
    $newHash = (Get-FileHash -LiteralPath $NewPath).Hash
    $request = @{ JobType = 'EP'; ArchivePath = $ArchivePath; NewPath = $NewPath; OutputPath = $output; WorkDirectory = $job; MatchingLanguageConfirmed = $true }
    $requestPath = Join-Path $job 'request.json'
    $resultPath = Join-Path $job 'result.json'
    [IO.File]::WriteAllText($requestPath, ($request | ConvertTo-Json), $utf8)
    $process = Start-Process -FilePath $exe -ArgumentList ('--worker "' + $requestPath + '" "' + $resultPath + '"') -WindowStyle Hidden -PassThru
    if (-not $process.WaitForExit(120000)) { throw ('EP test timed out; owned process records: ' + $job) }
    $result = Get-Content -LiteralPath $resultPath -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-True ($result.Success -eq $ExpectSuccess) ($Name + ': ' + $result.Error)
    Assert-True ((Get-FileHash -LiteralPath $ArchivePath).Hash -eq $archiveHash) 'EP archive changed'
    Assert-True ((Get-FileHash -LiteralPath $NewPath).Hash -eq $newHash) 'EP new input changed'
    $identity = Get-Content -LiteralPath (Join-Path $job 'word.process') -Raw | ConvertFrom-Json
    Assert-True ($null -eq (Get-Process -Id $identity.Id -ErrorAction SilentlyContinue)) 'Private EP Word process remains'
    if (-not $ExpectSuccess) {
        Assert-True (-not (Test-Path -LiteralPath $output)) 'Failed EP comparison published output'
        Write-Host ('PASS ' + $Name + ': unsafe section boundaries refused')
        return $result
    }
    Assert-True ($result.JobType -eq 'EP' -and $result.Sections.Count -eq 3) 'EP result must contain three scopes'
    Assert-True (($result.Sections.Scope -join '|') -eq 'Description|Full specification|Claims') 'EP scope order differs'
    Assert-True ($result.Sections[0].NewWords + $result.Sections[2].NewWords -eq $result.Sections[1].NewWords) 'EP new counts do not add up'
    Assert-True ($result.Sections[0].InsertedWords + $result.Sections[2].InsertedWords -eq $result.Sections[1].InsertedWords) 'EP insertion counts do not add up'
    $summary = Get-Content -LiteralPath ([IO.Path]::ChangeExtension($output, '.txt')) -Raw -Encoding UTF8
    foreach ($scope in @('Description', 'Full specification', 'Claims')) {
        Assert-True ($summary.Contains($scope)) ('Missing summary scope: ' + $scope)
    }
    $zip = [IO.Compression.ZipFile]::OpenRead($output)
    try {
        $reader = New-Object IO.StreamReader($zip.GetEntry('word/document.xml').Open())
        try { $xml = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
        $ns = New-Object Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        if ($result.InsertedWords -gt 0) {
            Assert-True ($xml.SelectNodes('//w:ins//w:r[w:t]/w:rPr/w:highlight[@w:val="yellow"]', $ns).Count -gt 0) 'EP insertion highlights missing'
        }
        if ($Name -eq 'scoped-edits') {
            Assert-True ($xml.SelectNodes('//w:del', $ns).Count -gt 0) 'EP deletion revisions missing'
            $text = $xml.SelectNodes('//w:t', $ns).InnerText -join ' '
            foreach ($excluded in @('COVEREXCLUDE', 'GERMANEXCLUDE', 'FRENCHEXCLUDE', 'DRAWINGEXCLUDE', 'REFERENCEEXCLUDE')) {
                Assert-True (-not $text.Contains($excluded)) ('Excluded EP text remains: ' + $excluded)
            }
        }
        if ($Name -eq 'scoped-moves') {
            Assert-True ($xml.SelectNodes('//w:moveFrom', $ns).Count -gt 0) 'EP move origin missing'
            Assert-True ($xml.SelectNodes('//w:moveTo', $ns).Count -gt 0) 'EP move destination missing'
        }
    } finally { $zip.Dispose() }
    Write-Host ('PASS ' + $Name + ': ' + (($result.Sections | ForEach-Object { $_.Scope + '=' + $_.NewWords + '/' + $_.InsertedWords }) -join ', '))
    return $result
}

$archive = Join-Path $fixtures 'ep-archive.docx'
$new = Join-Path $fixtures 'ep-new.docx'
New-Docx $archive @('COVEREXCLUDE', 'Description', 'Alpha beta gamma.', 'Keep this stable paragraph.', 'Old ending.', 'Claims', '1. A device with a blue housing.', '2. A process requiring heat.', 'Ansprüche', 'GERMANEXCLUDE', 'Revendications', 'FRENCHEXCLUDE', 'Drawing', 'DRAWINGEXCLUDE', 'References cited', 'REFERENCEEXCLUDE')
New-Docx $new @('Description', 'Alpha bright beta gamma.', 'Keep this stable paragraph.', 'Fresh ending.', 'Claims', '1. A device with a green housing.', '2. A process requiring heat.', '3. A system controlling pressure.')
$result = Invoke-EpCase 'scoped-edits' $archive $new
Assert-True ($result.Sections[0].NewWords -eq 10 -and $result.Sections[0].InsertedWords -eq 2) 'Description edit counts differ'
Assert-True ($result.Sections[2].NewWords -eq 17 -and $result.Sections[2].InsertedWords -eq 6) 'Claims edit counts differ'
Assert-True ([Math]::Abs($result.Sections[0].CrossoverPercent - 80) -lt 0.01) 'Description crossover differs'
Assert-True ([Math]::Abs($result.Sections[2].CrossoverPercent - (100.0 * 11 / 17)) -lt 0.01) 'Claims crossover differs'

$bad = Join-Path $fixtures 'ep-ambiguous.docx'
New-Docx $bad @('Description', 'A body paragraph.', 'Claims', '1. A claim.', 'Claims', '2. Another claim.')
$null = Invoke-EpCase 'ambiguous-claims' $bad $new $false

$opening = 'The opening section is unchanged and provides a clear starting point for this comparison.'
$moved = 'This entire paragraph describes the earlier archive process and will be moved to a different position without changing any of its original words.'
$anchor = 'The next section contains detailed information about project delivery schedules, review responsibilities, expected turnaround times, and the local document storage arrangements used by the team.'
$closing = 'The final fixed section explains the approval process in detail. Each reviewer checks the document carefully before the final version is released for use by the wider team.'
$moveArchive = Join-Path $fixtures 'ep-move-archive.docx'
$moveNew = Join-Path $fixtures 'ep-move-new.docx'
New-Docx $moveArchive @('Description', $opening, $moved, $anchor, $closing, 'Claims', '1. A device comprising an element.')
New-Docx $moveNew @('Description', $opening, $anchor, $closing, $moved, 'Claims', '1. A device comprising an element.')
$moveResult = Invoke-EpCase 'scoped-moves' $moveArchive $moveNew
foreach ($section in $moveResult.Sections) {
    Assert-True ($section.InsertedWords -eq 0 -and [Math]::Abs($section.CrossoverPercent - 100) -lt 0.01) 'Native EP move must remain retained'
}

if ($IncludeSamples) {
    $raw = $SampleArchivePath
    $prep = $SamplePreparedPath
    $sampleResult = Invoke-EpCase 'provided-raw-vs-prepared' $raw $prep
    foreach ($section in $sampleResult.Sections) {
        Assert-True ($section.InsertedWords -eq 0 -and [Math]::Abs($section.CrossoverPercent - 100) -lt 0.01) 'Provided raw/prepared section texts should compare identically'
    }
    # Samples and their derived documents contain user-provided content: remove this test's copies.
    $sampleCase = [IO.Path]::GetFullPath((Join-Path $fixtures 'provided-raw-vs-prepared'))
    $fixtureRoot = [IO.Path]::GetFullPath($fixtures).TrimEnd('\') + '\'
    if (-not $sampleCase.StartsWith($fixtureRoot, [StringComparison]::OrdinalIgnoreCase)) { throw 'Unexpected sample test cleanup path' }
    Remove-Item -LiteralPath $sampleCase -Recurse -Force
}
Write-Host ('EP synthetic test artifacts: ' + $fixtures)
