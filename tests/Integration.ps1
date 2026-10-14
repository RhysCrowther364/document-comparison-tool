param([switch]$DefineOnly)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$project = Split-Path -Parent $PSScriptRoot
$exe = Join-Path $project 'dist\LocalDocumentCompare\DocumentCompare.exe'
$fixtures = Join-Path $PSScriptRoot ('artifacts\integration-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $fixtures -Force | Out-Null
$utf8 = New-Object Text.UTF8Encoding($false)

function Add-ZipText($Zip, [string]$Name, [string]$Text) {
    $entry = $Zip.CreateEntry($Name)
    $writer = New-Object IO.StreamWriter($entry.Open(), $utf8)
    try { $writer.Write($Text) } finally { $writer.Dispose() }
}

function New-Docx([string]$Path, [string[]]$Paragraphs) {
    $zip = [IO.Compression.ZipFile]::Open($Path, [IO.Compression.ZipArchiveMode]::Create)
    try {
        Add-ZipText $zip '[Content_Types].xml' '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
        Add-ZipText $zip '_rels/.rels' '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
        $body = ($Paragraphs | ForEach-Object { '<w:p><w:r><w:t xml:space="preserve">' + [Security.SecurityElement]::Escape($_) + '</w:t></w:r></w:p>' }) -join ''
        Add-ZipText $zip 'word/document.xml' ('<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + $body + '<w:sectPr/></w:body></w:document>')
    } finally { $zip.Dispose() }
}

function Assert-True([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }

function Run-Case([string]$Name, [string[]]$Archive, [string[]]$New, $ExpectedCrossover, [int]$ExpectedWords, [int]$ExpectedInserted) {
    $case = Join-Path $fixtures $Name
    New-Item -ItemType Directory -Path $case -Force | Out-Null
    $oldPath = Join-Path $case 'Archive.docx'
    $newPath = Join-Path $case 'New.docx'
    New-Docx $oldPath $Archive
    New-Docx $newPath $New
    $oldHash = (Get-FileHash -LiteralPath $oldPath).Hash
    $newHash = (Get-FileHash -LiteralPath $newPath).Hash
    $job = Join-Path $case 'job'
    New-Item -ItemType Directory -Path $job -Force | Out-Null
    $output = Join-Path $case 'New vs Archive.docx'
    # These controlled fixtures have been checked as same-language English.
    $request = @{ ArchivePath = $oldPath; NewPath = $newPath; OutputPath = $output; WorkDirectory = $job; OcrArchive = $false; OcrNew = $false; MatchingLanguageConfirmed = $true }
    $requestPath = Join-Path $job 'request.json'
    [IO.File]::WriteAllText($requestPath, ($request | ConvertTo-Json), $utf8)
    $resultPath = Join-Path $job 'result.json'
    $arguments = '--worker "' + $requestPath + '" "' + $resultPath + '"'
    $process = Start-Process -FilePath $exe -ArgumentList $arguments -WindowStyle Hidden -PassThru
    if (-not $process.WaitForExit(90000)) { throw ('Timed out: ' + $Name + '; owned process records: ' + $job) }
    $result = Get-Content -LiteralPath $resultPath -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-True $result.Success ($Name + ': ' + $result.Error)
    Assert-True ($result.NewWords -eq $ExpectedWords) ($Name + ': expected ' + $ExpectedWords + ' words, got ' + $result.NewWords)
    Assert-True ($result.InsertedWords -eq $ExpectedInserted) ($Name + ': expected ' + $ExpectedInserted + ' inserted words, got ' + $result.InsertedWords)
    if ($null -eq $ExpectedCrossover) { Assert-True ($null -eq $result.CrossoverPercent) ($Name + ': expected unavailable crossover') }
    else { Assert-True ([Math]::Abs($result.CrossoverPercent - $ExpectedCrossover) -lt 0.01) ($Name + ': incorrect crossover: ' + $result.CrossoverPercent) }
    Assert-True ($null -eq $result.PSObject.Properties['EstimatedPages'] -and $null -eq $result.PSObject.Properties['DiscountPercent']) ($Name + ': retired metrics remain in worker result')
    Assert-True ((Get-FileHash -LiteralPath $oldPath).Hash -eq $oldHash) 'Archive was modified'
    Assert-True ((Get-FileHash -LiteralPath $newPath).Hash -eq $newHash) 'New file was modified'
    Assert-True (Test-Path -LiteralPath ([IO.Path]::ChangeExtension($output, '.txt'))) 'Missing summary'
    $zip = [IO.Compression.ZipFile]::OpenRead($output)
    try {
        $reader = New-Object IO.StreamReader($zip.GetEntry('word/document.xml').Open())
        try { $xml = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
        $ns = New-Object Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        if ($ExpectedInserted -gt 0) {
            Assert-True ($xml.SelectNodes('//w:ins//w:r[w:t]/w:rPr/w:highlight[@w:val="yellow"]', $ns).Count -gt 0) 'No yellow insertion highlights'
        }
        if ($Name -eq 'insert-delete') { Assert-True ($xml.SelectNodes('//w:del', $ns).Count -gt 0) 'Deletion revisions lost' }
    } finally { $zip.Dispose() }
    $recordPath = Join-Path $job 'word.process'
    if (Test-Path -LiteralPath $recordPath) {
        $owned = Get-Content -LiteralPath $recordPath -Raw | ConvertFrom-Json
        $stillRunning = Get-Process -Id $owned.Id -ErrorAction SilentlyContinue
        Assert-True ($null -eq $stillRunning) ('Owned Word process is still running: ' + $owned.Id)
    }
    Write-Host ('PASS ' + $Name + ': ' + $result.NewWords + ' words, ' + $result.InsertedWords + ' inserted, crossover ' + $result.CrossoverPercent)
    return $output
}

if ($DefineOnly) { return }
$first = Run-Case 'insert-delete' @('Alpha beta gamma.', 'Keep this stable paragraph.', 'Old ending.') @('Alpha bright beta gamma.', 'Keep this stable paragraph.', 'Fresh ending.') 80 10 2
$null = Run-Case 'identical' @('Alpha beta gamma.') @('Alpha beta gamma.') 100 3 0
$null = Run-Case 'deletion-only' @('Alpha beta gamma.') @('Alpha gamma.') 100 2 0
$null = Run-Case 'unrelated' @('Alpha beta gamma.') @('Fresh different wording.') 0 3 3
$null = Run-Case 'empty-new' @('Alpha beta gamma.') @('') $null 0 0

# Validate the post-processed DOCX opens in Word and still contains revisions.
$word = New-Object -ComObject Word.Application
$document = $null
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    $document = $word.Documents.Open($first, $false, $true, $false)
    Assert-True ($document.Revisions.Count -gt 0) 'Reopened document lost tracked changes'
    Write-Host 'PASS comparison reopens in Word with revisions'
} finally {
    $discard = 0
    if ($document) { $document.Close([ref]$discard); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    $word.Quit([ref]$discard)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
}
Write-Host ('Integration artifacts: ' + $fixtures)
