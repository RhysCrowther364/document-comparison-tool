$ErrorActionPreference = 'Stop'
$project = Split-Path -Parent $PSScriptRoot
$exe = Join-Path $project 'dist\LocalDocumentCompare\DocumentCompare.exe'
$desktop = [Environment]::GetFolderPath('DesktopDirectory')
if (-not (Test-Path -LiteralPath $desktop -PathType Container)) { throw 'Desktop folder unavailable' }
$id = [guid]::NewGuid().ToString('N')
$case = Join-Path $PSScriptRoot ('artifacts\desktop-output-' + $id)
New-Item -ItemType Directory -Path $case -Force | Out-Null
$job = Join-Path $case 'job'
New-Item -ItemType Directory -Path $job | Out-Null
$archive = Join-Path $case ('Desktop-check-archive-' + $id.Substring(0,8) + '.docx')
$new = Join-Path $case ('Desktop-check-new-' + $id.Substring(0,8) + '.docx')
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'PilotFixtures\WF-four-archive.docx') -Destination $archive
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'PilotFixtures\WF-five-added.docx') -Destination $new
$archiveHash = (Get-FileHash -LiteralPath $archive).Hash
$newHash = (Get-FileHash -LiteralPath $new).Hash
$output = Join-Path $desktop ([IO.Path]::GetFileNameWithoutExtension($new) + ' vs ' + [IO.Path]::GetFileNameWithoutExtension($archive) + '.docx')
$summary = [IO.Path]::ChangeExtension($output, '.txt')
if ((Test-Path -LiteralPath $output) -or (Test-Path -LiteralPath $summary)) { throw 'Test output already exists' }
$utf8 = New-Object Text.UTF8Encoding($false)
$request = @{ JobType = 'WF'; ArchivePath = $archive; NewPath = $new; OutputPath = $output; WorkDirectory = $job; MatchingLanguageConfirmed = $true }
$requestPath = Join-Path $job 'request.json'
$resultPath = Join-Path $job 'result.json'
[IO.File]::WriteAllText($requestPath, ($request | ConvertTo-Json), $utf8)
try {
    $worker = Start-Process -FilePath $exe -ArgumentList ('--worker "' + $requestPath + '" "' + $resultPath + '"') -WindowStyle Hidden -PassThru
    if (-not $worker.WaitForExit(90000)) { throw ('Test timed out; private process records at ' + $job) }
    $result = Get-Content -LiteralPath $resultPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if (-not $result.Success) { throw $result.Error }
    if ($result.NewWords -ne 5 -or $result.InsertedWords -ne 1 -or $result.CrossoverPercent -ne 80) { throw 'Desktop result figures differ' }
    if (-not (Test-Path -LiteralPath $output) -or -not (Test-Path -LiteralPath $summary)) { throw 'Desktop outputs missing' }
    $lines = @(Get-Content -LiteralPath $summary)
    if ($lines.Count -lt 3 -or $lines[0] -ne 'Crossover: 80.0%' -or $lines[1] -ne 'Total word count of new file: 5' -or $lines[2] -ne 'Highlighted insertion word count: 1') { throw 'Desktop summary differs' }
    if (-not ($lines -match 'Source language was confirmed manually')) { throw 'Manual language check note missing' }
    if ((Get-FileHash -LiteralPath $archive).Hash -ne $archiveHash -or (Get-FileHash -LiteralPath $new).Hash -ne $newHash) { throw 'Source changed' }
    $record = Get-Content -LiteralPath (Join-Path $job 'word.process') -Raw | ConvertFrom-Json
    if (Get-Process -Id $record.Id -ErrorAction SilentlyContinue) { throw 'Private Word process remains' }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($output)
    try {
        $reader = New-Object IO.StreamReader($zip.GetEntry('word/document.xml').Open())
        try { $xml = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
        $ns = New-Object Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        if ($xml.SelectNodes('//w:ins//w:r[w:t]/w:rPr/w:highlight[@w:val="yellow"]', $ns).Count -eq 0) { throw 'Desktop comparison has no yellow insertion' }
    } finally { $zip.Dispose() }
    Write-Output 'PASS desktop output in actual Word: DOCX and summary saved; 5 words, 1 insertion, 80.0%; source hashes unchanged and private Word closed.'
} finally {
    # These two unique synthetic output paths were checked absent before this
    # run. Resolve and check their Desktop boundary before removing them.
    $desktopRoot = [IO.Path]::GetFullPath($desktop).TrimEnd('\') + '\'
    foreach ($file in @($output,$summary)) {
        $resolved = [IO.Path]::GetFullPath($file)
        if (-not $resolved.StartsWith($desktopRoot, [StringComparison]::OrdinalIgnoreCase)) { throw 'Test cleanup path escaped Desktop' }
        if (Test-Path -LiteralPath $resolved -PathType Leaf) { Remove-Item -LiteralPath $resolved -Force }
    }
}
