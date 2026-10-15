#requires -Version 5.1
<#
Runs only against an explicitly selected local application build. All PDFs and
Word files contain synthetic text. The team package and installed app are not
changed. Native Word sessions run in bounded child processes; a timeout may
stop only a process whose exact PID and start time were recorded by this test
or the application's worker. Use Windows PowerShell 5.1 for COM inspection.
#>
[CmdletBinding()]
param(
    [string]$ExePath,
    [switch]$KeepFiles,
    [switch]$PrepareOnly,
    [string]$ReuseFixtureRoot,
    [switch]$InspectWord,
    [string]$DocumentPath,
    [string]$InspectionPath,
    [string]$IdentityPath,
    [string]$PreviewPdfPath
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$utf8 = New-Object Text.UTF8Encoding($false)

function Assert-True([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
    $script:checks++
}

if ($InspectWord) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class OcrLayoutWindowIdentity {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr window, out uint id);
}
'@
    $word = $null; $document = $null; $bootstrap = $null; $owned = $false
    $oldPrompt = $false; $oldLinks = $false; $optionsCaptured = $false
    $existing = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
    try {
        $word = New-Object -ComObject Word.Application
        $bootstrap = $word.Documents.Add()
        $window = $bootstrap.Windows.Item(1)
        try {
            $wordId = [uint32]0
            $null = [OcrLayoutWindowIdentity]::GetWindowThreadProcessId([IntPtr]$window.Hwnd, [ref]$wordId)
        }
        finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($window) }
        if ($wordId -eq 0 -or $existing -contains [int]$wordId) { throw 'Could not verify a separate private Word process.' }
        $owned = $true
        $process = Get-Process -Id $wordId
        [IO.File]::WriteAllText($IdentityPath, (@{ Id=[int]$wordId; StartTimeUtcTicks=$process.StartTime.ToUniversalTime().Ticks } | ConvertTo-Json -Compress), $utf8)
        $bootstrap.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($bootstrap); $bootstrap = $null
        $word.Visible = $false; $word.DisplayAlerts = 0; $word.AutomationSecurity = 3
        $oldPrompt = $word.Options.SaveNormalPrompt; $oldLinks = $word.Options.UpdateLinksAtOpen
        $optionsCaptured = $true
        $word.Options.SaveNormalPrompt = $false; $word.Options.UpdateLinksAtOpen = $false
        $document = $word.Documents.Open($DocumentPath, $false, $true, $false)
        if($PreviewPdfPath) { $document.ExportAsFixedFormat($PreviewPdfPath,17) }
        # Inspect final text, accepting revisions only in this unsaved private copy.
        $document.TrackRevisions = $false
        $document.Revisions.AcceptAll()
        $document.Repaginate()
        $markers = @{}
        foreach ($phrase in @('Document comparison system', 'The controller receives', 'The apparatus includes', 'Second page overview', 'Left column begins', 'Left column ends', 'Right column begins', 'Right column ends')) {
            $range = $document.Content.Duplicate
            try {
                $range.Find.ClearFormatting()
                $range.Find.MatchWildcards = $false; $range.Find.Forward = $true; $range.Find.Wrap = 0
                if (-not $range.Find.Execute($phrase)) { throw 'Native Word did not contain an expected synthetic marker.' }
                $paragraph = $range.ParagraphFormat
                try {
                    $markers[$phrase] = @{
                        Page=[int]$range.Information(3); X=[double]$range.Information(5); Y=[double]$range.Information(6)
                        FontSize=[double]$range.Font.Size; SpaceBefore=[double]$paragraph.SpaceBefore; SpaceAfter=[double]$paragraph.SpaceAfter; LeftIndent=[double]$paragraph.LeftIndent
                    }
                }
                finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($paragraph) }
            }
            finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($range) }
        }
        $inspection = @{
            Pages=[int]$document.ComputeStatistics(2); Width=[double]$document.PageSetup.PageWidth; Height=[double]$document.PageSetup.PageHeight
            Shapes=[int]$document.Shapes.Count; InlineShapes=[int]$document.InlineShapes.Count; Markers=$markers
        }
        [IO.File]::WriteAllText($InspectionPath, ($inspection | ConvertTo-Json -Depth 5), $utf8)
    }
    finally {
        if ($bootstrap) { try { $bootstrap.Close(0) } catch { }; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($bootstrap) }
        if ($document) { try { $document.Close(0) } catch { }; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
        if ($word) {
            if ($optionsCaptured) { try { $word.Options.SaveNormalPrompt = $oldPrompt; $word.Options.UpdateLinksAtOpen = $oldLinks } catch { } }
            if ($owned) { try { $word.Quit() } catch { } }
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
        }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
    exit 0
}

if (-not $ExePath -or -not [IO.File]::Exists($ExePath)) { throw 'Supply -ExePath for the private local build to test.' }
$ExePath = [IO.Path]::GetFullPath($ExePath)
$reuseResults = -not [string]::IsNullOrWhiteSpace($ReuseFixtureRoot)
if($reuseResults) {
    $fixtureRoot = [IO.Path]::GetFullPath($ReuseFixtureRoot)
    $boundary = [IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\')+'\'
    if(-not $fixtureRoot.StartsWith($boundary,[StringComparison]::OrdinalIgnoreCase) -or -not [IO.Path]::GetFileName($fixtureRoot).StartsWith('.ocr-layout-',[StringComparison]::Ordinal) -or -not [IO.Directory]::Exists($fixtureRoot)) { throw 'Only a previous synthetic OCR-layout fixture inside tests can be reused.' }
}
else {
    $fixtureRoot = Join-Path $PSScriptRoot ('.ocr-layout-' + [Guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $fixtureRoot
}
$powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$script:checks = 0
$ascii = [Text.Encoding]::ASCII
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Stop-VerifiedProcess([string]$RecordPath) {
    if (-not [IO.File]::Exists($RecordPath)) { return }
    $record = [IO.File]::ReadAllText($RecordPath) | ConvertFrom-Json
    $process = Get-Process -Id $record.Id -ErrorAction SilentlyContinue
    if ($process -and $process.StartTime.ToUniversalTime().Ticks -eq [long]$record.StartTimeUtcTicks) {
        Stop-Process -InputObject $process -Force
    }
}

function Write-Ascii([IO.Stream]$Stream, [string]$Text) {
    $bytes = $ascii.GetBytes($Text); $Stream.Write($bytes, 0, $bytes.Length)
}

function New-ScannedPdf([string]$Path, [bool]$Revised, [bool]$TextOnly=$false) {
    $pageOneLines = @(
        'The present invention describes a document comparison system.',
        $(if ($Revised) { 'The controller receives updated information from the sensor.' } else { 'The controller receives information from the sensor.' }),
        'The operator reviews each paragraph and checks the result.',
        'The apparatus includes a housing and a processor.',
        'The archive remains stored on this computer.',
        'The final document shows changes for the specialist.'
    )
    $left = @('Left column begins', 'The motor drives the wheel.', 'The sensor measures the speed.', 'The controller records each value.', 'The housing protects the system.', 'Left column ends')
    $right = @('Right column begins', 'The operator checks the display.', 'The device stores the result.', 'The method improves operation.', 'The processor handles the signal.', 'Right column ends')
    $pageTexts = @(
        ('Document comparison system ' + ($pageOneLines -join ' ')),
        ('Second page overview ' + ($left -join ' ') + ' ' + ($right -join ' '))
    )
    if($TextOnly) { return ($pageTexts -join "`f") }
    $images = New-Object 'Collections.Generic.List[byte[]]'
    foreach ($page in 1..2) {
        $bitmap = New-Object Drawing.Bitmap(1632, 2112)
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        $heading = New-Object Drawing.Font('Arial', 58, [Drawing.FontStyle]::Bold, [Drawing.GraphicsUnit]::Pixel)
        $body = New-Object Drawing.Font('Arial', 34, [Drawing.FontStyle]::Regular, [Drawing.GraphicsUnit]::Pixel)
        $jpeg = New-Object IO.MemoryStream
        try {
            $graphics.Clear([Drawing.Color]::White)
            $graphics.TextRenderingHint = [Drawing.Text.TextRenderingHint]::AntiAliasGridFit
            if ($page -eq 1) {
                $graphics.DrawString('Document comparison system', $heading, [Drawing.Brushes]::Black, [single]128, [single]128)
                for ($i=0; $i -lt $pageOneLines.Count; $i++) {
                    $y = 320 + 62*$i + $(if($i -ge 3) { 120 } else { 0 })
                    $graphics.DrawString($pageOneLines[$i], $body, [Drawing.Brushes]::Black, [single]128, [single]$y)
                }
            }
            else {
                $graphics.DrawString('Second page overview', $heading, [Drawing.Brushes]::Black, [single]128, [single]128)
                for ($i=0; $i -lt $left.Count; $i++) {
                    $graphics.DrawString($left[$i], $body, [Drawing.Brushes]::Black, [single]128, [single](340+64*$i))
                    $graphics.DrawString($right[$i], $body, [Drawing.Brushes]::Black, [single]865, [single](340+64*$i))
                }
            }
            $bitmap.Save($jpeg, [Drawing.Imaging.ImageFormat]::Jpeg)
            $images.Add($jpeg.ToArray())
        }
        finally { $jpeg.Dispose(); $heading.Dispose(); $body.Dispose(); $graphics.Dispose(); $bitmap.Dispose() }
    }
    $pdf = New-Object IO.MemoryStream
    $offsets = New-Object 'Collections.Generic.List[long]'
    function Add-PdfObject([int]$Number,[string]$Body) { $offsets.Add($pdf.Position); Write-Ascii $pdf ($Number.ToString()+" 0 obj`n"+$Body+"`nendobj`n") }
    try {
        Write-Ascii $pdf "%PDF-1.4`n"
        Add-PdfObject 1 '<< /Type /Catalog /Pages 2 0 R >>'
        Add-PdfObject 2 '<< /Type /Pages /Kids [3 0 R 6 0 R] /Count 2 >>'
        for ($page=0; $page -lt 2; $page++) {
            $number = 3+3*$page
            Add-PdfObject $number ('<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /XObject << /Scan '+($number+1)+' 0 R >> >> /Contents '+($number+2)+' 0 R >>')
            $offsets.Add($pdf.Position)
            Write-Ascii $pdf (($number+1).ToString()+' 0 obj << /Type /XObject /Subtype /Image /Width 1632 /Height 2112 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length '+$images[$page].Length+" >>`nstream`n")
            $pdf.Write($images[$page],0,$images[$page].Length); Write-Ascii $pdf "`nendstream`nendobj`n"
            $content = "q 612 0 0 792 0 0 cm /Scan Do Q`n"
            Add-PdfObject ($number+2) ('<< /Length '+$ascii.GetByteCount($content)+" >>`nstream`n"+$content+'endstream')
        }
        $xref = $pdf.Position; Write-Ascii $pdf "xref`n0 9`n0000000000 65535 f `n"
        foreach($offset in $offsets) { Write-Ascii $pdf ($offset.ToString('D10')+" 00000 n `n") }
        Write-Ascii $pdf ("trailer`n<< /Size 9 /Root 1 0 R >>`nstartxref`n"+$xref+"`n%%EOF`n")
        [IO.File]::WriteAllBytes($Path,$pdf.ToArray())
    }
    finally { $pdf.Dispose() }
    return ($pageTexts -join "`f")
}

function Invoke-Comparison([string]$Name,[string]$Archive,[string]$New,[int]$ExpectedWords,[int]$ExpectedInserted) {
    $job = Join-Path $fixtureRoot $Name
    if(-not $reuseResults) { $null = New-Item -ItemType Directory -Path $job }
    $output = Join-Path $job 'New vs Archive.docx'
    $requestPath = Join-Path $job 'request.json'; $resultPath = Join-Path $job 'result.json'
    $request = @{ ArchivePath=$Archive; NewPath=$New; OutputPath=$output; WorkDirectory=$job; OcrArchive=$true; OcrNew=$true; MatchingLanguageConfirmed=$true }
    $process = $null
    if(-not $reuseResults) {
        [IO.File]::WriteAllText($requestPath,($request | ConvertTo-Json),$utf8)
        $process = Start-Process -FilePath $ExePath -ArgumentList ('--worker "'+$requestPath+'" "'+$resultPath+'"') -WindowStyle Hidden -PassThru
    }
    try {
        if ($process -and -not $process.WaitForExit(240000)) {
            Stop-VerifiedProcess (Join-Path $job 'ocr.process'); Stop-VerifiedProcess (Join-Path $job 'word.process')
            if (-not $process.HasExited) { $process.Kill() }
            throw ('Bounded synthetic OCR comparison timed out: '+$Name)
        }
        Assert-True ([IO.File]::Exists($resultPath)) ($Name+': result missing.')
        $result = [IO.File]::ReadAllText($resultPath) | ConvertFrom-Json
        Assert-True $result.Success ($Name+': '+$result.Error)
        Assert-True ($result.NewWords -eq $ExpectedWords) ($Name+': OCR changed independent expected word count; expected '+$ExpectedWords+', actual '+$result.NewWords)
        Assert-True ($result.InsertedWords -eq $ExpectedInserted) ($Name+': inserted word count differs.')
        Assert-True ($null -ne $result.CrossoverPercent) ($Name+': strict crossover reconciliation failed.')
        Assert-True ([Math]::Abs($result.CrossoverPercent - 100.0*($ExpectedWords-$ExpectedInserted)/$ExpectedWords) -lt 0.000001) ($Name+': crossover differs.')
        foreach($label in @('archive','new')) {
            $status = [IO.File]::ReadAllText((Join-Path $job ($label+'-ocr-status.json'))) | ConvertFrom-Json
            Assert-True ($status.State -eq 'Complete' -and $status.PageCount -eq 2 -and $status.PagesCompleted -eq 2) 'OCR did not retain all source pages.'
        }
    }
    finally { if($process) { $process.Dispose() } }
    Write-Host ('PASS '+$Name+': '+$result.NewWords+' words, '+$result.InsertedWords+' inserted, crossover '+$result.CrossoverPercent)
    return $output
}

try {
    $archive = Join-Path $fixtureRoot 'Archive.pdf'; $new = Join-Path $fixtureRoot 'New.pdf'
    $null = New-ScannedPdf $archive $false $reuseResults
    $expectedText = New-ScannedPdf $new $true $reuseResults
    $expectedWords = ($expectedText -split '\s+' | Where-Object { $_ }).Count
    if($PrepareOnly) {
        Write-Output ('Synthetic two-page fixture prepared; new document has '+$expectedWords+' expected words and one added word.')
        return
    }
    $oldHash = (Get-FileHash -LiteralPath $archive).Hash; $newHash = (Get-FileHash -LiteralPath $new).Hash
    $changed = Invoke-Comparison 'changed' $archive $new $expectedWords 1
    $identical = Invoke-Comparison 'identical' $new $new $expectedWords 0
    $assembly = [Reflection.Assembly]::LoadFrom($ExePath)
    $contentType = $assembly.GetType('LocalDocumentCompare.DocxContent',$true)
    $textMethod = $contentType.GetMethod('ReadFinalText')
    foreach($document in @($changed,$identical)) {
        $parameters = New-Object object[] 1; $parameters[0] = [string]$document
        $finalText = [string]$textMethod.Invoke($null,$parameters)
        # Windows supplies reading order. Retain that established OCR sequence
        # while independently checking that every known source word appears once.
        $expectedTokens = @([regex]::Matches($expectedText,'[A-Za-z]+') | ForEach-Object Value | Sort-Object)
        $actualTokens = @([regex]::Matches($finalText,'[A-Za-z]+') | ForEach-Object Value | Sort-Object)
        Assert-True (($expectedTokens -join ' ') -ceq ($actualTokens -join ' ')) 'OCR output omitted or duplicated known synthetic source words.'
        $ocrText = [IO.File]::ReadAllText((Join-Path ([IO.Path]::GetDirectoryName($document)) 'new-ocr.txt'))
        Assert-True (($ocrText -replace '\s+',' ').Trim() -ceq ($finalText -replace '\s+',' ').Trim()) 'Layout reconstruction changed the established OCR reading order.'
        $zip = [IO.Compression.ZipFile]::OpenRead($document)
        try {
            $reader = New-Object IO.StreamReader($zip.GetEntry('word/document.xml').Open())
            try { $xml = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
            $ns = New-Object Xml.XmlNamespaceManager($xml.NameTable); $ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
            Assert-True ($xml.SelectNodes('//w:drawing|//w:pict|//w:vanish',$ns).Count -eq 0) 'OCR text was covered by a bitmap or duplicated in hidden text.'
            if($document -eq $changed) {
                Assert-True ($xml.SelectNodes('//w:ins//w:r[w:t]/w:rPr/w:highlight[@w:val="yellow"]',$ns).Count -gt 0) 'Inserted words lack yellow tracked markup.'
            }
        }
        finally { $zip.Dispose() }
    }
    $inspectionPath = Join-Path $fixtureRoot 'inspection.json'; $identity = Join-Path $fixtureRoot 'inspection-word.process'
    if([IO.File]::Exists($inspectionPath)) { [IO.File]::Delete($inspectionPath) }
    if([IO.File]::Exists($identity)) { [IO.File]::Delete($identity) }
    $stdout = Join-Path $fixtureRoot 'inspection.stdout'; $stderr = Join-Path $fixtureRoot 'inspection.stderr'
    $arguments = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "'+$PSCommandPath+'" -InspectWord -DocumentPath "'+$changed+'" -InspectionPath "'+$inspectionPath+'" -IdentityPath "'+$identity+'"'
    if($PreviewPdfPath) { $arguments += ' -PreviewPdfPath "'+[IO.Path]::GetFullPath($PreviewPdfPath)+'"' }
    # Start-Process -PassThru can return a null ExitCode under WindowsPS5.1
    # after waiting. Keep our own Process handle and drain both pipes instead.
    $processInfo = New-Object Diagnostics.ProcessStartInfo
    $processInfo.FileName=$powershell; $processInfo.Arguments=$arguments
    $processInfo.UseShellExecute=$false; $processInfo.CreateNoWindow=$true; $processInfo.WindowStyle=[Diagnostics.ProcessWindowStyle]::Hidden
    $processInfo.RedirectStandardOutput=$true; $processInfo.RedirectStandardError=$true
    $process = New-Object Diagnostics.Process; $process.StartInfo=$processInfo; $null=$process.Start()
    $stdoutTask=$process.StandardOutput.ReadToEndAsync(); $stderrTask=$process.StandardError.ReadToEndAsync()
    try {
        if(-not $process.WaitForExit(60000)) { Stop-VerifiedProcess $identity; if(-not $process.HasExited) { $process.Kill() }; throw 'Native Word layout inspection timed out.' }
        [IO.File]::WriteAllText($stdout,$stdoutTask.GetAwaiter().GetResult(),$utf8); [IO.File]::WriteAllText($stderr,$stderrTask.GetAwaiter().GetResult(),$utf8)
        Assert-True ($process.ExitCode -eq 0 -and [IO.File]::Exists($inspectionPath)) ('Word layout inspection failed: '+[IO.File]::ReadAllText($stderr))
    }
    finally { $process.Dispose() }
    $inspection = [IO.File]::ReadAllText($inspectionPath) | ConvertFrom-Json
    Assert-True ([Math]::Abs($inspection.Width-612) -lt 2 -and [Math]::Abs($inspection.Height-792) -lt 2) 'OCR lost the original physical page dimensions.'
    Assert-True ($inspection.Pages -eq 2) 'OCR did not retain the two-page source layout.'
    $markers = $inspection.Markers
    Assert-True ($markers.'Document comparison system'.Page -eq 1 -and $markers.'Second page overview'.Page -eq 2) 'OCR lost the source page boundary.'
    Assert-True ($markers.'Document comparison system'.FontSize -gt $markers.'The controller receives'.FontSize*1.25) 'OCR lost the heading/body size distinction.'
    Assert-True ([Math]::Abs($markers.'The controller receives'.SpaceAfter) -lt 0.1) 'Word default paragraph spacing distorted the recognised line spacing.'
    Assert-True ($markers.'The apparatus includes'.Y -gt $markers.'The controller receives'.Y+60) 'OCR flattened the visible gap between source paragraphs.'
    Assert-True ($markers.'Right column begins'.LeftIndent -gt $markers.'Left column begins'.LeftIndent+180) 'OCR flattened both source columns into one.'
    Assert-True ($inspection.Shapes -eq 0 -and $inspection.InlineShapes -eq 0) 'Comparison contains an OCR bitmap overlay.'
    $ownedWord = [IO.File]::ReadAllText($identity) | ConvertFrom-Json
    $stillRunning = Get-Process -Id $ownedWord.Id -ErrorAction SilentlyContinue
    Assert-True ($null -eq $stillRunning -or $stillRunning.StartTime.ToUniversalTime().Ticks -ne [long]$ownedWord.StartTimeUtcTicks) 'Private layout-inspection Word process is still running.'
    Assert-True ((Get-FileHash -LiteralPath $archive).Hash -eq $oldHash -and (Get-FileHash -LiteralPath $new).Hash -eq $newHash) 'Synthetic source PDFs changed.'
    Write-Output ('PASS: '+$script:checks+' synthetic OCR/Word layout checks; all pages, heading/body sizes, column placement and reading order, exact word counts, yellow insertions, unchanged sources and no duplicate image overlay.')
}
finally {
    if($KeepFiles -or $reuseResults) { Write-Output ('Synthetic fixture files: '+$fixtureRoot) }
    else {
        $resolved = [IO.Path]::GetFullPath($fixtureRoot); $allowed = [IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\')+'\'
        if(-not $resolved.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase) -or -not [IO.Path]::GetFileName($resolved).StartsWith('.ocr-layout-',[StringComparison]::Ordinal)) { throw 'Synthetic cleanup target is outside tests.' }
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}
