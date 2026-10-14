#requires -Version 5.1
[CmdletBinding()]
param([switch]$KeepFiles)

$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$fixtureRoot = Join-Path $PSScriptRoot ('.ocr-smoke-' + [Guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $fixtureRoot
$helper = Join-Path $root 'scripts\Convert-PdfLocal.ps1'
$pdfPath = Join-Path $fixtureRoot 'mixed.pdf'
$outputPath = Join-Path $fixtureRoot 'recognized.txt'
$statusPath = Join-Path $fixtureRoot 'status.json'
$ascii = [Text.Encoding]::ASCII

function Write-Ascii([IO.Stream]$Stream, [string]$Text) {
    $bytes = $ascii.GetBytes($Text)
    $Stream.Write($bytes, 0, $bytes.Length)
}

function Add-Object([IO.Stream]$Stream, [int]$Number, [string]$Body) {
    $offsets.Add($Stream.Position)
    Write-Ascii $Stream ($Number.ToString() + " 0 obj`n" + $Body + "`nendobj`n")
}

try {
    # Page one has real PDF text. Page two contains only a raster image, like a
    # scanned archive page. All content is synthetic, and everything stays local.
    Add-Type -AssemblyName System.Drawing
    $bitmap = New-Object Drawing.Bitmap(1600, 450)
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    $font = New-Object Drawing.Font('Arial', 42)
    $jpeg = New-Object IO.MemoryStream
    try {
        $graphics.Clear([Drawing.Color]::White)
        $graphics.TextRenderingHint = [Drawing.Text.TextRenderingHint]::AntiAliasGridFit
        $graphics.DrawString('Scanned archive page with local recognition.', $font, [Drawing.Brushes]::Black, [single]40, [single]80)
        $graphics.DrawString('Yellow apples and blue flowers.', $font, [Drawing.Brushes]::Black, [single]40, [single]180)
        $bitmap.Save($jpeg, [Drawing.Imaging.ImageFormat]::Jpeg)
        $imageBytes = $jpeg.ToArray()
    }
    finally {
        $jpeg.Dispose(); $font.Dispose(); $graphics.Dispose(); $bitmap.Dispose()
    }

    $pdf = New-Object IO.MemoryStream
    $offsets = New-Object 'Collections.Generic.List[long]'
    try {
        Write-Ascii $pdf "%PDF-1.4`n"
        Add-Object $pdf 1 '<< /Type /Catalog /Pages 2 0 R >>'
        Add-Object $pdf 2 '<< /Type /Pages /Kids [3 0 R 6 0 R] /Count 2 >>'
        Add-Object $pdf 3 '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>'
        Add-Object $pdf 4 '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'
        $pageOne = "BT /F1 20 Tf 50 710 Td (Local text recognition test.) Tj 0 -40 Td (This page contains selectable text.) Tj ET`n"
        Add-Object $pdf 5 ('<< /Length ' + $ascii.GetByteCount($pageOne) + " >>`nstream`n" + $pageOne + 'endstream')
        Add-Object $pdf 6 '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /XObject << /Scan 7 0 R >> >> /Contents 8 0 R >>'
        $offsets.Add($pdf.Position)
        Write-Ascii $pdf ('7 0 obj << /Type /XObject /Subtype /Image /Width 1600 /Height 450 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ' + $imageBytes.Length + " >>`nstream`n")
        $pdf.Write($imageBytes, 0, $imageBytes.Length)
        Write-Ascii $pdf "`nendstream`nendobj`n"
        $pageTwo = "q 540 0 0 151.875 36 550 cm /Scan Do Q`n"
        Add-Object $pdf 8 ('<< /Length ' + $ascii.GetByteCount($pageTwo) + " >>`nstream`n" + $pageTwo + 'endstream')
        $xrefPosition = $pdf.Position
        Write-Ascii $pdf "xref`n0 9`n0000000000 65535 f `n"
        foreach ($offset in $offsets) { Write-Ascii $pdf ($offset.ToString('D10') + " 00000 n `n") }
        Write-Ascii $pdf ("trailer`n<< /Size 9 /Root 1 0 R >>`nstartxref`n" + $xrefPosition + "`n%%EOF`n")
        $pdfBytes = $pdf.ToArray()
        [IO.File]::WriteAllBytes($pdfPath, $pdfBytes)
    }
    finally { $pdf.Dispose() }

    $powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    & $powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $helper -InputPath $pdfPath -OutputPath $outputPath -StatusPath $statusPath
    if ($LASTEXITCODE -ne 0) { throw 'OCR helper failed on the synthetic mixed PDF.' }
    $text = [IO.File]::ReadAllText($outputPath)
    $status = [IO.File]::ReadAllText($statusPath) | ConvertFrom-Json
    if ($status.State -ne 'Complete' -or $status.PageCount -ne 2 -or $status.PagesCompleted -ne 2 -or $status.BlankPages.Count -ne 0) {
        throw 'OCR progress metadata does not describe both pages.'
    }
    $pages = $text.Split([char]12)
    if ($pages.Count -ne 2) { throw 'OCR output did not preserve page boundaries.' }
    if ($pages[0] -notmatch 'Local text recognition test' -or $pages[1] -notmatch 'Scanned archive page with local recognition' -or $pages[1] -notmatch 'Yellow apples and blue flowers') {
        throw 'OCR did not recognize the expected synthetic text on both pages.'
    }
    # Refusal to overwrite output must also preserve the PDF if supplied as the
    # status path accidentally. Failures must not disclose document content.
    $beforeHash = (Get-FileHash -LiteralPath $pdfPath -Algorithm SHA256).Hash
    $processInfo = New-Object Diagnostics.ProcessStartInfo
    $processInfo.FileName = $powershell
    $processInfo.Arguments = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + $helper + '" -InputPath "' + $pdfPath + '" -OutputPath "' + $outputPath + '" -StatusPath "' + $pdfPath + '"'
    $processInfo.UseShellExecute = $false
    $processInfo.CreateNoWindow = $true
    $processInfo.RedirectStandardError = $true
    $processInfo.RedirectStandardOutput = $true
    $process = [Diagnostics.Process]::Start($processInfo)
    try {
        $failureText = $process.StandardError.ReadToEnd()
        $stdout = $process.StandardOutput.ReadToEnd()
        $process.WaitForExit()
        if ($process.ExitCode -eq 0) { throw 'OCR did not reject conflicting paths.' }
        if (($failureText + $stdout) -match 'Scanned archive|Yellow apples|selectable text') { throw 'OCR error disclosed document text.' }
    }
    finally { $process.Dispose() }
    if ((Get-FileHash -LiteralPath $pdfPath -Algorithm SHA256).Hash -ne $beforeHash) { throw 'OCR modified its source PDF.' }

    $invalidPdf = Join-Path $fixtureRoot 'invalid.pdf'
    $invalidOutput = Join-Path $fixtureRoot 'invalid.txt'
    $invalidStatus = Join-Path $fixtureRoot 'invalid-status.json'
    [IO.File]::WriteAllText($invalidPdf, 'SYNTHETIC CONFIDENTIAL CONTENT: this is not a PDF.')
    $processInfo.Arguments = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + $helper + '" -InputPath "' + $invalidPdf + '" -OutputPath "' + $invalidOutput + '" -StatusPath "' + $invalidStatus + '"'
    $process = [Diagnostics.Process]::Start($processInfo)
    try {
        $failureText = $process.StandardError.ReadToEnd()
        $stdout = $process.StandardOutput.ReadToEnd()
        $process.WaitForExit()
        if ($process.ExitCode -eq 0 -or [IO.File]::Exists($invalidOutput)) { throw 'OCR did not reject an invalid PDF cleanly.' }
        $failedStatus = [IO.File]::ReadAllText($invalidStatus) | ConvertFrom-Json
        if ($failedStatus.State -ne 'Error' -or $failedStatus.Error -notmatch 'could not open this PDF') { throw 'OCR did not report a useful invalid-PDF error.' }
        if (($failureText + $stdout + $failedStatus.Error) -match 'SYNTHETIC CONFIDENTIAL CONTENT') { throw 'OCR error disclosed input content.' }
    }
    finally { $process.Dispose() }

    # Preserve byte lengths and xref offsets while blanking page content. The
    # single-byte Latin-1 round trip also preserves the unused embedded JPEG.
    $latin1 = [Text.Encoding]::GetEncoding(28591)
    $pdfSource = $latin1.GetString($pdfBytes)
    $textAndBlankPdf = Join-Path $fixtureRoot 'text-and-blank.pdf'
    $textAndBlankOutput = Join-Path $fixtureRoot 'text-and-blank.txt'
    $textAndBlankStatus = Join-Path $fixtureRoot 'text-and-blank-status.json'
    [IO.File]::WriteAllBytes($textAndBlankPdf, $latin1.GetBytes($pdfSource.Replace($pageTwo, (' ' * $pageTwo.Length))))
    $processInfo.Arguments = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + $helper + '" -InputPath "' + $textAndBlankPdf + '" -OutputPath "' + $textAndBlankOutput + '" -StatusPath "' + $textAndBlankStatus + '"'
    $process = [Diagnostics.Process]::Start($processInfo)
    try {
        $failureText = $process.StandardError.ReadToEnd()
        $stdout = $process.StandardOutput.ReadToEnd()
        $process.WaitForExit()
        if ($process.ExitCode -ne 0) { throw 'OCR helper rejected a mixed text/blank PDF.' }
        $mixedStatus = [IO.File]::ReadAllText($textAndBlankStatus) | ConvertFrom-Json
        if ($mixedStatus.State -ne 'Complete' -or $mixedStatus.BlankPages.Count -ne 1 -or $mixedStatus.BlankPages[0] -ne 2) {
            throw 'OCR helper did not flag exactly the blank second page.'
        }
        $mixedPages = [IO.File]::ReadAllText($textAndBlankOutput).Split([char]12)
        if ($mixedPages.Count -ne 2 -or $mixedPages[0] -notmatch 'Local text recognition test' -or -not [string]::IsNullOrWhiteSpace($mixedPages[1])) {
            throw 'OCR helper did not preserve the mixed text/blank PDF boundary.'
        }
    }
    finally { $process.Dispose() }

    $blankPdf = Join-Path $fixtureRoot 'blank.pdf'
    $blankOutput = Join-Path $fixtureRoot 'blank.txt'
    $blankStatus = Join-Path $fixtureRoot 'blank-status.json'
    $blankSource = $pdfSource.Replace($pageOne, (' ' * $pageOne.Length)).Replace($pageTwo, (' ' * $pageTwo.Length))
    [IO.File]::WriteAllBytes($blankPdf, $latin1.GetBytes($blankSource))
    $processInfo.Arguments = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + $helper + '" -InputPath "' + $blankPdf + '" -OutputPath "' + $blankOutput + '" -StatusPath "' + $blankStatus + '"'
    $process = [Diagnostics.Process]::Start($processInfo)
    try {
        $failureText = $process.StandardError.ReadToEnd()
        $stdout = $process.StandardOutput.ReadToEnd()
        $process.WaitForExit()
        if ($process.ExitCode -eq 0 -or [IO.File]::Exists($blankOutput)) { throw 'OCR helper kept an output for an entirely blank PDF.' }
        $emptyStatus = [IO.File]::ReadAllText($blankStatus) | ConvertFrom-Json
        if ($emptyStatus.State -ne 'Error' -or $emptyStatus.BlankPages.Count -ne 2 -or $emptyStatus.Error -notmatch 'no readable text') {
            throw 'OCR helper did not explain the entirely blank PDF failure.'
        }
    }
    finally { $process.Dispose() }

    if (@(Get-ChildItem -LiteralPath $fixtureRoot -Filter '*.partial').Count -ne 0) { throw 'OCR left partial text files.' }
    Write-Output ('PASS: mixed text/scanned PDF recognized offline using ' + $status.Language + '; text/blank PDF reports blank page 2; entirely blank PDF fails without output; page boundaries and progress preserved; conflicting paths and invalid PDF refused without modifying source or disclosing content.')
}
finally {
    if ($KeepFiles) { Write-Output ('Synthetic fixture files: ' + $fixtureRoot) }
    else {
        $resolvedFixture = [IO.Path]::GetFullPath($fixtureRoot)
        $allowedRoot = [IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\') + '\'
        if (-not $resolvedFixture.StartsWith($allowedRoot, [StringComparison]::OrdinalIgnoreCase)) { throw 'Fixture cleanup target is outside the tests directory.' }
        Remove-Item -LiteralPath $resolvedFixture -Recurse -Force
    }
}
