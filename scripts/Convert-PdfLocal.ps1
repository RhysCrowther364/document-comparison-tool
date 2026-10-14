#requires -Version 5.1
<#
Runs Windows' installed OCR engine over every PDF page. No network access,
Office automation, external executables, or package downloads are used.
Output is UTF-8 text, with a form-feed character between pages.
Invoke in Windows PowerShell 5.1, not PowerShell 7.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$OutputPath,
    [string]$StatusPath
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$utf8 = New-Object System.Text.UTF8Encoding($false)
$status = [ordered]@{ State = 'Starting'; PageCount = 0; PagesCompleted = 0; BlankPages = @(); Language = ''; Error = '' }
$failureMessage = 'Local OCR could not start. Run this helper in Windows PowerShell 5.1 on Windows 10 or 11.'
$writer = $null
$partialPath = $null
$statusFullPath = $null
$createdOutput = $false
$exitCode = 1

function Write-Status {
    if ($statusFullPath) {
        # The parent reads this as a best-effort progress snapshot; it retries if
        # it reads during a write. Only metadata is written, never document text.
        [IO.File]::WriteAllText($statusFullPath, ($status | ConvertTo-Json -Compress -Depth 3), $utf8)
    }
}

function Get-LocalPath([string]$Path) {
    $full = [IO.Path]::GetFullPath($Path)
    if ($full.StartsWith('\\') -or $full.Substring(2).Contains(':')) {
        throw 'Only normal local file paths are supported.'
    }
    $drive = New-Object IO.DriveInfo([IO.Path]::GetPathRoot($full))
    if ($drive.DriveType -eq [IO.DriveType]::Network) { throw 'Network drives are not supported.' }
    return $full
}

function Wait-WinRtOperation($Operation, [Type]$ResultType) {
    $task = $script:asTaskOperation.MakeGenericMethod($ResultType).Invoke($null, @($Operation))
    if (-not $task.Wait(120000)) {
        try { $Operation.Cancel() } catch { }
        throw 'The local OCR operation timed out.'
    }
    return $task.GetAwaiter().GetResult()
}

function Wait-WinRtAction($Action) {
    $task = $script:asTaskAction.Invoke($null, @($Action))
    if (-not $task.Wait(120000)) {
        try { $Action.Cancel() } catch { }
        throw 'The PDF rendering operation timed out.'
    }
    $null = $task.GetAwaiter().GetResult()
}

try {
    if ($PSVersionTable.PSEdition -ne 'Desktop') { throw 'Windows PowerShell 5.1 is required.' }
    $failureMessage = 'Choose a readable PDF and new output files on a local disk. Output and status must have different paths from the PDF.'
    $inputFullPath = Get-LocalPath $InputPath
    $outputFullPath = Get-LocalPath $OutputPath
    $candidateStatusPath = $null
    if ($StatusPath) { $candidateStatusPath = Get-LocalPath $StatusPath }
    if ($inputFullPath -ieq $outputFullPath -or $inputFullPath -ieq $candidateStatusPath -or $outputFullPath -ieq $candidateStatusPath) {
        throw 'Input, output, and status paths must differ.'
    }
    $statusFullPath = $candidateStatusPath
    if (-not [IO.File]::Exists($inputFullPath) -or [IO.Path]::GetExtension($inputFullPath) -ine '.pdf') { throw 'PDF not found.' }
    if ([IO.File]::Exists($outputFullPath)) { throw 'Output already exists.' }
    if (-not [IO.Directory]::Exists([IO.Path]::GetDirectoryName($outputFullPath))) { throw 'Output directory is missing.' }
    Write-Status

    $failureMessage = 'Windows local PDF/OCR components are unavailable. Use Windows 10 or 11 and Windows PowerShell 5.1.'
    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
    $null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
    $null = [Windows.Data.Pdf.PdfPageRenderOptions, Windows.Data.Pdf, ContentType = WindowsRuntime]
    $null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
    $null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
    $null = [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
    $null = [Windows.Graphics.Imaging.BitmapPixelFormat, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
    $null = [Windows.Graphics.Imaging.BitmapAlphaMode, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
    $null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
    $null = [Windows.Media.Ocr.OcrResult, Windows.Foundation, ContentType = WindowsRuntime]
    $asTaskOperation = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
        $_.Name -eq 'AsTask' -and $_.IsGenericMethodDefinition -and $_.GetParameters().Count -eq 1 -and
        $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
    } | Select-Object -First 1
    $asTaskAction = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
        $_.Name -eq 'AsTask' -and -not $_.IsGenericMethod -and $_.GetParameters().Count -eq 1 -and
        $_.GetParameters()[0].ParameterType.FullName -eq 'Windows.Foundation.IAsyncAction'
    } | Select-Object -First 1

    $failureMessage = 'No Windows OCR language is installed. Ask IT to install the English OCR language feature in Windows Settings, then retry. This tool does not download language packs.'
    $languages = @([Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages)
    $preferred = $null
    foreach ($tag in @('en-GB', 'en-US')) {
        $preferred = $languages | Where-Object { $_.LanguageTag -eq $tag } | Select-Object -First 1
        if ($preferred) { break }
    }
    $engine = $null
    if ($preferred) { $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($preferred) }
    if (-not $engine) { $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages() }
    if (-not $engine -and $languages.Count -gt 0) { $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($languages[0]) }
    if (-not $engine) { throw 'No OCR engine available.' }
    $status.Language = $engine.RecognizerLanguage.LanguageTag

    $failureMessage = 'Windows could not open this PDF. Check that it is a valid, readable PDF and is not password protected.'
    $file = Wait-WinRtOperation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($inputFullPath)) ([Windows.Storage.StorageFile])
    $pdf = Wait-WinRtOperation ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
    $status.PageCount = [int]$pdf.PageCount
    if ($status.PageCount -eq 0) { throw 'The PDF has no pages.' }
    $status.State = 'Recognizing'
    Write-Status

    $failureMessage = 'Local OCR could not create its output. Check local disk space and folder permissions.'
    $partialPath = $outputFullPath + '.' + [Guid]::NewGuid().ToString('N') + '.partial'
    $writer = New-Object IO.StreamWriter($partialPath, $false, $utf8)
    $maxDimension = [double][Windows.Media.Ocr.OcrEngine]::MaxImageDimension
    for ($index = 0; $index -lt $status.PageCount; $index++) {
        $page = $null
        $stream = $null
        $bitmap = $null
        try {
            $failureMessage = 'Local OCR failed while rendering or reading PDF page ' + ($index + 1) + '. The PDF may be damaged or too complex. No partial output has been kept.'
            $page = $pdf.GetPage([uint32]$index)
            $width = [double]$page.Size.Width
            $height = [double]$page.Size.Height
            if ($width -le 0 -or $height -le 0) { throw 'Invalid page dimensions.' }
            # PdfPage.Size uses 96-DPI DIPs. Scale to approximately 200 DPI and
            # respect both Windows OCR's dimension cap and a 20M pixel budget.
            $scale = [Math]::Min(200.0 / 96.0, $maxDimension / [Math]::Max($width, $height))
            $scale = [Math]::Min($scale, [Math]::Sqrt(20000000.0 / ($width * $height)))
            $options = New-Object Windows.Data.Pdf.PdfPageRenderOptions
            $options.DestinationWidth = [uint32][Math]::Max(1, [Math]::Floor($width * $scale))
            $options.DestinationHeight = [uint32][Math]::Max(1, [Math]::Floor($height * $scale))
            $stream = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
            Wait-WinRtAction ($page.RenderToStreamAsync($stream, $options))
            $null = $stream.Seek(0)
            $decoder = Wait-WinRtOperation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
            $bitmap = Wait-WinRtOperation ($decoder.GetSoftwareBitmapAsync([Windows.Graphics.Imaging.BitmapPixelFormat]::Bgra8, [Windows.Graphics.Imaging.BitmapAlphaMode]::Ignore)) ([Windows.Graphics.Imaging.SoftwareBitmap])
            $result = Wait-WinRtOperation ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
            $lines = @($result.Lines | ForEach-Object { $_.Text })
            $pageText = $lines -join [Environment]::NewLine
            if ([string]::IsNullOrWhiteSpace($pageText)) { $status.BlankPages += ($index + 1) }
            if ($index -gt 0) { $writer.Write([char]12) }
            $writer.Write($pageText)
            $writer.Flush()
            $status.PagesCompleted = $index + 1
            Write-Status
        }
        finally {
            if ($bitmap) { $bitmap.Dispose() }
            if ($stream) { $stream.Dispose() }
            if ($page) { $page.Dispose() }
        }
    }
    $writer.Dispose()
    $writer = $null
    if ($status.BlankPages.Count -eq $status.PageCount) {
        $failureMessage = 'Windows OCR found no readable text. Check scan quality, page orientation, and the installed OCR language.'
        throw 'No readable text.'
    }
    $failureMessage = 'Local OCR could not finish saving its text output. Check local disk space and folder permissions.'
    [IO.File]::Move($partialPath, $outputFullPath)
    $createdOutput = $true
    $partialPath = $null
    $status.State = 'Complete'
    Write-Status
    $exitCode = 0
}
catch {
    # Native exception text may contain file paths or document data. Keep error
    # messages to a fixed, useful description of the failed processing stage.
    $status.State = 'Error'
    $status.Error = $failureMessage
    try { Write-Status } catch { }
    [Console]::Error.WriteLine($failureMessage)
}
finally {
    if ($writer) { $writer.Dispose() }
    if ($partialPath -and [IO.File]::Exists($partialPath)) { [IO.File]::Delete($partialPath) }
    # If the final status write failed after the atomic move, do not leave an
    # apparently usable result behind. This flag is set only for our own file.
    if ($exitCode -ne 0 -and $createdOutput -and [IO.File]::Exists($outputFullPath)) { [IO.File]::Delete($outputFullPath) }
}
exit $exitCode
