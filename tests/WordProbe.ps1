$ErrorActionPreference = 'Stop'
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class WordProbeNative {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
}
'@
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = $null
$document = $null
$owns = $false
try {
    'Create application'
    $word = New-Object -ComObject Word.Application
    'Read version'
    $word.Version
    'Read Documents'
    $documents = $word.Documents
    'Create blank document'
    $document = $documents.Add()
    'Read window'
    $window = $document.Windows.Item(1)
    $handle = $window.Hwnd
    [uint32]$ownedId = 0
    [void][WordProbeNative]::GetWindowThreadProcessId([IntPtr]$handle, [ref]$ownedId)
    if ($ownedId -and $before -notcontains $ownedId) { $owns = $true }
    'Verified owned process: ' + $owns + ', PID ' + $ownedId
} catch {
    $_.Exception.GetType().FullName
    $_.Exception.Message
    $_.Exception.HResult.ToString('X8')
    exit 1
} finally {
    $discard = 0
    if ($document) { try { $document.Close([ref]$discard) } catch { }; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    if ($owns) { $word.Quit([ref]$discard) }
    if ($word) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) }
}
