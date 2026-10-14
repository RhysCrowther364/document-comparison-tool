$ErrorActionPreference = 'Stop'
$projectRoot = $PSScriptRoot
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path -LiteralPath $compiler)) { $compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe' }
$dist = Join-Path $projectRoot 'dist\LocalDocumentCompare'
New-Item -ItemType Directory -Path $dist -Force | Out-Null
$buildDirectory = Join-Path $projectRoot '.build'
New-Item -ItemType Directory -Path $buildDirectory -Force | Out-Null
$buildId = [guid]::NewGuid().ToString('N')
$compiledPath = Join-Path $buildDirectory ('DocumentCompare.' + $buildId + '.exe')
$sourceFiles = @(Get-ChildItem -LiteralPath (Join-Path $projectRoot 'src') -Filter '*.cs' | ForEach-Object FullName)
$references = @('System.dll', 'System.Core.dll', 'System.Drawing.dll', 'System.Windows.Forms.dll', 'System.Web.Extensions.dll', 'Microsoft.CSharp.dll', 'System.Xml.dll', 'System.Xml.Linq.dll', 'System.IO.Compression.dll', 'System.IO.Compression.FileSystem.dll')
$compilerArgs = @('/nologo', '/target:winexe', '/platform:anycpu', '/optimize+', '/utf8output', ('/out:' + $compiledPath))
$compilerArgs += '/win32icon:' + (Join-Path $projectRoot 'src\assets\document-compare.ico')
$compilerArgs += '/resource:' + (Join-Path $projectRoot 'src\assets\rws-logo.png') + ',LocalDocumentCompare.RwsLogo.png'
$compilerArgs += '/resource:' + (Join-Path $projectRoot 'src\assets\document-compare.ico') + ',LocalDocumentCompare.ComparisonIcon.ico'
foreach ($reference in $references) { $compilerArgs += '/reference:' + $reference }
& $compiler @compilerArgs @sourceFiles
if ($LASTEXITCODE -ne 0) { throw 'Compilation failed.' }
$publishedPath = Join-Path $dist 'DocumentCompare.exe'
if (Test-Path -LiteralPath $publishedPath) { Move-Item -LiteralPath $publishedPath -Destination (Join-Path $buildDirectory ('previous.' + $buildId + '.exe')) }
Move-Item -LiteralPath $compiledPath -Destination $publishedPath
Copy-Item -LiteralPath (Join-Path $projectRoot 'src\DocumentCompare.exe.config') -Destination $dist -Force
Copy-Item -LiteralPath (Join-Path $projectRoot 'src\assets\document-compare.ico') -Destination (Join-Path $dist 'DocumentCompare-icon-1.0.8.ico') -Force
New-Item -ItemType Directory -Path (Join-Path $dist 'scripts') -Force | Out-Null
Get-ChildItem -LiteralPath (Join-Path $projectRoot 'scripts') -File | Copy-Item -Destination (Join-Path $dist 'scripts') -Force
foreach ($file in @('README.md', 'VALIDATION.md', 'Install.cmd', 'Uninstall.cmd', 'DocumentCompare-QuickGuide.docx')) {
    $sourcePath = Join-Path $projectRoot $file
    if (Test-Path -LiteralPath $sourcePath) { Copy-Item -LiteralPath $sourcePath -Destination $dist -Force }
}
New-Item -ItemType Directory -Path (Join-Path $dist 'tests') -Force | Out-Null
foreach ($testFile in @('AcceptanceChecklist.md', 'PilotTestPlan.md', 'PilotTestPlan.docx', 'PilotResults.csv')) {
    Copy-Item -LiteralPath (Join-Path (Join-Path $projectRoot 'tests') $testFile) -Destination (Join-Path $dist 'tests') -Force
}
New-Item -ItemType Directory -Path (Join-Path $dist 'tests\PilotFixtures') -Force | Out-Null
Get-ChildItem -LiteralPath (Join-Path $projectRoot 'tests\PilotFixtures') -File | Copy-Item -Destination (Join-Path $dist 'tests\PilotFixtures') -Force
Write-Output ('Built ' + (Join-Path $dist 'DocumentCompare.exe'))
