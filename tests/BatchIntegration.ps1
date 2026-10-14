param([switch]$SkipBuild)
$ErrorActionPreference = 'Stop'
$project = Split-Path -Parent $PSScriptRoot
if (-not $SkipBuild) { & (Join-Path $project 'build.ps1') }
$release = Join-Path $project 'dist\LocalDocumentCompare\DocumentCompare.exe'
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$artifactRoot = Join-Path $PSScriptRoot ('artifacts\batch-' + [guid]::NewGuid().ToString('N'))
$testExe = Join-Path (Join-Path $project '.build') ('BatchUiIntegration-' + [guid]::NewGuid().ToString('N') + '.exe')
$references = @('System.dll', 'System.Core.dll', 'System.Drawing.dll', 'System.Windows.Forms.dll', 'System.IO.Compression.dll', 'System.IO.Compression.FileSystem.dll', $release)
$arguments = @('/nologo', '/target:exe', '/platform:anycpu', '/utf8output', ('/out:' + $testExe))
foreach ($reference in $references) { $arguments += '/reference:' + $reference }
& $compiler @arguments (Join-Path $PSScriptRoot 'BatchUiIntegration.cs')
if ($LASTEXITCODE -ne 0) { throw 'Batch integration test compilation failed.' }
# .NET resolves the shipped types from the test executable's directory.
# Copy only to this workspace build directory; no installed files are changed.
Copy-Item -LiteralPath $release -Destination (Join-Path $project '.build\DocumentCompare.exe') -Force
Copy-Item -LiteralPath (Join-Path (Split-Path -Parent $release) 'DocumentCompare.exe.config') -Destination (Join-Path $project '.build\DocumentCompare.exe.config') -Force
New-Item -ItemType Directory -Path (Join-Path $project '.build\scripts') -Force | Out-Null
Copy-Item -LiteralPath (Join-Path (Split-Path -Parent $release) 'scripts\Convert-PdfLocal.ps1') -Destination (Join-Path $project '.build\scripts\Convert-PdfLocal.ps1') -Force
& $testExe $release $artifactRoot
if ($LASTEXITCODE -ne 0) { throw ('Batch UI integration failed. Artifacts: ' + $artifactRoot) }
