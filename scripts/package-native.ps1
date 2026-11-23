param()

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$outputRoot = Join-Path $projectRoot 'outputs'
$portable = Join-Path $outputRoot 'LocalPDF-portable'
$stage = Join-Path $projectRoot 'work\native-installer-stage'
$appOutput = Join-Path $projectRoot 'app\bin\Release\net48'
$makensis = Join-Path $env:LOCALAPPDATA 'tauri\NSIS\makensis.exe'
$icon = Join-Path $projectRoot 'app\Assets\icon.ico'
$installer = Join-Path $outputRoot 'LocalPDF_1.0.0_x64-setup.exe'

& (Join-Path $PSScriptRoot 'build-native.ps1') -Configuration Release
if (-not (Test-Path -LiteralPath $makensis)) { throw 'NSIS compiler is unavailable.' }

foreach ($directory in @($portable, $stage)) {
    $full = [IO.Path]::GetFullPath($directory)
    if (-not $full.StartsWith([IO.Path]::GetFullPath($projectRoot) + [IO.Path]::DirectorySeparatorChar)) {
        throw "Refusing to replace directory outside the project: $full"
    }
    if (Test-Path -LiteralPath $full) { Remove-Item -LiteralPath $full -Recurse -Force }
    New-Item -ItemType Directory -Path $full | Out-Null
}

$releaseFiles = Get-ChildItem -LiteralPath $appOutput -File | Where-Object {
    $_.Extension -in @('.exe', '.dll', '.config')
}
foreach ($file in $releaseFiles) {
    Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $stage $file.Name) -Force
    Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $portable $file.Name) -Force
}

$zip = Join-Path $outputRoot 'LocalPDF-portable.zip'
if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }
Compress-Archive -Path (Join-Path $portable '*') -DestinationPath $zip -CompressionLevel Optimal

& $makensis "/DSTAGE=$stage" "/DOUTPUT=$installer" "/DICON=$icon" (Join-Path $projectRoot 'installer\LocalPDF.nsi')
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $installer)) { throw 'NSIS packaging failed.' }

Write-Host "Portable build: $portable"
Write-Host "Installer: $installer"
