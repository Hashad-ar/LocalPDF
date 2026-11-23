param(
    [ValidateSet('Debug', 'Release')][string]$Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$cargo = Join-Path $env:USERPROFILE '.cargo\bin\cargo.exe'
$targetRoot = Join-Path $env:TEMP 'LocalPDF-build-target'
$referenceRoot = Join-Path $env:USERPROFILE '.nuget\packages\microsoft.netframework.referenceassemblies.net48\1.0.3\build\.NETFramework\v4.8'

if (-not (Test-Path -LiteralPath $cargo)) { throw 'The Rust toolchain is not installed.' }
$env:CARGO_TARGET_DIR = $targetRoot
$env:FrameworkPathOverride = $referenceRoot

if ($Configuration -eq 'Release') {
    & $cargo +1.88.0 build --manifest-path (Join-Path $projectRoot 'word-helper\Cargo.toml') --release
    $cargoFolder = 'release'
} else {
    & $cargo +1.88.0 build --manifest-path (Join-Path $projectRoot 'word-helper\Cargo.toml')
    $cargoFolder = 'debug'
}
if ($LASTEXITCODE -ne 0) { throw 'The Word helper build failed.' }

& dotnet build (Join-Path $projectRoot 'app\LocalPDF.csproj') -c $Configuration --no-incremental
if ($LASTEXITCODE -ne 0) { throw 'The WPF build failed.' }

$appOutput = Join-Path $projectRoot "app\bin\$Configuration\net48"
Copy-Item -LiteralPath (Join-Path $targetRoot "$cargoFolder\LocalPDF-WordHelper.exe") -Destination (Join-Path $appOutput 'LocalPDF-WordHelper.exe') -Force
Write-Host "Native LocalPDF build ready: $appOutput"
