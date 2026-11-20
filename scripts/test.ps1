$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$exe = Join-Path $root 'dist/unsigned/AutoPaperPDF.exe'
if (-not (Test-Path -LiteralPath $exe)) { throw 'Run scripts/build.ps1 first.' }
$previous = $env:AUTOPAPER_TEST_EXE
try {
    $env:AUTOPAPER_TEST_EXE = $exe
    & (Join-Path $root 'tests/core.tests.ps1')
    & (Join-Path $root 'tests/margins.tests.ps1')
} finally { $env:AUTOPAPER_TEST_EXE = $previous }
