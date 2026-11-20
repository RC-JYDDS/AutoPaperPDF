$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$libDir = Join-Path $root 'vendor/lib'
$packages = Get-Content -LiteralPath (Join-Path $root 'vendor/dependencies.lock.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$expected = @()
foreach ($package in $packages) {
    foreach ($name in $package.files) {
        if ([IO.Path]::GetFileName($name) -ne $name) { throw 'Invalid dependency filename.' }
        $path = Join-Path $libDir $name
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Missing dependency: $name" }
        $hash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        if ($hash -ne $package.sha256.PSObject.Properties[$name].Value) { throw "Dependency hash mismatch: $name" }
        $expected += $name
    }
}
foreach ($file in Get-ChildItem -LiteralPath $libDir -File) {
    if ($expected -notcontains $file.Name) { throw "Unlisted dependency file: $($file.Name)" }
}
Write-Output "Verified $($expected.Count) dependency files against the lock file."
