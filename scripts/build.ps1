param(
    [string]$FrameworkReferencePath = "${env:ProgramFiles(x86)}\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.7.2"
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2
$root = Split-Path $PSScriptRoot -Parent
$source = Join-Path $root 'src'
$libs = Join-Path $root 'vendor/lib'
$csc = Join-Path $env:WINDIR 'Microsoft.NET/Framework/v4.0.30319/csc.exe'
if (-not (Test-Path -LiteralPath $csc)) { throw 'Windows .NET Framework C# compiler is required.' }
if (-not (Test-Path -LiteralPath (Join-Path $FrameworkReferencePath 'Facades/netstandard.dll'))) {
    throw '.NET Framework 4.7.2 Developer Pack is required. Supply -FrameworkReferencePath if installed elsewhere.'
}
& (Join-Path $PSScriptRoot 'verify-dependencies.ps1')
$info = Get-Content -LiteralPath (Join-Path $source 'AppInfo.cs') -Raw -Encoding UTF8
$version = [regex]::Match($info, 'const string Version = "([0-9.]+)"').Groups[1].Value
$fileVersion = [regex]::Match($info, 'const string AssemblyVersion = "([0-9.]+)"').Groups[1].Value
$date = [regex]::Match($info, 'const string ReleaseDate = "([0-9]+)"').Groups[1].Value
if (-not $version -or -not $fileVersion -or -not $date) { throw 'Missing version metadata.' }
$outDir = Join-Path $root 'dist/unsigned'
New-Item -ItemType Directory -Path $outDir -Force | Out-Null
$output = Join-Path $outDir 'AutoPaperPDF.exe'
$arguments = @('/nologo','/target:winexe','/platform:anycpu','/optimize+','/nostdlib+','/langversion:5','/codepage:65001',
    "/out:$output", "/win32manifest:$source\app.manifest", "/win32icon:$source\assets\app.ico")
foreach ($name in @('mscorlib','System','System.Core','System.Drawing','System.Windows.Forms','Microsoft.CSharp','System.Runtime.Serialization','System.Xml','Facades/netstandard')) {
    $arguments += '/reference:' + (Join-Path $FrameworkReferencePath ($name + '.dll'))
}
foreach ($file in Get-ChildItem -LiteralPath $libs -Filter *.dll | Sort-Object Name) {
    $arguments += '/reference:' + $file.FullName
    $arguments += '/resource:' + $file.FullName + ',AutoPaper.Dependencies.' + $file.Name
}
$arguments += '/resource:' + (Join-Path $root 'THIRD_PARTY_NOTICES.txt') + ',AutoPaper.ThirdPartyNotices.txt'
foreach ($file in Get-ChildItem -LiteralPath $source -Filter *.cs | Sort-Object Name) { $arguments += $file.FullName }
& $csc @arguments
if ($LASTEXITCODE -ne 0) { throw 'Compilation failed.' }
$metadata = [Diagnostics.FileVersionInfo]::GetVersionInfo($output)
if ($metadata.FileVersion -ne $fileVersion -or $metadata.ProductVersion -ne "$version+$date" -or $metadata.ProductName -ne '通用表格 PDF 自动选纸工具') {
    throw 'Executable version or product metadata does not match source.'
}
$hash = (Get-FileHash -LiteralPath $output -Algorithm SHA256).Hash.ToLowerInvariant()
"$hash  AutoPaperPDF.exe" | Set-Content -LiteralPath (Join-Path $root 'dist/SHA256SUMS-unsigned.txt') -Encoding UTF8
if ($env:GITHUB_OUTPUT) {
    "version=$version" | Add-Content -LiteralPath $env:GITHUB_OUTPUT -Encoding UTF8
    "file_version=$fileVersion" | Add-Content -LiteralPath $env:GITHUB_OUTPUT -Encoding UTF8
    "product_version=$version+$date" | Add-Content -LiteralPath $env:GITHUB_OUTPUT -Encoding UTF8
}
Write-Output "Built unsigned executable: $output"
