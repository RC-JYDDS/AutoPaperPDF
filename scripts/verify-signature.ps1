param([Parameter(Mandatory=$true)][string]$Path)
$ErrorActionPreference = 'Stop'
$signature = Get-AuthenticodeSignature -LiteralPath $Path
if ($signature.Status -ne 'Valid') { throw "Signature is not valid: $($signature.Status)" }
if (-not $signature.SignerCertificate -or $signature.SignerCertificate.Subject -notmatch 'SignPath Foundation') {
    throw 'The certificate is not issued to SignPath Foundation.'
}
Write-Output 'Valid SignPath Foundation signature.'
