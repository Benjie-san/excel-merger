param(
  [string]$ReleaseDir = "release",
  [switch]$FailOnUnsigned
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-LatestFile {
  param(
    [string]$PathPattern
  )

  $file = Get-ChildItem -Path $PathPattern -File -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

  return $file
}

function Get-FileSecuritySummary {
  param(
    [System.IO.FileInfo]$File
  )

  if (-not $File) {
    return $null
  }

  $hash = Get-FileHash -Path $File.FullName -Algorithm SHA256
  $sig = Get-AuthenticodeSignature -FilePath $File.FullName
  $sizeMb = [Math]::Round(($File.Length / 1MB), 2)
  $signer = ""
  if ($sig.SignerCertificate) {
    $signer = $sig.SignerCertificate.Subject
  }

  return [pscustomobject]@{
    FileName = $File.Name
    Path = $File.FullName
    SizeMB = $sizeMb
    SHA256 = $hash.Hash
    SignatureStatus = [string]$sig.Status
    Signer = $signer
    LastWriteTime = $File.LastWriteTime
  }
}

if (-not (Test-Path $ReleaseDir)) {
  throw "Release directory not found: $ReleaseDir"
}

$installer = Resolve-LatestFile -PathPattern (Join-Path $ReleaseDir "*Setup*.exe")
$appExe = Resolve-LatestFile -PathPattern (Join-Path $ReleaseDir "win-unpacked\\*.exe")

if (-not $installer) {
  throw "Installer not found in $ReleaseDir. Run npm run desktop:dist first."
}

if (-not $appExe) {
  throw "Packed app exe not found in $ReleaseDir\\win-unpacked."
}

$results = @(
  Get-FileSecuritySummary -File $installer
  Get-FileSecuritySummary -File $appExe
) | Where-Object { $_ -ne $null }

Write-Host ""
Write-Host "Desktop Release Verification"
Write-Host "============================"
Write-Host ""
$results | Format-Table FileName, SizeMB, SignatureStatus, LastWriteTime -AutoSize

Write-Host ""
foreach ($entry in $results) {
  Write-Host "$($entry.FileName)"
  Write-Host "  Path: $($entry.Path)"
  Write-Host "  SHA256: $($entry.SHA256)"
  if ($entry.Signer) {
    Write-Host "  Signer: $($entry.Signer)"
  } else {
    Write-Host "  Signer: (none)"
  }
  Write-Host ""
}

# Corruption guardrail: tiny EXEs are likely broken artifacts.
$tooSmall = @($results | Where-Object { $_.SizeMB -lt 5 })
if ($tooSmall.Count -gt 0) {
  throw "Potential corruption detected: one or more EXEs are smaller than 5 MB."
}

if ($FailOnUnsigned) {
  $unsigned = @($results | Where-Object { $_.SignatureStatus -ne "Valid" })
  if ($unsigned.Count -gt 0) {
    throw "Signature check failed: one or more files are not Authenticode Valid."
  }
}

Write-Host "Verification complete."
