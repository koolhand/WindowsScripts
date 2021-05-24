# GetAuthenticodeOrHashes.ps1
# Luke Morey 2021-03-28

# For a given folder, export a report of...
# - authenticode signers where available
# - SHA-256 hash of executabl files where Authenticode signature is missing, excluding zero byte files
# - list of zero byte files

param (
  [string]$ScanDirectory = 'C:\Program Files\Microsoft Office\',
  [string[]]$IncludeFiles = @(
  '*.exe', '*.dll', '*.ocx', '*.sys',
  '*.ps1', '*.vbs', '*.js', '*.wsf',
  '*.msi', '*.msp', '*.mst',
  '*.appx',
  '*.cab'
  ),
  [string]$CSVSigners       = 'Authenticode-Signers.csv',
  [string]$CSVHashes        = 'NoAuthenticode-Hashes.csv',
  [string]$CSVZeroByteFiles = 'ZeroByteFiles.csv',
  [string]$ReportDirectory  = '.\'
)

# Signatures

Write-Host 'Finding valid Authenticode signers...'

Get-ChildItem -Recurse -File -Path $ScanDirectory -Include $IncludeFiles |
  Where-Object { $_.Length -ne 0 } |                 # exclude zero-byte files
  ForEach-Object { Get-AuthenticodeSignature $_ } |
  Where-Object { $_.status -eq "Valid" } |
  Select-Object -Unique `
      @{ Name='SignerIssuer';       Expression={($_.SignerCertificate.Issuer)} },
      @{ Name='SignerSerialNumber'; Expression={($_.SignerCertificate.SerialNumber)} },
      @{ Name='SignerThumbprint';   Expression={($_.SignerCertificate.Thumbprint)} } |
  Sort-Object SignerIssuer,SignerSerialNumber,SignerThumprint |
  Export-Csv (Join-Path $ReportDirectory $CSVSigners) -UseCulture -NoTypeInformation

Write-Host '...saved to' (Join-Path $ReportDirectory $CSVSigners)

# No valid signature

Write-Host ''
Write-Host 'Generating SHA-256 Hashes for files without valid Authenticode signatures...'

Get-ChildItem -Recurse -File -Path $ScanDirectory -Include $IncludeFiles |
  Where-Object { $_.Length -ne 0 } |                  # exclude zero-byte files
  ForEach-Object { Get-AuthenticodeSignature $_ } |
  Where-Object { $_.status -ne "Valid" } |
  Select-Object Path |
  ForEach-Object { Get-FileHash -Path $_.Path } |
  Select-Object Hash,Path |
  Sort-Object Path,Hash |
  Export-Csv (Join-Path $ReportDirectory $CSVHashes) -UseCulture -NoTypeInformation

Write-Host '...saved to' (Join-Path $ReportDirectory $CSVHashes)

# zero-byte files

Write-Host ''
Write-Host 'Listing zero-byte files...'

Get-ChildItem -Recurse -File -Path $ScanDirectory -Include $IncludeFiles |
  Where-Object { $_.Length -eq 0 } |
  Select-Object FullName |
  Export-Csv (Join-Path $ReportDirectory $CSVZeroByteFiles) -UseCulture -NoTypeInformation

Write-Host '...saved to' (Join-Path $ReportDirectory $CSVZeroByteFiles)
