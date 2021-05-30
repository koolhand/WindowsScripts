# GetAuthenticodeOrHashes.ps1
# Luke Morey 2021-03-28

# Scan a folder for executable files and export CSV reports of...
# - authenticode signers, where available
# - where executable file is missing Authenticode signature, SHA-256 hash of each file (excluding zero byte files)
# - list of zero byte executable files

param (
  [string]$Directory    = 'C:\Program Files\Microsoft Office\',
  [string[]]$Extensions = @(
  '*.exe', '*.dll', '*.ocx', '*.sys',
  '*.ps1', '*.vbs', '*.wsf', # '*.js'
  '*.msi', '*.msp', '*.mst',
  '*.appx',
  '*.cab'
  ),
  [string]$OutputSigners       = 'Authenticode-Signers.csv',
  [string]$OutputHashes        = 'NoAuthenticode-Hashes.csv',
  [string]$OutputZeroByteFiles = 'ZeroByteFiles.csv',
  [string]$OutputDirectory     = '.\'
)

# Signatures

Write-Host 'Finding valid Authenticode signers...'

Get-ChildItem -Recurse -File -Path $Directory -Include $Extensions |
  Where-Object { $_.Length -ne 0 } |                 # exclude zero-byte files
  ForEach-Object { Get-AuthenticodeSignature $_ } |
  Where-Object { $_.status -eq "Valid" } |
  Select-Object -Unique `
      @{ Name='SignerIssuer';       Expression={($_.SignerCertificate.Issuer)} },
      @{ Name='SignerSerialNumber'; Expression={($_.SignerCertificate.SerialNumber)} },
      @{ Name='SignerThumbprint';   Expression={($_.SignerCertificate.Thumbprint)} } |
  Sort-Object SignerIssuer,SignerSerialNumber,SignerThumprint |
  Export-Csv (Join-Path $OutputDirectory $OutputSigners) -UseCulture -NoTypeInformation

Write-Host '...saved to' (Join-Path $OutputDirectory $OutputSigners)

# No valid signature

Write-Host ''
Write-Host 'Generating SHA-256 Hashes for files without valid Authenticode signatures...'

Get-ChildItem -Recurse -File -Path $Directory -Include $Extensions |
  Where-Object { $_.Length -ne 0 } |                  # exclude zero-byte files
  ForEach-Object { Get-AuthenticodeSignature $_ } |
  Where-Object { $_.status -ne "Valid" } |
  Select-Object Path |
  ForEach-Object { Get-FileHash -Path $_.Path } |
  Select-Object Hash,Path |
  Sort-Object Path,Hash |
  Export-Csv (Join-Path $OutputDirectory $OutputHashes) -UseCulture -NoTypeInformation

Write-Host '...saved to' (Join-Path $OutputDirectory $OutputHashes)

# zero-byte files

Write-Host ''
Write-Host 'Listing zero-byte files...'

Get-ChildItem -Recurse -File -Path $Directory -Include $Extensions |
  Where-Object { $_.Length -eq 0 } |
  Select-Object FullName |
  Export-Csv (Join-Path $OutputDirectory $OutputZeroByteFiles) -UseCulture -NoTypeInformation

Write-Host '...saved to' (Join-Path $OutputDirectory $OutputZeroByteFiles)
