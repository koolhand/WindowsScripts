<#
.SYNOPSIS
  Moves PNG screenshots created on a given date into a yyyy-MM-dd subfolder (e.g., 2025-09-13)
  under the Screenshots directory.

.DESCRIPTION
  - Screenshots directory is derived authoritatively from the Pictures known folder:
      [Environment]::GetFolderPath('MyPictures')  ->  <Pictures>\Screenshots
    with fallback to the User Shell Folders registry value for Pictures if necessary.
  - Creates the date subfolder if missing.
  - Moves only PNG files in the top-level Screenshots folder whose CreationTime matches the date.
  - Handles filename collisions by suffixing " (n)".

.PARAMETER Date
  The date whose files you want to move. Defaults to 2025-09-13.

.PARAMETER ScreenshotsRoot
  Optional explicit path to the Screenshots folder, which overrides the derived location.

.EXAMPLE
  .\Move-ScreenshotsByDate.ps1 -Date '2025-09-13' -WhatIf

.EXAMPLE
  .\Move-ScreenshotsByDate.ps1 -Date '2025-09-13'
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param(
  [Parameter(Mandatory = $false)]
  [DateTime]$Date = [DateTime]::Parse('2025-09-13'),

  [Parameter(Mandatory = $false)]
  [string]$ScreenshotsRoot
)

function Get-PicturesPath {
  # 1) Preferred: Known Folder (respects user redirection, including OneDrive Pictures)
  $path = [Environment]::GetFolderPath('MyPictures')
  if (-not [string]::IsNullOrWhiteSpace($path)) {
    $expanded = [Environment]::ExpandEnvironmentVariables($path)
    if (Test-Path -LiteralPath $expanded) { return (Resolve-Path -LiteralPath $expanded).Path }
  }

  # 2) Fallback: Registry (User Shell Folders)
  $regKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
  $candidates = @(
    '{33E28130-4E1E-4676-835A-98395C3BC3BB}', # Pictures KNOWNFOLDERID
    'My Pictures'
  )

  foreach ($name in $candidates) {
    try {
      $val = (Get-ItemProperty -LiteralPath $regKey -ErrorAction Stop).$name
      if ([string]::IsNullOrWhiteSpace($val)) { continue }
      $expanded = [Environment]::ExpandEnvironmentVariables($val)
      if (Test-Path -LiteralPath $expanded) { return (Resolve-Path -LiteralPath $expanded).Path }
    } catch { }
  }

  throw "Unable to resolve the Pictures folder from Known Folders or registry."
}

function Resolve-ScreenshotsRoot {
  param([string]$Override)

  if (-not [string]::IsNullOrWhiteSpace($Override)) {
    if (-not (Test-Path -LiteralPath $Override)) {
      throw "Provided ScreenshotsRoot does not exist: $Override"
    }
    return (Resolve-Path -LiteralPath $Override).Path
  }

  $pictures = Get-PicturesPath
  $screens = Join-Path $pictures 'Screenshots'
  if (-not (Test-Path -LiteralPath $screens)) {
    throw "Screenshots folder not found at: $screens. Create it (e.g., take a Win+PrtScn screenshot once) or specify -ScreenshotsRoot."
  }
  return (Resolve-Path -LiteralPath $screens).Path
}

try {
  $root = Resolve-ScreenshotsRoot -Override $ScreenshotsRoot
  Write-Host "Using Screenshots folder: $root" -ForegroundColor Cyan
  Write-Host "Target date: $($Date.ToString('yyyy-MM-dd'))" -ForegroundColor Cyan

  # Ensure the target yyyy-MM-dd subfolder exists (create if missing).
  $targetFolderName = $Date.ToString('yyyy-MM-dd')
  $targetPath = Join-Path $root $targetFolderName

  if (-not (Test-Path -LiteralPath $targetPath)) {
    if ($PSCmdlet.ShouldProcess($targetPath, "Create folder")) {
      New-Item -ItemType Directory -Path $targetPath -Force | Out-Null
      Write-Host "Created folder: $targetPath" -ForegroundColor Green
    } else {
      Write-Host "[WhatIf] Would create folder: $targetPath" -ForegroundColor Yellow
    }
  } else {
    Write-Host "Folder already exists: $targetPath" -ForegroundColor DarkGray
  }

  # Enumerate PNGs in Screenshots root (non-recursive) matching CreationTime.Date.
  $pngs = Get-ChildItem -LiteralPath $root -Filter *.png -File |
          Where-Object { $_.CreationTime.Date -eq $Date.Date }

  if (-not $pngs) {
    Write-Host "No PNG files with CreationTime on $($Date.ToString('yyyy-MM-dd')) found in: $root" -ForegroundColor Yellow
    return
  }

  $moved = 0
  foreach ($f in $pngs) {
    # Skip if already in the target folder.
    if ($f.DirectoryName -eq $targetPath) {
      Write-Host "Already in target folder: $($f.Name)" -ForegroundColor DarkGray
      continue
    }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
    $ext      = $f.Extension
    $destPath = Join-Path $targetPath $f.Name

    # Resolve name collisions by adding (n) suffix.
    $i = 1
    while (Test-Path -LiteralPath $destPath) {
      $destPath = Join-Path $targetPath ("{0} ({1}){2}" -f $baseName, $i, $ext)
      $i++
    }

    if ($PSCmdlet.ShouldProcess($destPath, "Move '$($f.FullName)'")) {
      Move-Item -LiteralPath $f.FullName -Destination $destPath
      Write-Host "Moved: $($f.Name)  ->  $targetFolderName\$([IO.Path]::GetFileName($destPath))" -ForegroundColor Green
      $moved++
    } else {
      Write-Host "[WhatIf] Move: '$($f.FullName)' -> '$destPath'" -ForegroundColor Yellow
    }
  }

  Write-Host ""
  Write-Host ("Done. {0} file(s) matched; {1} moved." -f $pngs.Count, $moved) -ForegroundColor Cyan

} catch {
  Write-Error $_
  exit 1
}
