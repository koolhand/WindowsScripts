<#
.SYNOPSIS
  Organizes PNG screenshots into yyyy-MM-dd subfolders based on file date.

.DESCRIPTION
  - Default behavior (no parameters): For each PNG in the Screenshots folder root,
    determine the CreationTime.Date, create a yyyy-MM-dd subfolder if missing, and
    move the file into that folder. Skips files already in the correct folder.
  - Screenshots directory is derived authoritatively from the Pictures known folder:
      [Environment]::GetFolderPath('MyPictures')  ->  <Pictures>\Screenshots
    with fallback to the User Shell Folders registry value for Pictures if necessary.
  - Supports optional limiting to one date, recursion, using LastWriteTime instead
    of CreationTime, and an explicit Screenshots path override.

.PARAMETER Date
  Optional. If provided, only files whose date equals this value are moved.
  Date is compared at day precision (yyyy-MM-dd). If omitted, all PNG files
  are grouped by their respective dates.

.PARAMETER ScreenshotsRoot
  Optional explicit path to the Screenshots folder, which overrides the derived location.

.PARAMETER Recurse
  Include PNG files from subfolders as well.

.PARAMETER UseLastWriteTime
  Use LastWriteTime.Date instead of CreationTime.Date.

.EXAMPLE
  # Default: organize all PNGs in the root by CreationTime.Date
  .\Organize-ScreenshotsByDate.ps1

.EXAMPLE
  # Dry run (see actions without making changes)
  .\Organize-ScreenshotsByDate.ps1 -WhatIf

.EXAMPLE
  # Only organize files for 2025-09-13
  .\Organize-ScreenshotsByDate.ps1 -Date '2025-09-13'

.EXAMPLE
  # Recurse into subfolders and use LastWriteTime
  .\Organize-ScreenshotsByDate.ps1 -Recurse -UseLastWriteTime

.EXAMPLE
  # Override the screenshots folder
  .\Organize-ScreenshotsByDate.ps1 -ScreenshotsRoot 'D:\Media\Screenshots'
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param(
  [Parameter(Mandatory = $false)]
  [Nullable[DateTime]]$Date,  # Null => organize all by their own dates

  [Parameter(Mandatory = $false)]
  [string]$ScreenshotsRoot,

  [switch]$Recurse,

  [switch]$UseLastWriteTime
)

function Get-PicturesPath {
  # 1) Preferred: Known Folder (respects user redirection, including OneDrive KFM)
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
    throw "Screenshots folder not found at: $screens. Create it (take a Win+PrtScn screenshot once) or specify -ScreenshotsRoot."
  }
  return (Resolve-Path -LiteralPath $screens).Path
}

function Get-DateForFile {
  param([System.IO.FileInfo]$File, [switch]$UseLastWriteTime)
  if ($UseLastWriteTime) { return $File.LastWriteTime.Date }
  return $File.CreationTime.Date
}

function Get-UniqueDestinationPath {
  param([string]$Folder, [string]$FileName)
  $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
  $ext      = [System.IO.Path]::GetExtension($FileName)
  $destPath = Join-Path $Folder $FileName
  $i = 1
  while (Test-Path -LiteralPath $destPath) {
    $destPath = Join-Path $Folder ("{0} ({1}){2}" -f $baseName, $i, $ext)
    $i++
  }
  return $destPath
}

try {
  $root = Resolve-ScreenshotsRoot -Override $ScreenshotsRoot
  Write-Host "Using Screenshots folder: $root" -ForegroundColor Cyan

  # Gather candidate PNGs. Default is non-recursive to avoid reorganizing already-sorted items.
  $files = Get-ChildItem -LiteralPath $root -Filter *.png -File -Recurse:$Recurse

  if (-not $files) {
    Write-Host "No PNG files found in: $root" -ForegroundColor Yellow
    return
  }

  $total = 0
  $moved = 0
  $skippedAlreadyCorrect = 0
  $perDateCounts = @{}

  foreach ($f in $files) {
    $fileDate = Get-DateForFile -File $f -UseLastWriteTime:$UseLastWriteTime
    if ($Date.HasValue -and $fileDate -ne $Date.Value.Date) {
      continue # honoring a user-specified single date
    }

    $folderName = $fileDate.ToString('yyyy-MM-dd')
    $targetFolder = Join-Path $root $folderName

    # If the file is already in the correct date folder, skip
    if ($f.DirectoryName -eq $targetFolder) {
      $skippedAlreadyCorrect++
      continue
    }

    # Ensure date folder exists
    if (-not (Test-Path -LiteralPath $targetFolder)) {
      if ($PSCmdlet.ShouldProcess($targetFolder, "Create folder")) {
        New-Item -ItemType Directory -Path $targetFolder -Force | Out-Null
        Write-Host "Created folder: $targetFolder" -ForegroundColor Green
      } else {
        Write-Host "[WhatIf] Would create folder: $targetFolder" -ForegroundColor Yellow
      }
    }

    $destPath = Get-UniqueDestinationPath -Folder $targetFolder -FileName $f.Name

    if ($PSCmdlet.ShouldProcess($destPath, "Move '$($f.FullName)'")) {
      Move-Item -LiteralPath $f.FullName -Destination $destPath
      Write-Host ("Moved: {0}  ->  {1}\{2}" -f $f.Name, [IO.Path]::GetFileName($targetFolder), [IO.Path]::GetFileName($destPath)) -ForegroundColor Green
      $moved++

      if ($perDateCounts.ContainsKey($folderName)) { $perDateCounts[$folderName]++ }
      else { $perDateCounts[$folderName] = 1 }
    } else {
      Write-Host ("[WhatIf] Move: '{0}' -> '{1}'" -f $f.FullName, $destPath) -ForegroundColor Yellow
    }

    $total++
  }

  Write-Host ""
  if ($Date) {
    Write-Host ("Summary for {0}:" -f $Date.Value.ToString('yyyy-MM-dd')) -ForegroundColor Cyan
  } else {
    Write-Host "Summary (all dates):" -ForegroundColor Cyan
  }

  Write-Host ("Processed candidates: {0}" -f $total)
  Write-Host ("Moved: {0}" -f $moved)
  Write-Host ("Skipped (already in correct folder): {0}" -f $skippedAlreadyCorrect)

  if ($perDateCounts.Keys.Count -gt 0) {
    Write-Host ""
    Write-Host "Moved by date:" -ForegroundColor Cyan
    $perDateCounts.GetEnumerator() | Sort-Object Name | ForEach-Object {
      Write-Host ("  {0} : {1}" -f $_.Name, $_.Value)
    }
  }

} catch {
  Write-Error $_
  exit 1
}
