<#
WinImgNormalizer.ps1
Non-destructive image normalizer + archive prep for mixed photo folders

Author: Janne Vuorela
Target OS: Windows 10/11
PowerShell: Windows PowerShell 5.1+, also works on PowerShell 7
Dependencies: ImageMagick 7+ (magick.exe in PATH), optional: .bat wrapper for drag-and-drop

SYNOPSIS
    Recursively mirrors a source folder into a safe copy under the user’s Pictures folder,
    converts all images to JPEG ≤ 1 MB (auto-orient, strip metadata, flatten alpha),
    copies videos as-is, skips duplicates by (filename + LastWriteTimeUtc), shows progress in terminal,
    and writes a detailed log file.

WHAT THIS IS (AND ISN’T)
    - Personal, purpose-built tool for my archive workflow.
      It trades knobs for reliability, speed, and repeatability.
    - Designed for drag-and-drop via the .bat wrapper, also works from PowerShell directly.
    - Not a deduplication/content-hashing system, duplicate detection is lightweight
      (filename + timestamp), good enough for typical camera roll structures.

FEATURES
    - Non-destructive: creates "<SourceName>_WinImgNormalized_<yyyyMMdd_HHmmss>" under Pictures.
    - Exact tree mirror: same subfolders, image files become .jpeg (same base names).
    - Supported images via ImageMagick: JPG/JPEG/PNG/BMP/TIF/TIFF/GIF/HEIC/HEIF/WebP.
    - Videos copied as-is (mp4/mov/mkv/avi/m4v/wmv/webm/mts/m2ts/3gp/3g2).
    - JPEG ≤ 1 MB targeting with progressive scaling (100→50%) and quality sizing (jpeg:extent).
    - EXIF auto-orientation, metadata stripped, alpha flattened to white.
    - Duplicate skip: (filename lowercase + LastWriteTimeUtc ticks) anywhere in the tree.
    - Progress bar in terminal; detailed timestamped log in the destination folder.

MY INTENDED USAGE
    - I drag a mixed photo/video/whatever folder onto WinImgNormalizer.bat.
    - The script writes a normalized copy to %USERPROFILE%\Pictures\… and a log file.
    - I keep the original source intact.

SETUP
    1) Install ImageMagick 7+ for Windows and ensure `magick.exe` is on PATH.
    2) Keep these two files together (same base name):
         • WinImgNormalizer.ps1
         • WinImgNormalizer.bat  (enables drag-and-drop)
    3) Optional: run in PowerShell 7 for slightly better performance; 5.1 is supported.

USAGE
    A) Drag & Drop (recommended)
       - Drag a folder onto WinImgNormalizer.bat.
       - Output: %USERPROFILE%\Pictures\<Source>_WinImgNormalized_<timestamp>
    B) Direct PowerShell (positional args only; avoids PS 5.1 param-set quirks)
       - .\WinImgNormalizer.ps1 "D:\Photos\2024"
       - .\WinImgNormalizer.ps1 "D:\Photos\2024" 1048576
        custom size cap (bytes)

NOTES
    - HEIC/WebP support depends on ImageMagick build/codecs.
    - “Invalid SOS parameters for sequential JPEG” and similar libjpeg warnings are handled
      (suppressed via -quiet, script treats them as non-fatal).
    - Multi-frame GIF/TIFF are flattened to a single frame.
    - Filenames are preserved, only the extension changes to .jpeg for images.
    - Timestamps on outputs are set to the source file times.

LIMITATIONS
    - Size targeting is best-effort, extremely noisy or huge images may not reach ≤ 1 MB even at 50%.
    - Duplicate detection is not content-hash based.
    - Metadata is stripped by design, this is an archive/preview-friendly normalization pass.

TROUBLESHOOTING
    - "magick not found": install ImageMagick; ensure magick.exe is in PATH (check `magick -version`).
    - PS 5.1 “Parameter set cannot be resolved”: this script uses positional arguments by design.
      Always launch via the provided .bat or pass the folder path positionally.
    - Excessive JPEG warnings: the script runs ImageMagick with `-quiet` and ignores warnings;
      results are still size-checked and logged.
    - No outputs created: check the log file in the destination for per-file errors.

LICENSE / WARRANTY
    - Personal tool; provided as-is, without warranty. Use at your own risk.

#>


$ErrorActionPreference = 'Stop'

# --------- Args, positional only ---
$Source   = $null
$MaxBytes = 1MB
if ($args.Count -ge 1) { $Source   = $args[0] }
if ($args.Count -ge 2) { $MaxBytes = [int64]$args[1] }

if ([string]::IsNullOrWhiteSpace($Source)) { Write-Host "Usage: WinImgNormalizer.ps1 <sourceFolder> [maxBytes]"; exit 1 }
if (-not (Test-Path -LiteralPath $Source -PathType Container)) { Write-Error "Source folder not found: $Source"; exit 1 }
$srcRoot = (Get-Item -LiteralPath $Source).FullName.TrimEnd('\','/')

# --------- Logger, literal-safe ---
$script:LogPath = $null
function Write-Log {
  param([string]$Message, [ValidateSet('INFO','OK','SKIP','WARN','ERR')]$Level='INFO')
  $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
  switch ($Level) { 'ERR'{Write-Host $line -ForegroundColor Red}; 'WARN'{Write-Host $line -ForegroundColor Yellow}; 'OK'{Write-Host $line -ForegroundColor Green}; default{Write-Host $line} }
  if ($script:LogPath) { try { Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8 } catch {} }
}

# --------- Resolve Pictures + dest root ----
try { $pictures = [Environment]::GetFolderPath('MyPictures') } catch { $pictures = $null }
if ([string]::IsNullOrWhiteSpace($pictures) -or -not (Test-Path -LiteralPath $pictures)) {
  $pictures = Join-Path $env:USERPROFILE 'Pictures'
  if (-not (Test-Path -LiteralPath $pictures)) { New-Item -ItemType Directory -Path $pictures -Force | Out-Null }
}

$base  = [System.IO.Path]::GetFileName($srcRoot.TrimEnd('\','/'))
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$destRoot = [System.IO.Path]::Combine($pictures, "${base}_WinImgNormalized_${stamp}")
New-Item -ItemType Directory -Path $destRoot -Force | Out-Null

# --------- Log file ---
try {
  $script:LogPath = [System.IO.Path]::Combine($destRoot, "WinImgNormalizer_${stamp}.log")
  "WinImgNormalizer started $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Set-Content -LiteralPath $script:LogPath -Encoding UTF8
} catch {
  $script:LogPath = [System.IO.Path]::Combine($env:TEMP, "WinImgNormalizer_${stamp}.log")
  "WinImgNormalizer started $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') (TEMP fallback log)" | Set-Content -LiteralPath $script:LogPath -Encoding UTF8
}
Write-Log "Source: $srcRoot"
Write-Log "Destination: $destRoot"
Write-Log ("MaxBytes: {0:n0} ({1} MB)" -f $MaxBytes, [Math]::Round($MaxBytes/1MB,2))

# --------- ImageMagick presence ---
$MagickCmd = $null
$cmd = Get-Command magick -ErrorAction SilentlyContinue
if ($cmd) { $MagickCmd = $cmd.Source; if (-not $MagickCmd) { $MagickCmd = $cmd.Path }; if (-not $MagickCmd) { $MagickCmd = $cmd.Definition } }
if (-not $MagickCmd) { $MagickCmd = 'magick' }
try { $null = & $MagickCmd -version 2>$null } catch { Write-Error "ImageMagick 'magick' not found on PATH."; exit 1 }

# --------- Mirror tree, non-destructive ---
try {
  Get-ChildItem -LiteralPath $srcRoot -Recurse -Directory | ForEach-Object {
    $rel = $_.FullName.Substring($srcRoot.Length).TrimStart('\','/')
    $target = [System.IO.Path]::Combine($destRoot, $rel)
    New-Item -ItemType Directory -Path $target -Force | Out-Null
  }
} catch { Write-Log "WARN: Could not fully pre-create directory tree ($($_.Exception.Message))" 'WARN' }

# --------- File sets ---
$imgExts   = '.jpg','.jpeg','.png','.bmp','.tif','.tiff','.gif','.heic','.heif','.webp'
$videoExts = '.mp4','.mov','.mkv','.avi','.m4v','.wmv','.webm','.mts','.m2ts','.3gp','.3g2'

$allFiles = Get-ChildItem -LiteralPath $srcRoot -Recurse -File | Where-Object {
  $e = $_.Extension.ToLowerInvariant()
  ($imgExts -contains $e) -or ($videoExts -contains $e)
}
$total = $allFiles.Count
if ($total -eq 0) { Write-Log "No images or videos found." 'WARN'; exit 0 }

# --------- Dedupe + stats ---
$seen = New-Object 'System.Collections.Generic.HashSet[string]'
$keyFirst = @{}
$stats = [ordered]@{ Converted=0; CopiedVideo=0; SkippedDuplicate=0; Unsupported=0; Errors=0 }

# --------- IM helpers ---
function Get-ExtentString([long]$Bytes) {
  if ($Bytes -ge 1MB -and ($Bytes % 1MB) -eq 0) { return "{0}MB" -f [int]($Bytes/1MB) }
  elseif ($Bytes -ge 1KB -and ($Bytes % 1KB) -eq 0) { return "{0}KB" -f [int]($Bytes/1KB) }
  else { return "{0}B" -f $Bytes }
}
function Convert-ImageMagick {
  param(
    [string]$SourcePath,
    [string]$DestPath,
    [long]$MaxBytes
  )

  $extent = Get-ExtentString $MaxBytes
  $scales = 100,90,80,70,60,50

  foreach ($p in $scales) {
    # Ensure destination directory exists
    $destDir = [System.IO.Path]::GetDirectoryName($DestPath)
    if ($destDir -and -not (Test-Path -LiteralPath $destDir)) {
      New-Item -ItemType Directory -Path $destDir -Force | Out-Null
    }

    # Primary attempt, handles alpha -> white
    $args = @(
      '-quiet',
      $SourcePath,
      '-auto-orient','-strip','-colorspace','sRGB',
      '-sampling-factor','4:2:0','-interlace','Line',
      '-background','white','-alpha','remove','-alpha','off',
      '-resize', "$p%",
      '-define', "jpeg:extent=$extent",
      ('JPEG:' + $DestPath)
    )

    # Run ImageMagick without escalating warnings to terminating errors
    $prevEAP = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
      $LASTEXITCODE = 0
      & $MagickCmd @args 1>$null 2>$null
    } finally {
      $ErrorActionPreference = $prevEAP
    }

    # Retry without alpha flags if magick errored or produced nothing
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $DestPath)) {
      $args2 = @(
        '-quiet',
        $SourcePath,
        '-auto-orient','-strip','-colorspace','sRGB',
        '-sampling-factor','4:2:0','-interlace','Line',
        '-resize', "$p%",
        '-define', "jpeg:extent=$extent",
        ('JPEG:' + $DestPath)
      )
      $prevEAP = $ErrorActionPreference
      $ErrorActionPreference = 'Continue'
      try {
        $LASTEXITCODE = 0
        & $MagickCmd @args2 1>$null 2>$null
      } finally {
        $ErrorActionPreference = $prevEAP
      }
    }

    # Check size target
    if (Test-Path -LiteralPath $DestPath) {
      $len = (Get-Item -LiteralPath $DestPath).Length
      if ($len -le $MaxBytes) {
        return @{ Status='Converted'; BytesOut=$len; Scale=$p; Note=$null }
      }
      # Too big? try next smaller scale.
    }
  }

  # Best-effort fallback if we have an output but couldn't hit the cap
  if (Test-Path -LiteralPath $DestPath) {
    $len = (Get-Item -LiteralPath $DestPath).Length
    return @{ Status='Converted'; BytesOut=$len; Scale=50; Note='WARN: Could not reach target; best-effort saved' }
  } else {
    return @{ Status='Error'; Note='ImageMagick conversion failed' }
  }
}

# --------- Main loop ---
[int]$i = 0
foreach ($f in $allFiles) {
  $i++
  $rel = $f.FullName.Substring($srcRoot.Length).TrimStart('\','/')
  $ext = $f.Extension.ToLowerInvariant()
  Write-Progress -Activity "WinImgNormalizer" -Status "$i / $total : $rel" -PercentComplete ([int]($i*100/$total))

  $dupKey = ('{0}|{1}' -f $f.Name.ToLowerInvariant(), $f.LastWriteTimeUtc.Ticks)
  if (-not $seen.Add($dupKey)) { $stats.SkippedDuplicate++; $firstRel = $keyFirst[$dupKey]; Write-Log "Duplicate skipped: $rel (first seen at: $firstRel)" 'SKIP'; continue } else { $keyFirst[$dupKey] = $rel }

  $destRel  = if ($imgExts -contains $ext) { [System.IO.Path]::ChangeExtension($rel, '.jpeg') } else { $rel }
  $destPath = [System.IO.Path]::Combine($destRoot, $destRel)

  try {
    if ($imgExts -contains $ext) {
      $res = Convert-ImageMagick -SourcePath $f.FullName -DestPath $destPath -MaxBytes $MaxBytes
      if ($res.Status -eq 'Converted') {
        $stats.Converted++
        try { (Get-Item -LiteralPath $destPath).LastWriteTimeUtc = $f.LastWriteTimeUtc; (Get-Item -LiteralPath $destPath).CreationTimeUtc = $f.CreationTimeUtc } catch {}
        $note = if ($res.Note) { " ($($res.Note))" } else { "" }
        Write-Log ("OK IMG: {0} -> {1} [{2:n0} bytes, Scale={3}%]{4}" -f $rel, $destRel, $res.BytesOut, $res.Scale, $note) 'OK'
      } else {
        $stats.Errors++; Write-Log "ERR IMG: $rel ($($res.Note))" 'ERR'
      }
    }
    elseif ($videoExts -contains $ext) {
      $destDir = [System.IO.Path]::GetDirectoryName($destPath)
      if ($destDir -and -not (Test-Path -LiteralPath $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
      Copy-Item -LiteralPath $f.FullName -Destination $destPath -Force
      try { (Get-Item -LiteralPath $destPath).LastWriteTimeUtc = $f.LastWriteTimeUtc; (Get-Item -LiteralPath $destPath).CreationTimeUtc = $f.CreationTimeUtc } catch {}
      $bytesOut = (Get-Item -LiteralPath $destPath).Length
      $stats.CopiedVideo++; Write-Log ("OK VID: {0} -> {1} [{2:n0} bytes]" -f $rel, $destRel, $bytesOut) 'OK'
    }
    else {
      $stats.Unsupported++; Write-Log "Unsupported skipped: $rel" 'WARN'
    }
  } catch {
    $stats.Errors++; Write-Log "Exception processing: $rel ($($_.Exception.Message))" 'ERR'
  }
}

Write-Progress -Activity "WinImgNormalizer" -Completed

Write-Host ""
Write-Host "Summary:" -ForegroundColor Cyan
Write-Host ("  Converted images : {0}" -f $stats.Converted)
Write-Host ("  Copied videos    : {0}" -f $stats.CopiedVideo)
Write-Host ("  Duplicates       : {0}" -f $stats.SkippedDuplicate)
Write-Host ("  Unsupported      : {0}" -f $stats.Unsupported)
Write-Host ("  Errors           : {0}" -f $stats.Errors)
Write-Host ("  Log file         : {0}" -f $script:LogPath)

Write-Log ("SUMMARY ConvertedImages={0} CopiedVideos={1} Duplicates={2} Unsupported={3} Errors={4}" -f $stats.Converted,$stats.CopiedVideo,$stats.SkippedDuplicate,$stats.Unsupported,$stats.Errors)
Write-Log "Completed $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
