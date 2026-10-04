<#
WinImgNormalizer.ps1
Non-destructive image normalizer + archive prep for mixed photo folders

Author: Janne Vuorela
Target OS: Windows 10/11
PowerShell: Windows PowerShell 5.1+, also works on PowerShell 7
Dependencies: ImageMagick 7.1.2-32+ supported 7.x (magick.exe in PATH), optional: .bat wrapper

SYNOPSIS
    Recursively mirrors a source folder into a safe copy under the user’s Pictures folder,
    converts all images to JPEG ≤ 1 MB (auto-orient, strip metadata, flatten alpha),
    copies videos as-is, skips heuristic duplicates by filename/time plus equal length,
    shows progress in terminal,
    and writes a detailed log file.

WHAT THIS IS (AND ISN’T)
    - Personal, purpose-built tool for my archive workflow.
      It trades knobs for reliability, speed, and repeatability.
    - Designed for drag-and-drop via the .bat wrapper, also works from PowerShell directly.
    - Not a deduplication/content-hashing system, duplicate detection is lightweight
      (filename + timestamp + equal byte length); different content can still match.

FEATURES
    - Non-destructive: creates "<SourceName>_WinImgNormalized_<yyyyMMdd_HHmmss>_<runId>" under Pictures.
    - Exact tree mirror: same subfolders, images become .jpeg (same base names when unique).
    - Deterministic collision suffixes; complete output plan; no replacement of arriving targets.
    - Supported images via ImageMagick: JPG/JPEG/PNG/BMP/TIF/TIFF/GIF/HEIC/HEIF/WebP.
    - Videos copied as-is (mp4/mov/mkv/avi/m4v/wmv/webm/mts/m2ts/3gp/3g2).
    - JPEG ≤ 1 MB targeting with progressive scaling (100→50%) and quality sizing (jpeg:extent).
    - EXIF auto-orientation, metadata stripped, alpha flattened to white.
    - Heuristic skips match lowercase filename + LastWriteTimeUtc + input byte length.
      Only finalized successes register; every skip links its retained source and output.
    - Progress bar in terminal; detailed timestamped log in the destination folder.

MY INTENDED USAGE
    - I drag a mixed photo/video/whatever folder onto WinImgNormalizer.bat.
    - The script writes a normalized copy to %USERPROFILE%\Pictures\… and a log file.
    - I keep the original source intact.

SETUP
    1) Install ImageMagick 7.1.2-32 or newer supported 7.x; ensure `magick.exe` is on PATH.
    2) Keep these two files together (same base name):
         • WinImgNormalizer.ps1
         • WinImgNormalizer.bat  (enables drag-and-drop)
    3) Optional: run in PowerShell 7 for slightly better performance; 5.1 is supported.

USAGE
    A) Drag & Drop (recommended)
       - Drag a folder onto WinImgNormalizer.bat.
       - Output: %USERPROFILE%\Pictures\<Source>_WinImgNormalized_<timestamp>_<runId>
    B) Direct PowerShell (positional args only; avoids PS 5.1 param-set quirks)
       - .\WinImgNormalizer.ps1 "D:\Photos\2024"
       - .\WinImgNormalizer.ps1 "D:\Photos\2024" 1048576
        custom size cap (bytes)

NOTES
    - HEIC/WebP support depends on ImageMagick build/codecs.
    - Each fresh conversion attempt must exit successfully and fully decode as
      a nonempty single-frame JPEG with positive dimensions before finalization.
    - Animations become their first displayed frame; TIFF uses its first page.
      HEIC/HEIF uses the decoder's primary/first image. Counts/omissions are logged.
    - Unique image basenames are preserved; conflicts use __sourceext and checked __N suffixes.
    - Timestamps on outputs are set to the source file times.

LIMITATIONS
    - Size targeting is best-effort, extremely noisy or huge images may not reach ≤ 1 MB even at 50%.
    - Duplicate detection is not content-hash based.
    - Metadata is stripped by design, this is an archive/preview-friendly normalization pass.

TROUBLESHOOTING
    - "magick not found": install ImageMagick; ensure magick.exe is in PATH (check `magick -version`).
    - PS 5.1 “Parameter set cannot be resolved”: this script uses positional arguments by design.
      Always launch via the provided .bat or pass the folder path positionally.
    - Conversion uses `-quiet`; candidate validation promotes decode warnings to
      failure so recovered/truncated JPEGs cannot be finalized.
    - No outputs created: check the log file in the destination for per-file errors.

LICENSE / WARRANTY
    - Personal tool; provided as-is, without warranty. Use at your own risk.

#>


# Preflight helpers never create a run directory. Keep roots absolute, including
# their separator: C:\ is a root, while C: means the drive's current directory.
function Normalize-WinImgRootPath {
  param([string]$Path)
  if ([string]::IsNullOrWhiteSpace($Path) -or $Path -match '^[A-Za-z]:($|[^\\/])') {
    throw 'Use a directory path, not a drive-relative path such as C:.'
  }
  $full = [IO.Path]::GetFullPath($Path)
  $root = [IO.Path]::GetPathRoot($full)
  if ($full.TrimEnd('\','/').Equals($root.TrimEnd('\','/'), [StringComparison]::OrdinalIgnoreCase)) {
    return $root.TrimEnd('\','/') + [IO.Path]::DirectorySeparatorChar
  }
  return $full.TrimEnd('\','/')
}

# Windows path comparisons deliberately use OrdinalIgnoreCase, including on a
# directory configured for case sensitivity: a conservative containment policy.
function Test-WinImgPathContained {
  param([string]$Root, [string]$Path)
  $rootPath = Normalize-WinImgRootPath $Root
  $candidate = Normalize-WinImgRootPath $Path
  $prefix = $rootPath.TrimEnd('\','/') + [IO.Path]::DirectorySeparatorChar
  return $candidate.Equals($rootPath, [StringComparison]::OrdinalIgnoreCase) -or
    $candidate.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)
}

function Initialize-WinImgDirectoryApi {
  if ('WinImgNormalizer.NativeDirectory' -as [type]) { return }
  # Loaded only when the application runs; importing definitions has no effects.
  # Both APIs exist on the supported Windows/Windows PowerShell versions.
  Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;
namespace WinImgNormalizer {
  public static class NativeDirectory {
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern SafeFileHandle CreateFileW(string path, uint access, uint share,
      IntPtr security, uint creation, uint flags, IntPtr template);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern uint GetFinalPathNameByHandleW(SafeFileHandle handle,
      StringBuilder path, uint capacity, uint flags);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CreateDirectoryW(string path, IntPtr security);
    private static string NativePath(string path) {
      // Generated descendants can exceed MAX_PATH even with ordinary input roots.
      // Prefix only internally validated absolute paths; do not shorten user names.
      if (path.StartsWith(@"\\", StringComparison.Ordinal)) return @"\\?\UNC\" + path.Substring(2);
      return @"\\?\" + path;
    }
    public static string CanonicalPath(string path) {
      using (SafeFileHandle handle = CreateFileW(NativePath(path), 0, 7, IntPtr.Zero, 3, 0x02000000, IntPtr.Zero)) {
        if (handle.IsInvalid) throw new Win32Exception(Marshal.GetLastWin32Error());
        StringBuilder buffer = new StringBuilder(512);
        uint length = GetFinalPathNameByHandleW(handle, buffer, (uint)buffer.Capacity, 0);
        if (length == 0) throw new Win32Exception(Marshal.GetLastWin32Error());
        if (length >= buffer.Capacity) {
          buffer = new StringBuilder(checked((int)length + 1));
          length = GetFinalPathNameByHandleW(handle, buffer, (uint)buffer.Capacity, 0);
          if (length == 0 || length >= buffer.Capacity) throw new Win32Exception(Marshal.GetLastWin32Error());
        }
        string result = buffer.ToString();
        if (result.StartsWith(@"\\?\UNC\", StringComparison.OrdinalIgnoreCase)) return @"\\" + result.Substring(8);
        if (result.StartsWith(@"\\?\", StringComparison.Ordinal)) return result.Substring(4);
        throw new InvalidOperationException("Windows returned an unsupported canonical directory path.");
      }
    }
    public static bool CreateExclusive(string path) {
      if (CreateDirectoryW(NativePath(path), IntPtr.Zero)) return true;
      int error = Marshal.GetLastWin32Error();
      if (error == 183 || error == 80) return false;
      throw new Win32Exception(error);
    }
  }
}
'@
}

function Assert-WinImgNoReparseAncestors {
  param([string]$Path)
  $current = Normalize-WinImgRootPath $Path
  while ($current) {
    $item = $null
    try { $item = Get-Item -LiteralPath $current -Force -ErrorAction Stop }
    catch {
      # Test-Path can hide a dangling link because its target does not exist.
      # Inspect the entry itself; only truly missing components may be appended.
      if ($_.CategoryInfo.Category -ne [Management.Automation.ErrorCategory]::ObjectNotFound) { throw }
    }
    if ($null -ne $item) {
      if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw 'A linked path or ancestor is not supported; select the actual directory.'
      }
      if ($item -isnot [IO.DirectoryInfo]) { throw 'The directory path has a non-directory component.' }
    }
    $parent = [IO.Directory]::GetParent($current)
    if ($null -eq $parent) { break }
    $current = $parent.FullName
  }
}

function Resolve-WinImgCanonicalDirectory {
  param([string]$Path)
  $full = Normalize-WinImgRootPath $Path
  # Avoid Win32 aliases caused by trailing dots/spaces and special device syntax.
  if ($full -match '^\\\\[?.]\\' -or
      $full.Substring([IO.Path]::GetPathRoot($full).Length) -match '(^|[\\/])[^\\/]*[. ]([\\/]|$)|:') {
    throw 'Use an ordinary directory path without device syntax or trailing dots/spaces.'
  }
  Assert-WinImgNoReparseAncestors $full
  $ancestor = $full
  $missing = New-Object 'System.Collections.Generic.Stack[string]'
  while (-not (Test-Path -LiteralPath $ancestor -ErrorAction Stop)) {
    $missing.Push([IO.Path]::GetFileName($ancestor))
    $parent = [IO.Directory]::GetParent($ancestor)
    if ($null -eq $parent) { throw 'The destination volume/share does not exist.' }
    $ancestor = $parent.FullName
  }
  Initialize-WinImgDirectoryApi
  $canonical = [WinImgNormalizer.NativeDirectory]::CanonicalPath($ancestor)
  while ($missing.Count -gt 0) { $canonical = [IO.Path]::Combine($canonical, $missing.Pop()) }
  $canonical = Normalize-WinImgRootPath $canonical
  Assert-WinImgNoReparseAncestors $canonical
  return $canonical
}

function Assert-WinImgSafeDestination {
  param([string]$SourceRoot, [string]$OutputParent)
  $sourcePath = Resolve-WinImgCanonicalDirectory $SourceRoot
  $parentPath = Resolve-WinImgCanonicalDirectory $OutputParent
  if (Test-WinImgPathContained -Root $sourcePath -Path $parentPath) {
    throw 'The destination is inside or equal to the source. Choose a source subfolder outside the destination tree (for example, a subfolder of Pictures).'
  }
  return $parentPath
}

function Get-WinImgRelativePath {
  param([string]$Root, [string]$Path)
  if (-not (Test-WinImgPathContained -Root $Root -Path $Path) -or
      (Normalize-WinImgRootPath $Path).Equals((Normalize-WinImgRootPath $Root), [StringComparison]::OrdinalIgnoreCase)) {
    throw 'An enumerated path escaped its source root.'
  }
  return $Path.Substring($Root.TrimEnd('\','/').Length).TrimStart('\','/')
}

function Get-WinImgSourceTree {
  param([string]$SourceRoot)
  $files = New-Object 'System.Collections.Generic.List[System.IO.FileInfo]'
  $directories = New-Object 'System.Collections.Generic.List[System.IO.DirectoryInfo]'
  $names = New-Object 'System.Collections.Generic.List[string]'
  $warnings = New-Object 'System.Collections.Generic.List[object]'
  $pending = New-Object 'System.Collections.Generic.Stack[string]'
  $pending.Push($SourceRoot)
  while ($pending.Count -gt 0) {
    $directory = $pending.Pop()
    try {
      Assert-WinImgNoReparseAncestors $directory
      $paths = New-Object 'System.Collections.Generic.List[string]'
      foreach ($entry in @(Get-ChildItem -LiteralPath $directory -Force -ErrorAction Stop)) { $paths.Add($entry.FullName) }
      $paths.Sort([StringComparer]::OrdinalIgnoreCase)
      foreach ($path in $paths) {
        try {
          $entry = Get-Item -LiteralPath $path -Force -ErrorAction Stop
          $null = Get-WinImgRelativePath -Root $SourceRoot -Path $entry.FullName
          if ($directory.Equals($SourceRoot, [StringComparison]::OrdinalIgnoreCase)) { $names.Add($entry.Name) }
          if (($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            $warnings.Add([pscustomobject]@{ Path = $entry.FullName; Reason = 'Reparse point skipped; links and junctions are not followed.' })
          } elseif ($entry -is [IO.DirectoryInfo]) {
            $directories.Add($entry)
            $pending.Push($entry.FullName)
          } elseif ($entry -is [IO.FileInfo]) { $files.Add($entry) }
        } catch { $warnings.Add([pscustomobject]@{ Path = $path; Reason = 'Could not inspect source entry: ' + $_.Exception.Message }) }
      }
    } catch { $warnings.Add([pscustomobject]@{ Path = $directory; Reason = 'Incomplete source scan: ' + $_.Exception.Message }) }
  }
  return [pscustomobject]@{ Files = $files.ToArray(); Directories = $directories.ToArray(); TopLevelNames = $names.ToArray(); Warnings = $warnings.ToArray() }
}

function Get-WinImgGeneratedName {
  param([string[]]$TopLevelNames)
  $reserved = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
  foreach ($name in $TopLevelNames) { $null = $reserved.Add($name) }
  $candidate = '.WinImgNormalizer'
  $suffix = 2
  while ($reserved.Contains($candidate)) { $candidate = '.WinImgNormalizer__' + $suffix; $suffix++ }
  return $candidate
}

function Get-WinImgOutputPlan {
  param(
    [string]$SourceRoot,
    [object]$SourceTree,
    [string]$GeneratedName,
    [string[]]$ImageExtensions = @('.jpg','.jpeg','.png','.bmp','.tif','.tiff','.gif','.heic','.heif','.webp'),
    [string[]]$VideoExtensions = @('.mp4','.mov','.mkv','.avi','.m4v','.wmv','.webm','.mts','.m2ts','.3gp','.3g2')
  )

  # Plan the whole namespace before assigning suffixes. Reserve unique legacy
  # names first so a suffix cannot steal a later source's nonconflicting name.
  $reserved = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
  foreach ($directory in $SourceTree.Directories) {
    $null = $reserved.Add((Get-WinImgRelativePath -Root $SourceRoot -Path $directory.FullName))
  }
  foreach ($path in @($GeneratedName, [IO.Path]::Combine($GeneratedName, 'work'), [IO.Path]::Combine($GeneratedName, 'reports'))) {
    $null = $reserved.Add($path)
  }
  $rows = New-Object 'Collections.Generic.Dictionary[string,object]' ([StringComparer]::Ordinal)
  $preferredCounts = New-Object 'Collections.Generic.Dictionary[string,int]' ([StringComparer]::OrdinalIgnoreCase)
  $orderedPaths = New-Object 'Collections.Generic.List[string]'
  foreach ($file in $SourceTree.Files) {
    $extension = $file.Extension.ToLowerInvariant()
    if ($ImageExtensions -notcontains $extension -and $VideoExtensions -notcontains $extension) { continue }
    $relative = Get-WinImgRelativePath -Root $SourceRoot -Path $file.FullName
    $kind = if ($ImageExtensions -contains $extension) { 'Image' } else { 'Video' }
    $preferred = if ($kind -eq 'Image') { [IO.Path]::ChangeExtension($relative, '.jpeg') } else { $relative }
    $rows.Add($relative, [pscustomobject]@{
      Source = $file; SourceRelativePath = $relative; OutputRelativePath = $preferred
      Kind = $kind; NamingReason = 'Legacy'
    })
    $orderedPaths.Add($relative)
    if (-not $preferredCounts.ContainsKey($preferred)) { $preferredCounts.Add($preferred, 0) }
    $preferredCounts[$preferred]++
  }
  # Ordinal is a total order, including case-only variants, independent of locale
  # and the filesystem's enumeration order in either supported PowerShell host.
  $orderedPaths.Sort([StringComparer]::Ordinal)
  $collisions = New-Object 'Collections.Generic.List[object]'
  foreach ($relative in $orderedPaths) {
    $row = $rows[$relative]
    if ($preferredCounts[$row.OutputRelativePath] -eq 1 -and $reserved.Add($row.OutputRelativePath)) { continue }
    $collisions.Add($row)
  }
  foreach ($row in $collisions) {
    $directory = [IO.Path]::GetDirectoryName($row.SourceRelativePath)
    $stem = [IO.Path]::GetFileNameWithoutExtension($row.SourceRelativePath)
    $extension = $row.Source.Extension.ToLowerInvariant()
    $outputExtension = if ($row.Kind -eq 'Image') { '.jpeg' } else { $extension }
    $suffixStem = $stem + '__' + $extension.TrimStart('.')
    $candidate = [IO.Path]::Combine($directory, $suffixStem + $outputExtension)
    $row.NamingReason = 'ExtensionSuffix'
    $number = 2
    while (-not $reserved.Add($candidate)) {
      $candidate = [IO.Path]::Combine($directory, $suffixStem + '__' + $number + $outputExtension)
      $row.NamingReason = 'NumericSuffix'
      $number++
    }
    $row.OutputRelativePath = $candidate
  }
  foreach ($relative in $orderedPaths) { $rows[$relative] }
}

function Assert-WinImgOutputAvailable {
  param([string]$Path)
  Assert-WinImgNoReparseAncestors ([IO.Path]::GetDirectoryName($Path))
  if (Test-Path -LiteralPath $Path) { throw 'The planned output is already occupied; the existing entry was preserved.' }
}

function New-WinImgImageCandidate {
  param([string]$WorkRoot)
  Assert-WinImgNoReparseAncestors $WorkRoot
  for ($attempt = 0; $attempt -lt 8; $attempt++) {
    $directory = [IO.Path]::Combine($WorkRoot, [guid]::NewGuid().ToString('N'))
    if (New-WinImgExclusiveDirectory $directory) {
      return [IO.Path]::Combine($directory, 'image.jpeg')
    }
  }
  throw 'Could not allocate an exclusive image scratch directory after eight name collisions.'
}

function Get-WinImgNativeOutputPath {
  param([string]$Path)
  # Only internal, absolute candidate paths reach this helper. ImageMagick's
  # Windows file APIs need the extended prefix for long generated scratch paths.
  if ($Path.Length -lt 260) { return $Path }
  if ($Path.StartsWith('\\', [StringComparison]::Ordinal)) { return '\\?\UNC\' + $Path.Substring(2) }
  return '\\?\' + $Path
}

function Get-WinImgDuplicateKey {
  param([string]$Name, [datetime]$LastWriteTimeUtc, [long]$Length)
  # The existing filename/time heuristic gains an equal-length guard. The
  # composite key retains separate first successes for each input byte length.
  return '{0}|{1}|{2}' -f $Name.ToLowerInvariant(),
    $LastWriteTimeUtc.Ticks.ToString([Globalization.CultureInfo]::InvariantCulture),
    $Length.ToString([Globalization.CultureInfo]::InvariantCulture)
}

function Move-WinImgPlannedImage {
  param([string]$CandidatePath, [string]$DestinationPath)
  Assert-WinImgNoReparseAncestors ([IO.Path]::GetDirectoryName($CandidatePath))
  if (-not [IO.Path]::GetPathRoot($CandidatePath).Equals([IO.Path]::GetPathRoot($DestinationPath), [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Candidate and final output must be on the same volume.'
  }
  Assert-WinImgOutputAvailable $DestinationPath
  # The two-argument overload also refuses an arrival after the preceding check.
  [IO.File]::Move($CandidatePath, $DestinationPath)
}

function Copy-WinImgPlannedVideo {
  param([string]$SourcePath, [string]$DestinationPath, [string]$WorkRoot)
  if (-not $WorkRoot) { $WorkRoot = [IO.Path]::GetDirectoryName($DestinationPath) }
  $allocated = New-WinImgImageCandidate -WorkRoot $WorkRoot
  $candidate = [IO.Path]::Combine([IO.Path]::GetDirectoryName($allocated), 'video.partial')
  $cleanupWarning = $null
  $fileOwned = $false
  try {
    Assert-WinImgNoReparseAncestors ([IO.Path]::GetDirectoryName($SourcePath))
    $before = Get-Item -LiteralPath $SourcePath -Force -ErrorAction Stop
    if ($before.PSIsContainer -or ($before.Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Video source is not a regular file.' }
    $length = $before.Length
    $modified = $before.LastWriteTimeUtc
    $created = $before.CreationTimeUtc
    $reservation = [IO.FileStream]::new($candidate, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    $fileOwned = $true
    $reservation.Dispose()
    Copy-WinImgVideoToCandidate -SourcePath $SourcePath -CandidatePath $candidate
    $after = Get-Item -LiteralPath $SourcePath -Force -ErrorAction Stop
    $partial = Get-Item -LiteralPath $candidate -Force -ErrorAction Stop
    if ($after.PSIsContainer -or ($after.Attributes -band [IO.FileAttributes]::ReparsePoint) -or
        $partial.PSIsContainer -or ($partial.Attributes -band [IO.FileAttributes]::ReparsePoint) -or
        $after.Length -ne $length -or $after.LastWriteTimeUtc -ne $modified -or $partial.Length -ne $length) {
      throw 'Video copy is incomplete or the source changed during copying; no final output was committed.'
    }
    Move-WinImgPlannedImage -CandidatePath $candidate -DestinationPath $DestinationPath
  } finally {
    try { Remove-WinImgOwnedCandidate -CandidatePath $candidate -FileOwned $fileOwned }
    catch { $cleanupWarning = $_.Exception.Message; Write-Warning ('Could not remove owned video scratch: ' + $cleanupWarning) }
  }
  return [pscustomobject]@{ BytesOut = $length; LastWriteTimeUtc = $modified; CreationTimeUtc = $created; CleanupWarning = $cleanupWarning }
}

function Copy-WinImgVideoToCandidate {
  param([string]$SourcePath, [string]$CandidatePath)
  # Read sharing blocks cooperative writers while copying. Both streams must
  # finish/dispose before the caller verifies length and source timestamps.
  $inputStream = [IO.FileStream]::new($SourcePath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
  try {
    $outputStream = [IO.FileStream]::new($CandidatePath, [IO.FileMode]::Open, [IO.FileAccess]::Write, [IO.FileShare]::None)
    try {
      if ($outputStream.Length -ne 0) { throw 'The reserved video partial is unexpectedly nonempty.' }
      $inputStream.CopyTo($outputStream); $outputStream.Flush()
    }
    finally { $outputStream.Dispose() }
  } finally { $inputStream.Dispose() }
}

function Remove-WinImgOwnedCandidate {
  param([string]$CandidatePath, [bool]$FileOwned = $true)
  # Call only for an exact path in an exclusively allocated attempt directory.
  # Unknown siblings, including numbered native outputs, are never deleted.
  Assert-WinImgNoReparseAncestors ([IO.Path]::GetDirectoryName($CandidatePath))
  if ($FileOwned -and [IO.File]::Exists($CandidatePath)) { [IO.File]::Delete($CandidatePath) }
  [IO.Directory]::Delete([IO.Path]::GetDirectoryName($CandidatePath), $false)
}

function Test-WinImgImageCandidate {
  param([string]$CandidatePath, [string]$MagickPath)
  Assert-WinImgNoReparseAncestors ([IO.Path]::GetDirectoryName($CandidatePath))
  $before = Get-Item -LiteralPath $CandidatePath -Force -ErrorAction Stop
  if ($before.PSIsContainer -or ($before.Attributes -band [IO.FileAttributes]::ReparsePoint) -or $before.Length -le 0) {
    throw 'ImageMagick produced no nonempty regular candidate.'
  }
  $length = $before.Length
  $modified = $before.LastWriteTimeUtc
  $nativePath = Get-WinImgNativeOutputPath $CandidatePath
  $previousPreference = $ErrorActionPreference
  $ErrorActionPreference = 'Continue'
  try {
    # Explicit +ping decodes pixels; -regard-warnings rejects recovered/truncated
    # JPEGs. Do not force an input coder: inspect the actual stored format.
    $metadata = @(& $MagickPath identify +ping -regard-warnings -format '%m|%w|%h|%n' $nativePath 2>&1)
    $decodeExit = $LASTEXITCODE
  } finally { $ErrorActionPreference = $previousPreference }
  $text = $metadata -join "`n"
  if ($decodeExit -ne 0 -or $text -notmatch '^JPEG\|([1-9][0-9]*)\|([1-9][0-9]*)\|1$') {
    throw 'Candidate failed full single-frame JPEG decode or dimension validation.'
  }
  $width = [long]$Matches[1]
  $height = [long]$Matches[2]
  $after = Get-Item -LiteralPath $CandidatePath -Force -ErrorAction Stop
  if ($after.Length -ne $length -or $after.LastWriteTimeUtc -ne $modified -or ($after.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
    throw 'Candidate changed during JPEG validation.'
  }
  return [pscustomobject]@{ BytesOut = $length; Width = $width; Height = $height }
}

# Multi-image inspection and selection read the same private, stable copy. No
# selector is appended to a user filename; only the decoded first image is kept.
function New-WinImgSourceSnapshot {
  param([string]$SourcePath, [string]$WorkRoot, [long]$ExpectedLength,
    [datetime]$ExpectedModified, [System.Collections.Generic.List[object]]$OwnedCandidates)

  Assert-WinImgNoReparseAncestors ([IO.Path]::GetDirectoryName($SourcePath))
  $before = Get-Item -LiteralPath $SourcePath -Force -ErrorAction Stop
  if ($before.PSIsContainer -or ($before.Attributes -band [IO.FileAttributes]::ReparsePoint) -or
      $before.Length -ne $ExpectedLength -or $before.LastWriteTimeUtc -ne $ExpectedModified) {
    throw 'Image source changed before frame/page inspection.'
  }
  $allocated = New-WinImgImageCandidate -WorkRoot $WorkRoot
  $snapshot = [IO.Path]::Combine([IO.Path]::GetDirectoryName($allocated), ('source' + $before.Extension.ToLowerInvariant()))
  $ownership = [pscustomobject]@{ Path = $snapshot; FileOwned = $false; Removed = $false }
  $OwnedCandidates.Add($ownership)
  $inputStream = [IO.FileStream]::new($SourcePath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
  try {
    $outputStream = [IO.FileStream]::new($snapshot, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    $ownership.FileOwned = $true
    try { $inputStream.CopyTo($outputStream); $outputStream.Flush() }
    finally { $outputStream.Dispose() }
  } finally { $inputStream.Dispose() }
  $after = Get-Item -LiteralPath $SourcePath -Force -ErrorAction Stop
  $copied = Get-Item -LiteralPath $snapshot -Force -ErrorAction Stop
  if ($after.PSIsContainer -or ($after.Attributes -band [IO.FileAttributes]::ReparsePoint) -or
      $after.Length -ne $ExpectedLength -or $after.LastWriteTimeUtc -ne $ExpectedModified -or
      $copied.PSIsContainer -or ($copied.Attributes -band [IO.FileAttributes]::ReparsePoint) -or
      $copied.Length -ne $ExpectedLength) {
    throw 'Image snapshot is incomplete or the source changed during copying.'
  }
  return $snapshot
}

function Get-WinImgSourceImageInfo {
  param([string]$SnapshotPath, [string]$MagickPath)
  Assert-WinImgNoReparseAncestors ([IO.Path]::GetDirectoryName($SnapshotPath))
  $nativePath = Get-WinImgNativeOutputPath $SnapshotPath
  $previousPreference = $ErrorActionPreference
  $ErrorActionPreference = 'Continue'
  try {
    # Count every image exposed by this decoder without applying the first-image
    # define. Header inspection is not the final JPEG's full-decode validation.
    $metadata = @(& $MagickPath identify -ping -regard-warnings -format '%m|%n|%w|%h\n' $nativePath 2>&1)
    $inspectExit = $LASTEXITCODE
  } finally { $ErrorActionPreference = $previousPreference }
  if ($inspectExit -ne 0 -or $metadata.Count -eq 0) { throw 'Source frame/page inspection failed; no image was selected.' }
  $count = $metadata.Count
  $decoder = $null
  foreach ($line in $metadata) {
    if ([string]$line -notmatch '^([A-Z0-9]+)\|([1-9][0-9]*)\|[1-9][0-9]*\|[1-9][0-9]*$' -or
        $Matches[2] -ne $count.ToString([Globalization.CultureInfo]::InvariantCulture)) {
      throw 'Source frame/page count or dimensions were ambiguous; no image was selected.'
    }
    if ($null -eq $decoder) { $decoder = $Matches[1] }
    elseif ($decoder -ne $Matches[1]) { throw 'Source inspection returned mixed image formats.' }
  }
  $unit = 'Images'; $policy = 'FirstImage'
  if ($decoder -in @('GIF', 'WEBP')) { $unit = 'Frames'; $policy = 'FirstDisplayedFrame' }
  elseif ($decoder -eq 'TIFF') { $unit = 'Pages'; $policy = 'FirstPage' }
  elseif ($decoder -in @('HEIC', 'HEIF')) { $policy = 'DecoderPrimaryOrFirstImage' }
  return [pscustomobject]@{ SourceCount = $count; Omitted = $count - 1; Unit = $unit; Policy = $policy; Decoder = $decoder }
}

function New-WinImgExclusiveDirectory {
  param([string]$Path)
  Initialize-WinImgDirectoryApi
  return [WinImgNormalizer.NativeDirectory]::CreateExclusive($Path)
}

function New-WinImgRunDirectory {
  param([string]$OutputParent, [string]$BaseName, [string]$Stamp)
  Assert-WinImgNoReparseAncestors $OutputParent
  [IO.Directory]::CreateDirectory($OutputParent) | Out-Null
  for ($attempt = 0; $attempt -lt 8; $attempt++) {
    $candidate = [IO.Path]::Combine($OutputParent, ('{0}_WinImgNormalized_{1}_{2}' -f $BaseName, $Stamp, [guid]::NewGuid().ToString('N')))
    if (New-WinImgExclusiveDirectory $candidate) { return $candidate }
  }
  throw 'Could not allocate an exclusive run directory after eight name collisions.'
}

function Resolve-WinImgSourcePath {
  param([string]$Path)
  if ([string]::IsNullOrWhiteSpace($Path)) { throw 'A source folder is required.' }
  if ($Path -match '^[A-Za-z]:($|[^\\/])') { throw 'Use an absolute drive path such as C:\, not C:.' }
  $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
  if ($item.PSProvider.Name -ne 'FileSystem' -or $item -isnot [IO.DirectoryInfo]) {
    throw 'The source must be a FileSystem directory.'
  }
  if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
    throw 'A linked source root is not supported; select the actual source directory.'
  }
  return Resolve-WinImgCanonicalDirectory $item.FullName
}

function ConvertTo-WinImgByteCap {
  param([object]$Value)
  $number = 0L
  $text = [string]$Value
  if ($text -notmatch '^[0-9]+$' -or
      -not [long]::TryParse($text, [Globalization.NumberStyles]::None, [Globalization.CultureInfo]::InvariantCulture, [ref]$number) -or
      $number -le 0) {
    throw 'maxBytes must be a positive whole number from 1 to 9223372036854775807.'
  }
  return $number
}

function Resolve-WinImgMagickApplication {
  param([string]$MagickPath)
  $name = if ($MagickPath) { $MagickPath } else { 'magick.exe' }
  $application = Get-Command $name -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
  if (-not $application -or [IO.Path]::GetExtension($application.Path) -ine '.exe') {
    throw "ImageMagick magick.exe was not found. Install a supported build and put its executable on PATH."
  }
  return (Get-Item -LiteralPath $application.Path -ErrorAction Stop).FullName
}

# Version/format queries need their output and exit status, independently of the
# existing conversion seam. Bound query time and drain both pipes in PS 5.1/7.
function Invoke-WinImgPreflightProcess {
  param([string]$Executable, [string[]]$Arguments)
  if (($Arguments -join ' ') -notin @('-version', '-list format')) { throw 'Unexpected preflight query.' }
  $start = New-Object Diagnostics.ProcessStartInfo
  $start.FileName = $Executable
  $start.Arguments = $Arguments -join ' '
  $start.UseShellExecute = $false
  $start.CreateNoWindow = $true
  $start.RedirectStandardOutput = $true
  $start.RedirectStandardError = $true
  $process = New-Object Diagnostics.Process
  $process.StartInfo = $start
  try {
    if (-not $process.Start()) { throw 'ImageMagick query could not start.' }
    $stdout = $process.StandardOutput.ReadToEndAsync()
    $stderr = $process.StandardError.ReadToEndAsync()
    $timedOut = -not $process.WaitForExit(15000)
    if ($timedOut) {
      & (Join-Path $env:SystemRoot 'System32\taskkill.exe') /PID $process.Id /T /F 1>$null 2>$null
      if (-not $process.WaitForExit(5000)) { $process.Kill() }
    }
    if (-not $stdout.Wait(5000) -or -not $stderr.Wait(5000)) { throw 'ImageMagick query streams did not complete.' }
    return [pscustomobject]@{ ExitCode = $process.ExitCode; StdOut = $stdout.GetAwaiter().GetResult(); StdErr = $stderr.GetAwaiter().GetResult(); TimedOut = $timedOut }
  } finally { $process.Dispose() }
}

function Get-WinImgMagickInfo {
  param([string]$MagickPath, [scriptblock]$PreflightRunner)
  $executable = Resolve-WinImgMagickApplication $MagickPath
  if (-not $PreflightRunner) { $PreflightRunner = { param($Executable, $Arguments) Invoke-WinImgPreflightProcess $Executable $Arguments } }
  $version = & $PreflightRunner $executable @('-version')
  if ($version.TimedOut -or $version.ExitCode -ne 0) { throw 'ImageMagick version query failed or timed out; no files were processed.' }
  $match = [regex]::Match($version.StdOut, '(?m)^Version: ImageMagick (7)\.(\d+)\.(\d+)-(\d+)\b')
  if (-not $match.Success) { throw 'Unrecognized ImageMagick version; use a supported ImageMagick 7 build.' }
  $build = [version]($match.Groups[1].Value + '.' + $match.Groups[2].Value + '.' + $match.Groups[3].Value + '.' + $match.Groups[4].Value)
  # Reviewed 2026-10-04: 7.1.2-32 includes jpeg:extent hang and JPEG/GIF/XMP fixes.
  if ($build -lt [version]'7.1.2.32') { throw 'ImageMagick 7.1.2-32 or newer supported 7.x build is required; update the dependency before running.' }
  $formatResult = & $PreflightRunner $executable @('-list','format')
  if ($formatResult.TimedOut -or $formatResult.ExitCode -ne 0) { throw 'ImageMagick format query failed or timed out; no files were processed.' }
  $formats = @{}
  foreach ($line in ($formatResult.StdOut -split '\r?\n')) {
    if ($line -match '^\s*([A-Z0-9]+)\*?\s+(?:\S+\s+)?([r-][w-][+-])\s') {
      $formats[$Matches[1]] = @{ Read = ($Matches[2][0] -eq 'r'); Write = ($Matches[2][1] -eq 'w') }
    }
  }
  if ($formats.Count -eq 0) { throw 'ImageMagick returned no usable format capability information.' }
  $delegates = [regex]::Match($version.StdOut, '(?m)^Delegates[^:]*:\s*(.*)$').Groups[1].Value.Trim()
  return [pscustomobject]@{ Path = $executable; Version = $build; VersionText = $version.StdOut.Trim(); Delegates = $delegates; Formats = $formats }
}

function Test-WinImgDestinationWritable {
  param([string]$Path)
  Assert-WinImgNoReparseAncestors $Path
  # Probe only the nearest existing directory; missing parents are created later.
  $ancestor = $Path
  while (-not (Test-Path -LiteralPath $ancestor)) {
    $parent = [IO.Directory]::GetParent($ancestor)
    if ($null -eq $parent) { throw 'The destination volume/share does not exist.' }
    $ancestor = $parent.FullName
  }
  $item = Get-Item -LiteralPath $ancestor -Force
  if ($item.PSProvider.Name -ne 'FileSystem' -or $item -isnot [IO.DirectoryInfo]) { throw 'The destination parent must be a FileSystem directory.' }
  $probe = Join-Path $ancestor ('.WinImgNormalizer-write-probe-' + [guid]::NewGuid().ToString('N'))
  # DeleteOnClose also checks deletion rights when opening, so a destination that
  # permits creating but forbids deleting cannot leave a probe behind.
  $stream = [IO.FileStream]::new($probe, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write,
    [IO.FileShare]::None, 4096, [IO.FileOptions]::DeleteOnClose)
  try { $stream.WriteByte(0) } finally { $stream.Dispose() }
}

function Get-WinImgAvailableBytes {
  param([string]$Path)
  try { return ([IO.DriveInfo]::new([IO.Path]::GetPathRoot($Path))).AvailableFreeSpace } catch { return $null }
}

function Get-WinImgDestinationInfo {
  param([string]$Path, [object[]]$Files, [long]$MaxBytes)
  $full = Normalize-WinImgRootPath $Path
  Test-WinImgDestinationWritable $full
  # Best-effort estimate: videos + bounded image budgets + two largest-input
  # scratch copies + 64 MiB reserve. A large cap is an upper limit, not required
  # output allocation. Decimal arithmetic prevents Int64 wrapping.
  [decimal]$required = 64MB
  [decimal]$largestImage = 0
  foreach ($file in $Files) {
    if ($file.Extension.ToLowerInvariant() -in @('.jpg','.jpeg','.png','.bmp','.tif','.tiff','.gif','.heic','.heif','.webp')) {
      $required += [Math]::Min([decimal]$MaxBytes, [Math]::Max([decimal]1MB, 4 * [decimal]$file.Length))
      $largestImage = [Math]::Max($largestImage, [decimal]$file.Length)
    } else { $required += [decimal]$file.Length }
  }
  $required += 2 * $largestImage
  $available = Get-WinImgAvailableBytes $full
  if ($null -ne $available -and [decimal]$available -lt $required) {
    throw ('Insufficient destination space: estimated {0} bytes required, {1} available.' -f $required, $available)
  }
  return [pscustomobject]@{ Path = $full; RequiredBytes = $required; AvailableBytes = $available }
}

# Dot-sourcing defines the callable boundary only. Internal seams are for tests;
# the public script and batch launcher retain their positional interface.
function Invoke-WinImgNormalizer {
  param(
    [string]$Source,
    [object]$MaxBytes = 1MB,
    [string]$OutputParent,
    [string]$MagickPath,
    [scriptblock]$PreflightRunner,
    [scriptblock]$ProcessRunner = {
      param([string]$Executable, [string[]]$Arguments)
      $LASTEXITCODE = 0
      & $Executable @Arguments 1>$null 2>$null
      return $LASTEXITCODE
    }
  )

  $ErrorActionPreference = 'Stop'

  # Complete basic setup before creating Pictures, run folders, logs or mirrors.
  try {
    $MaxBytes = ConvertTo-WinImgByteCap $MaxBytes
    $srcRoot = Resolve-WinImgSourcePath $Source
    $pictures = $OutputParent
    if (-not $pictures) {
      $pictures = [Environment]::GetFolderPath('MyPictures')
      if ([string]::IsNullOrWhiteSpace($pictures)) { $pictures = Join-Path $env:USERPROFILE 'Pictures' }
    }
    # This precedes enumeration and, critically, the destination write probe.
    $pictures = Assert-WinImgSafeDestination -SourceRoot $srcRoot -OutputParent $pictures
    $magickInfo = Get-WinImgMagickInfo -MagickPath $MagickPath -PreflightRunner $PreflightRunner
    $MagickCmd = $magickInfo.Path
    $imgExts = '.jpg','.jpeg','.png','.bmp','.tif','.tiff','.gif','.heic','.heif','.webp'
    $videoExts = '.mp4','.mov','.mkv','.avi','.m4v','.wmv','.webm','.mts','.m2ts','.3gp','.3g2'
    $imageCoders = @{ '.jpg'='JPEG'; '.jpeg'='JPEG'; '.png'='PNG'; '.bmp'='BMP'; '.tif'='TIFF'; '.tiff'='TIFF'; '.gif'='GIF'; '.heic'='HEIC'; '.heif'='HEIF'; '.webp'='WEBP' }
    $sourceTree = Get-WinImgSourceTree -SourceRoot $srcRoot
    $generatedName = Get-WinImgGeneratedName -TopLevelNames $sourceTree.TopLevelNames
    $outputPlan = @(Get-WinImgOutputPlan -SourceRoot $srcRoot -SourceTree $sourceTree -GeneratedName $generatedName -ImageExtensions $imgExts -VideoExtensions $videoExts)
    $allFiles = @($outputPlan | ForEach-Object { $_.Source })
    $missingCoders = @{}
    foreach ($f in $allFiles) {
      $coder = $imageCoders[$f.Extension.ToLowerInvariant()]
      if ($coder -and (-not $magickInfo.Formats.ContainsKey($coder) -or -not $magickInfo.Formats[$coder].Read)) { $missingCoders[$coder] = $true }
    }
    $readableImages = @($allFiles | Where-Object {
      $coder = $imageCoders[$_.Extension.ToLowerInvariant()]
      $coder -and -not $missingCoders.ContainsKey($coder)
    })
    if ($readableImages.Count -gt 0 -and (-not $magickInfo.Formats.ContainsKey('JPEG') -or -not $magickInfo.Formats['JPEG'].Write)) {
      throw 'This ImageMagick build has no JPEG encoder; install a build with JPEG write support.'
    }
    $spaceFiles = @($allFiles | Where-Object {
      $coder = $imageCoders[$_.Extension.ToLowerInvariant()]
      -not $coder -or -not $missingCoders.ContainsKey($coder)
    })
    $destinationInfo = Get-WinImgDestinationInfo -Path $pictures -Files $spaceFiles -MaxBytes $MaxBytes
    $pictures = $destinationInfo.Path
  } catch {
    Write-Host ('Setup error: ' + $_.Exception.Message) -ForegroundColor Red
    return 1
  }

  # --------- Logger, literal-safe ---
  $LogPath = $null
  function Write-Log {
    param([string]$Message, [ValidateSet('INFO','OK','SKIP','WARN','ERR')]$Level='INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    switch ($Level) { 'ERR'{Write-Host $line -ForegroundColor Red}; 'WARN'{Write-Host $line -ForegroundColor Yellow}; 'OK'{Write-Host $line -ForegroundColor Green}; default{Write-Host $line} }
    if ($LogPath) { try { Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8 } catch {} }
  }

  # --------- Resolve Pictures + dest root ----
  $base  = [System.IO.Path]::GetFileName($srcRoot.TrimEnd('\','/'))
  if ([string]::IsNullOrWhiteSpace($base)) { $base = $srcRoot.TrimEnd('\','/').TrimEnd(':') }
  $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
  try {
    $pictures = Assert-WinImgSafeDestination -SourceRoot $srcRoot -OutputParent $pictures
    $destRoot = New-WinImgRunDirectory -OutputParent $pictures -BaseName $base -Stamp $stamp
    $generatedRoot = [IO.Path]::Combine($destRoot, $generatedName)
    $workRoot = [IO.Path]::Combine($generatedRoot, 'work')
    $reportRoot = [IO.Path]::Combine($generatedRoot, 'reports')
    foreach ($path in @($generatedRoot, $workRoot, $reportRoot)) {
      if (-not (New-WinImgExclusiveDirectory $path)) { throw 'A generated namespace was unexpectedly occupied; the run was not adopted.' }
    }
  } catch {
    Write-Host ('Setup error: could not create the destination directory. ' + $_.Exception.Message) -ForegroundColor Red
    return 1
  }

  # --------- Log file ---
  try {
    $LogPath = [System.IO.Path]::Combine($reportRoot, "WinImgNormalizer_${stamp}.log")
    $stream = [IO.FileStream]::new($LogPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
    try {
      $writer = [IO.StreamWriter]::new($stream, [Text.UTF8Encoding]::new($true))
      try { $writer.WriteLine("WinImgNormalizer started $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')") }
      finally { $writer.Dispose() }
    } finally { $stream.Dispose() }
  } catch {
    Write-Host ('Setup error: could not create the run-owned log. ' + $_.Exception.Message) -ForegroundColor Red
    return 1
  }
  Write-Log "Source: $srcRoot"
  Write-Log "Destination: $destRoot"
  Write-Log "Generated work/report directory: $generatedName"
  foreach ($row in $outputPlan) {
    Write-Log ('PLAN {0}: {1} -> {2} ({3})' -f $row.Kind, $row.SourceRelativePath, $row.OutputRelativePath, $row.NamingReason)
  }
  Write-Log ("MaxBytes: {0:n0} ({1} MB)" -f $MaxBytes, [Math]::Round($MaxBytes/1MB,2))

  Write-Log ('ImageMagick executable: ' + $MagickCmd)
  foreach ($line in ($magickInfo.VersionText -split '\r?\n')) { Write-Log $line }
  Write-Log ('Relevant formats: ' + (($imageCoders.Values | Select-Object -Unique | Sort-Object | ForEach-Object {
    $mode = $magickInfo.Formats[$_]
    '{0}:read={1},write={2}' -f $_, [bool]$mode.Read, [bool]$mode.Write
  }) -join '; '))
  Write-Log ('Destination space estimate: {0} bytes; available: {1}' -f $destinationInfo.RequiredBytes, $destinationInfo.AvailableBytes)
  if ($null -eq $destinationInfo.AvailableBytes) { Write-Log 'Free space could not be measured for this destination; writes may still fail.' 'WARN' }

  # --------- Mirror tree, non-destructive ---
  $hasTraversalWarnings = $sourceTree.Warnings.Count -gt 0
  $hasNamingWarnings = $false
  foreach ($warning in $sourceTree.Warnings) { Write-Log ("Source scan: {0} ({1})" -f $warning.Path, $warning.Reason) 'WARN' }
  foreach ($directory in $sourceTree.Directories) {
    try {
      Assert-WinImgNoReparseAncestors $directory.FullName
      $rel = Get-WinImgRelativePath -Root $srcRoot -Path $directory.FullName
      $target = [System.IO.Path]::Combine($destRoot, $rel)
      Assert-WinImgNoReparseAncestors $target
      [IO.Directory]::CreateDirectory($target) | Out-Null
    } catch {
      $hasTraversalWarnings = $true
      Write-Log ("Could not mirror directory: {0} ({1})" -f $directory.FullName, $_.Exception.Message) 'WARN'
    }
  }

  # --------- File sets ---
  $total = $allFiles.Count
  if ($total -eq 0) {
    Write-Log "No images or videos found." 'WARN'
    if ($hasTraversalWarnings) { return 2 }
    return 0
  }

  # --------- Dedupe + stats ---
  $retained = New-Object 'Collections.Generic.Dictionary[string,object]' ([StringComparer]::Ordinal)
  $hasDuplicateWarnings = $false
  $stats = [ordered]@{ Converted=0; CopiedVideo=0; SkippedDuplicate=0; Unsupported=0; Errors=0 }
  Write-Log 'Heuristic duplicate matching uses lowercase filename, LastWriteTimeUtc and equal input byte length after successful finalization; same-key/same-length content can still differ.'

  # --------- IM helpers ---
  function Get-ExtentString([long]$Bytes) {
    if ($Bytes -ge 1MB -and ($Bytes % 1MB) -eq 0) { return "{0}MB" -f [long]([decimal]$Bytes/1MB) }
    elseif ($Bytes -ge 1KB -and ($Bytes % 1KB) -eq 0) { return "{0}KB" -f [long]([decimal]$Bytes/1KB) }
    else { return "{0}B" -f $Bytes }
  }
  function Convert-ImageMagick {
    param([string]$SourcePath, [string]$DestPath, [long]$MaxBytes,
      [string]$WorkRoot, [System.Collections.Generic.List[object]]$OwnedCandidates, [object]$SourceInfo)

    $extent = Get-ExtentString $MaxBytes
    $scales = 100,90,80,70,60,50
    $lastFailure = 'ImageMagick conversion failed'
    $attemptCount = 0
    foreach ($p in $scales) {
      foreach ($alpha in @($true, $false)) {
        # Every native attempt has a new exact path, including alpha fallback.
        if ($attemptCount -gt 0) {
          $previous = $OwnedCandidates[$OwnedCandidates.Count - 1]
          try { Remove-WinImgOwnedCandidate -CandidatePath $previous.Path -FileOwned $previous.FileOwned; $previous.Removed = $true }
          catch { Write-Log ('Could not remove superseded image scratch: ' + $_.Exception.Message) 'WARN' }
          $DestPath = New-WinImgImageCandidate -WorkRoot $WorkRoot
          $OwnedCandidates.Add([pscustomobject]@{ Path = $DestPath; FileOwned = $false; Removed = $false })
        }
        $ownership = $OwnedCandidates[$OwnedCandidates.Count - 1]
        $stream = [IO.FileStream]::new($DestPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        $ownership.FileOwned = $true
        $stream.Dispose()
        $attemptCount++
        $nativeDestPath = Get-WinImgNativeOutputPath $DestPath
        $nativeArguments = @('-quiet')
        if ($SourceInfo) { $nativeArguments += @('-define', 'image:frames=0') }
        $nativeArguments += $(if ($SourceInfo) { Get-WinImgNativeOutputPath $SourcePath } else { $SourcePath })
        # Coalesce only the selected first animation frame onto its logical
        # canvas. Document pages/HEIC images must never be overlaid together.
        if ($SourceInfo -and $SourceInfo.Policy -eq 'FirstDisplayedFrame') { $nativeArguments += '-coalesce' }
        if ($SourceInfo) { $nativeArguments += '+repage' }
        $nativeArguments += @('-auto-orient','-strip','-colorspace','sRGB',
          '-sampling-factor','4:2:0','-interlace','Line')
        if ($alpha) { $nativeArguments += @('-background','white','-alpha','remove','-alpha','off') }
        $nativeArguments += @('-resize', "$p%", '-define', "jpeg:extent=$extent", ('JPEG:' + $nativeDestPath))
        $previousPreference = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try { $nativeResult = & $ProcessRunner $MagickCmd $nativeArguments }
        finally { $ErrorActionPreference = $previousPreference }
        # Keep the inherited integer seam and allow a future structured native
        # result. Missing, ambiguous and timeout/cancellation results fail closed.
        $successful = (($nativeResult -is [int] -or $nativeResult -is [long]) -and $nativeResult -eq 0)
        if ($null -ne $nativeResult -and $nativeResult -isnot [array] -and
            $null -ne $nativeResult.PSObject.Properties['ExitCode']) {
          $successful = (($nativeResult.ExitCode -is [int] -or $nativeResult.ExitCode -is [long]) -and
            $nativeResult.ExitCode -eq 0 -and -not $nativeResult.TimedOut -and -not $nativeResult.Cancelled)
        }
        if (-not $successful) { $lastFailure = 'ImageMagick attempt did not return a successful native outcome'; continue }
        try { $validation = Test-WinImgImageCandidate -CandidatePath $DestPath -MagickPath $MagickCmd }
        catch { $lastFailure = $_.Exception.Message; continue }
        if ($validation.BytesOut -le $MaxBytes -or $p -eq 50) {
          $note = if ($validation.BytesOut -gt $MaxBytes) { 'WARN: Could not reach target; best-effort saved' } else { $null }
          return @{ Status='Converted'; CandidatePath=$DestPath; BytesOut=$validation.BytesOut;
            Width=$validation.Width; Height=$validation.Height; Scale=$p; Attempts=$attemptCount; Note=$note }
        }
        # This valid above-target attempt is superseded. A later failure may not
        # reuse it or mislabel it with a later scale. Only the last attempt counts.
        break
      }
    }
    return @{ Status='Error'; Attempts=$attemptCount; Note=$lastFailure }
  }

  # --------- Main loop ---
  [int]$i = 0
  foreach ($row in $outputPlan) {
    $f = $row.Source
    $i++
    $rel = $row.SourceRelativePath
    $ext = $f.Extension.ToLowerInvariant()
    Write-Progress -Activity "WinImgNormalizer" -Status "$i / $total : $rel" -PercentComplete ([int]($i*100/$total))

    $coder = $imageCoders[$ext]
    if ($coder -and $missingCoders.ContainsKey($coder)) {
      $stats.Errors++
      Write-Log "ERR IMG: $rel (Missing $coder decoder in this ImageMagick build; no conversion attempted. Install a build with $coder read support.)" 'ERR'
      continue
    }

    $destRel = $row.OutputRelativePath
    $destPath = [System.IO.Path]::Combine($destRoot, $destRel)

    # A source entry can change after inventory. Recheck links before reading;
    # concurrent hostile filesystem mutation is not a complete sandbox boundary.
    try {
      Assert-WinImgNoReparseAncestors $f.DirectoryName
      $f = Get-Item -LiteralPath $f.FullName -Force -ErrorAction Stop
      if (($f.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw 'Source entry became a reparse point after inventory; no read attempted.'
      }
      if ($f.PSIsContainer) {
        throw 'Source entry became a directory after inventory; no read attempted.'
      }
      # Inventory FileInfo values may be cached. Snapshot live metadata before
      # lookup, then reuse those immutable values for the processed input.
      $sourceLength = $f.Length
      $sourceModified = $f.LastWriteTimeUtc
      $dupKey = Get-WinImgDuplicateKey -Name $f.Name -LastWriteTimeUtc $sourceModified -Length $sourceLength
      if ($retained.ContainsKey($dupKey)) {
        $first = $retained[$dupKey]
        $stats.SkippedDuplicate++
        Write-Log ('Heuristic duplicate skipped: {0} (retained source: {1}; retained output: {2}; retained status: {3})' -f
          $rel, $first.SourceRelativePath, $first.OutputRelativePath, $first.Status) 'SKIP'
        continue
      }
      Assert-WinImgOutputAvailable $destPath
    } catch {
      $hasTraversalWarnings = $true
      $stats.Errors++
      Write-Log "Source/destination safety check failed: $rel ($($_.Exception.Message))" 'ERR'
      continue
    }
    $ownedCandidates = New-Object 'System.Collections.Generic.List[object]'
    try {
      if ($imgExts -contains $ext) {
        $imageSource = $f.FullName
        $sourceInfo = $null
        if ($ext -in @('.gif','.tif','.tiff','.webp','.heic','.heif')) {
          $imageSource = New-WinImgSourceSnapshot -SourcePath $f.FullName -WorkRoot $workRoot -ExpectedLength $sourceLength -ExpectedModified $sourceModified -OwnedCandidates $ownedCandidates
          $sourceInfo = Get-WinImgSourceImageInfo -SnapshotPath $imageSource -MagickPath $MagickCmd
          Write-Log ('SOURCE IMG: {0} (SourceCount={1}; Selected=1; Omitted={2}; Unit={3}; Policy={4}; Decoder={5})' -f
            $rel, $sourceInfo.SourceCount, $sourceInfo.Omitted, $sourceInfo.Unit, $sourceInfo.Policy, $sourceInfo.Decoder)
        }
        try { $candidatePath = New-WinImgImageCandidate -WorkRoot $workRoot }
        catch { $hasNamingWarnings = $true; throw }
        $ownedCandidates.Add([pscustomobject]@{ Path = $candidatePath; FileOwned = $false; Removed = $false })
        $res = Convert-ImageMagick -SourcePath $imageSource -DestPath $candidatePath -MaxBytes $MaxBytes -WorkRoot $workRoot -OwnedCandidates $ownedCandidates -SourceInfo $sourceInfo
        if ($res.Status -eq 'Converted') {
          try { Move-WinImgPlannedImage -CandidatePath $res.CandidatePath -DestinationPath $destPath }
          catch { $hasNamingWarnings = $true; throw }
          $stats.Converted++
          try {
            Assert-WinImgNoReparseAncestors $f.DirectoryName
            $current = Get-Item -LiteralPath $f.FullName -Force -ErrorAction Stop
            if ($current.PSIsContainer -or ($current.Attributes -band [IO.FileAttributes]::ReparsePoint) -or
                $current.Length -ne $sourceLength -or $current.LastWriteTimeUtc -ne $sourceModified) {
              throw 'Source metadata changed during image processing.'
            }
            if (-not $retained.ContainsKey($dupKey)) {
              $retained.Add($dupKey, [pscustomobject]@{
                SourceRelativePath = $rel; OutputRelativePath = $destRel
                Status = $(if ($res.Note) { 'ConvertedWithWarning' } else { 'Converted' })
              })
            }
          } catch {
            $hasDuplicateWarnings = $true
            Write-Log ('Finalized image kept without heuristic registration: {0} ({1})' -f $rel, $_.Exception.Message) 'WARN'
          }
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
        try { $videoResult = Copy-WinImgPlannedVideo -SourcePath $f.FullName -DestinationPath $destPath -WorkRoot $workRoot }
        catch { $hasNamingWarnings = $true; throw }
        # The copy helper's stable snapshot describes the bytes actually copied,
        # even if source metadata changed between lookup and staging.
        $videoKey = Get-WinImgDuplicateKey -Name $f.Name -LastWriteTimeUtc $videoResult.LastWriteTimeUtc -Length $videoResult.BytesOut
        if (-not $retained.ContainsKey($videoKey)) {
          $retained.Add($videoKey, [pscustomobject]@{ SourceRelativePath = $rel; OutputRelativePath = $destRel; Status = 'CopiedVideo' })
        }
        if ($videoResult.CleanupWarning) { $hasNamingWarnings = $true; Write-Log ('Video scratch cleanup: ' + $videoResult.CleanupWarning) 'WARN' }
        try { (Get-Item -LiteralPath $destPath).LastWriteTimeUtc = $videoResult.LastWriteTimeUtc; (Get-Item -LiteralPath $destPath).CreationTimeUtc = $videoResult.CreationTimeUtc } catch {}
        $bytesOut = $videoResult.BytesOut
        $stats.CopiedVideo++; Write-Log ("OK VID: {0} -> {1} [{2:n0} bytes]" -f $rel, $destRel, $bytesOut) 'OK'
      }
      else {
        $stats.Unsupported++; Write-Log "Unsupported skipped: $rel" 'WARN'
      }
    } catch {
      $stats.Errors++; Write-Log "Exception processing: $rel ($($_.Exception.Message))" 'ERR'
    } finally {
      foreach ($owned in $ownedCandidates) {
        if ($owned.Removed) { continue }
        try { Remove-WinImgOwnedCandidate -CandidatePath $owned.Path -FileOwned $owned.FileOwned }
        catch {
          $hasNamingWarnings = $true
          Write-Log ('Could not remove owned image scratch: ' + $_.Exception.Message) 'WARN'
        }
      }
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
  Write-Host ("  Log file         : {0}" -f $LogPath)

  Write-Log ("SUMMARY ConvertedImages={0} CopiedVideos={1} Duplicates={2} Unsupported={3} Errors={4}" -f $stats.Converted,$stats.CopiedVideo,$stats.SkippedDuplicate,$stats.Unsupported,$stats.Errors)
  Write-Log "Completed $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

  if ($stats.Errors -gt 0 -or $missingCoders.Count -gt 0 -or $hasTraversalWarnings -or $hasNamingWarnings -or $hasDuplicateWarnings) { return 2 }
  return 0
}

function Invoke-WinImgNormalizerCommand {
  param(
    [object[]]$Arguments,
    [string]$OutputParent,
    [string]$MagickPath,
    [scriptblock]$PreflightRunner,
    [scriptblock]$ProcessRunner
  )

  # Preserve the two existing positional forms; reject ambiguous/ignored extras.
  $ErrorActionPreference = 'Stop'
  $Source = $null
  $MaxBytes = 1MB
  if (@($Arguments).Count -lt 1 -or @($Arguments).Count -gt 2) {
    Write-Host 'Usage: WinImgNormalizer.ps1 <sourceFolder> [maxBytes] (one folder only)' -ForegroundColor Red
    return 1
  }
  $Source = $Arguments[0]
  if ($Arguments.Count -eq 2) { $MaxBytes = $Arguments[1] }
  $invoke = @{ Source = $Source; MaxBytes = $MaxBytes }
  if ($OutputParent) { $invoke.OutputParent = $OutputParent }
  if ($ProcessRunner) { $invoke.ProcessRunner = $ProcessRunner }
  if ($MagickPath) { $invoke.MagickPath = $MagickPath }
  if ($PreflightRunner) { $invoke.PreflightRunner = $PreflightRunner }
  return Invoke-WinImgNormalizer @invoke
}

if ($MyInvocation.InvocationName -ne '.') {
  exit (Invoke-WinImgNormalizerCommand -Arguments $args)
}
