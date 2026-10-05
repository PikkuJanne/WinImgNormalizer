<#
WinImgNormalizer.ps1
Non-destructive image normalizer + archive prep for mixed photo folders

Author: Janne Vuorela
Target OS: Windows 10/11
PowerShell: Windows PowerShell 5.1+, also works on PowerShell 7
Dependencies: ImageMagick 7.1.2-32+ supported 7.x (magick.exe in PATH), optional: .bat wrapper

SYNOPSIS
    Recursively mirrors a source folder into a safe copy under the user’s Pictures folder,
    converts images to JPEG targeting 1,048,576 bytes (1 MiB), with auto-orient,
    metadata stripping and white alpha flattening,
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
    - JPEG targeting 1,048,576 bytes (1 MiB) with scaling (100→50%) and jpeg:extent.
      Actual byte length determines compliance; valid above-target fallback is a warning.
    - EXIF auto-orientation; profiled colour to sRGB before white alpha flattening and metadata removal.
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
    - Output timestamp restoration is best effort; failed fields warn and retain valid files.

LIMITATIONS
    - Size targeting is best-effort; noisy or huge images may exceed the cap at 50%.
      Valid above-target JPEGs are retained with a warning and application exit 2.
    - Duplicate detection is not content-hash based.
    - Untagged RGB is assumed sRGB; untagged CMYK/unknown colour is rejected rather than guessed.
    - Final known-sRGB JPEGs omit ICC and source metadata; profiled originals remain untouched.

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
  # These APIs exist on the supported Windows/Windows PowerShell versions.
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
    [StructLayout(LayoutKind.Sequential)]
    private struct FileIdInfo {
      public ulong VolumeSerialNumber;
      public ulong FileIdLow;
      public ulong FileIdHigh;
    }
    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetFileInformationByHandleEx(SafeFileHandle handle,
      int informationClass, out FileIdInfo information, uint size);
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
    public static string Identity(string path) {
      using (SafeFileHandle handle = CreateFileW(NativePath(path), 0, 7, IntPtr.Zero, 3, 0x02000000, IntPtr.Zero)) {
        if (handle.IsInvalid) throw new Win32Exception(Marshal.GetLastWin32Error());
        FileIdInfo information;
        // FileIdInfo (18) supplies a 128-bit ID, including on ReFS. An unsupported
        // provider cannot establish safe alias boundaries and fails setup.
        if (!GetFileInformationByHandleEx(handle, 18, out information, (uint)Marshal.SizeOf(typeof(FileIdInfo))))
          throw new InvalidOperationException("The directory provider cannot establish safe alias boundaries with a 128-bit file identity.",
            new Win32Exception(Marshal.GetLastWin32Error()));
        if (information.FileIdLow == 0 && information.FileIdHigh == 0)
          throw new InvalidOperationException("The directory provider returned no usable file identity.");
        return information.VolumeSerialNumber.ToString("X16") + ":" +
          information.FileIdHigh.ToString("X16") + information.FileIdLow.ToString("X16");
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
  $contained = Test-WinImgPathContained -Root $sourcePath -Path $parentPath
  if (-not $contained) {
    # A drive path and a share alias can spell the same physical directory
    # differently. Compare the existing destination ancestors before any write.
    $sourceIdentity = [WinImgNormalizer.NativeDirectory]::Identity($sourcePath)
    $ancestor = $parentPath
    while ($ancestor) {
      if (Test-Path -LiteralPath $ancestor -ErrorAction Stop) {
        if ([WinImgNormalizer.NativeDirectory]::Identity($ancestor) -eq $sourceIdentity) {
          $contained = $true
          break
        }
      }
      $parent = [IO.Directory]::GetParent($ancestor)
      if ($null -eq $parent) { break }
      $ancestor = $parent.FullName
    }
  }
  if ($contained) {
    throw 'The destination is inside or equal to the source, including a directory alias. Choose a source subfolder outside the destination tree (for example, a subfolder of Pictures).'
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

function Test-WinImgCancellationException {
  param([object]$ErrorOrException)
  $exception = if ($ErrorOrException -is [Management.Automation.ErrorRecord]) { $ErrorOrException.Exception } else { $ErrorOrException }
  while ($exception -and $exception.InnerException) { $exception = $exception.InnerException }
  return ($exception -is [OperationCanceledException])
}

function Assert-WinImgCancellation {
  # Dot-sourced helper tests need not initialize the native API or a run session.
  if ('WinImgNormalizer.CancellationSession' -as [type]) {
    $state = [WinImgNormalizer.CancellationSession]::Current
    if ($state) { $state.ThrowIfRequested() }
  }
}

function Invoke-WinImgRunStage {
  param([string]$Stage, [string]$RelativePath, [string]$Path)
  Assert-WinImgCancellation
  if ($script:WinImgCancellationContext -and $script:WinImgCancellationContext.RunStageObserver) {
    & $script:WinImgCancellationContext.RunStageObserver $Stage ([WinImgNormalizer.CancellationSession]::Current) $RelativePath $Path
  }
  Assert-WinImgCancellation
}

function Copy-WinImgCancellableStream {
  param([IO.Stream]$InputStream, [IO.Stream]$OutputStream, [string]$SourcePath, [string]$CandidatePath)
  $buffer = New-Object byte[] (256KB)
  [long]$copied = 0
  while ($true) {
    Assert-WinImgCancellation
    $read = $InputStream.Read($buffer, 0, $buffer.Length)
    Assert-WinImgCancellation
    if ($read -eq 0) { break }
    $OutputStream.Write($buffer, 0, $read)
    $copied += $read
    if ($script:WinImgCancellationContext -and $script:WinImgCancellationContext.CopyProgressObserver) {
      & $script:WinImgCancellationContext.CopyProgressObserver $SourcePath $CandidatePath $copied ([WinImgNormalizer.CancellationSession]::Current)
    }
    Assert-WinImgCancellation
  }
  $OutputStream.Flush()
  Assert-WinImgCancellation
}

function Get-WinImgSourceTree {
  param([string]$SourceRoot, [object]$ReportState)
  $files = New-Object 'System.Collections.Generic.List[System.IO.FileInfo]'
  $directories = New-Object 'System.Collections.Generic.List[System.IO.DirectoryInfo]'
  $names = New-Object 'System.Collections.Generic.List[string]'
  $warnings = New-Object 'System.Collections.Generic.List[object]'
  if ($ReportState) { $warnings = $ReportState.ScanIssues }
  $pending = New-Object 'System.Collections.Generic.Stack[string]'
  $pending.Push($SourceRoot)
  $inaccessibleDirectories = 0
  $uninspectableEntries = 0
  $skippedLinks = 0
  while ($pending.Count -gt 0) {
    Assert-WinImgCancellation
    $directory = $pending.Pop()
    try {
      Assert-WinImgNoReparseAncestors $directory
      $paths = New-Object 'System.Collections.Generic.List[string]'
      foreach ($entry in @(Get-ChildItem -LiteralPath $directory -Force -ErrorAction Stop)) { $paths.Add($entry.FullName) }
      $paths.Sort([StringComparer]::OrdinalIgnoreCase)
      foreach ($path in $paths) {
        Assert-WinImgCancellation
        try {
          $entry = Get-Item -LiteralPath $path -Force -ErrorAction Stop
          $null = Get-WinImgRelativePath -Root $SourceRoot -Path $entry.FullName
          if ($directory.Equals($SourceRoot, [StringComparison]::OrdinalIgnoreCase)) { $names.Add($entry.Name) }
          if (($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            $skippedLinks++
            if ($ReportState) { $ReportState.SkippedLinks = $skippedLinks }
            $warnings.Add([pscustomobject]@{ Path = $entry.FullName; Kind = 'SkippedReparsePoint'; Reason = 'Reparse point skipped; links and junctions are not followed.' })
          } elseif ($entry -is [IO.DirectoryInfo]) {
            $directories.Add($entry)
            $pending.Push($entry.FullName)
          } elseif ($entry -is [IO.FileInfo]) {
            if ($ReportState) { Add-WinImgDiscoveredFile -State $ReportState -SourceRoot $SourceRoot -File $entry }
            $files.Add($entry)
          }
        } catch [Management.Automation.PipelineStoppedException] { throw }
        catch { if (Test-WinImgCancellationException $_) { throw };
          $uninspectableEntries++
          if ($ReportState) { $ReportState.UninspectableEntries = $uninspectableEntries }
          $warnings.Add([pscustomobject]@{ Path = $path; Kind = 'EntryInspectionFailed'; Reason = 'Could not inspect source entry: ' + $_.Exception.Message })
        }
      }
    } catch [Management.Automation.PipelineStoppedException] { throw }
    catch { if (Test-WinImgCancellationException $_) { throw };
      $inaccessibleDirectories++
      if ($ReportState) { $ReportState.IncompleteDirectories = $inaccessibleDirectories }
      $warnings.Add([pscustomobject]@{ Path = $directory; Kind = 'DirectoryEnumerationFailed'; Reason = 'Incomplete source scan: ' + $_.Exception.Message })
    }
  }
  if ($ReportState) { $ReportState.ScanComplete = ($inaccessibleDirectories -eq 0 -and $uninspectableEntries -eq 0) }
  return [pscustomobject]@{
    Files = $files.ToArray(); Directories = $directories.ToArray(); TopLevelNames = $names.ToArray(); Warnings = $warnings.ToArray()
    ScanComplete = ($inaccessibleDirectories -eq 0 -and $uninspectableEntries -eq 0)
    InaccessibleDirectoryCount = $inaccessibleDirectories; UninspectableEntryCount = $uninspectableEntries; SkippedReparsePointCount = $skippedLinks
  }
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
    Assert-WinImgCancellation
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
    Assert-WinImgCancellation
    $row = $rows[$relative]
    if ($preferredCounts[$row.OutputRelativePath] -eq 1 -and $reserved.Add($row.OutputRelativePath)) { continue }
    $collisions.Add($row)
  }
  foreach ($row in $collisions) {
    Assert-WinImgCancellation
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

function Initialize-WinImgProcessApi {
  if ('WinImgNormalizer.NativeProcess' -as [type]) { return }
  # Fixed buffers on two readers avoid both pipe deadlocks and unbounded
  # ReadToEnd/ReadLine allocations (a native diagnostic need not contain LF).
  Add-Type -TypeDefinition @'
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
namespace WinImgNormalizer {
  // A console callback signals only this managed state. Filesystem work and
  // reporting stay on the calling PowerShell thread.
  public sealed class CancellationState {
    private int requested, consoleControlCount;
    private readonly object commitGate = new object();
    public bool Requested { get { return Interlocked.CompareExchange(ref requested,0,0)!=0; } }
    public int ConsoleControlCount { get { return Interlocked.CompareExchange(ref consoleControlCount,0,0); } }
    public void Request() { Interlocked.Exchange(ref requested,1); }
    internal void RequestFromConsole() { Interlocked.Increment(ref consoleControlCount); Request(); }
    public void ThrowIfRequested() { if (Requested) throw new OperationCanceledException("WinImgNormalizer cancellation requested."); }
    public void Wait(int milliseconds) {
      if(milliseconds<0) throw new ArgumentOutOfRangeException("milliseconds");
      Stopwatch clock=Stopwatch.StartNew();
      while(true) { ThrowIfRequested(); int remaining=(int)Math.Max(0L,milliseconds-clock.ElapsedMilliseconds); if(remaining==0) return; Thread.Sleep(Math.Min(50,remaining)); }
    }
    public IDisposable EnterCommit() {
      Monitor.Enter(commitGate);
      try { ThrowIfRequested(); return new CommitReservation(commitGate); }
      catch { Monitor.Exit(commitGate); throw; }
    }
    private sealed class CommitReservation : IDisposable {
      private object gate;
      internal CommitReservation(object value) { gate=value; }
      public void Dispose() { object value=Interlocked.Exchange(ref gate,null); if(value!=null) Monitor.Exit(value); }
    }
  }
  public sealed class CancellationSession : IDisposable {
    [UnmanagedFunctionPointer(CallingConvention.Winapi)]
    [return:MarshalAs(UnmanagedType.Bool)]
    private delegate bool Handler(uint kind);
    [DllImport("kernel32.dll",SetLastError=true)]
    [return:MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetConsoleCtrlHandler(Handler handler,[MarshalAs(UnmanagedType.Bool)] bool add);
    private static readonly object gate=new object();
    // If Windows refuses unregistration, retain the inactive delegate so a
    // later native callback cannot target garbage-collected managed code.
    private static readonly List<CancellationSession> retainedHandlers=new List<CancellationSession>();
    private static CancellationSession current;
    private readonly CancellationSession previous;
    private readonly Handler handler;
    private int active=1, disposed;
    public CancellationState State { get; private set; }
    public bool Registered { get; private set; }
    public int RegistrationErrorCode { get; private set; }
    public int UnregistrationErrorCode { get; private set; }
    public static CancellationState Current { get { lock(gate) { return current==null?null:current.State; } } }
    private CancellationSession(CancellationState state,bool capture) {
      State=state;
      previous=current;
      handler=Handle;
      if(capture) {
        Registered=SetConsoleCtrlHandler(handler,true);
        if(!Registered) RegistrationErrorCode=Marshal.GetLastWin32Error();
      }
      current=this;
    }
    public static CancellationSession Begin(CancellationState state,bool captureConsoleControl) {
      if(state==null) throw new ArgumentNullException("state");
      lock(gate) { return new CancellationSession(state,captureConsoleControl); }
    }
    private bool Handle(uint kind) {
      if(kind!=0 || Interlocked.CompareExchange(ref active,0,0)==0) return false;
      // No PowerShell callback, logging, process kill, locks, or filesystem I/O.
      State.RequestFromConsole(); return true;
    }
    public void Dispose() {
      lock(gate) {
        if(Interlocked.CompareExchange(ref disposed,0,0)!=0) return;
        if(current!=this) throw new InvalidOperationException("Cancellation sessions must be disposed in reverse scope order.");
        Interlocked.Exchange(ref disposed,1); Interlocked.Exchange(ref active,0);
        if(Registered && !SetConsoleCtrlHandler(handler,false)) { UnregistrationErrorCode=Marshal.GetLastWin32Error(); retainedHandlers.Add(this); }
        current=previous;
      }
    }
  }

  public sealed class NativeResult {
    public int? ExitCode;
    public string StdOut = "", StdErr = "", StartError = "";
    public long StdOutCharacters, StdErrCharacters, ElapsedMilliseconds;
    public bool StdOutTruncated, StdErrTruncated, TimedOut;
    public bool Cancelled { get; internal set; }
    public bool StreamsComplete = true, JobAssigned, TreeTerminated, DrainTimedOut;
    public int Win32ErrorCode, CaptureLimit, ProcessId;
  }
  internal sealed class BoundedCapture {
    private readonly char[] ring, head;
    private int next;
    private long total;
    private readonly object gate = new object();
    public string ReadError = "";
    public BoundedCapture(int limit) { ring = new char[limit]; head = new char[limit / 4]; }
    public void Append(char[] block, int count) {
      lock (gate) {
        for (int i = 0; i < count; i++) {
          if (total < head.Length) head[(int)total] = block[i];
          ring[next] = block[i]; next = (next + 1) % ring.Length; total++;
        }
      }
    }
    public long Count { get { lock (gate) { return total; } } }
    public bool Truncated { get { lock (gate) { return total > ring.Length; } } }
    public string Text() {
      lock (gate) {
        if (total <= ring.Length) return new string(ring, 0, (int)total);
        string marker = "\n[... truncated; total " + total.ToString(System.Globalization.CultureInfo.InvariantCulture) + " characters ...]\n";
        int tailLength = ring.Length - head.Length - marker.Length;
        StringBuilder text = new StringBuilder(ring.Length);
        text.Append(head); text.Append(marker);
        int start = (next - tailLength + ring.Length) % ring.Length;
        for (int i = 0; i < tailLength; i++) text.Append(ring[(start + i) % ring.Length]);
        return text.ToString();
      }
    }
  }
  internal sealed class PipeDrain {
    private IntPtr pipe;
    internal IntPtr ThreadHandle;
    private readonly object handleGate = new object();
    private volatile bool stopping;
    internal bool Started;
    internal readonly ManualResetEvent Ready = new ManualResetEvent(false);
    internal readonly Thread Worker;
    internal readonly BoundedCapture Capture;
    internal PipeDrain(IntPtr handle, BoundedCapture capture) {
      pipe = handle; Capture = capture;
      Worker = new Thread(Read); Worker.IsBackground = true;
    }
    internal void Start() { Worker.Start(); Started=true; }
    private void Read() {
      byte[] bytes = new byte[4096]; char[] chars = new char[4098];
      Decoder decoder = new UTF8Encoding(false, false).GetDecoder();
      try {
        lock(handleGate) {
          ThreadHandle = NativeProcess.OpenThread(1, false, NativeProcess.GetCurrentThreadId());
          if (ThreadHandle == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
        }
        Ready.Set();
        while (!stopping) {
          uint count;
          bool ok = NativeProcess.ReadFile(pipe, bytes, (uint)bytes.Length, out count, IntPtr.Zero);
          if (!ok) {
            int error = Marshal.GetLastWin32Error();
            if (error == 109) break; // All whitelisted pipe writers have closed.
            throw new Win32Exception(error);
          }
          if (count == 0) break;
          int decoded = decoder.GetChars(bytes, 0, (int)count, chars, 0, false);
          Capture.Append(chars, decoded);
        }
        Capture.Append(chars, decoder.GetChars(bytes, 0, 0, chars, 0, true));
      } catch (Exception error) { Capture.ReadError = error.GetType().Name; }
      finally { Ready.Set(); NativeProcess.Close(ref pipe); lock(handleGate) { NativeProcess.Close(ref ThreadHandle); } }
    }
    internal void Cancel() {
      stopping=true;
      lock(handleGate) { if (ThreadHandle != IntPtr.Zero) NativeProcess.CancelSynchronousIo(ThreadHandle); }
    }
    internal void FinishHandles() {
      // A stuck background reader retains ownership of its pipe; never call
      // StreamReader.Close on another thread or free a pending read's buffer.
      if (!Started) NativeProcess.Close(ref pipe);
      if (!Worker.IsAlive) Ready.Close();
    }
  }
  public static class NativeProcess {
    [StructLayout(LayoutKind.Sequential)] private struct SECURITY_ATTRIBUTES { public int length; public IntPtr descriptor; [MarshalAs(UnmanagedType.Bool)] public bool inherit; }
    [StructLayout(LayoutKind.Sequential)] private struct STARTUPINFO { public int cb; public IntPtr reserved, desktop, title; public uint x, y, xSize, ySize, xCountChars, yCountChars, fillAttribute, flags; public ushort showWindow, cbReserved2; public IntPtr reserved2, stdin, stdout, stderr; }
    [StructLayout(LayoutKind.Sequential)] private struct STARTUPINFOEX { public STARTUPINFO info; public IntPtr attributes; }
    [StructLayout(LayoutKind.Sequential)] private struct PROCESS_INFORMATION { public IntPtr process, thread; public uint processId, threadId; }
    [StructLayout(LayoutKind.Sequential)] private struct JOB_BASIC_LIMIT { public long processTime, jobTime; public uint flags; public UIntPtr minWorkingSet, maxWorkingSet; public uint activeProcessLimit; public UIntPtr affinity; public uint priority, scheduling; }
    [StructLayout(LayoutKind.Sequential)] private struct IO_COUNTERS { public ulong readOps, writeOps, otherOps, readBytes, writeBytes, otherBytes; }
    [StructLayout(LayoutKind.Sequential)] private struct JOB_EXTENDED_LIMIT { public JOB_BASIC_LIMIT basic; public IO_COUNTERS io; public UIntPtr processMemory, jobMemory, peakProcessMemory, peakJobMemory; }
    [StructLayout(LayoutKind.Sequential)] private struct JOB_ACCOUNTING { public long userTime, kernelTime, periodUserTime, periodKernelTime; public uint pageFaults, totalProcesses, activeProcesses, terminatedProcesses; }
    [DllImport("kernel32.dll", SetLastError=true)] private static extern IntPtr CreateJobObjectW(IntPtr attributes, IntPtr name);
    [DllImport("kernel32.dll", SetLastError=true)] private static extern bool SetInformationJobObject(IntPtr job, int kind, ref JOB_EXTENDED_LIMIT value, uint length);
    [DllImport("kernel32.dll", SetLastError=true)] private static extern bool QueryInformationJobObject(IntPtr job, int kind, out JOB_ACCOUNTING value, uint length, IntPtr returned);
    [DllImport("kernel32.dll", SetLastError=true)] private static extern bool TerminateJobObject(IntPtr job, uint exitCode);
    [DllImport("kernel32.dll", SetLastError=true)] private static extern bool IsProcessInJob(IntPtr process, IntPtr job, out bool inJob);
    [DllImport("kernel32.dll", SetLastError=true)] private static extern bool CreatePipe(out IntPtr read, out IntPtr write, ref SECURITY_ATTRIBUTES attributes, uint size);
    [DllImport("kernel32.dll", SetLastError=true)] private static extern bool SetHandleInformation(IntPtr handle, uint mask, uint flags);
    [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)] private static extern IntPtr CreateFileW(string path, uint access, uint share, ref SECURITY_ATTRIBUTES attributes, uint disposition, uint flags, IntPtr template);
    [DllImport("kernel32.dll", SetLastError=true)] private static extern bool InitializeProcThreadAttributeList(IntPtr list, int count, int flags, ref IntPtr size);
    [DllImport("kernel32.dll", SetLastError=true)] private static extern bool UpdateProcThreadAttribute(IntPtr list, uint flags, IntPtr attribute, IntPtr value, IntPtr size, IntPtr previous, IntPtr returned);
    [DllImport("kernel32.dll")] private static extern void DeleteProcThreadAttributeList(IntPtr list);
    [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)] private static extern bool CreateProcessW(string application, StringBuilder command, IntPtr processAttributes, IntPtr threadAttributes, bool inheritHandles, uint flags, IntPtr environment, string directory, ref STARTUPINFOEX startup, out PROCESS_INFORMATION information);
    [DllImport("kernel32.dll", SetLastError=true)] private static extern uint ResumeThread(IntPtr thread);
    [DllImport("kernel32.dll", SetLastError=true)] private static extern uint WaitForSingleObject(IntPtr handle, uint milliseconds);
    [DllImport("kernel32.dll", SetLastError=true)] private static extern bool GetExitCodeProcess(IntPtr process, out uint exitCode);
    [DllImport("kernel32.dll", SetLastError=true)] internal static extern bool ReadFile(IntPtr handle, byte[] buffer, uint requested, out uint read, IntPtr overlapped);
    [DllImport("kernel32.dll", SetLastError=true)] internal static extern IntPtr OpenThread(uint access, bool inherit, uint threadId);
    [DllImport("kernel32.dll")] internal static extern uint GetCurrentThreadId();
    [DllImport("kernel32.dll", SetLastError=true)] internal static extern bool CancelSynchronousIo(IntPtr thread);
    [DllImport("kernel32.dll", SetLastError=true)] private static extern bool CloseHandle(IntPtr handle);
    internal static void Close(ref IntPtr handle) { if (handle != IntPtr.Zero && handle != new IntPtr(-1)) { CloseHandle(handle); handle = IntPtr.Zero; } }
    private static void Check(bool ok) { if (!ok) throw new Win32Exception(Marshal.GetLastWin32Error()); }
    public static string Quote(string value) {
      if (value == null) value = "";
      StringBuilder text = new StringBuilder(); text.Append('"'); int slashes = 0;
      foreach (char ch in value) {
        if (ch == '\\') { slashes++; continue; }
        if (ch == '"') { text.Append('\\', slashes * 2 + 1); text.Append(ch); }
        else { text.Append('\\', slashes); text.Append(ch); }
        slashes = 0;
      }
      text.Append('\\', slashes * 2); text.Append('"'); return text.ToString();
    }
    private static IntPtr EnvironmentBlock(string[] names, string[] values) {
      if (names == null || values == null || names.Length != values.Length) throw new ArgumentException("Environment override arrays must have matching lengths.");
      if (names.Length == 0) return IntPtr.Zero;
      SortedDictionary<string,string> environment = new SortedDictionary<string,string>(StringComparer.OrdinalIgnoreCase);
      foreach (DictionaryEntry pair in Environment.GetEnvironmentVariables()) environment[(string)pair.Key] = (string)pair.Value;
      for (int i=0; i<names.Length; i++) {
        if (String.IsNullOrEmpty(names[i]) || names[i].IndexOf('=') >= 0 || names[i].IndexOf('\0') >= 0 || (values[i] != null && values[i].IndexOf('\0') >= 0)) throw new ArgumentException("Invalid environment override.");
        if (values[i] == null) environment.Remove(names[i]); else environment[names[i]] = values[i];
      }
      StringBuilder block = new StringBuilder();
      foreach (KeyValuePair<string,string> pair in environment) { block.Append(pair.Key); block.Append('='); block.Append(pair.Value); block.Append('\0'); }
      block.Append('\0'); return Marshal.StringToHGlobalUni(block.ToString());
    }
    private static bool EmptyJob(IntPtr job) {
      JOB_ACCOUNTING accounting;
      Check(QueryInformationJobObject(job, 1, out accounting, (uint)Marshal.SizeOf(typeof(JOB_ACCOUNTING)), IntPtr.Zero));
      return accounting.activeProcesses == 0;
    }
    private static int Remaining(Stopwatch clock, int deadline) { return (int)Math.Max(0L, deadline-clock.ElapsedMilliseconds); }
    private static bool CancellationRequested(CancellationState state,NativeResult result) {
      if(state==null || !state.Requested) return false;
      result.Cancelled=true; return true;
    }
    private static bool ReaderReady(PipeDrain reader,Stopwatch clock,int timeout,CancellationState state,NativeResult result) {
      Stopwatch readiness=Stopwatch.StartNew();
      while(true) {
        if(CancellationRequested(state,result)) return false;
        int remaining=Math.Min(Remaining(clock,timeout),Remaining(readiness,2000));
        if(reader.Ready.WaitOne(0)) return true;
        if(remaining==0) return false;
        reader.Ready.WaitOne(Math.Min(50,remaining));
      }
    }
    private static void Error(NativeResult result, Exception error) {
      string message = error.Message; result.StartError = message.Length > 512 ? message.Substring(0,512) : message;
      Win32Exception native = error as Win32Exception; if (native != null) result.Win32ErrorCode = native.NativeErrorCode;
    }
    public static NativeResult Run(string executable, string[] arguments, int limit, int timeout, string[] environmentNames, string[] environmentValues) {
      NativeResult result = new NativeResult(); result.CaptureLimit = limit;
      // Capture one batch state; later scopes must not retarget this launch.
      CancellationState cancellation=CancellationSession.Current;
      Stopwatch clock = Stopwatch.StartNew();
      IntPtr job=IntPtr.Zero, outRead=IntPtr.Zero, outWrite=IntPtr.Zero, errRead=IntPtr.Zero, errWrite=IntPtr.Zero, input=IntPtr.Zero;
      IntPtr attributes=IntPtr.Zero, jobValue=IntPtr.Zero, handleValues=IntPtr.Zero, environment=IntPtr.Zero;
      PROCESS_INFORMATION process = new PROCESS_INFORMATION();
      bool attributesReady=false, launched=false;
      PipeDrain stdout=null, stderr=null;
      BoundedCapture outCapture=null, errCapture=null;
      try {
        if (limit < 1024 || limit > 262144 || timeout <= 0) throw new ArgumentException("Bounded native capture requires a valid limit and positive deadline.");
        if(CancellationRequested(cancellation,result)) return result;
        outCapture = new BoundedCapture(limit); errCapture = new BoundedCapture(limit);
        job = CreateJobObjectW(IntPtr.Zero,IntPtr.Zero); Check(job != IntPtr.Zero);
        Check(SetHandleInformation(job,1,0));
        JOB_EXTENDED_LIMIT jobLimits = new JOB_EXTENDED_LIMIT(); jobLimits.basic.flags=0x2000; // KILL_ON_JOB_CLOSE; no breakaway.
        Check(SetInformationJobObject(job,9,ref jobLimits,(uint)Marshal.SizeOf(typeof(JOB_EXTENDED_LIMIT))));
        SECURITY_ATTRIBUTES security = new SECURITY_ATTRIBUTES(); security.length=Marshal.SizeOf(typeof(SECURITY_ATTRIBUTES)); security.inherit=true;
        Check(CreatePipe(out outRead,out outWrite,ref security,0)); Check(SetHandleInformation(outRead,1,0));
        Check(CreatePipe(out errRead,out errWrite,ref security,0)); Check(SetHandleInformation(errRead,1,0));
        input=CreateFileW("NUL",0x80000000,3,ref security,3,0,IntPtr.Zero); Check(input != new IntPtr(-1));
        IntPtr size=IntPtr.Zero; InitializeProcThreadAttributeList(IntPtr.Zero,2,0,ref size);
        if (size==IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
        attributes=Marshal.AllocHGlobal(size); Check(InitializeProcThreadAttributeList(attributes,2,0,ref size)); attributesReady=true;
        jobValue=Marshal.AllocHGlobal(IntPtr.Size); Marshal.WriteIntPtr(jobValue,job);
        Check(UpdateProcThreadAttribute(attributes,0,new IntPtr(0x2000D),jobValue,new IntPtr(IntPtr.Size),IntPtr.Zero,IntPtr.Zero));
        handleValues=Marshal.AllocHGlobal(3*IntPtr.Size);
        Marshal.WriteIntPtr(handleValues,0,input); Marshal.WriteIntPtr(handleValues,IntPtr.Size,outWrite); Marshal.WriteIntPtr(handleValues,2*IntPtr.Size,errWrite);
        Check(UpdateProcThreadAttribute(attributes,0,new IntPtr(0x20002),handleValues,new IntPtr(3*IntPtr.Size),IntPtr.Zero,IntPtr.Zero));
        environment=EnvironmentBlock(environmentNames,environmentValues);
        STARTUPINFOEX startup=new STARTUPINFOEX(); startup.info.cb=Marshal.SizeOf(typeof(STARTUPINFOEX)); startup.info.flags=0x100;
        startup.info.stdin=input; startup.info.stdout=outWrite; startup.info.stderr=errWrite; startup.attributes=attributes;
        StringBuilder command=new StringBuilder(Quote(executable));
        foreach(string argument in arguments) { command.Append(' '); command.Append(Quote(argument)); }
        if(CancellationRequested(cancellation,result)) return result;
        // JOB_LIST assigns atomically before the initial thread can execute.
        Check(CreateProcessW(executable,command,IntPtr.Zero,IntPtr.Zero,true,0x08080404,environment,null,ref startup,out process));
        launched=true; result.ProcessId=(int)process.processId;
        bool inJob; Check(IsProcessInJob(process.process,job,out inJob));
        if (!inJob) throw new InvalidOperationException("Native process was not assigned to its private job.");
        result.JobAssigned=true;
        Close(ref outWrite); Close(ref errWrite); Close(ref input);
        stdout=new PipeDrain(outRead,outCapture); outRead=IntPtr.Zero; stdout.Start();
        stderr=new PipeDrain(errRead,errCapture); errRead=IntPtr.Zero; stderr.Start();
        bool readersReady=ReaderReady(stdout,clock,timeout,cancellation,result) && ReaderReady(stderr,clock,timeout,cancellation,result);
        if(CancellationRequested(cancellation,result)) return result;
        if(!readersReady || outCapture.ReadError.Length!=0 || errCapture.ReadError.Length!=0) throw new InvalidOperationException("Native output readers could not initialize.");
        if(Remaining(clock,timeout)==0) result.TimedOut=true;
        else {
          if(CancellationRequested(cancellation,result)) return result;
          uint previous=ResumeThread(process.thread); if(previous==0xFFFFFFFF) throw new Win32Exception(Marshal.GetLastWin32Error());
          while(true) {
            if(CancellationRequested(cancellation,result)) break;
            int remaining=Remaining(clock,timeout);
            if(remaining==0) { result.TimedOut=true; break; }
            uint wait=WaitForSingleObject(process.process,(uint)Math.Min(50,remaining));
            if(wait==0) break;
            if(wait!=0x102) throw new Win32Exception(Marshal.GetLastWin32Error());
          }
        }
      } catch(Exception error) { Error(result,error); result.StreamsComplete=false; }
      finally {
        // Even a root that exits normally may leave descendants holding pipes.
        // Termination is restricted to this private job, never PID snapshots.
        Stopwatch cleanup=Stopwatch.StartNew();
        if(launched && job!=IntPtr.Zero) {
          try {
            Check(TerminateJobObject(job,0xE0000001));
            while(!EmptyJob(job) && cleanup.ElapsedMilliseconds<2000) Thread.Sleep(10);
            result.TreeTerminated=EmptyJob(job);
            if(!result.TreeTerminated) { result.StreamsComplete=false; result.StartError="Owned native job did not become empty within termination grace."; }
            if(WaitForSingleObject(process.process,0)==0) { uint code; Check(GetExitCodeProcess(process.process,out code)); result.ExitCode=unchecked((int)code); }
          } catch(Exception error) { Error(result,error); result.StreamsComplete=false; }
        }
        Close(ref outWrite); Close(ref errWrite); Close(ref input);
        if(stdout!=null || stderr!=null) {
          bool outDone=stdout!=null && stdout.Started && stdout.Worker.Join(Remaining(cleanup,2500));
          bool errDone=stderr!=null && stderr.Started && stderr.Worker.Join(Remaining(cleanup,2500));
          if(!outDone || !errDone) {
            result.DrainTimedOut=true;
            // A cancel can race the reader's next ReadFile. Mark it stopping
            // and repeat cancellation during the same finite grace budget.
            while(cleanup.ElapsedMilliseconds<3000 && (!outDone || !errDone)) {
              if(stdout!=null && stdout.Started) { stdout.Cancel(); outDone=stdout.Worker.Join(Math.Min(10,Remaining(cleanup,3000))); }
              if(stderr!=null && stderr.Started) { stderr.Cancel(); errDone=stderr.Worker.Join(Math.Min(10,Remaining(cleanup,3000))); }
            }
          }
          result.StreamsComplete=result.StreamsComplete && !result.DrainTimedOut && outDone && errDone && outCapture.ReadError.Length==0 && errCapture.ReadError.Length==0 && result.TreeTerminated;
        } else if(launched) result.StreamsComplete=false;
        Close(ref process.thread); Close(ref process.process); Close(ref job);
        Close(ref outRead); Close(ref errRead);
        if(stdout!=null) stdout.FinishHandles(); if(stderr!=null) stderr.FinishHandles();
        if(attributesReady) DeleteProcThreadAttributeList(attributes);
        if(attributes!=IntPtr.Zero) Marshal.FreeHGlobal(attributes);
        if(jobValue!=IntPtr.Zero) Marshal.FreeHGlobal(jobValue);
        if(handleValues!=IntPtr.Zero) Marshal.FreeHGlobal(handleValues);
        if(environment!=IntPtr.Zero) Marshal.FreeHGlobal(environment);
        if(outCapture!=null) { result.StdOut=outCapture.Text(); result.StdOutCharacters=outCapture.Count; result.StdOutTruncated=outCapture.Truncated; }
        if(errCapture!=null) { result.StdErr=errCapture.Text(); result.StdErrCharacters=errCapture.Count; result.StdErrTruncated=errCapture.Truncated; }
        result.ElapsedMilliseconds=clock.ElapsedMilliseconds;
        // A request during termination/drain still prevents callers accepting
        // a successful native exit or finalizing its candidate.
        CancellationRequested(cancellation,result);
      }
      return result;
    }
  }
}
'@
}

function Invoke-WinImgNativeProcess {
  param([string]$Executable, [string[]]$Arguments,
    [ValidateRange(1024,262144)][int]$OutputLimit = 16384,
    [ValidateRange(1,2147483647)][int]$TimeoutMilliseconds = 15000,
    [hashtable]$Environment = @{})
  Initialize-WinImgProcessApi
  $names = [string[]]@($Environment.Keys | Sort-Object)
  $values = [string[]]@($names | ForEach-Object { [string]$Environment[$_] })
  return [WinImgNormalizer.NativeProcess]::Run($Executable, $Arguments, $OutputLimit, $TimeoutMilliseconds, $names, $values)
}

# Internal policies may lower these ceilings for synthetic controls, never raise
# them. The public positional interface and installed policy stay unchanged.
function New-WinImgExecutionPolicy {
  param([ValidateRange(1,120000)][int]$TimeoutMilliseconds = 120000,
    [ValidateRange(0,536870912)][long]$MemoryBytes = 536870912,
    [ValidateRange(0,1073741824)][long]$MapBytes = 1073741824,
    [ValidateRange(0,2147483648)][long]$DiskBytes = 2147483648,
    [ValidateRange(1,2)][int]$Threads = 2)
  $values = New-Object 'Collections.Generic.Dictionary[string,long]' ([StringComparer]::Ordinal)
  $values.Add('TimeoutMilliseconds', $TimeoutMilliseconds)
  $values.Add('MemoryBytes', $MemoryBytes); $values.Add('MapBytes', $MapBytes)
  $values.Add('DiskBytes', $DiskBytes); $values.Add('Threads', $Threads)
  return ,([Collections.ObjectModel.ReadOnlyDictionary[string,long]]::new($values))
}

function Resolve-WinImgNativeTemporaryRoot {
  param([string]$Path, [string]$SourceRoot)
  if (-not $Path) { $Path = [IO.Path]::GetTempPath() }
  $root = Normalize-WinImgRootPath ([IO.Path]::GetFullPath($Path))
  Assert-WinImgNoReparseAncestors $root
  if ($SourceRoot) { $root = Assert-WinImgSafeDestination -SourceRoot $SourceRoot -OutputParent $root }
  else { $root = Resolve-WinImgCanonicalDirectory $root }
  # ImageMagick's internal cache path normalization does not preserve extended
  # prefixes. Reserve room for the owned directory plus its magick-* filenames.
  if ($root.Length -gt 160 -or $root.StartsWith('\\?\', [StringComparison]::Ordinal)) {
    throw 'Native temporary root is too long for ImageMagick cache filenames; use a shorter process TMP/TEMP directory.'
  }
  Assert-WinImgNoReparseAncestors $root
  $item = Get-Item -LiteralPath $root -Force -ErrorAction Stop
  if (-not $item.PSIsContainer) { throw 'Native temporary root must be an existing regular directory.' }
  return $root
}

function New-WinImgImageContext {
  param([object]$Policy, [string]$WorkRoot, [scriptblock]$NativeProcessObserver, [string]$TemporaryRoot)
  Assert-WinImgCancellation
  if (-not $Policy) { $Policy = New-WinImgExecutionPolicy }
  # Validate/copy even an explicitly supplied test policy.
  $policyCopy = New-WinImgExecutionPolicy -TimeoutMilliseconds $Policy.TimeoutMilliseconds -MemoryBytes $Policy.MemoryBytes -MapBytes $Policy.MapBytes -DiskBytes $Policy.DiskBytes -Threads $Policy.Threads
  $root = Resolve-WinImgNativeTemporaryRoot $TemporaryRoot
  $temporary = $null
  for ($attempt = 0; $attempt -lt 8; $attempt++) {
    $path = [IO.Path]::Combine($root, ('WinImgNormalizer-cache-' + [guid]::NewGuid().ToString('N')))
    if (New-WinImgExclusiveDirectory $path) { $temporary = $path; break }
  }
  if (-not $temporary) { throw 'Could not allocate an exclusive native cache directory.' }
  return [pscustomobject]@{ Policy=$policyCopy; Clock=[Diagnostics.Stopwatch]::StartNew(); TemporaryPath=$temporary; NativeProcessObserver=$NativeProcessObserver; CleanupSafe=$true }
}

function Get-WinImgRemainingTime {
  param([object]$Context)
  Assert-WinImgCancellation
  $remaining = [long]$Context.Policy.TimeoutMilliseconds - $Context.Clock.ElapsedMilliseconds
  if ($remaining -le 0) {
    $exception = [TimeoutException]::new('Per-image runtime budget exhausted (Category=Timeout); no further native work or finalization attempted.')
    $exception.Data['WinImgNativeDetail'] = 'Exit=none; Category=Timeout; Shared per-image deadline exhausted.'
    throw $exception
  }
  return [int]$remaining
}

function Invoke-WinImgImageProcess {
  param([string]$Executable, [string[]]$Arguments, [object]$Context, [scriptblock]$ProcessRunner)
  Assert-WinImgCancellation
  $policy = if ($Context) { $Context.Policy } else { New-WinImgExecutionPolicy }
  $remaining = if ($Context) { Get-WinImgRemainingTime $Context } else { [int]$policy.TimeoutMilliseconds }
  if ($ProcessRunner) {
    # Preserve the established conversion fixture seam. The public command never
    # supplies it; real execution always uses the bounded wrapper below.
    $result = & $ProcessRunner $Executable $Arguments
  } else {
    $invariant = [Globalization.CultureInfo]::InvariantCulture
    $limits = @('-limit','memory',($policy.MemoryBytes.ToString($invariant)+'B'),
      '-limit','map',($policy.MapBytes.ToString($invariant)+'B'),
      '-limit','disk',($policy.DiskBytes.ToString($invariant)+'B'),
      '-limit','thread',$policy.Threads.ToString($invariant),
      '-limit','time',([Math]::Ceiling($remaining / 1000.0)).ToString($invariant))
    # identify's subcommand must precede its options; limits precede any decode.
    $tokens = if ($Arguments[0] -eq 'identify') { @('identify') + $limits + @($Arguments | Select-Object -Skip 1) } else { $limits + $Arguments }
    $environment = @{}
    if ($Context) {
      Assert-WinImgNoReparseAncestors $Context.TemporaryPath
      foreach ($key in @('TEMP','TMP','MAGICK_TEMPORARY_PATH')) { $environment[$key] = $Context.TemporaryPath }
    }
    if ($Context -and $Context.NativeProcessObserver) {
      $result = & $Context.NativeProcessObserver $Executable $tokens $remaining $environment
    } else {
      $result = Invoke-WinImgNativeProcess -Executable $Executable -Arguments $tokens -TimeoutMilliseconds $remaining -Environment $environment
    }
  }
  if ($Context -and $result.PSObject.Properties['ProcessId'] -and $result.ProcessId -gt 0 -and
      $result.PSObject.Properties['TreeTerminated'] -and -not $result.TreeTerminated) { $Context.CleanupSafe = $false }
  if ($result.PSObject.Properties['Cancelled'] -and $result.Cancelled) {
    $state = [WinImgNormalizer.CancellationSession]::Current
    if ($state) { $state.Request() }
    throw [OperationCanceledException]::new('Owned native work was cancelled; no finalization attempted.')
  }
  Assert-WinImgCancellation
  if ($Context -and $Context.Clock.ElapsedMilliseconds -ge $policy.TimeoutMilliseconds -and -not $result.TimedOut) {
    # A successful native exit arriving beyond the shared deadline cannot revive
    # an earlier candidate or authorize another phase.
    $null = Get-WinImgRemainingTime $Context
  }
  return $result
}

function Remove-WinImgImageContext {
  param([object]$Context)
  if (-not $Context) { return }
  $Context.Clock.Stop()
  if (-not $Context.CleanupSafe) { throw 'Native tree termination is unconfirmed; cache and item scratch were preserved.' }
  Assert-WinImgNoReparseAncestors $Context.TemporaryPath
  # The native job has stopped before this runs. Only ImageMagick's regular
  # magick-* cache files in this exclusively allocated directory are eligible.
  # Foreign names, subdirectories and reparse arrivals survive with a warning.
  foreach ($entry in [IO.Directory]::EnumerateFileSystemEntries($Context.TemporaryPath)) {
    $item = Get-Item -LiteralPath $entry -Force -ErrorAction Stop
    if ($item.PSIsContainer -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -or $item.Name -notmatch '^magick-[A-Za-z0-9_-]+$') {
      throw 'Owned native temporary directory contains an unexpected entry; it was preserved.'
    }
    [IO.File]::Delete($entry)
  }
  [IO.Directory]::Delete($Context.TemporaryPath, $false)
}

function Get-WinImgBoundedText {
  param([string]$Text, [int]$Limit = 16384)
  if ($Text.Length -le $Limit) { return $Text }
  return $Text.Substring(0, $Limit / 4) + "`n[... truncated ...]`n" + $Text.Substring($Text.Length - ($Limit * 3 / 4 - 32))
}

function Get-WinImgNativeOutcome {
  param([object]$Result, [switch]$AllowStdOut, [switch]$Strict)
  $code = $null; $stdout = ''; $stderr = ''; $startError = ''; $win32 = 0
  $category = 'InvalidResult'; $warning = $false; $acceptable = $false; $retryable = $false
  $outCount = 0L; $errCount = 0L; $overflow = $false; $incomplete = $false; $captureLimit = 16384
  if ($Result -is [int] -or $Result -is [long]) { $code = $Result; $category = 'NativeFailure' }
  elseif ($null -ne $Result -and $Result -isnot [array] -and $Result.PSObject.Properties['ExitCode']) {
    $value = $Result.ExitCode
    if ($value -is [int] -or $value -is [long]) { $code = $value; $category = 'NativeFailure' }
    if ($Result.PSObject.Properties['StdOut']) { $stdout = [string]$Result.StdOut }
    if ($Result.PSObject.Properties['StdErr']) { $stderr = [string]$Result.StdErr }
    elseif ($Result.PSObject.Properties['DiagnosticOutput']) { $stderr = [string]$Result.DiagnosticOutput }
    $outCount = $stdout.Length; $errCount = $stderr.Length
    if ($Result.PSObject.Properties['StdOutCharacters']) { $outCount = [long]$Result.StdOutCharacters }
    if ($Result.PSObject.Properties['StdErrCharacters']) { $errCount = [long]$Result.StdErrCharacters }
    if ($Result.PSObject.Properties['CaptureLimit'] -and $Result.CaptureLimit -ge 1024 -and $Result.CaptureLimit -le 262144) { $captureLimit = [int]$Result.CaptureLimit }
    $overflow = $stdout.Length -gt $captureLimit -or $stderr.Length -gt $captureLimit -or $Result.StdOutTruncated -or $Result.StdErrTruncated
    $stdout = Get-WinImgBoundedText $stdout $captureLimit; $stderr = Get-WinImgBoundedText $stderr $captureLimit
    $startError = Get-WinImgBoundedText ([string]$Result.StartError) 512
    if ($Result.PSObject.Properties['Win32ErrorCode']) { $win32 = [int]$Result.Win32ErrorCode }
    $incomplete = ($Result.PSObject.Properties['StreamsComplete'] -and -not $Result.StreamsComplete) -or
      ($Result.PSObject.Properties['ProcessId'] -and $Result.ProcessId -gt 0 -and $Result.PSObject.Properties['TreeTerminated'] -and -not $Result.TreeTerminated)
    if ($Result.TimedOut) { $category = 'Timeout' }
    elseif ($Result.Cancelled) { $category = 'Cancelled' }
    elseif ($overflow) { $category = 'OutputLimit' }
    elseif ($startError) { $category = if ($Result.PSObject.Properties['ProcessId'] -and $Result.ProcessId -gt 0) { 'NativeFailure' } elseif ($win32 -in @(32,33)) { 'TransientIO' } elseif ($win32 -eq 5) { 'AccessDenied' } else { 'StartFailure' } }
    elseif ($incomplete) { $category = 'NativeFailure' }
  }
  $terminal = $category -in @('InvalidResult','Timeout','Cancelled','OutputLimit','StartFailure') -or
    $incomplete
  if (-not $terminal -and -not $startError) {
    $diagnostics = $stderr + "`n" + $(if (-not $AllowStdOut) { $stdout } else { '' })
    $lines = @($stderr -split '\r?\n' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    # Only the pinned JPEG writer's intended lossless-to-lossy derivative notice
    # is allowed, and only at native exit zero. ICC/decode/unknown messages fail.
    $benign = $lines.Count -gt 0 -and @($lines | Where-Object {
      $_ -notmatch "^magick(?:\.exe)?: lossless to lossy JPEG conversion [\x60'].*' @ warning/jpeg\.c/WriteJPEGImage_/\d+\.$"
    }).Count -eq 0
    if ($code -eq 0 -and ([string]::IsNullOrWhiteSpace($stdout) -or $AllowStdOut) -and [string]::IsNullOrWhiteSpace($stderr)) { $category = 'Success'; $acceptable = $true }
    elseif ($code -eq 0 -and $benign -and -not $Strict -and ([string]::IsNullOrWhiteSpace($stdout) -or $AllowStdOut)) { $category = 'Warning'; $acceptable = $true; $warning = $true }
    elseif ($diagnostics -match '(?i)cache resources exhausted|memory allocation failed|no space left on device|disk quota exceeded|exceeds (?:user )?limit|time limit exceeded') { $category = 'ResourceExhaustion' }
    elseif ($diagnostics -match '(?i)permission denied|access is denied|operation not permitted|not authorized|security policy') { $category = 'AccessDenied' }
    elseif ($diagnostics -match '(?i)no (?:decode|encode) delegate for this image format|unable to load module|delegate library support not built-in') { $category = 'MissingCodec' }
    elseif ($diagnostics -match '(?i)improper image header|corrupt image|insufficient image data|unexpected end.of.file|premature end|invalid (?:image|profile)|colorspacecolorprofilemismatch|unable to create color transform|profile.*(?:mismatch|invalid|corrupt)|CRC error') { $category = 'DamagedInput' }
    elseif ($code -ne 0 -and [string]::IsNullOrWhiteSpace($stdout) -and $lines.Count -gt 0 -and @($lines | Where-Object {
      # Match a reason in a bounded whole native line, never a quoted filename.
      $_ -notmatch "^magick(?:\.exe)?: (?:(?:sharing|lock) violation|unable to open image [\x60'][^\r\n]*': (?:sharing violation|lock violation|The process cannot access the file because it is being used by another process)\.?) @ error/blob\.c/OpenBlob/\d+\.$"
    }).Count -eq 0) { $category = 'TransientIO' }
    else { $category = if ($code -eq 0) { 'DiagnosticFailure' } else { 'NativeFailure' } }
  }
  $retryable = $category -eq 'TransientIO'
  $displayCode = if ($null -eq $code) { 'none' } else { $code.ToString([Globalization.CultureInfo]::InvariantCulture) }
  $detail = 'Exit=' + $displayCode + '; Category=' + $category + '; StdOutCharacters=' + $outCount + '; StdErrCharacters=' + $errCount + '; Truncated=' + $overflow
  if ($null -ne $Result -and $Result.PSObject.Properties['JobAssigned']) { $detail += '; JobAssigned=' + $Result.JobAssigned + '; TreeTerminated=' + $Result.TreeTerminated + '; DrainTimedOut=' + $Result.DrainTimedOut + '; ElapsedMs=' + $Result.ElapsedMilliseconds }
  if ($startError) { $detail += "`nStartError: " + $startError + '; Win32ErrorCode=' + $win32 }
  if ($stdout) { $detail += "`nStdOut: " + $stdout }
  if ($stderr) { $detail += "`nStdErr: " + $stderr }
  return [pscustomobject]@{ ExitCode=$code; Category=$category; Acceptable=$acceptable; Retryable=$retryable; Warning=$warning; DiagnosticText=$detail }
}

function Assert-WinImgNativeQuery {
  param([object]$Result, [string]$Context, [switch]$AllowStdOut)
  if ($Result -and $Result.PSObject.Properties['Cancelled'] -and $Result.Cancelled) {
    $state = if ('WinImgNormalizer.CancellationSession' -as [type]) { [WinImgNormalizer.CancellationSession]::Current } else { $null }
    if ($state) { $state.Request() }
    throw [OperationCanceledException]::new('Native query was cancelled.')
  }
  Assert-WinImgCancellation
  $outcome = Get-WinImgNativeOutcome -Result $Result -AllowStdOut:$AllowStdOut -Strict
  if (-not $outcome.Acceptable) {
    $exception = [InvalidOperationException]::new($Context + ' failed (Category=' + $outcome.Category + '; Exit=' + $outcome.ExitCode + '); see native details.')
    $exception.Data['WinImgNativeDetail'] = $outcome.DiagnosticText
    throw $exception
  }
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
  Assert-WinImgCancellation
  $state = if ('WinImgNormalizer.CancellationSession' -as [type]) { [WinImgNormalizer.CancellationSession]::Current } else { $null }
  # A fully validated move reserved before Request may finish. Never hold the
  # callback's state lock across filesystem I/O or reserve an incomplete file.
  $commit = if ($state) { $state.EnterCommit() } else { $null }
  try { [IO.File]::Move($CandidatePath, $DestinationPath) }
  finally { if ($commit) { $commit.Dispose() } }
}

function Copy-WinImgPlannedVideo {
  param([string]$SourcePath, [string]$DestinationPath, [string]$WorkRoot, [object]$ReportRow)
  Assert-WinImgCancellation
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
    # Record the linearized final move before owned scratch cleanup can fail.
    if ($ReportRow) {
      $ReportRow.InputBytes=[long]$length; $ReportRow.OutputBytes=[long]$length
      $ReportRow.OutputRelativePath=$ReportRow.PlannedOutputRelativePath; $ReportRow.Attempts=1
      $ReportRow.Reason='Stable byte-for-byte video copy finalized'; $ReportRow.Status='CopiedVideo'
    }
  } finally {
    try { Remove-WinImgOwnedCandidate -CandidatePath $candidate -FileOwned $fileOwned }
    catch { if (Test-WinImgCancellationException $_) { throw }; $cleanupWarning = $_.Exception.Message; Write-Warning ('Could not remove owned video scratch: ' + $cleanupWarning) }
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
      Copy-WinImgCancellableStream -InputStream $inputStream -OutputStream $outputStream -SourcePath $SourcePath -CandidatePath $CandidatePath
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
  param([string]$CandidatePath, [string]$MagickPath, [object]$NativeContext)
  Assert-WinImgNoReparseAncestors ([IO.Path]::GetDirectoryName($CandidatePath))
  $before = Get-Item -LiteralPath $CandidatePath -Force -ErrorAction Stop
  if ($before.PSIsContainer -or ($before.Attributes -band [IO.FileAttributes]::ReparsePoint) -or $before.Length -le 0) {
    throw 'ImageMagick produced no nonempty regular candidate.'
  }
  $length = $before.Length
  $modified = $before.LastWriteTimeUtc
  $nativePath = Get-WinImgNativeOutputPath $CandidatePath
  # Explicit +ping fully decodes pixels; warnings remain fatal for validation.
  $result = Invoke-WinImgImageProcess -Executable $MagickPath -Arguments @('identify','+ping','-regard-warnings','-define','registry:filename:literal=true','-format','%m|%w|%h|%n',$nativePath) -Context $NativeContext
  Assert-WinImgNativeQuery -Result $result -Context 'Candidate full JPEG decode' -AllowStdOut
  if ($result.StdOut -notmatch '^JPEG\|([1-9][0-9]*)\|([1-9][0-9]*)\|1$') {
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

# Colour inspection and conversion read the same private, stable copy. No
# selector is appended to a user filename; only the decoded first image is kept.
function New-WinImgSourceSnapshot {
  param([string]$SourcePath, [string]$WorkRoot, [long]$ExpectedLength,
    [datetime]$ExpectedModified, [System.Collections.Generic.List[object]]$OwnedCandidates)

  Assert-WinImgCancellation
  Assert-WinImgNoReparseAncestors ([IO.Path]::GetDirectoryName($SourcePath))
  $before = Get-Item -LiteralPath $SourcePath -Force -ErrorAction Stop
  if ($before.PSIsContainer -or ($before.Attributes -band [IO.FileAttributes]::ReparsePoint) -or
      $before.Length -ne $ExpectedLength -or $before.LastWriteTimeUtc -ne $ExpectedModified) {
    throw 'Image source changed before frame/page or colour inspection.'
  }
  $allocated = New-WinImgImageCandidate -WorkRoot $WorkRoot
  $snapshot = [IO.Path]::Combine([IO.Path]::GetDirectoryName($allocated), ('source' + $before.Extension.ToLowerInvariant()))
  $ownership = [pscustomobject]@{ Path = $snapshot; FileOwned = $false; Removed = $false }
  $OwnedCandidates.Add($ownership)
  $inputStream = [IO.FileStream]::new($SourcePath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
  try {
    $outputStream = [IO.FileStream]::new($snapshot, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    $ownership.FileOwned = $true
    try { Copy-WinImgCancellableStream -InputStream $inputStream -OutputStream $outputStream -SourcePath $SourcePath -CandidatePath $snapshot }
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
  param([string]$SnapshotPath, [string]$MagickPath, [object]$NativeContext)
  Assert-WinImgNoReparseAncestors ([IO.Path]::GetDirectoryName($SnapshotPath))
  $nativePath = Get-WinImgNativeOutputPath $SnapshotPath
  # Count all images before the selected first-image read. This is header
  # inspection; the final JPEG separately receives a full pixel decode.
  $result = Invoke-WinImgImageProcess -Executable $MagickPath -Arguments @('identify','-ping','-regard-warnings','-define','registry:filename:literal=true','-format','%m|%n|%w|%h\n',$nativePath) -Context $NativeContext
  Assert-WinImgNativeQuery -Result $result -Context 'Source frame/page inspection' -AllowStdOut
  $text = $result.StdOut -replace '\r?\n$', ''
  if ([string]::IsNullOrEmpty($text)) { throw 'Source frame/page inspection failed; no image was selected.' }
  $metadata = @($text -split '\r?\n')
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

# Inspect the selected snapshot before stripping. Clear only free-form names
# that could mask built-in metadata; the real ICC profile remains attached.
function Get-WinImgColourInfo {
  param([string]$SnapshotPath, [string]$MagickPath, [object]$NativeContext)
  Assert-WinImgNoReparseAncestors ([IO.Path]::GetDirectoryName($SnapshotPath))
  $nativePath = Get-WinImgNativeOutputPath $SnapshotPath
  # profiles=none is an option fallback; leave the real ICC attached.
  $result = Invoke-WinImgImageProcess -Executable $MagickPath -Arguments @('identify','-ping','-regard-warnings','-define','registry:filename:literal=true','-define','image:frames=0','-define','profiles=none','+set','profiles','+set','colorspace','-format','%[colorspace]|%[profiles]',$nativePath) -Context $NativeContext
  Assert-WinImgNativeQuery -Result $result -Context 'Colour/profile inspection' -AllowStdOut
  if ($result.StdOut -notmatch '^([A-Za-z][A-Za-z0-9]*)\|([A-Za-z0-9, _:-]+)$') {
    throw 'Colour/profile inspection failed or was ambiguous; no accurate conversion was claimed.'
  }
  $colourSpace = $Matches[1]
  $profiles = @($Matches[2].ToLowerInvariant().Split(',') | ForEach-Object { $_.Trim() })
  $hasIcc = $profiles -contains 'icc' -or $profiles -contains 'icm'
  if ($hasIcc) { $policy = 'ProfileToSrgb' }
  elseif ($colourSpace -eq 'sRGB') { $policy = 'AssumeSrgb' }
  elseif ($colourSpace -eq 'RGB') { $policy = 'ConvertLinearRgb' }
  elseif ($colourSpace -in @('Gray', 'LinearGray')) { $policy = 'ConvertGray' }
  elseif ($colourSpace -eq 'CMYK') {
    throw 'Untagged CMYK has no source ICC characterization; accurate sRGB colour cannot be inferred. Supply a correctly profiled source.'
  } else {
    throw ('Unprofiled ' + $colourSpace + ' colour is unsupported; supply a correctly profiled source instead of guessing its characterization.')
  }
  return [pscustomobject]@{ ColourSpace = $colourSpace; HasIcc = [bool]$hasIcc; Policy = $policy }
}

function New-WinImgSourceIccProfile {
  param([string]$SnapshotPath, [string]$MagickPath, [string]$ColourSpace,
    [string]$WorkRoot, [System.Collections.Generic.List[object]]$OwnedCandidates, [object]$NativeContext)
  $allocated = New-WinImgImageCandidate -WorkRoot $WorkRoot
  $profilePath = [IO.Path]::Combine([IO.Path]::GetDirectoryName($allocated), 'source.icc')
  $ownership = [pscustomobject]@{ Path = $profilePath; FileOwned = $false; Removed = $false }
  $OwnedCandidates.Add($ownership)
  $reservation = [IO.FileStream]::new($profilePath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
  $ownership.FileOwned = $true
  $reservation.Dispose()
  $nativeSource = Get-WinImgNativeOutputPath $SnapshotPath
  $nativeProfile = Get-WinImgNativeOutputPath $profilePath
  # Extract only the selected embedded ICC; every diagnostic is fatal here.
  $result = Invoke-WinImgImageProcess -Executable $MagickPath -Arguments @('-ping','-regard-warnings','-define','registry:filename:literal=true','-define','image:frames=0',$nativeSource,('ICC:' + $nativeProfile)) -Context $NativeContext
  Assert-WinImgNativeQuery -Result $result -Context 'Embedded ICC extraction'
  Assert-WinImgNoReparseAncestors ([IO.Path]::GetDirectoryName($profilePath))
  $item = Get-Item -LiteralPath $profilePath -Force -ErrorAction Stop
  if ($item.PSIsContainer -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -or $item.Length -lt 132) {
    throw 'Embedded ICC is missing or truncated; no accurate colour conversion was claimed.'
  }
  $length = $item.Length
  $stream = [IO.FileStream]::new($profilePath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
  try {
    $header = New-Object byte[] 132
    if ($stream.Read($header, 0, 132) -ne 132) { throw 'Embedded ICC header could not be read completely.' }
    function Read-IccUInt32([byte[]]$Bytes, [int]$Offset) {
      return [long]$Bytes[$Offset] * 16777216 + [long]$Bytes[$Offset + 1] * 65536 + [long]$Bytes[$Offset + 2] * 256 + [long]$Bytes[$Offset + 3]
    }
    $declared = Read-IccUInt32 $header 0
    $signature = [Text.Encoding]::ASCII.GetString($header, 36, 4)
    $model = [Text.Encoding]::ASCII.GetString($header, 16, 4).Trim()
    $tags = Read-IccUInt32 $header 128
    $tableEnd = 132 + 12 * $tags
    if ($declared -ne $length -or $signature -ne 'acsp' -or $tags -eq 0 -or $tableEnd -gt $length) {
      throw 'Embedded ICC size, signature or tag table is invalid; no accurate colour conversion was claimed.'
    }
    $allowedModels = switch ($ColourSpace) {
      'sRGB' { 'RGB' }; 'RGB' { 'RGB' }; 'CMYK' { 'CMYK' }
      'Gray' { 'GRAY'; 'RGB' }; 'LinearGray' { 'GRAY'; 'RGB' }
      default { @() }
    }
    if ($model -notin @($allowedModels)) {
      throw ('Embedded ICC model ' + $model + ' does not match supported decoded ' + $ColourSpace + ' colour; no accurate conversion was claimed.')
    }
    $tag = New-Object byte[] 12
    for ($index = 0; $index -lt $tags; $index++) {
      if ($stream.Read($tag, 0, 12) -ne 12) { throw 'Embedded ICC tag table is truncated.' }
      $offset = Read-IccUInt32 $tag 4
      $size = Read-IccUInt32 $tag 8
      # Shared tag payloads occur in compact profiles. Only
      # enforce table/payload bounds; LCMS and native diagnostics check semantics.
      if ($offset -lt $tableEnd -or $size -lt 8 -or $offset + $size -gt $length) {
        throw 'Embedded ICC tag payload lies outside the profile; no accurate colour conversion was claimed.'
      }
    }
  } finally { $stream.Dispose() }
  return $profilePath
}

function New-WinImgSrgbProfile {
  param([string]$WorkRoot, [System.Collections.Generic.List[object]]$OwnedCandidates)
  # Compact-ICC-Profiles sRGB-v4, CC0-1.0, immutable upstream commit
  # bdd84663061bc4ae95ca70decff54f581e27f702. RGB/XYZ matrix target, 480 bytes.
  # SHA256 c56e1685d888f5edb92fe07f2750f387f8fe8e91b32ff8fb0b56bfbbb9458353
  # https://github.com/saucecontrol/Compact-ICC-Profiles (license/provenance in tests/fixtures/colour).
  # Embedded bytes keep the standalone script independent of installed ICC paths.
  $bytes = [Convert]::FromBase64String('AAAB4GxjbXMEIAAAbW50clJHQiBYWVogB+IAAwAUAAkADgAdYWNzcE1TRlQAAAAAc2F3c2N0cmwAAAAAAAAAAAAAAAAAAPbWAAEAAAAA0y1oYW5keem/Vlo+AbaDI4VVRvdPqgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKZGVzYwAAAPwAAAAkY3BydAAAASAAAAAid3RwdAAAAUQAAAAUY2hhZAAAAVgAAAAsclhZWgAAAYQAAAAUZ1hZWgAAAZgAAAAUYlhZWgAAAawAAAAUclRSQwAAAcAAAAAgZ1RSQwAAAcAAAAAgYlRSQwAAAcAAAAAgbWx1YwAAAAAAAAABAAAADGVuVVMAAAAIAAAAHABzAFIARwBCbWx1YwAAAAAAAAABAAAADGVuVVMAAAAGAAAAHABDAEMAMAAAWFlaIAAAAAAAAPbWAAEAAAAA0y1zZjMyAAAAAAABDD8AAAXd///zJgAAB5AAAP2S///7of///aIAAAPcAADAcVhZWiAAAAAAAABvoAAAOPIAAAOPWFlaIAAAAAAAAGKWAAC3iQAAGNpYWVogAAAAAAAAJKAAAA+FAAC2xHBhcmEAAAAAAAMAAAACZmkAAPKnAAANWQAAE9AAAApb')
  $allocated = New-WinImgImageCandidate -WorkRoot $WorkRoot
  $profilePath = [IO.Path]::Combine([IO.Path]::GetDirectoryName($allocated), 'sRGB.icc')
  $ownership = [pscustomobject]@{ Path = $profilePath; FileOwned = $false; Removed = $false }
  $OwnedCandidates.Add($ownership)
  $stream = [IO.FileStream]::new($profilePath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
  $ownership.FileOwned = $true
  try { $stream.Write($bytes, 0, $bytes.Length); $stream.Flush() }
  finally { $stream.Dispose() }
  return $profilePath
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
  # Format tables have a larger, still bounded budget than per-item details.
  return Invoke-WinImgNativeProcess -Executable $Executable -Arguments $Arguments -OutputLimit 262144 -TimeoutMilliseconds 15000
}

function Get-WinImgMagickInfo {
  param([string]$MagickPath, [scriptblock]$PreflightRunner)
  $executable = Resolve-WinImgMagickApplication $MagickPath
  if (-not $PreflightRunner) { $PreflightRunner = { param($Executable, $Arguments) Invoke-WinImgPreflightProcess $Executable $Arguments } }
  $version = & $PreflightRunner $executable @('-version')
  Assert-WinImgNativeQuery -Result $version -Context 'ImageMagick version query; no files were processed' -AllowStdOut
  $match = [regex]::Match($version.StdOut, '(?m)^Version: ImageMagick (7)\.(\d+)\.(\d+)-(\d+)\b')
  if (-not $match.Success) { throw 'Unrecognized ImageMagick version; use a supported ImageMagick 7 build.' }
  $build = [version]($match.Groups[1].Value + '.' + $match.Groups[2].Value + '.' + $match.Groups[3].Value + '.' + $match.Groups[4].Value)
  # Reviewed 2026-10-04: 7.1.2-32 includes jpeg:extent hang and JPEG/GIF/XMP fixes.
  if ($build -lt [version]'7.1.2.32') { throw 'ImageMagick 7.1.2-32 or newer supported 7.x build is required; update the dependency before running.' }
  $formatResult = & $PreflightRunner $executable @('-list','format')
  Assert-WinImgNativeQuery -Result $formatResult -Context 'ImageMagick format query; no files were processed' -AllowStdOut
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
  try { return ([IO.DriveInfo]::new([IO.Path]::GetPathRoot($Path))).AvailableFreeSpace } catch { if (Test-WinImgCancellationException $_) { throw }; return $null }
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
# Reporting must not retry a broken sink or recursively log its own failure.
function ConvertTo-WinImgLogText {
  param([string]$Text)
  return [regex]::Replace($Text, '[\p{Cc}\p{Cf}\p{Cs}\p{Zl}\p{Zp}]', {
    param($match)
    return ('\u{0:x4}' -f [int][char]$match.Value)
  })
}

function Write-WinImgEmergencyReport {
  param([string]$Message)
  # The PowerShell pipeline may already be stopped. stderr is best effort only;
  # an absent/closed console cannot be repaired by recursively calling Write-Log.
  try { [Console]::Error.WriteLine($Message) } catch { if (Test-WinImgCancellationException $_) { throw };}
}

function New-WinImgLogState {
  param([string]$Path)
  return [pscustomobject]@{
    Path = $Path; DiskEnabled = $false; Degraded = $false; FailureCount = 0
    FailurePhase = $null; FailureReason = $null
    Fallback = [Text.StringBuilder]::new(); FallbackDroppedLines = [long]0; FallbackTruncatedLines = [long]0
    MaxFallbackCharacters = 8192; MaxFallbackLineCharacters = 1024; FallbackEmitted = $false
  }
}

function Add-WinImgLogFallback {
  param([object]$State, [string]$Line)
  if ($Line.Length -gt $State.MaxFallbackLineCharacters) { $State.FallbackTruncatedLines++ }
  $bounded = (Get-WinImgBoundedText $Line $State.MaxFallbackLineCharacters) + [Environment]::NewLine
  if ($State.Fallback.Length + $bounded.Length -le $State.MaxFallbackCharacters) {
    $null = $State.Fallback.Append($bounded)
  } else { $State.FallbackDroppedLines++ }
}

function Set-WinImgLogFailure {
  param([object]$State, [string]$Phase, [string]$Reason)
  if ($State.Degraded) { return }
  $State.DiskEnabled = $false; $State.Degraded = $true; $State.FailureCount++
  $State.FailurePhase = $Phase
  $State.FailureReason = Get-WinImgBoundedText (ConvertTo-WinImgLogText $Reason) 512
  $notice = 'Log warning: {0} failed ({1}); disk logging disabled. Continuing with a bounded console fallback; the run is degraded.' -f $Phase, $State.FailureReason
  Add-WinImgLogFallback -State $State -Line $notice
  Write-WinImgEmergencyReport $notice
  try { Write-Host $notice -ForegroundColor Yellow }
  catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog $State; throw }
  catch { if (Test-WinImgCancellationException $_) { throw }; Write-WinImgEmergencyReport 'The normal console report also failed; stderr fallback is best effort.' }
}

function New-WinImgRunLogFile {
  param([string]$Path, [string]$Header)
  $stream = [IO.FileStream]::new($Path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
  try {
    $writer = [IO.StreamWriter]::new($stream, [Text.UTF8Encoding]::new($true))
    try { $writer.WriteLine($Header) } finally { $writer.Dispose() }
  } finally { $stream.Dispose() }
}

function Initialize-WinImgRunLog {
  param([object]$State)
  try {
    New-WinImgRunLogFile -Path $State.Path -Header ("WinImgNormalizer started $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $State.DiskEnabled = $true
  } catch [Management.Automation.PipelineStoppedException] {
    Write-WinImgEmergencyReport 'Reporting interrupted while opening the run log.'; throw
  } catch { if (Test-WinImgCancellationException $_) { throw }; Set-WinImgLogFailure -State $State -Phase 'Creation' -Reason $_.Exception.Message }
}

function Add-WinImgRunLogLine {
  param([string]$Path, [string]$Line)
  Add-Content -LiteralPath $Path -Value $Line -Encoding UTF8 -ErrorAction Stop
}

function Write-WinImgRunLog {
  param([object]$State, [string]$Message, [ValidateSet('INFO','OK','SKIP','WARN','ERR')]$Level='INFO', [switch]$Quiet)
  $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, (ConvertTo-WinImgLogText $Message)
  if ($State.DiskEnabled) {
    try { Add-WinImgRunLogLine -Path $State.Path -Line $line }
    catch [Management.Automation.PipelineStoppedException] {
      Write-WinImgEmergencyReport 'Reporting interrupted; inspect the existing run log.'; throw
    } catch { if (Test-WinImgCancellationException $_) { throw }; Set-WinImgLogFailure -State $State -Phase 'Append' -Reason $_.Exception.Message }
  }
  if ($State.Degraded) { Add-WinImgLogFallback -State $State -Line $line }
  if (-not $Quiet) {
    try {
      switch ($Level) {
        'ERR' { Write-Host $line -ForegroundColor Red }
        'WARN' { Write-Host $line -ForegroundColor Yellow }
        'OK' { Write-Host $line -ForegroundColor Green }
        default { Write-Host $line }
      }
    } catch [Management.Automation.PipelineStoppedException] {
      Complete-WinImgRunLog $State; Write-WinImgEmergencyReport 'Reporting interrupted.'; throw
    } catch { if (Test-WinImgCancellationException $_) { throw }; Write-WinImgEmergencyReport (Get-WinImgBoundedText $line 1024) }
  }
}

function Complete-WinImgRunLog {
  param([object]$State)
  if ($null -eq $State -or -not $State.Degraded -or $State.FallbackEmitted) { return }
  $State.FallbackEmitted = $true
  Write-WinImgEmergencyReport ('Bounded log fallback (LogWarnings={4}; DiskLogIncomplete={5}; UTF-16 characters={0}/{1}; dropped lines={2}; truncated lines={3}):' -f
    $State.Fallback.Length, $State.MaxFallbackCharacters, $State.FallbackDroppedLines, $State.FallbackTruncatedLines, $State.FailureCount, $State.Degraded)
  Write-WinImgEmergencyReport ($State.Fallback.ToString())
  Write-WinImgEmergencyReport 'End of bounded log fallback. Required disk log is incomplete; completed files are retained.'
}

function Set-WinImgOutputTimestampValue {
  param([string]$Path, [ValidateSet('LastWriteTimeUtc','CreationTimeUtc')][string]$Name, [datetime]$Value)
  Assert-WinImgNoReparseAncestors ([IO.Path]::GetDirectoryName($Path))
  $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
  if ($item -isnot [IO.FileInfo] -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
    throw 'The finalized timestamp target is not a regular file.'
  }
  $item.$Name = $Value
}

function Set-WinImgOutputTimestamps {
  param([string]$Path, [datetime]$LastWriteTimeUtc, [datetime]$CreationTimeUtc)
  $failed = New-Object 'Collections.Generic.List[object]'
  # These are separate filesystem writes. Attempt both and expose partial failure.
  foreach ($field in @('LastWriteTimeUtc','CreationTimeUtc')) {
    $value = if ($field -eq 'LastWriteTimeUtc') { $LastWriteTimeUtc } else { $CreationTimeUtc }
    try { Set-WinImgOutputTimestampValue -Path $Path -Name $field -Value $value }
    catch [Management.Automation.PipelineStoppedException] {
      Write-WinImgEmergencyReport 'Timestamp reporting interrupted after finalization; valid output is retained.'; throw
    } catch { if (Test-WinImgCancellationException $_) { throw }; $failed.Add([pscustomobject]@{ Name = $field; Reason = Get-WinImgBoundedText (ConvertTo-WinImgLogText $_.Exception.Message) 512 }) }
  }
  return [pscustomobject]@{ Succeeded = ($failed.Count -eq 0); FailedFields = $failed.ToArray() }
}

function New-WinImgReportState {
  return [pscustomobject]@{
    Rows = New-Object 'Collections.Generic.List[object]'
    BySource = New-Object 'Collections.Generic.Dictionary[string,object]' ([StringComparer]::Ordinal)
    ScanIssues = New-Object 'Collections.Generic.List[object]'
    ImageExtensions = @(); VideoExtensions = @(); SourceRoot = $null; DestinationRoot = $null
    ScanComplete = $false; IncompleteDirectories = 0; UninspectableEntries = 0; SkippedLinks = 0
    Clock = [Diagnostics.Stopwatch]::StartNew(); StartedAtUtc = [DateTime]::UtcNow
    CsvPath = $null; ReportAttempted = $false; ReportComplete = $false; ReportWarnings = 0
    ReportFailurePhase = $null; ReportFailureReason = $null; ConsoleWarnings = 0; ConsoleFailureReason = $null; ProgressWarnings = 0; ProgressEnabled = $true
    RunState = 'Incomplete'; ProcessingWarnings = $false; Summary = $null; Observed = $false
  }
}

function Add-WinImgDiscoveredFile {
  param([object]$State, [string]$SourceRoot, [IO.FileInfo]$File)
  $relative = Get-WinImgRelativePath -Root $SourceRoot -Path $File.FullName
  $extension = $File.Extension.ToLowerInvariant()
  $kind = if ($State.ImageExtensions -contains $extension) { 'Image' }
    elseif ($State.VideoExtensions -contains $extension) { 'Video' } else { 'Ignored' }
  $row = [pscustomobject]@{
    SourceRelativePath = $relative; PlannedOutputRelativePath = ''; OutputRelativePath = ''
    Kind = $kind; Status = $(if ($kind -eq 'Ignored') { 'Ignored' } else { 'NotStarted' })
    Reason = $(if ($kind -eq 'Ignored') { 'Unsupported file extension; intentionally not processed' } else { 'Not started' })
    InputBytes = [long]$File.Length; OutputBytes = $null; ScalePercent = $null; Width = $null; Height = $null; Attempts = 0
    SizeWarning = $false; NativeWarning = $false; TimestampWarning = $false; FramesOmitted = 0
    AncillaryWarning = $false; NamingReason = ''; RetainedSourceRelativePath = ''; RetainedOutputRelativePath = ''; RetainedStatus = ''; Started = $false
  }
  $State.BySource.Add($relative, $row); $State.Rows.Add($row)
}

function Set-WinImgReportWarning {
  param([object]$Row, [ValidateSet('SizeWarning','NativeWarning','TimestampWarning','AncillaryWarning')][string]$Flag, [string]$Reason)
  if (-not $Row) { return }
  $Row.$Flag = $true
  if ($Reason) { $Row.Reason = Get-WinImgBoundedText ((@($Row.Reason, $Reason) | Where-Object { $_ }) -join '; ') 2048 }
}

function Get-WinImgReportSummary {
  param([object]$State)
  $counts = [ordered]@{ Discovered=$State.Rows.Count; Converted=0; CopiedVideo=0; SkippedDuplicate=0; Ignored=0; Error=0; Cancelled=0; NotStarted=0 }
  [decimal]$inputBytes = 0; [decimal]$outputBytes = 0; [decimal]$videoBytes = 0
  $warnings = [ordered]@{ SizeWarnings=0; NativeWarnings=0; TimestampWarnings=0; AncillaryWarnings=0; FramesOmitted=0 }
  foreach ($row in $State.Rows) {
    if ($row.Status -notin @('Converted','CopiedVideo','SkippedDuplicate','Ignored','Error','Cancelled','NotStarted')) { throw 'Unknown file outcome; cannot report a reconciled partition.' }
    $counts[$row.Status]++
    if ($row.Status -eq 'Converted') { $inputBytes += [decimal]$row.InputBytes; $outputBytes += [decimal]$row.OutputBytes }
    elseif ($row.Status -eq 'CopiedVideo') { $videoBytes += [decimal]$row.OutputBytes }
    foreach ($pair in @(@('SizeWarning','SizeWarnings'),@('NativeWarning','NativeWarnings'),@('TimestampWarning','TimestampWarnings'),@('AncillaryWarning','AncillaryWarnings'))) {
      if ($row.($pair[0])) { $warnings[$pair[1]]++ }
    }
    $warnings.FramesOmitted += $row.FramesOmitted
  }
  $partition = $counts.Converted + $counts.CopiedVideo + $counts.SkippedDuplicate + $counts.Ignored + $counts.Error + $counts.Cancelled + $counts.NotStarted
  if ($partition -ne $counts.Discovered) { throw 'Known discovered-file outcomes do not reconcile.' }
  [decimal]$savings = $inputBytes - $outputBytes
  $seconds = $State.Clock.Elapsed.TotalSeconds
  return [pscustomobject]@{
    Counts=$counts; WarningCounts=$warnings; PartitionBalanced=$true
    ImageInputBytes=$inputBytes; ImageOutputBytes=$outputBytes; ImageSavingsBytes=$savings
    ImageSavingsPercent=$(if ($inputBytes -gt 0) { 100 * ($savings / $inputBytes) } else { $null })
    VideoBytes=$videoBytes; ElapsedSeconds=$seconds
    FilesPerSecond=$(if ($seconds -gt 0) { ($counts.Converted + $counts.CopiedVideo) / $seconds } else { $null })
  }
}

function Sync-WinImgReportStats {
  param([object]$State, [object]$Stats)
  $summary = Get-WinImgReportSummary $State
  $Stats.Converted=$summary.Counts.Converted; $Stats.CopiedVideo=$summary.Counts.CopiedVideo
  $Stats.SkippedDuplicate=$summary.Counts.SkippedDuplicate; $Stats.Unsupported=$summary.Counts.Ignored; $Stats.Errors=$summary.Counts.Error
  $Stats.SizeWarnings=$summary.WarningCounts.SizeWarnings; $Stats.NativeWarnings=$summary.WarningCounts.NativeWarnings; $Stats.TimestampWarnings=$summary.WarningCounts.TimestampWarnings
}

function ConvertTo-WinImgReportText {
  param([AllowEmptyString()][string]$Text)
  # A visible prefix protects every text value, including whitespace/full-width
  # formula openers. Escaping is reversible, including literal \u and surrogates.
  $result = [Text.StringBuilder]::new('text:')
  foreach ($character in $Text.ToCharArray()) {
    $category = [char]::GetUnicodeCategory($character)
    if ($character -eq [char]92) { $null = $result.Append('\\') }
    elseif ($category -in @([Globalization.UnicodeCategory]::Control,[Globalization.UnicodeCategory]::Format,[Globalization.UnicodeCategory]::LineSeparator,[Globalization.UnicodeCategory]::ParagraphSeparator,[Globalization.UnicodeCategory]::Surrogate)) {
      $null = $result.Append(('\u{0:x4}' -f [int]$character))
    } else { $null = $result.Append($character) }
  }
  return $result.ToString()
}

function ConvertTo-WinImgCsvField {
  param([object]$Value, [switch]$Number)
  if ($Number) {
    $text = if ($null -eq $Value) { '' } else { [Convert]::ToString($Value, [Globalization.CultureInfo]::InvariantCulture) }
    if ($text -and $text -notmatch '^-?[0-9]+(?:\.[0-9]+)?$') { throw 'Report numeric fields must contain generated invariant numbers.' }
  } else { $text = ConvertTo-WinImgReportText ([string]$Value) }
  return '"' + $text.Replace('"','""') + '"'
}

function New-WinImgRunReportFile {
  param([string]$Path, [string]$Header)
  Assert-WinImgNoReparseAncestors ([IO.Path]::GetDirectoryName($Path))
  $stream = [IO.FileStream]::new($Path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
  $writer = $null
  try {
    $writer = [IO.StreamWriter]::new($stream, [Text.UTF8Encoding]::new($true, $true))
    $writer.NewLine = "`r`n"; $writer.WriteLine($Header)
    return $writer
  } catch { if ($writer) { $writer.Dispose() } else { $stream.Dispose() }; throw }
}

function Add-WinImgRunReportLine { param([IO.TextWriter]$Writer, [string]$Line) $Writer.WriteLine($Line) }
function Complete-WinImgRunReportFile { param([IO.TextWriter]$Writer) $Writer.Dispose() }

function Write-WinImgOutcomeCsv {
  param([object]$State)
  $columns = @('RecordType','SourceRelativePath','PlannedOutputRelativePath','OutputRelativePath','Kind','Status','Reason','InputBytes','OutputBytes','ImageSavingsBytes','ScalePercent','Width','Height','Attempts','SizeWarning','NativeWarning','TimestampWarning','FramesOmitted','AncillaryWarning','NamingReason','RetainedSourceRelativePath','RetainedOutputRelativePath','RetainedStatus','Started')
  $numbers = @('InputBytes','OutputBytes','ImageSavingsBytes','ScalePercent','Width','Height','Attempts','SizeWarning','NativeWarning','TimestampWarning','FramesOmitted','AncillaryWarning','Started')
  $writer = $null; $phase = 'Creation'
  try {
    $writer = New-WinImgRunReportFile -Path $State.CsvPath -Header ($columns -join ',')
    $phase = 'Write'
    $paths = New-Object 'Collections.Generic.List[string]'
    foreach ($key in $State.BySource.Keys) { $paths.Add($key) }
    $paths.Sort([StringComparer]::Ordinal)
    foreach ($key in $paths) {
      $row = $State.BySource[$key]; $cells = New-Object 'Collections.Generic.List[string]'
      foreach ($column in $columns) {
        $value = if ($column -eq 'RecordType') { 'File' }
          elseif ($column -eq 'ImageSavingsBytes') { if ($row.Status -eq 'Converted') { [decimal]$row.InputBytes - [decimal]$row.OutputBytes } else { $null } }
          else { $row.$column }
        if ($value -is [bool]) { $value = [int]$value }
        $cells.Add((ConvertTo-WinImgCsvField -Value $value -Number:($numbers -contains $column)))
      }
      Add-WinImgRunReportLine -Writer $writer -Line ($cells -join ',')
    }
    foreach ($issue in $State.ScanIssues) {
      $relative = if ($issue.Path.Equals($State.SourceRoot, [StringComparison]::OrdinalIgnoreCase)) { '.' } else { Get-WinImgRelativePath -Root $State.SourceRoot -Path $issue.Path }
      $values = @{ RecordType='ScanIssue'; SourceRelativePath=$relative; Kind=$issue.Kind; Status='Warning'; Reason=$issue.Reason }
      $cells = New-Object 'Collections.Generic.List[string]'
      foreach ($column in $columns) { $cells.Add((ConvertTo-WinImgCsvField -Value $values[$column] -Number:($numbers -contains $column))) }
      Add-WinImgRunReportLine -Writer $writer -Line ($cells -join ',')
    }
    $phase = 'Close'
    Complete-WinImgRunReportFile -Writer $writer; $writer = $null
    $State.ReportComplete = $true
  } catch [Management.Automation.PipelineStoppedException] { throw }
  catch {
    if (Test-WinImgCancellationException $_) { throw }
    $State.ReportWarnings = 1; $State.ReportFailurePhase = $phase
    $State.ReportFailureReason = Get-WinImgBoundedText (ConvertTo-WinImgLogText $_.Exception.Message) 512
  } finally {
    if ($writer) {
      try { $writer.Dispose() }
      catch { if ($State.ReportWarnings -eq 0) { $State.ReportWarnings=1; $State.ReportFailurePhase='Close'; $State.ReportFailureReason=Get-WinImgBoundedText (ConvertTo-WinImgLogText $_.Exception.Message) 512 } }
    }
  }
}

function Write-WinImgReportProgress {
  param([object]$State, [string]$Status, [int]$PercentComplete, [switch]$Completed)
  if (-not $State.ProgressEnabled) { return }
  try {
    if ($Completed) { Write-Progress -Activity 'WinImgNormalizer' -Completed }
    else { Write-Progress -Activity 'WinImgNormalizer' -Status $Status -PercentComplete $PercentComplete }
  }
  catch [Management.Automation.PipelineStoppedException] { throw }
  catch {
    if (Test-WinImgCancellationException $_) { throw }
    $State.ProgressEnabled=$false; $State.ProgressWarnings=1
    Write-WinImgEmergencyReport 'Progress warning: progress display failed; processing continues and the run is degraded.'
  }
}

function Complete-WinImgReport {
  param([object]$State, [object]$Stats, [object]$LogState, [scriptblock]$Observer, [switch]$Interrupted, [switch]$Incomplete, [bool]$ProcessingWarnings = $false)
  if ($ProcessingWarnings) { $State.ProcessingWarnings=$true }
  if ($Interrupted) {
    foreach ($row in $State.Rows) {
      if ($row.Status -eq 'NotStarted' -and $row.Started) { $row.Status='Cancelled'; $row.Reason='Cooperative cancellation before finalization' }
    }
  } elseif ($Incomplete) {
    foreach ($row in $State.Rows) {
      if ($row.Status -eq 'NotStarted' -and $row.Started) { $row.Status='Error'; $row.Reason='Run ended before an outcome was committed' }
    }
  }
  Sync-WinImgReportStats -State $State -Stats $Stats
  if (-not $State.ReportAttempted) {
    $State.ReportAttempted = $true
    # Do not poll the already-requested cancellation token while recording the
    # known terminal outcomes. A closed/force-killed host may bypass this attempt.
    if ($State.CsvPath) { Write-WinImgOutcomeCsv $State }
    if ($State.ReportWarnings) {
      $notice = 'CSV report warning: {0} failed ({1}); report is incomplete. Completed media are retained; the run is degraded.' -f $State.ReportFailurePhase,$State.ReportFailureReason
      Write-WinImgEmergencyReport $notice
      if ($LogState) { Write-WinImgRunLog -State $LogState -Message $notice -Level WARN -Quiet }
    }
  }
  $State.Summary = Get-WinImgReportSummary $State
  $c = $State.Summary.Counts; $s = $State.Summary; $w = $s.WarningCounts
  $degraded = $Incomplete -or $State.ProcessingWarnings -or -not $State.ScanComplete -or $State.SkippedLinks -or $c.Error -or $c.NotStarted -or $State.ReportWarnings -or $State.ProgressWarnings -or $State.ConsoleWarnings -or $w.SizeWarnings -or $w.NativeWarnings -or $w.TimestampWarnings -or $w.AncillaryWarnings -or ($LogState -and $LogState.Degraded)
  $State.RunState = if ($Interrupted) { 'Interrupted' } elseif ($degraded) { 'Partial' } else { 'Completed' }
  $line = 'ACCOUNTING Discovered={0} Converted={1} CopiedVideo={2} SkippedDuplicate={3} Ignored={4} Error={5} Cancelled={6} NotStarted={7} Balanced=True ScanComplete={8} IncompleteDirectories={9} UninspectableEntries={10} SkippedLinks={11} ImageInputBytes={12} ImageOutputBytes={13} ImageSavingsBytes={14} VideoBytes={15} ElapsedSeconds={16} FinalizedFilesPerSecond={17} ReportWarnings={18} ReportComplete={19} ProgressWarnings={20} ImageSavingsPercent={21} SizeWarnings={22} NativeWarnings={23} TimestampWarnings={24} AncillaryWarnings={25} FramesOmitted={26} ConsoleWarnings={27} ProcessingWarnings={28}' -f
    $c.Discovered,$c.Converted,$c.CopiedVideo,$c.SkippedDuplicate,$c.Ignored,$c.Error,$c.Cancelled,$c.NotStarted,$State.ScanComplete,$State.IncompleteDirectories,$State.UninspectableEntries,$State.SkippedLinks,
    $s.ImageInputBytes.ToString([Globalization.CultureInfo]::InvariantCulture),$s.ImageOutputBytes.ToString([Globalization.CultureInfo]::InvariantCulture),$s.ImageSavingsBytes.ToString([Globalization.CultureInfo]::InvariantCulture),$s.VideoBytes.ToString([Globalization.CultureInfo]::InvariantCulture),
    $s.ElapsedSeconds.ToString('F3',[Globalization.CultureInfo]::InvariantCulture),$(if ($null -ne $s.FilesPerSecond) { $s.FilesPerSecond.ToString('F3',[Globalization.CultureInfo]::InvariantCulture) } else { 'unknown' }),$State.ReportWarnings,$State.ReportComplete,$State.ProgressWarnings,
    $(if ($null -ne $s.ImageSavingsPercent) { $s.ImageSavingsPercent.ToString('F3',[Globalization.CultureInfo]::InvariantCulture) } else { 'unknown' }),$w.SizeWarnings,$w.NativeWarnings,$w.TimestampWarnings,$w.AncillaryWarnings,$w.FramesOmitted,$State.ConsoleWarnings,$State.ProcessingWarnings
  $line=Get-WinImgBoundedText $line 2048
  $csvHint=Get-WinImgBoundedText (ConvertTo-WinImgLogText ([string]$State.CsvPath)) 512
  if ($LogState) { Write-WinImgRunLog -State $LogState -Message $line -Quiet }
  # The aggregate line's own log write can fail; observe that degradation last.
  if ($LogState -and $LogState.Degraded -and -not $Interrupted) { $State.RunState='Partial' }
  try { Write-Host $line; Write-Host ('Final report outcome: State={0} ReportWarnings={1} ReportComplete={2} LogWarnings={3}; CSV={4}' -f $State.RunState,$State.ReportWarnings,$State.ReportComplete,$(if ($LogState) { $LogState.FailureCount } else { 0 }),$csvHint) }
  catch [Management.Automation.PipelineStoppedException] { throw }
  catch {
    if (Test-WinImgCancellationException $_) { throw }
    $State.ConsoleWarnings=1; $State.ConsoleFailureReason=Get-WinImgBoundedText (ConvertTo-WinImgLogText $_.Exception.Message) 512
    if (-not $Interrupted) { $State.RunState='Partial' }
    $notice='Console report warning: final presentation failed ({0}); ConsoleWarnings=1 State={1}. Completed media are retained.' -f $State.ConsoleFailureReason,$State.RunState
    Write-WinImgEmergencyReport $notice; Write-WinImgEmergencyReport (Get-WinImgBoundedText $line 2048)
    if ($LogState) { Write-WinImgRunLog -State $LogState -Message $notice -Level WARN -Quiet }
  }
  # Interrupted callers append their bounded INTERRUPTED marker before emitting.
  if ($LogState -and -not $Interrupted) { Complete-WinImgRunLog $LogState }
  if ($Observer -and -not $State.Observed) { $State.Observed=$true; $null = & $Observer $State }
}

function Invoke-WinImgNormalizer {
  param(
    [string]$Source, [object]$MaxBytes = 1MB, [string]$OutputParent,
    [string]$MagickPath, [scriptblock]$PreflightRunner, [scriptblock]$ProcessRunner,
    [object]$ExecutionPolicy, [scriptblock]$NativeProcessObserver, [string]$NativeTemporaryRoot,
    [object]$CancellationState, [bool]$CaptureConsoleControl = $false,
    [scriptblock]$CopyProgressObserver, [scriptblock]$RunStageObserver, [scriptblock]$ReportObserver
  )
  $ErrorActionPreference = 'Stop'
  $previousContext = $script:WinImgCancellationContext
  $context = [pscustomobject]@{
    Stats = [ordered]@{ Converted=0; CopiedVideo=0; SkippedDuplicate=0; Unsupported=0; Errors=0; SizeWarnings=0; NativeWarnings=0; TimestampWarnings=0 }
    Total=0; Started=0; Current=$null; ScanComplete=$false; Destination=$null; LogState=$null
    CopyProgressObserver=$CopyProgressObserver; RunStageObserver=$RunStageObserver
    ReportState = New-WinImgReportState
  }
  $session = $null
  try {
    Initialize-WinImgProcessApi
    if (-not $CancellationState) { $CancellationState = [WinImgNormalizer.CancellationState]::new() }
    $session = [WinImgNormalizer.CancellationSession]::Begin($CancellationState, $CaptureConsoleControl)
    $script:WinImgCancellationContext = $context
    if ($CaptureConsoleControl -and -not $session.Registered) {
      Write-Host ('Console Ctrl+C handler unavailable (Win32={0}); host termination may bypass cooperative cleanup.' -f $session.RegistrationErrorCode) -ForegroundColor Yellow
    }
    $CancellationContext = $context

  $logState = $null
  $stats = $CancellationContext.Stats
  $reportState = $CancellationContext.ReportState
  Assert-WinImgCancellation

  # Complete basic setup before creating Pictures, run folders, logs or mirrors.
  try {
    $MaxBytes = ConvertTo-WinImgByteCap $MaxBytes
    if (-not $ExecutionPolicy) { $ExecutionPolicy = New-WinImgExecutionPolicy }
    $ExecutionPolicy = New-WinImgExecutionPolicy -TimeoutMilliseconds $ExecutionPolicy.TimeoutMilliseconds -MemoryBytes $ExecutionPolicy.MemoryBytes -MapBytes $ExecutionPolicy.MapBytes -DiskBytes $ExecutionPolicy.DiskBytes -Threads $ExecutionPolicy.Threads
    $srcRoot = Resolve-WinImgSourcePath $Source
    $reportState.SourceRoot = $srcRoot
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
    $reportState.ImageExtensions = $imgExts; $reportState.VideoExtensions = $videoExts
    $sourceTree = Get-WinImgSourceTree -SourceRoot $srcRoot -ReportState $reportState
    $generatedName = Get-WinImgGeneratedName -TopLevelNames $sourceTree.TopLevelNames
    $outputPlan = @(Get-WinImgOutputPlan -SourceRoot $srcRoot -SourceTree $sourceTree -GeneratedName $generatedName -ImageExtensions $imgExts -VideoExtensions $videoExts)
    foreach ($planned in $outputPlan) {
      $record = $reportState.BySource[$planned.SourceRelativePath]
      $record.PlannedOutputRelativePath = $planned.OutputRelativePath; $record.NamingReason = $planned.NamingReason
    }
    $allFiles = @($outputPlan | ForEach-Object { $_.Source })
    $CancellationContext.Total = $allFiles.Count
    $CancellationContext.ScanComplete = $sourceTree.ScanComplete
    Assert-WinImgCancellation
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
    if ($readableImages.Count -gt 0) { $NativeTemporaryRoot = Resolve-WinImgNativeTemporaryRoot -Path $NativeTemporaryRoot -SourceRoot $srcRoot }
    Assert-WinImgCancellation
    $destinationInfo = Get-WinImgDestinationInfo -Path $pictures -Files $spaceFiles -MaxBytes $MaxBytes
    $pictures = $destinationInfo.Path
  } catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog -State $logState; throw }
  catch { if (Test-WinImgCancellationException $_) { throw };
    Write-Host ('Setup error: ' + $_.Exception.Message) -ForegroundColor Red
    if ($_.Exception.Data['WinImgNativeDetail']) {
      # No run log exists before setup succeeds. Keep its native reason visible
      # within a small console bound without creating output during preflight.
      Write-Host ('Native setup details: ' + (Get-WinImgBoundedText ([string]$_.Exception.Data['WinImgNativeDetail']) 512)) -ForegroundColor Red
    }
    Assert-WinImgCancellation
    return 1
  }

  Assert-WinImgCancellation

  # --------- Logger, literal-safe and explicitly degraded after sink failure ---
  $LogPath = $null
  $logState = $null
  function Write-Log {
    param([string]$Message, [ValidateSet('INFO','OK','SKIP','WARN','ERR')]$Level='INFO', [switch]$Quiet)
    Write-WinImgRunLog -State $logState -Message $Message -Level $Level -Quiet:$Quiet
  }

  # --------- Resolve Pictures + dest root ----
  $base  = [System.IO.Path]::GetFileName($srcRoot.TrimEnd('\','/'))
  if ([string]::IsNullOrWhiteSpace($base)) { $base = $srcRoot.TrimEnd('\','/').TrimEnd(':') }
  $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
  try {
    $pictures = Assert-WinImgSafeDestination -SourceRoot $srcRoot -OutputParent $pictures
    Assert-WinImgCancellation
    $destRoot = New-WinImgRunDirectory -OutputParent $pictures -BaseName $base -Stamp $stamp
    $CancellationContext.Destination = $destRoot
    $reportState.DestinationRoot = $destRoot
    $generatedRoot = [IO.Path]::Combine($destRoot, $generatedName)
    $workRoot = [IO.Path]::Combine($generatedRoot, 'work')
    $reportRoot = [IO.Path]::Combine($generatedRoot, 'reports')
    foreach ($path in @($generatedRoot, $workRoot, $reportRoot)) {
      if (-not (New-WinImgExclusiveDirectory $path)) { throw 'A generated namespace was unexpectedly occupied; the run was not adopted.' }
    }
  } catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog -State $logState; throw }
  catch { if (Test-WinImgCancellationException $_) { throw };
    Write-Host ('Setup error: could not create the destination directory. ' + $_.Exception.Message) -ForegroundColor Red
    Assert-WinImgCancellation
    return 1
  }

  # A failed required log is visible but does not discard valid media work.
  $reportState.CsvPath = [IO.Path]::Combine($reportRoot, "WinImgNormalizer_${stamp}.csv")
  $LogPath = [System.IO.Path]::Combine($reportRoot, "WinImgNormalizer_${stamp}.log")
  $logState = New-WinImgLogState -Path $LogPath
  $CancellationContext.LogState = $logState
  Initialize-WinImgRunLog -State $logState
  Write-Log "Source: $srcRoot"
  Write-Log "Destination: $destRoot"
  Write-Log "Generated work/report directory: $generatedName"
  foreach ($row in $outputPlan) {
    Assert-WinImgCancellation
    Write-Log ('PLAN {0}: {1} -> {2} ({3})' -f $row.Kind, $row.SourceRelativePath, $row.OutputRelativePath, $row.NamingReason)
  }
  Write-Log ("MaxBytes: {0} bytes ({1} MiB; best-effort target)" -f $MaxBytes.ToString([Globalization.CultureInfo]::InvariantCulture), [Math]::Round($MaxBytes/1MB,2))

  Write-Log ('Per-image limits: RuntimeMs={0}; MemoryBytes={1}; MapBytes={2}; DiskBytes={3}; Threads={4}; stricter installed policies remain effective; timeout/resource failure ends this item and continues siblings.' -f $ExecutionPolicy.TimeoutMilliseconds,$ExecutionPolicy.MemoryBytes,$ExecutionPolicy.MapBytes,$ExecutionPolicy.DiskBytes,$ExecutionPolicy.Threads)
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
  Write-Log ('SCAN ScanComplete={0} IncompleteDirectories={1} UninspectableEntries={2} SkippedLinks={3}; unknown files in unreadable locations are not counted.' -f
    $sourceTree.ScanComplete, $sourceTree.InaccessibleDirectoryCount, $sourceTree.UninspectableEntryCount, $sourceTree.SkippedReparsePointCount)
  foreach ($warning in $sourceTree.Warnings) { Write-Log ("Source scan: {0} ({1})" -f $warning.Path, $warning.Reason) 'WARN' }
  foreach ($directory in $sourceTree.Directories) {
    Assert-WinImgCancellation
    try {
      Assert-WinImgNoReparseAncestors $directory.FullName
      $rel = Get-WinImgRelativePath -Root $srcRoot -Path $directory.FullName
      $target = [System.IO.Path]::Combine($destRoot, $rel)
      Assert-WinImgNoReparseAncestors $target
      [IO.Directory]::CreateDirectory($target) | Out-Null
    } catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog -State $logState; throw }
    catch { if (Test-WinImgCancellationException $_) { throw };
      $hasTraversalWarnings = $true
      Write-Log ("Could not mirror directory: {0} ({1})" -f $directory.FullName, $_.Exception.Message) 'WARN'
    }
  }

  # --------- File sets ---
  $total = $allFiles.Count
  if ($total -eq 0) {
    Sync-WinImgReportStats -State $reportState -Stats $stats
    Write-Log "No images or videos found." 'WARN'
    Write-Log ('SUMMARY ConvertedImages=0 CopiedVideos=0 Duplicates=0 Unsupported={6} Errors=0 SizeWarnings=0 NativeWarnings=0 TimestampWarnings=0 ScanComplete={0} IncompleteDirectories={1} UninspectableEntries={2} SkippedLinks={3} LogWarnings={4} FallbackDropped={5}' -f
      $sourceTree.ScanComplete, $sourceTree.InaccessibleDirectoryCount, $sourceTree.UninspectableEntryCount, $sourceTree.SkippedReparsePointCount, $logState.FailureCount, $logState.FallbackDroppedLines, $stats.Unsupported)
    Write-Host ('Final reporting state: LogWarnings={0} DiskLogIncomplete={1} FallbackDropped={2}' -f $logState.FailureCount, $logState.Degraded, $logState.FallbackDroppedLines)
    Assert-WinImgCancellation
    Complete-WinImgReport -State $reportState -Stats $stats -LogState $logState -Observer $ReportObserver -ProcessingWarnings $hasTraversalWarnings
    if ($hasTraversalWarnings -or $logState.Degraded -or $reportState.RunState -eq 'Partial') { return 2 }
    return 0
  }

  # --------- Dedupe + stats ---
  $retained = New-Object 'Collections.Generic.Dictionary[string,object]' ([StringComparer]::Ordinal)
  $hasDuplicateWarnings = $false
  Write-Log 'Heuristic duplicate matching uses lowercase filename, LastWriteTimeUtc and equal input byte length after successful finalization; same-key/same-length content can still differ.'

  # --------- IM helpers ---
  function Get-ExtentString([long]$Bytes) {
    # ImageMagick KB/MB are decimal; PowerShell 1KB/1MB are binary.
    # Emit invariant decimal bytes. Its encoder search uses floating point;
    # the validated file's Int64 length below remains the compliance authority.
    return $Bytes.ToString([Globalization.CultureInfo]::InvariantCulture) + 'B'
  }
  function Convert-ImageMagick {
    param([string]$SourcePath, [string]$DestPath, [long]$MaxBytes,
      [string]$WorkRoot, [System.Collections.Generic.List[object]]$OwnedCandidates, [object]$SourceInfo,
      [object]$ColourInfo, [string]$SrgbProfilePath, [object]$NativeContext, [object]$ReportRow)

    $extent = Get-ExtentString $MaxBytes
    $scales = 100,90,80,70,60,50
    $attemptCount = 0; $transientRetries = 0
    foreach ($p in $scales) {
      while ($true) {
        Assert-WinImgCancellation
        # Every size attempt or diagnosed transient retry gets a new candidate.
        if ($attemptCount -gt 0) {
          $previous = $OwnedCandidates[$OwnedCandidates.Count - 1]
          try { Remove-WinImgOwnedCandidate -CandidatePath $previous.Path -FileOwned $previous.FileOwned; $previous.Removed = $true }
          catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog -State $logState; throw }
          catch { if (Test-WinImgCancellationException $_) { throw }; Write-Log ('Could not remove superseded image scratch: ' + $_.Exception.Message) 'WARN' }
          $DestPath = New-WinImgImageCandidate -WorkRoot $WorkRoot
          $OwnedCandidates.Add([pscustomobject]@{ Path = $DestPath; FileOwned = $false; Removed = $false })
        }
        $ownership = $OwnedCandidates[$OwnedCandidates.Count - 1]
        $stream = [IO.FileStream]::new($DestPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        $ownership.FileOwned = $true; $stream.Dispose(); $attemptCount++
        $ReportRow.Attempts = $attemptCount
        $nativeDestPath = Get-WinImgNativeOutputPath $DestPath
        $nativeArguments = @('-quiet', '-regard-warnings', '-define', 'registry:filename:literal=true')
        if ($SourceInfo) { $nativeArguments += @('-define', 'image:frames=0') }
        $nativeArguments += Get-WinImgNativeOutputPath $SourcePath
        if ($SourceInfo -and $SourceInfo.Policy -eq 'FirstDisplayedFrame') { $nativeArguments += '-coalesce' }
        if ($SourceInfo) { $nativeArguments += '+repage' }
        $nativeArguments += '-auto-orient'
        if ($ColourInfo.HasIcc) {
          $nativeArguments += @('+black-point-compensation', '-intent', 'Relative', '-profile', (Get-WinImgNativeOutputPath $SrgbProfilePath))
        } else { $nativeArguments += @('-colorspace', 'sRGB') }
        # White composition and all JPEG/colour flags are identical on retries.
        $nativeArguments += @('-background','white','-alpha','remove','-alpha','off',
          '-strip','-sampling-factor','4:2:0','-interlace','Line')
        $nativeArguments += @('-resize', "$p%", '-define', "jpeg:extent=$extent", ('JPEG:' + $nativeDestPath))
        try { $nativeResult = Invoke-WinImgImageProcess -Executable $MagickCmd -Arguments $nativeArguments -Context $NativeContext -ProcessRunner $ProcessRunner }
        catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog -State $logState; throw }
        catch { if (Test-WinImgCancellationException $_) { throw };
          $nativeException = $_.Exception
          while ($nativeException.InnerException) { $nativeException = $nativeException.InnerException }
          $win32 = if ($nativeException -is [ComponentModel.Win32Exception]) { $nativeException.NativeErrorCode }
            elseif ($nativeException -is [IO.IOException] -and ($nativeException.HResult -band 0xFFFF0000L) -eq 0x80070000L) { $nativeException.HResult -band 0xFFFF }
            else { 0 }
          $nativeResult = [pscustomobject]@{ ExitCode=$null; StartError=Get-WinImgBoundedText $nativeException.Message 512; Win32ErrorCode=$win32; TimedOut=($nativeException -is [TimeoutException]) }
        }
        $outcome = Get-WinImgNativeOutcome -Result $nativeResult
        $exit = if ($null -eq $outcome.ExitCode) { 'none' } else { [string]$outcome.ExitCode }
        Write-Log ('NATIVE IMG: Attempt={0}; Scale={1}%; Category={2}; Exit={3}; TransientRetries={4}/2' -f $attemptCount,$p,$outcome.Category,$exit,$transientRetries) -Quiet
        if (-not $outcome.Acceptable -or $outcome.Warning) { Write-Log ('NATIVE DETAILS: ' + $outcome.DiagnosticText) -Quiet }
        if (-not $outcome.Acceptable) {
          if ($outcome.Retryable -and $transientRetries -lt 2) {
            $transientRetries++; $delay = 100 * $transientRetries
            Write-Log ('RETRY IMG: Category=TransientIO; Attempt={0}; Scale={1}%; Retry={2}/2; DelayMs={3}; Reason=sharing or lock violation; unchanged colour/alpha policy' -f $attemptCount,$p,$transientRetries,$delay) 'WARN'
            $state = [WinImgNormalizer.CancellationSession]::Current
            if ($state) { $state.Wait($delay); $state.ThrowIfRequested() }
            else { Start-Sleep -Milliseconds $delay }
            continue
          }
          return @{ Status='Error'; Attempts=$attemptCount; Note=('Native conversion stopped: Category={0}; Exit={1}; TransientRetries={2}/2; see NATIVE DETAILS in log' -f $outcome.Category,$exit,$transientRetries) }
        }
        try { $validation = Test-WinImgImageCandidate -CandidatePath $DestPath -MagickPath $MagickCmd -NativeContext $NativeContext }
        catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog -State $logState; throw }
        catch { if (Test-WinImgCancellationException $_) { throw };
          if ($_.Exception.Data['WinImgNativeDetail']) { Write-Log ('NATIVE DETAILS: ' + $_.Exception.Data['WinImgNativeDetail']) -Quiet }
          return @{ Status='Error'; Attempts=$attemptCount; Note=('Candidate validation failed; no retry or earlier candidate retained: ' + $_.Exception.Message) }
        }
        if ($validation.BytesOut -le $MaxBytes -or $p -eq 50) {
          $aboveTarget = $validation.BytesOut -gt $MaxBytes
          $notes = @()
          if ($aboveTarget) { $notes += 'Could not reach target; best-effort saved' }
          if ($outcome.Warning) { $notes += 'Accepted native lossless-to-lossy JPEG notice; see NATIVE DETAILS in log' }
          return @{ Status=$(if ($aboveTarget -or $outcome.Warning) { 'ConvertedWithWarning' } else { 'Converted' });
            CandidatePath=$DestPath; BytesOut=$validation.BytesOut; Width=$validation.Width; Height=$validation.Height;
            Scale=$p; Attempts=$attemptCount; SizeWarning=$aboveTarget; NativeWarning=$outcome.Warning; Note=($notes -join '; ') }
        }
        # A valid above-cap result advances the size sequence. Native failures
        # never advance it or recover a previously superseded image.
        break
      }
    }
    return @{ Status='Error'; Attempts=$attemptCount; Note='No valid final candidate was retained.' }
  }

  # --------- Main loop ---
  [int]$i = 0
  foreach ($row in $outputPlan) {
    $outcomeRow = $reportState.BySource[$row.SourceRelativePath]
    Invoke-WinImgRunStage -Stage 'BeforeItem' -RelativePath $row.SourceRelativePath -Path ([IO.Path]::Combine($destRoot, $row.OutputRelativePath))
    $f = $row.Source
    $i++
    $CancellationContext.Started = $i
    $CancellationContext.Current = $row.SourceRelativePath
    $outcomeRow.Started = $true; $outcomeRow.Reason = 'Processing'
    $rel = $row.SourceRelativePath
    $ext = $f.Extension.ToLowerInvariant()
    Write-WinImgReportProgress -State $reportState -Status "$i / $total : $rel" -PercentComplete ([int]($i*100/$total))

    $coder = $imageCoders[$ext]
    if ($coder -and $missingCoders.ContainsKey($coder)) {
      $outcomeRow.Status = 'Error'; $outcomeRow.Reason = "Missing $coder decoder; no conversion attempted"
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
      $outcomeRow.InputBytes = [long]$sourceLength
      $sourceModified = $f.LastWriteTimeUtc
      $sourceCreated = $f.CreationTimeUtc
      $dupKey = Get-WinImgDuplicateKey -Name $f.Name -LastWriteTimeUtc $sourceModified -Length $sourceLength
      if ($retained.ContainsKey($dupKey)) {
        $first = $retained[$dupKey]
        $outcomeRow.Status='SkippedDuplicate'; $outcomeRow.Reason='Heuristic filename/time/length match; content equality is not verified'
        $outcomeRow.RetainedSourceRelativePath=$first.SourceRelativePath; $outcomeRow.RetainedOutputRelativePath=$first.OutputRelativePath; $outcomeRow.RetainedStatus=$first.Status
        $stats.SkippedDuplicate++
        Write-Log ('Heuristic duplicate skipped: {0} (retained source: {1}; retained output: {2}; retained status: {3})' -f
          $rel, $first.SourceRelativePath, $first.OutputRelativePath, $first.Status) 'SKIP'
        continue
      }
      Assert-WinImgOutputAvailable $destPath
    } catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog -State $logState; throw }
    catch { if (Test-WinImgCancellationException $_) { throw };
      $hasTraversalWarnings = $true
      $outcomeRow.Status='Error'; $outcomeRow.Reason=Get-WinImgBoundedText $_.Exception.Message 2048
      $stats.Errors++
      Write-Log "Source/destination safety check failed: $rel ($($_.Exception.Message))" 'ERR'
      continue
    }
    $ownedCandidates = New-Object 'System.Collections.Generic.List[object]'
    $nativeContext = $null
    try {
      if ($imgExts -contains $ext) {
        # Inspection and every attempt use the same profile-bearing bytes.
        $nativeContext = New-WinImgImageContext -Policy $ExecutionPolicy -WorkRoot $workRoot -NativeProcessObserver $NativeProcessObserver -TemporaryRoot $NativeTemporaryRoot
        $imageSource = New-WinImgSourceSnapshot -SourcePath $f.FullName -WorkRoot $workRoot -ExpectedLength $sourceLength -ExpectedModified $sourceModified -OwnedCandidates $ownedCandidates
        $sourceInfo = $null
        if ($ext -in @('.gif','.tif','.tiff','.webp','.heic','.heif')) {
          $sourceInfo = Get-WinImgSourceImageInfo -SnapshotPath $imageSource -MagickPath $MagickCmd -NativeContext $nativeContext
          $outcomeRow.FramesOmitted = $sourceInfo.Omitted
          Write-Log ('SOURCE IMG: {0} (SourceCount={1}; Selected=1; Omitted={2}; Unit={3}; Policy={4}; Decoder={5})' -f
            $rel, $sourceInfo.SourceCount, $sourceInfo.Omitted, $sourceInfo.Unit, $sourceInfo.Policy, $sourceInfo.Decoder)
        }
        $colourInfo = Get-WinImgColourInfo -SnapshotPath $imageSource -MagickPath $MagickCmd -NativeContext $nativeContext
        $srgbProfilePath = $null
        if ($colourInfo.HasIcc) {
          $null = New-WinImgSourceIccProfile -SnapshotPath $imageSource -MagickPath $MagickCmd -ColourSpace $colourInfo.ColourSpace -WorkRoot $workRoot -OwnedCandidates $ownedCandidates -NativeContext $nativeContext
          $srgbProfilePath = New-WinImgSrgbProfile -WorkRoot $workRoot -OwnedCandidates $ownedCandidates
        }
        Write-Log ('COLOUR IMG: {0} (SourceSpace={1}; SourceICC={2}; Policy={3}; Intent={4}; Alpha=WhiteAfterSrgb; OutputICC=None)' -f
          $rel, $colourInfo.ColourSpace, $colourInfo.HasIcc, $colourInfo.Policy, $(if ($colourInfo.HasIcc) { 'Relative' } else { 'None' }))
        try { $candidatePath = New-WinImgImageCandidate -WorkRoot $workRoot }
        catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog -State $logState; throw }
        catch { if (Test-WinImgCancellationException $_) { throw }; $hasNamingWarnings = $true; throw }
        $ownedCandidates.Add([pscustomobject]@{ Path = $candidatePath; FileOwned = $false; Removed = $false })
        $res = Convert-ImageMagick -SourcePath $imageSource -DestPath $candidatePath -MaxBytes $MaxBytes -WorkRoot $workRoot -OwnedCandidates $ownedCandidates -SourceInfo $sourceInfo -ColourInfo $colourInfo -SrgbProfilePath $srgbProfilePath -NativeContext $nativeContext -ReportRow $outcomeRow
        if ($res.Status -in @('Converted', 'ConvertedWithWarning')) {
          $null = Get-WinImgRemainingTime $nativeContext
          Invoke-WinImgRunStage -Stage 'CandidateValidated' -RelativePath $rel -Path $res.CandidatePath
          try { Move-WinImgPlannedImage -CandidatePath $res.CandidatePath -DestinationPath $destPath }
          catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog -State $logState; throw }
          catch { if (Test-WinImgCancellationException $_) { throw }; $hasNamingWarnings = $true; throw }
          # Commit the single retained outcome before ancillary timestamp/log
          # operations. Cancellation can no longer turn this final file into Error.
          $outcomeRow.InputBytes=[long]$sourceLength; $outcomeRow.OutputBytes=[long]$res.BytesOut
          $outcomeRow.OutputRelativePath=$destRel; $outcomeRow.Width=$res.Width; $outcomeRow.Height=$res.Height; $outcomeRow.ScalePercent=$res.Scale
          $outcomeRow.SizeWarning=[bool]$res.SizeWarning; $outcomeRow.NativeWarning=[bool]$res.NativeWarning
          $outcomeRow.Reason=$(if ($res.Note) { $res.Note } else { 'Validated JPEG finalized' }); $outcomeRow.Status='Converted'
          $stats.Converted++
          if ($res.SizeWarning) { $stats.SizeWarnings++ }
          if ($res.NativeWarning) { $stats.NativeWarnings++ }
          $timestampResult = Set-WinImgOutputTimestamps -Path $destPath -LastWriteTimeUtc $sourceModified -CreationTimeUtc $sourceCreated
          if (-not $timestampResult.Succeeded) {
            Set-WinImgReportWarning -Row $outcomeRow -Flag TimestampWarning -Reason 'Timestamp restoration incomplete; valid output retained'
            $stats.TimestampWarnings++
            $res.Status = 'ConvertedWithWarning'
            $res.Note = (@($res.Note, 'Timestamp restoration incomplete; valid output retained') | Where-Object { $_ }) -join '; '
            foreach ($failure in $timestampResult.FailedFields) { Write-Log ('Timestamp warning IMG: {0} ({1}: {2})' -f $destRel, $failure.Name, $failure.Reason) 'WARN' }
          }
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
                Status = $res.Status
              })
            }
          } catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog -State $logState; throw }
          catch { if (Test-WinImgCancellationException $_) { throw };
            $hasDuplicateWarnings = $true
            Set-WinImgReportWarning -Row $outcomeRow -Flag AncillaryWarning -Reason 'Source metadata changed after finalization; heuristic registration omitted'
            Write-Log ('Finalized image kept without heuristic registration: {0} ({1})' -f $rel, $_.Exception.Message) 'WARN'
          }
          $note = if ($res.Note) { " ($($res.Note))" } else { "" }
          $level = if ($res.Status -eq 'ConvertedWithWarning') { 'WARN' } else { 'OK' }
          Write-Log ("{0} IMG: {1} -> {2} [{3} bytes, MaxBytes={4}, Width={5}, Height={6}, Scale={7}%]{8}" -f
            $level, $rel, $destRel, $res.BytesOut.ToString([Globalization.CultureInfo]::InvariantCulture),
            $MaxBytes.ToString([Globalization.CultureInfo]::InvariantCulture), $res.Width, $res.Height, $res.Scale, $note) $level
          Invoke-WinImgRunStage -Stage 'AfterFinalization' -RelativePath $rel -Path $destPath
        } else {
          $outcomeRow.Status='Error'; $outcomeRow.Reason=$res.Note; $outcomeRow.Attempts=$res.Attempts
          $stats.Errors++; Write-Log "ERR IMG: $rel ($($res.Note))" 'ERR'
        }
      }
      elseif ($videoExts -contains $ext) {
        $destDir = [System.IO.Path]::GetDirectoryName($destPath)
        if ($destDir) { [IO.Directory]::CreateDirectory($destDir) | Out-Null }
        try { $videoResult = Copy-WinImgPlannedVideo -SourcePath $f.FullName -DestinationPath $destPath -WorkRoot $workRoot -ReportRow $outcomeRow }
        catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog -State $logState; throw }
        catch { if (Test-WinImgCancellationException $_) { throw }; $hasNamingWarnings = $true; throw }
        # The copy helper's stable snapshot describes the bytes actually copied,
        # even if source metadata changed between lookup and staging.
        $outcomeRow.InputBytes=[long]$videoResult.BytesOut; $outcomeRow.OutputBytes=[long]$videoResult.BytesOut
        $outcomeRow.OutputRelativePath=$destRel; $outcomeRow.Attempts=1; $outcomeRow.Reason='Stable byte-for-byte video copy finalized'; $outcomeRow.Status='CopiedVideo'
        $stats.CopiedVideo++
        $videoKey = Get-WinImgDuplicateKey -Name $f.Name -LastWriteTimeUtc $videoResult.LastWriteTimeUtc -Length $videoResult.BytesOut
        $timestampResult = Set-WinImgOutputTimestamps -Path $destPath -LastWriteTimeUtc $videoResult.LastWriteTimeUtc -CreationTimeUtc $videoResult.CreationTimeUtc
        $videoStatus = 'CopiedVideo'
        if (-not $timestampResult.Succeeded) {
          Set-WinImgReportWarning -Row $outcomeRow -Flag TimestampWarning -Reason 'Timestamp restoration incomplete; valid output retained'
          $stats.TimestampWarnings++
          $videoStatus = 'CopiedVideoWithWarning'
          foreach ($failure in $timestampResult.FailedFields) { Write-Log ('Timestamp warning VID: {0} ({1}: {2})' -f $destRel, $failure.Name, $failure.Reason) 'WARN' }
        }
        if (-not $retained.ContainsKey($videoKey)) {
          $retained.Add($videoKey, [pscustomobject]@{ SourceRelativePath = $rel; OutputRelativePath = $destRel; Status = $videoStatus })
        }
        if ($videoResult.CleanupWarning) { Set-WinImgReportWarning -Row $outcomeRow -Flag AncillaryWarning -Reason $videoResult.CleanupWarning; $hasNamingWarnings = $true; Write-Log ('Video scratch cleanup: ' + $videoResult.CleanupWarning) 'WARN' }
        $bytesOut = $videoResult.BytesOut
        $videoLevel = if ($timestampResult.Succeeded) { 'OK' } else { 'WARN' }
        Write-Log ("{0} VID: {1} -> {2} [{3:n0} bytes]" -f $videoLevel, $rel, $destRel, $bytesOut) $videoLevel
        Invoke-WinImgRunStage -Stage 'AfterFinalization' -RelativePath $rel -Path $destPath
      }
      else {
        $outcomeRow.Status='Ignored'; $outcomeRow.Reason='Unsupported extension; intentionally not processed'
        $stats.Unsupported++; Write-Log "Unsupported skipped: $rel" 'WARN'
      }
    } catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog -State $logState; throw }
    catch { if (Test-WinImgCancellationException $_) { throw };
      if ($_.Exception.Data['WinImgNativeDetail']) { Write-Log ('NATIVE DETAILS: ' + $_.Exception.Data['WinImgNativeDetail']) -Quiet }
      if ($outcomeRow.Status -in @('Converted','CopiedVideo')) {
        Set-WinImgReportWarning -Row $outcomeRow -Flag AncillaryWarning -Reason ('Ancillary failure after finalization: ' + $_.Exception.Message)
        Write-Log "Ancillary warning after finalization: $rel ($($_.Exception.Message)); valid output retained" 'WARN'
      } else {
        $outcomeRow.Status='Error'; $outcomeRow.Reason=Get-WinImgBoundedText $_.Exception.Message 2048
        $stats.Errors++; Write-Log "Exception processing: $rel ($($_.Exception.Message))" 'ERR'
      }
    } finally {
      try { Remove-WinImgImageContext $nativeContext }
      catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog -State $logState; throw }
      catch { if (Test-WinImgCancellationException $_) { throw }; Set-WinImgReportWarning -Row $outcomeRow -Flag AncillaryWarning -Reason $_.Exception.Message; $hasNamingWarnings = $true; Write-Log ('Could not remove owned native cache: ' + $_.Exception.Message) 'WARN' }
      foreach ($owned in $ownedCandidates) {
        if ($nativeContext -and -not $nativeContext.CleanupSafe) { continue }
        if ($owned.Removed) { continue }
        try { Remove-WinImgOwnedCandidate -CandidatePath $owned.Path -FileOwned $owned.FileOwned }
        catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog -State $logState; throw }
        catch { if (Test-WinImgCancellationException $_) { throw };
          $hasNamingWarnings = $true
          Set-WinImgReportWarning -Row $outcomeRow -Flag AncillaryWarning -Reason $_.Exception.Message
          Write-Log ('Could not remove owned image scratch: ' + $_.Exception.Message) 'WARN'
        }
      }
    }
  }

  $CancellationContext.Current = $null
  Assert-WinImgCancellation
  Write-WinImgReportProgress -State $reportState -Completed
  Sync-WinImgReportStats -State $reportState -Stats $stats

  Write-Host ""
  Write-Host "Summary:" -ForegroundColor Cyan
  Write-Host ("  Converted images : {0}" -f $stats.Converted)
  Write-Host ("  Above byte target: {0} (included in converted images)" -f $stats.SizeWarnings)
  Write-Host ("  Native warnings: {0} (included in converted images)" -f $stats.NativeWarnings)
  Write-Host ("  Timestamp warnings: {0} (included in finalized images/videos)" -f $stats.TimestampWarnings)
  Write-Host ("  Scan complete: {0}; incomplete directories: {1}; uninspectable entries: {2}; skipped links: {3}" -f $sourceTree.ScanComplete, $sourceTree.InaccessibleDirectoryCount, $sourceTree.UninspectableEntryCount, $sourceTree.SkippedReparsePointCount)
  Write-Host ("  Copied videos    : {0}" -f $stats.CopiedVideo)
  Write-Host ("  Duplicates       : {0}" -f $stats.SkippedDuplicate)
  Write-Host ("  Ignored files    : {0}" -f $stats.Unsupported)
  Write-Host ("  Errors           : {0}" -f $stats.Errors)
  Write-Host ("  Log file         : {0}" -f $LogPath)

  Write-Log ("SUMMARY ConvertedImages={0} CopiedVideos={1} Duplicates={2} Unsupported={3} Errors={4} SizeWarnings={5} NativeWarnings={6} TimestampWarnings={7} ScanComplete={8} IncompleteDirectories={9} UninspectableEntries={10} SkippedLinks={11} LogWarnings={12} FallbackDropped={13}" -f
    $stats.Converted,$stats.CopiedVideo,$stats.SkippedDuplicate,$stats.Unsupported,$stats.Errors,$stats.SizeWarnings,$stats.NativeWarnings,$stats.TimestampWarnings,
    $sourceTree.ScanComplete,$sourceTree.InaccessibleDirectoryCount,$sourceTree.UninspectableEntryCount,$sourceTree.SkippedReparsePointCount,$logState.FailureCount,$logState.FallbackDroppedLines)
  Write-Log "Processing ended $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
  Complete-WinImgReport -State $reportState -Stats $stats -LogState $logState -Observer $ReportObserver -ProcessingWarnings ([bool]($hasTraversalWarnings -or $hasNamingWarnings -or $hasDuplicateWarnings))
  Write-Host ('Final reporting state: LogWarnings={0} DiskLogIncomplete={1} FallbackDropped={2}' -f $logState.FailureCount, $logState.Degraded, $logState.FallbackDroppedLines)
  Complete-WinImgRunLog -State $logState

  Assert-WinImgCancellation
  if ($stats.Errors -gt 0 -or $stats.SizeWarnings -gt 0 -or $stats.NativeWarnings -gt 0 -or $stats.TimestampWarnings -gt 0 -or $logState.Degraded -or $missingCoders.Count -gt 0 -or $hasTraversalWarnings -or $hasNamingWarnings -or $hasDuplicateWarnings -or $reportState.RunState -eq 'Partial') { return 2 }
  return 0
  } catch {
    if (-not (Test-WinImgCancellationException $_)) { throw }
    # Item finally blocks have finished their exact owned cleanup before this
    # boundary. A closed/force-killed host cannot guarantee this record or exit.
    $stats = $context.Stats
    Complete-WinImgReport -State $context.ReportState -Stats $stats -LogState $context.LogState -Observer $ReportObserver -Interrupted
    $current = if ($context.Current) { ConvertTo-WinImgLogText $context.Current } else { 'none' }
    $summary = 'INTERRUPTED Mode=Cooperative ExitCode=130 ConvertedImages={0} CopiedVideos={1} Started={2} Unstarted={3} Current={4} ScanComplete={5} CompletedOutputs=Retained Cleanup=BestEffortOwnedOnly' -f $stats.Converted,$stats.CopiedVideo,$context.Started,([Math]::Max(0,$context.Total-$context.Started)),$current,$context.ScanComplete
    $summary = Get-WinImgBoundedText $summary 2048
    Write-WinImgReportProgress -State $context.ReportState -Completed
    try { Write-Host $summary -ForegroundColor Yellow }
    catch [Management.Automation.PipelineStoppedException] { throw }
    catch { if (Test-WinImgCancellationException $_) { throw }; Write-WinImgEmergencyReport $summary }
    if ($context.LogState) { Write-WinImgRunLog -State $context.LogState -Message $summary -Level WARN -Quiet }
    Complete-WinImgRunLog -State $context.LogState
    return 130
  } finally {
    if ($context.ReportState.CsvPath -and -not $context.ReportState.ReportAttempted) {
      try { Complete-WinImgReport -State $context.ReportState -Stats $context.Stats -LogState $context.LogState -Observer $ReportObserver -Incomplete }
      catch { Write-WinImgEmergencyReport 'Final report could not complete after a stopped run; completed media are retained.' }
    }
    $script:WinImgCancellationContext = $previousContext
    if ($session) { $session.Dispose() }
  }
}

function Invoke-WinImgNormalizerCommand {
  param(
    [object[]]$Arguments,
    [string]$OutputParent,
    [string]$MagickPath,
    [scriptblock]$PreflightRunner,
    [scriptblock]$ProcessRunner,
    [object]$ExecutionPolicy, [scriptblock]$NativeProcessObserver, [string]$NativeTemporaryRoot,
    [object]$CancellationState, [bool]$CaptureConsoleControl = $true,
    [scriptblock]$CopyProgressObserver, [scriptblock]$RunStageObserver, [scriptblock]$ReportObserver
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
  $invoke.CaptureConsoleControl = $CaptureConsoleControl
  foreach ($key in @('ExecutionPolicy','NativeProcessObserver','NativeTemporaryRoot','CancellationState','CopyProgressObserver','RunStageObserver','ReportObserver')) {
    if ($PSBoundParameters.ContainsKey($key)) { $invoke[$key] = $PSBoundParameters[$key] }
  }
  try { return Invoke-WinImgNormalizer @invoke }
  catch [Management.Automation.PipelineStoppedException] { throw }
  catch {
    # Ordinary run-level failures have a stable command result. Cooperative
    # cancellation is handled by the active run; stopped/force-killed hosts may
    # bypass this boundary and have no promised report or application exit code.
    $message = 'Run error: ' + (Get-WinImgBoundedText (ConvertTo-WinImgLogText $_.Exception.Message) 1024)
    try { Write-Host $message -ForegroundColor Red }
    catch [Management.Automation.PipelineStoppedException] { throw }
    catch { Write-WinImgEmergencyReport $message }
    return 1
  }
}

if ($MyInvocation.InvocationName -ne '.') {
  exit (Invoke-WinImgNormalizerCommand -Arguments $args)
}
