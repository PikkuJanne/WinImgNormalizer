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

function Get-WinImgSourceTree {
  param([string]$SourceRoot)
  $files = New-Object 'System.Collections.Generic.List[System.IO.FileInfo]'
  $directories = New-Object 'System.Collections.Generic.List[System.IO.DirectoryInfo]'
  $names = New-Object 'System.Collections.Generic.List[string]'
  $warnings = New-Object 'System.Collections.Generic.List[object]'
  $pending = New-Object 'System.Collections.Generic.Stack[string]'
  $pending.Push($SourceRoot)
  $inaccessibleDirectories = 0
  $uninspectableEntries = 0
  $skippedLinks = 0
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
            $skippedLinks++
            $warnings.Add([pscustomobject]@{ Path = $entry.FullName; Kind = 'SkippedReparsePoint'; Reason = 'Reparse point skipped; links and junctions are not followed.' })
          } elseif ($entry -is [IO.DirectoryInfo]) {
            $directories.Add($entry)
            $pending.Push($entry.FullName)
          } elseif ($entry -is [IO.FileInfo]) { $files.Add($entry) }
        } catch [Management.Automation.PipelineStoppedException] { throw }
        catch {
          $uninspectableEntries++
          $warnings.Add([pscustomobject]@{ Path = $path; Kind = 'EntryInspectionFailed'; Reason = 'Could not inspect source entry: ' + $_.Exception.Message })
        }
      }
    } catch [Management.Automation.PipelineStoppedException] { throw }
    catch {
      $inaccessibleDirectories++
      $warnings.Add([pscustomobject]@{ Path = $directory; Kind = 'DirectoryEnumerationFailed'; Reason = 'Incomplete source scan: ' + $_.Exception.Message })
    }
  }
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

function Initialize-WinImgProcessApi {
  if ('WinImgNormalizer.NativeProcess' -as [type]) { return }
  # Fixed buffers on two readers avoid both pipe deadlocks and unbounded
  # ReadToEnd/ReadLine allocations (a native diagnostic need not contain LF).
  Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
namespace WinImgNormalizer {
  public sealed class NativeResult {
    public int? ExitCode;
    public string StdOut = "", StdErr = "", StartError = "";
    public long StdOutCharacters, StdErrCharacters;
    public bool StdOutTruncated, StdErrTruncated, TimedOut, Cancelled;
    public bool StreamsComplete = true;
    public int Win32ErrorCode, CaptureLimit;
  }
  internal sealed class BoundedCapture {
    private readonly char[] ring, head;
    private int next;
    private long total;
    private readonly object gate = new object();
    public string ReadError = "";
    public BoundedCapture(int limit) { ring = new char[limit]; head = new char[limit / 4]; }
    public void Drain(StreamReader reader) {
      char[] block = new char[4096];
      try {
        int count;
        while ((count = reader.Read(block, 0, block.Length)) != 0) {
          lock (gate) {
            for (int i = 0; i < count; i++) {
              if (total < head.Length) head[(int)total] = block[i];
              ring[next] = block[i]; next = (next + 1) % ring.Length; total++;
            }
          }
        }
      } catch (Exception error) { ReadError = error.GetType().Name; }
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
  public static class NativeProcess {
    // The Windows CRT grammar is shared by .NET Framework and modern .NET.
    // Always quote, doubling backslashes before quotes and the closing quote.
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
    public static NativeResult Run(string executable, string[] arguments, int limit, int timeout, bool killTree) {
      NativeResult result = new NativeResult();
      result.CaptureLimit = limit;
      BoundedCapture stdout = new BoundedCapture(limit), stderr = new BoundedCapture(limit);
      using (Process process = new Process()) {
        ProcessStartInfo info = new ProcessStartInfo();
        info.FileName = executable; info.UseShellExecute = false; info.CreateNoWindow = true;
        info.RedirectStandardOutput = true; info.RedirectStandardError = true;
        info.StandardOutputEncoding = new UTF8Encoding(false, false);
        info.StandardErrorEncoding = new UTF8Encoding(false, false);
        StringBuilder command = new StringBuilder();
        foreach (string argument in arguments) { if (command.Length != 0) command.Append(' '); command.Append(Quote(argument)); }
        info.Arguments = command.ToString(); process.StartInfo = info;
        try { if (!process.Start()) throw new InvalidOperationException("Native process did not start."); }
        catch (Exception error) {
          string message = error.Message; result.StartError = message.Length > 512 ? message.Substring(0, 512) : message;
          Win32Exception native = error as Win32Exception;
          if (native != null) result.Win32ErrorCode = native.NativeErrorCode;
          return result;
        }
        StreamReader outStream = process.StandardOutput, errStream = process.StandardError;
        Thread outReader = new Thread(delegate() { stdout.Drain(outStream); });
        Thread errReader = new Thread(delegate() { stderr.Drain(errStream); });
        outReader.IsBackground = true; errReader.IsBackground = true; outReader.Start(); errReader.Start();
        try {
          if (timeout == 0) process.WaitForExit();
          else if (!process.WaitForExit(timeout)) {
            result.TimedOut = true;
            // Preserve the existing preflight-only deadline and owned PID tree
            // termination. Conversion deadlines/cancellation remain later work.
            if (killTree) {
              try {
                string taskkill = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "taskkill.exe");
                Run(taskkill, new string[] { "/PID", process.Id.ToString(System.Globalization.CultureInfo.InvariantCulture), "/T", "/F" }, 8192, 5000, false);
              } catch { }
            }
            if (!process.WaitForExit(5000)) { try { process.Kill(); } catch { } }
            process.WaitForExit(5000);
          }
          if (process.HasExited) result.ExitCode = process.ExitCode;
          bool outComplete = outReader.Join(5000), errComplete = errReader.Join(5000);
          result.StreamsComplete = outComplete && errComplete && stdout.ReadError.Length == 0 && stderr.ReadError.Length == 0;
        } finally {
          process.StandardOutput.Close(); process.StandardError.Close();
          outReader.Join(1000); errReader.Join(1000);
        }
        result.StdOut = stdout.Text(); result.StdErr = stderr.Text();
        result.StdOutCharacters = stdout.Count; result.StdErrCharacters = stderr.Count;
        result.StdOutTruncated = stdout.Truncated; result.StdErrTruncated = stderr.Truncated;
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
    [ValidateRange(0,2147483647)][int]$TimeoutMilliseconds = 0)
  Initialize-WinImgProcessApi
  return [WinImgNormalizer.NativeProcess]::Run($Executable, $Arguments, $OutputLimit, $TimeoutMilliseconds, $true)
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
    if ($Result.TimedOut) { $category = 'Timeout' }
    elseif ($Result.Cancelled) { $category = 'Cancelled' }
    elseif ($overflow) { $category = 'OutputLimit' }
    elseif ($Result.PSObject.Properties['StreamsComplete'] -and -not $Result.StreamsComplete) { $category = 'NativeFailure'; $incomplete = $true }
    elseif ($startError) { $category = if ($win32 -in @(32,33)) { 'TransientIO' } elseif ($win32 -eq 5) { 'AccessDenied' } else { 'StartFailure' } }
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
  if ($startError) { $detail += "`nStartError: " + $startError + '; Win32ErrorCode=' + $win32 }
  if ($stdout) { $detail += "`nStdOut: " + $stdout }
  if ($stderr) { $detail += "`nStdErr: " + $stderr }
  return [pscustomobject]@{ ExitCode=$code; Category=$category; Acceptable=$acceptable; Retryable=$retryable; Warning=$warning; DiagnosticText=$detail }
}

function Assert-WinImgNativeQuery {
  param([object]$Result, [string]$Context, [switch]$AllowStdOut)
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
  # Explicit +ping fully decodes pixels; warnings remain fatal for validation.
  $result = Invoke-WinImgNativeProcess -Executable $MagickPath -Arguments @('identify','+ping','-regard-warnings','-define','registry:filename:literal=true','-format','%m|%w|%h|%n',$nativePath)
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
  # Count all images before the selected first-image read. This is header
  # inspection; the final JPEG separately receives a full pixel decode.
  $result = Invoke-WinImgNativeProcess -Executable $MagickPath -Arguments @('identify','-ping','-regard-warnings','-define','registry:filename:literal=true','-format','%m|%n|%w|%h\n',$nativePath)
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
  param([string]$SnapshotPath, [string]$MagickPath)
  Assert-WinImgNoReparseAncestors ([IO.Path]::GetDirectoryName($SnapshotPath))
  $nativePath = Get-WinImgNativeOutputPath $SnapshotPath
  # profiles=none is an option fallback; leave the real ICC attached.
  $result = Invoke-WinImgNativeProcess -Executable $MagickPath -Arguments @('identify','-ping','-regard-warnings','-define','registry:filename:literal=true','-define','image:frames=0','-define','profiles=none','+set','profiles','+set','colorspace','-format','%[colorspace]|%[profiles]',$nativePath)
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
    [string]$WorkRoot, [System.Collections.Generic.List[object]]$OwnedCandidates)
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
  $result = Invoke-WinImgNativeProcess -Executable $MagickPath -Arguments @('-ping','-regard-warnings','-define','registry:filename:literal=true','-define','image:frames=0',$nativeSource,('ICC:' + $nativeProfile))
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
# Reporting must not retry a broken sink or recursively log its own failure.
function ConvertTo-WinImgLogText {
  param([string]$Text)
  return [regex]::Replace($Text, '[\x00-\x1f\x7f]', {
    param($match)
    return ('\u{0:x4}' -f [int][char]$match.Value)
  })
}

function Write-WinImgEmergencyReport {
  param([string]$Message)
  # The PowerShell pipeline may already be stopped. stderr is best effort only;
  # an absent/closed console cannot be repaired by recursively calling Write-Log.
  try { [Console]::Error.WriteLine($Message) } catch {}
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
  catch { Write-WinImgEmergencyReport 'The normal console report also failed; stderr fallback is best effort.' }
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
  } catch { Set-WinImgLogFailure -State $State -Phase 'Creation' -Reason $_.Exception.Message }
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
    } catch { Set-WinImgLogFailure -State $State -Phase 'Append' -Reason $_.Exception.Message }
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
    } catch { Write-WinImgEmergencyReport (Get-WinImgBoundedText $line 1024) }
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
    } catch { $failed.Add([pscustomobject]@{ Name = $field; Reason = Get-WinImgBoundedText (ConvertTo-WinImgLogText $_.Exception.Message) 512 }) }
  }
  return [pscustomobject]@{ Succeeded = ($failed.Count -eq 0); FailedFields = $failed.ToArray() }
}

function Invoke-WinImgNormalizer {
  param(
    [string]$Source,
    [object]$MaxBytes = 1MB,
    [string]$OutputParent,
    [string]$MagickPath,
    [scriptblock]$PreflightRunner,
    [scriptblock]$ProcessRunner = {
      param([string]$Executable, [string[]]$Arguments)
      Invoke-WinImgNativeProcess -Executable $Executable -Arguments $Arguments
    }
  )

  $ErrorActionPreference = 'Stop'
  $logState = $null

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
  } catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog -State $logState; throw }
  catch {
    Write-Host ('Setup error: ' + $_.Exception.Message) -ForegroundColor Red
    if ($_.Exception.Data['WinImgNativeDetail']) {
      # No run log exists before setup succeeds. Keep its native reason visible
      # within a small console bound without creating output during preflight.
      Write-Host ('Native setup details: ' + (Get-WinImgBoundedText ([string]$_.Exception.Data['WinImgNativeDetail']) 512)) -ForegroundColor Red
    }
    return 1
  }

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
    $destRoot = New-WinImgRunDirectory -OutputParent $pictures -BaseName $base -Stamp $stamp
    $generatedRoot = [IO.Path]::Combine($destRoot, $generatedName)
    $workRoot = [IO.Path]::Combine($generatedRoot, 'work')
    $reportRoot = [IO.Path]::Combine($generatedRoot, 'reports')
    foreach ($path in @($generatedRoot, $workRoot, $reportRoot)) {
      if (-not (New-WinImgExclusiveDirectory $path)) { throw 'A generated namespace was unexpectedly occupied; the run was not adopted.' }
    }
  } catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog -State $logState; throw }
  catch {
    Write-Host ('Setup error: could not create the destination directory. ' + $_.Exception.Message) -ForegroundColor Red
    return 1
  }

  # A failed required log is visible but does not discard valid media work.
  $LogPath = [System.IO.Path]::Combine($reportRoot, "WinImgNormalizer_${stamp}.log")
  $logState = New-WinImgLogState -Path $LogPath
  Initialize-WinImgRunLog -State $logState
  Write-Log "Source: $srcRoot"
  Write-Log "Destination: $destRoot"
  Write-Log "Generated work/report directory: $generatedName"
  foreach ($row in $outputPlan) {
    Write-Log ('PLAN {0}: {1} -> {2} ({3})' -f $row.Kind, $row.SourceRelativePath, $row.OutputRelativePath, $row.NamingReason)
  }
  Write-Log ("MaxBytes: {0} bytes ({1} MiB; best-effort target)" -f $MaxBytes.ToString([Globalization.CultureInfo]::InvariantCulture), [Math]::Round($MaxBytes/1MB,2))

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
    try {
      Assert-WinImgNoReparseAncestors $directory.FullName
      $rel = Get-WinImgRelativePath -Root $srcRoot -Path $directory.FullName
      $target = [System.IO.Path]::Combine($destRoot, $rel)
      Assert-WinImgNoReparseAncestors $target
      [IO.Directory]::CreateDirectory($target) | Out-Null
    } catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog -State $logState; throw }
    catch {
      $hasTraversalWarnings = $true
      Write-Log ("Could not mirror directory: {0} ({1})" -f $directory.FullName, $_.Exception.Message) 'WARN'
    }
  }

  # --------- File sets ---
  $total = $allFiles.Count
  if ($total -eq 0) {
    Write-Log "No images or videos found." 'WARN'
    Write-Log ('SUMMARY ConvertedImages=0 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=0 SizeWarnings=0 NativeWarnings=0 TimestampWarnings=0 ScanComplete={0} IncompleteDirectories={1} UninspectableEntries={2} SkippedLinks={3} LogWarnings={4} FallbackDropped={5}' -f
      $sourceTree.ScanComplete, $sourceTree.InaccessibleDirectoryCount, $sourceTree.UninspectableEntryCount, $sourceTree.SkippedReparsePointCount, $logState.FailureCount, $logState.FallbackDroppedLines)
    Write-Host ('Final reporting state: LogWarnings={0} DiskLogIncomplete={1} FallbackDropped={2}' -f $logState.FailureCount, $logState.Degraded, $logState.FallbackDroppedLines)
    Complete-WinImgRunLog -State $logState
    if ($hasTraversalWarnings -or $logState.Degraded) { return 2 }
    return 0
  }

  # --------- Dedupe + stats ---
  $retained = New-Object 'Collections.Generic.Dictionary[string,object]' ([StringComparer]::Ordinal)
  $hasDuplicateWarnings = $false
  $stats = [ordered]@{ Converted=0; CopiedVideo=0; SkippedDuplicate=0; Unsupported=0; Errors=0; SizeWarnings=0; NativeWarnings=0; TimestampWarnings=0 }
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
      [object]$ColourInfo, [string]$SrgbProfilePath)

    $extent = Get-ExtentString $MaxBytes
    $scales = 100,90,80,70,60,50
    $attemptCount = 0; $transientRetries = 0
    foreach ($p in $scales) {
      while ($true) {
        # Every size attempt or diagnosed transient retry gets a new candidate.
        if ($attemptCount -gt 0) {
          $previous = $OwnedCandidates[$OwnedCandidates.Count - 1]
          try { Remove-WinImgOwnedCandidate -CandidatePath $previous.Path -FileOwned $previous.FileOwned; $previous.Removed = $true }
          catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog -State $logState; throw }
          catch { Write-Log ('Could not remove superseded image scratch: ' + $_.Exception.Message) 'WARN' }
          $DestPath = New-WinImgImageCandidate -WorkRoot $WorkRoot
          $OwnedCandidates.Add([pscustomobject]@{ Path = $DestPath; FileOwned = $false; Removed = $false })
        }
        $ownership = $OwnedCandidates[$OwnedCandidates.Count - 1]
        $stream = [IO.FileStream]::new($DestPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        $ownership.FileOwned = $true; $stream.Dispose(); $attemptCount++
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
        try { $nativeResult = & $ProcessRunner $MagickCmd $nativeArguments }
        catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog -State $logState; throw }
        catch {
          $error = $_.Exception
          while ($error.InnerException) { $error = $error.InnerException }
          $win32 = if ($error -is [ComponentModel.Win32Exception]) { $error.NativeErrorCode }
            elseif ($error -is [IO.IOException] -and ($error.HResult -band 0xFFFF0000L) -eq 0x80070000L) { $error.HResult -band 0xFFFF }
            else { 0 }
          $nativeResult = [pscustomobject]@{ ExitCode=$null; StartError=Get-WinImgBoundedText $error.Message 512; Win32ErrorCode=$win32 }
        }
        $outcome = Get-WinImgNativeOutcome -Result $nativeResult
        $exit = if ($null -eq $outcome.ExitCode) { 'none' } else { [string]$outcome.ExitCode }
        Write-Log ('NATIVE IMG: Attempt={0}; Scale={1}%; Category={2}; Exit={3}; TransientRetries={4}/2' -f $attemptCount,$p,$outcome.Category,$exit,$transientRetries) -Quiet
        if (-not $outcome.Acceptable -or $outcome.Warning) { Write-Log ('NATIVE DETAILS: ' + $outcome.DiagnosticText) -Quiet }
        if (-not $outcome.Acceptable) {
          if ($outcome.Retryable -and $transientRetries -lt 2) {
            $transientRetries++; $delay = 100 * $transientRetries
            Write-Log ('RETRY IMG: Category=TransientIO; Attempt={0}; Scale={1}%; Retry={2}/2; DelayMs={3}; Reason=sharing or lock violation; unchanged colour/alpha policy' -f $attemptCount,$p,$transientRetries,$delay) 'WARN'
            Start-Sleep -Milliseconds $delay
            continue
          }
          return @{ Status='Error'; Attempts=$attemptCount; Note=('Native conversion stopped: Category={0}; Exit={1}; TransientRetries={2}/2; see NATIVE DETAILS in log' -f $outcome.Category,$exit,$transientRetries) }
        }
        try { $validation = Test-WinImgImageCandidate -CandidatePath $DestPath -MagickPath $MagickCmd }
        catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog -State $logState; throw }
        catch {
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
      $sourceCreated = $f.CreationTimeUtc
      $dupKey = Get-WinImgDuplicateKey -Name $f.Name -LastWriteTimeUtc $sourceModified -Length $sourceLength
      if ($retained.ContainsKey($dupKey)) {
        $first = $retained[$dupKey]
        $stats.SkippedDuplicate++
        Write-Log ('Heuristic duplicate skipped: {0} (retained source: {1}; retained output: {2}; retained status: {3})' -f
          $rel, $first.SourceRelativePath, $first.OutputRelativePath, $first.Status) 'SKIP'
        continue
      }
      Assert-WinImgOutputAvailable $destPath
    } catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog -State $logState; throw }
    catch {
      $hasTraversalWarnings = $true
      $stats.Errors++
      Write-Log "Source/destination safety check failed: $rel ($($_.Exception.Message))" 'ERR'
      continue
    }
    $ownedCandidates = New-Object 'System.Collections.Generic.List[object]'
    try {
      if ($imgExts -contains $ext) {
        # Inspection and every attempt use the same profile-bearing bytes.
        $imageSource = New-WinImgSourceSnapshot -SourcePath $f.FullName -WorkRoot $workRoot -ExpectedLength $sourceLength -ExpectedModified $sourceModified -OwnedCandidates $ownedCandidates
        $sourceInfo = $null
        if ($ext -in @('.gif','.tif','.tiff','.webp','.heic','.heif')) {
          $sourceInfo = Get-WinImgSourceImageInfo -SnapshotPath $imageSource -MagickPath $MagickCmd
          Write-Log ('SOURCE IMG: {0} (SourceCount={1}; Selected=1; Omitted={2}; Unit={3}; Policy={4}; Decoder={5})' -f
            $rel, $sourceInfo.SourceCount, $sourceInfo.Omitted, $sourceInfo.Unit, $sourceInfo.Policy, $sourceInfo.Decoder)
        }
        $colourInfo = Get-WinImgColourInfo -SnapshotPath $imageSource -MagickPath $MagickCmd
        $srgbProfilePath = $null
        if ($colourInfo.HasIcc) {
          $null = New-WinImgSourceIccProfile -SnapshotPath $imageSource -MagickPath $MagickCmd -ColourSpace $colourInfo.ColourSpace -WorkRoot $workRoot -OwnedCandidates $ownedCandidates
          $srgbProfilePath = New-WinImgSrgbProfile -WorkRoot $workRoot -OwnedCandidates $ownedCandidates
        }
        Write-Log ('COLOUR IMG: {0} (SourceSpace={1}; SourceICC={2}; Policy={3}; Intent={4}; Alpha=WhiteAfterSrgb; OutputICC=None)' -f
          $rel, $colourInfo.ColourSpace, $colourInfo.HasIcc, $colourInfo.Policy, $(if ($colourInfo.HasIcc) { 'Relative' } else { 'None' }))
        try { $candidatePath = New-WinImgImageCandidate -WorkRoot $workRoot }
        catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog -State $logState; throw }
        catch { $hasNamingWarnings = $true; throw }
        $ownedCandidates.Add([pscustomobject]@{ Path = $candidatePath; FileOwned = $false; Removed = $false })
        $res = Convert-ImageMagick -SourcePath $imageSource -DestPath $candidatePath -MaxBytes $MaxBytes -WorkRoot $workRoot -OwnedCandidates $ownedCandidates -SourceInfo $sourceInfo -ColourInfo $colourInfo -SrgbProfilePath $srgbProfilePath
        if ($res.Status -in @('Converted', 'ConvertedWithWarning')) {
          try { Move-WinImgPlannedImage -CandidatePath $res.CandidatePath -DestinationPath $destPath }
          catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog -State $logState; throw }
          catch { $hasNamingWarnings = $true; throw }
          $stats.Converted++
          if ($res.SizeWarning) { $stats.SizeWarnings++ }
          if ($res.NativeWarning) { $stats.NativeWarnings++ }
          $timestampResult = Set-WinImgOutputTimestamps -Path $destPath -LastWriteTimeUtc $sourceModified -CreationTimeUtc $sourceCreated
          if (-not $timestampResult.Succeeded) {
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
          catch {
            $hasDuplicateWarnings = $true
            Write-Log ('Finalized image kept without heuristic registration: {0} ({1})' -f $rel, $_.Exception.Message) 'WARN'
          }
          $note = if ($res.Note) { " ($($res.Note))" } else { "" }
          $level = if ($res.Status -eq 'ConvertedWithWarning') { 'WARN' } else { 'OK' }
          Write-Log ("{0} IMG: {1} -> {2} [{3} bytes, MaxBytes={4}, Width={5}, Height={6}, Scale={7}%]{8}" -f
            $level, $rel, $destRel, $res.BytesOut.ToString([Globalization.CultureInfo]::InvariantCulture),
            $MaxBytes.ToString([Globalization.CultureInfo]::InvariantCulture), $res.Width, $res.Height, $res.Scale, $note) $level
        } else {
          $stats.Errors++; Write-Log "ERR IMG: $rel ($($res.Note))" 'ERR'
        }
      }
      elseif ($videoExts -contains $ext) {
        $destDir = [System.IO.Path]::GetDirectoryName($destPath)
        if ($destDir) { [IO.Directory]::CreateDirectory($destDir) | Out-Null }
        try { $videoResult = Copy-WinImgPlannedVideo -SourcePath $f.FullName -DestinationPath $destPath -WorkRoot $workRoot }
        catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog -State $logState; throw }
        catch { $hasNamingWarnings = $true; throw }
        # The copy helper's stable snapshot describes the bytes actually copied,
        # even if source metadata changed between lookup and staging.
        $videoKey = Get-WinImgDuplicateKey -Name $f.Name -LastWriteTimeUtc $videoResult.LastWriteTimeUtc -Length $videoResult.BytesOut
        $timestampResult = Set-WinImgOutputTimestamps -Path $destPath -LastWriteTimeUtc $videoResult.LastWriteTimeUtc -CreationTimeUtc $videoResult.CreationTimeUtc
        $videoStatus = 'CopiedVideo'
        if (-not $timestampResult.Succeeded) {
          $stats.TimestampWarnings++
          $videoStatus = 'CopiedVideoWithWarning'
          foreach ($failure in $timestampResult.FailedFields) { Write-Log ('Timestamp warning VID: {0} ({1}: {2})' -f $destRel, $failure.Name, $failure.Reason) 'WARN' }
        }
        if (-not $retained.ContainsKey($videoKey)) {
          $retained.Add($videoKey, [pscustomobject]@{ SourceRelativePath = $rel; OutputRelativePath = $destRel; Status = $videoStatus })
        }
        if ($videoResult.CleanupWarning) { $hasNamingWarnings = $true; Write-Log ('Video scratch cleanup: ' + $videoResult.CleanupWarning) 'WARN' }
        $bytesOut = $videoResult.BytesOut
        $stats.CopiedVideo++
        $videoLevel = if ($timestampResult.Succeeded) { 'OK' } else { 'WARN' }
        Write-Log ("{0} VID: {1} -> {2} [{3:n0} bytes]" -f $videoLevel, $rel, $destRel, $bytesOut) $videoLevel
      }
      else {
        $stats.Unsupported++; Write-Log "Unsupported skipped: $rel" 'WARN'
      }
    } catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog -State $logState; throw }
    catch {
      if ($_.Exception.Data['WinImgNativeDetail']) { Write-Log ('NATIVE DETAILS: ' + $_.Exception.Data['WinImgNativeDetail']) -Quiet }
      $stats.Errors++; Write-Log "Exception processing: $rel ($($_.Exception.Message))" 'ERR'
    } finally {
      foreach ($owned in $ownedCandidates) {
        if ($owned.Removed) { continue }
        try { Remove-WinImgOwnedCandidate -CandidatePath $owned.Path -FileOwned $owned.FileOwned }
        catch [Management.Automation.PipelineStoppedException] { Complete-WinImgRunLog -State $logState; throw }
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
  Write-Host ("  Above byte target: {0} (included in converted images)" -f $stats.SizeWarnings)
  Write-Host ("  Native warnings: {0} (included in converted images)" -f $stats.NativeWarnings)
  Write-Host ("  Timestamp warnings: {0} (included in finalized images/videos)" -f $stats.TimestampWarnings)
  Write-Host ("  Scan complete: {0}; incomplete directories: {1}; uninspectable entries: {2}; skipped links: {3}" -f $sourceTree.ScanComplete, $sourceTree.InaccessibleDirectoryCount, $sourceTree.UninspectableEntryCount, $sourceTree.SkippedReparsePointCount)
  Write-Host ("  Copied videos    : {0}" -f $stats.CopiedVideo)
  Write-Host ("  Duplicates       : {0}" -f $stats.SkippedDuplicate)
  Write-Host ("  Unsupported      : {0}" -f $stats.Unsupported)
  Write-Host ("  Errors           : {0}" -f $stats.Errors)
  Write-Host ("  Log file         : {0}" -f $LogPath)

  Write-Log ("SUMMARY ConvertedImages={0} CopiedVideos={1} Duplicates={2} Unsupported={3} Errors={4} SizeWarnings={5} NativeWarnings={6} TimestampWarnings={7} ScanComplete={8} IncompleteDirectories={9} UninspectableEntries={10} SkippedLinks={11} LogWarnings={12} FallbackDropped={13}" -f
    $stats.Converted,$stats.CopiedVideo,$stats.SkippedDuplicate,$stats.Unsupported,$stats.Errors,$stats.SizeWarnings,$stats.NativeWarnings,$stats.TimestampWarnings,
    $sourceTree.ScanComplete,$sourceTree.InaccessibleDirectoryCount,$sourceTree.UninspectableEntryCount,$sourceTree.SkippedReparsePointCount,$logState.FailureCount,$logState.FallbackDroppedLines)
  Write-Log "Processing ended $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
  Write-Host ('Final reporting state: LogWarnings={0} DiskLogIncomplete={1} FallbackDropped={2}' -f $logState.FailureCount, $logState.Degraded, $logState.FallbackDroppedLines)
  Complete-WinImgRunLog -State $logState

  if ($stats.Errors -gt 0 -or $stats.SizeWarnings -gt 0 -or $stats.NativeWarnings -gt 0 -or $stats.TimestampWarnings -gt 0 -or $logState.Degraded -or $missingCoders.Count -gt 0 -or $hasTraversalWarnings -or $hasNamingWarnings -or $hasDuplicateWarnings) { return 2 }
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
