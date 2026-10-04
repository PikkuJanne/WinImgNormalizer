<#
Runs a reviewed legacy snapshot in an already-isolated child PowerShell process.
The parent runner owns the timeout and termination of this entire process tree.
Only the Pictures resolution block is replaced; the application is never imported
or executed from its original repository path. Raw evidence remains in scratch.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Repository,
    [Parameter(Mandatory = $true)][string]$ScratchRoot,
    [Parameter(Mandatory = $true)][string]$Source,
    [Parameter(Mandatory = $true)][string]$PicturesRoot,
    [long]$MaxBytes = 1048576,
    [string]$MagickPath,
    [string]$FakeMode,
    [ValidatePattern('^[0-9A-Fa-f]{64}$')]
    [string]$ExpectedLegacySha256 = '3D31EAEA3D926C2D2A4F75DC32856E6681FCF03F2F29AC89020AA1A4A907133A'
)

$ErrorActionPreference = 'Stop'

function Get-Sha256 {
    param([string]$Path)
    $hasher = [Security.Cryptography.SHA256]::Create()
    $stream = [IO.File]::OpenRead($Path)
    try { return ([BitConverter]::ToString($hasher.ComputeHash($stream))).Replace('-', '') }
    finally { $stream.Dispose(); $hasher.Dispose() }
}

function Get-CanonicalDirectoryPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not [IO.Path]::IsPathRooted($Path)) {
        throw 'Containment requires an explicit absolute path.'
    }
    $full = [IO.Path]::GetFullPath($Path)
    $pathRoot = [IO.Path]::GetPathRoot($full)
    if ($full.TrimEnd('\', '/') -eq $pathRoot.TrimEnd('\', '/')) {
        throw 'A drive or share root is not a disposable test directory.'
    }
    return $full.TrimEnd('\', '/')
}

function Test-ContainsPath {
    param([string]$Parent, [string]$Child)
    return $Child.Equals($Parent, [StringComparison]::OrdinalIgnoreCase) -or
        $Child.StartsWith($Parent + [IO.Path]::DirectorySeparatorChar,
            [StringComparison]::OrdinalIgnoreCase)
}

function Assert-NoReparseAncestors {
    param([string]$Path)
    $current = $Path
    while (-not [string]::IsNullOrEmpty($current)) {
        if (Test-Path -LiteralPath $current) {
            $entry = Get-Item -LiteralPath $current -Force
            if (($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw 'A reparse point occurs in a containment path.'
            }
        }
        $parent = [IO.Directory]::GetParent($current)
        if ($null -eq $parent) { break }
        $current = $parent.FullName
    }
}

function Assert-NoReparseTree {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { return }
    $pending = New-Object 'System.Collections.Generic.Queue[string]'
    $pending.Enqueue($Path)
    $count = 0
    while ($pending.Count -gt 0) {
        $directory = $pending.Dequeue()
        foreach ($entry in Get-ChildItem -LiteralPath $directory -Force) {
            $count++
            if ($count -gt 10000) { throw 'Disposable test tree exceeds its safety bound.' }
            if (($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw 'A reparse point occurs inside the disposable test tree.'
            }
            if ($entry.PSIsContainer) { $pending.Enqueue($entry.FullName) }
        }
    }
}

$repositoryFull = Get-CanonicalDirectoryPath $Repository
$scratchFull = Get-CanonicalDirectoryPath $ScratchRoot
$sourceFull = Get-CanonicalDirectoryPath $Source
$picturesFull = Get-CanonicalDirectoryPath $PicturesRoot
if ($MaxBytes -le 0) { throw 'The characterization byte limit must be positive.' }
if (-not (Test-Path -LiteralPath $repositoryFull -PathType Container) -or
    -not (Test-Path -LiteralPath $scratchFull -PathType Container) -or
    -not (Test-Path -LiteralPath $sourceFull -PathType Container)) {
    throw 'Repository, scratch, and source directories must already exist.'
}
foreach ($ownedPath in @($sourceFull, $picturesFull)) {
    if ($ownedPath.Equals($scratchFull, [StringComparison]::OrdinalIgnoreCase) -or
        -not (Test-ContainsPath $scratchFull $ownedPath)) {
        throw 'Source and injected Pictures must be distinct descendants of scratch.'
    }
}
if ((Test-ContainsPath $sourceFull $picturesFull) -or
    (Test-ContainsPath $picturesFull $sourceFull)) {
    throw 'Nested or overlapping source and output trees are never executed.'
}
foreach ($guardedPath in @($repositoryFull, $scratchFull, $sourceFull, $picturesFull)) {
    Assert-NoReparseAncestors $guardedPath
}
$marker = Join-Path $scratchFull '.winimg-fixture-root'
if (-not (Test-Path -LiteralPath $marker -PathType Leaf)) {
    throw 'Scratch is missing the explicit synthetic-fixture ownership marker.'
}
Assert-NoReparseTree $scratchFull

# Known Folders is read only for this guard; USERPROFILE is never redefined.
$actualPictures = [Environment]::GetFolderPath('MyPictures')
$protectedPaths = @($actualPictures)
if (-not [string]::IsNullOrWhiteSpace($env:USERPROFILE)) {
    $protectedPaths += Join-Path $env:USERPROFILE 'Pictures'
    $profileFull = Get-CanonicalDirectoryPath $env:USERPROFILE
    if (Test-ContainsPath $scratchFull $profileFull) {
        throw 'Scratch cannot contain the real user profile.'
    }
}
foreach ($protectedPath in $protectedPaths) {
    if ([string]::IsNullOrWhiteSpace($protectedPath)) { continue }
    $protectedFull = Get-CanonicalDirectoryPath $protectedPath
    if ((Test-ContainsPath $scratchFull $protectedFull) -or
        (Test-ContainsPath $protectedFull $scratchFull)) {
        throw 'Scratch cannot overlap either real Pictures location.'
    }
}

# Scratch inside the repository must already be ignored. Outside-repository
# scratch is inherently outside the tracked checkout.
if (Test-ContainsPath $repositoryFull $scratchFull) {
    $git = Get-Command git -CommandType Application -ErrorAction Stop | Select-Object -First 1
    $ignoreProbe = Join-Path $scratchFull 'legacy-snapshot-ignore-probe'
    & $git.Source -C $repositoryFull check-ignore --quiet --no-index -- $ignoreProbe
    if ($LASTEXITCODE -ne 0) { throw 'Repository scratch must be ignored before launch.' }
}

$legacyPath = Join-Path $repositoryFull 'WinImgNormalizer.ps1'
if (-not (Test-Path -LiteralPath $legacyPath -PathType Leaf)) {
    throw 'The legacy application is missing.'
}
Assert-NoReparseAncestors $legacyPath
$legacyHash = Get-Sha256 $legacyPath
if (-not $legacyHash.Equals($ExpectedLegacySha256, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'The legacy application hash differs from the explicitly reviewed source.'
}
$legacyText = [IO.File]::ReadAllText($legacyPath)
$picturesBlock = @'
try { $pictures = [Environment]::GetFolderPath('MyPictures') } catch { $pictures = $null }
if ([string]::IsNullOrWhiteSpace($pictures) -or -not (Test-Path -LiteralPath $pictures)) {
  $pictures = Join-Path $env:USERPROFILE 'Pictures'
  if (-not (Test-Path -LiteralPath $pictures)) { New-Item -ItemType Directory -Path $pictures -Force | Out-Null }
}
'@
$lineEnding = if ($legacyText.Contains("`r`n")) { "`r`n" } else { "`n" }
$picturesBlock = $picturesBlock.Replace("`r`n", "`n").Replace("`n", $lineEnding)
$blockStart = $legacyText.IndexOf($picturesBlock, [StringComparison]::Ordinal)
if ($blockStart -lt 0 -or
    $legacyText.IndexOf($picturesBlock, $blockStart + $picturesBlock.Length,
        [StringComparison]::Ordinal) -ge 0) {
    throw 'Expected legacy Pictures block is absent or ambiguous.'
}
$picturesEncoded = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($picturesFull))
# Preserve the legacy file encoding while keeping the injected code ASCII-safe
# for Windows PowerShell 5.1, including a Unicode scratch/output path.
$replacement = "# M0-T02 harness-only Pictures injection; legacy processing is unchanged." +
    $lineEnding + '$pictures = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String(' +
    "'" + $picturesEncoded + "'))"
$snapshotText = $legacyText.Substring(0, $blockStart) + $replacement +
    $legacyText.Substring($blockStart + $picturesBlock.Length)
$tokens = $null
$parseErrors = $null
$null = [Management.Automation.Language.Parser]::ParseInput($snapshotText,
    [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -ne 0) { throw 'The contained legacy snapshot does not parse.' }
if ((Get-Sha256 $legacyPath) -ne $legacyHash) {
    throw 'Legacy application changed while its snapshot was being prepared.'
}

$snapshotDirectory = Join-Path $scratchFull ('legacy-snapshot-' + [Guid]::NewGuid().ToString('N'))
if (Test-Path -LiteralPath $snapshotDirectory) { throw 'Snapshot ownership path already exists.' }
$null = [IO.Directory]::CreateDirectory($snapshotDirectory)
$temporaryDirectory = Join-Path $snapshotDirectory 'temporary'
$null = [IO.Directory]::CreateDirectory($temporaryDirectory)
if (Test-Path -LiteralPath $picturesFull) {
    if (-not (Test-Path -LiteralPath $picturesFull -PathType Container)) {
        throw 'Injected Pictures exists but is not a directory.'
    }
    if (@(Get-ChildItem -LiteralPath $picturesFull -Force).Count -ne 0) {
        throw 'Injected Pictures must be absent or empty to prevent stale-output reuse.'
    }
} else {
    $null = [IO.Directory]::CreateDirectory($picturesFull)
}
Assert-NoReparseAncestors $picturesFull
Assert-NoReparseTree $picturesFull
$snapshotPath = Join-Path $snapshotDirectory 'WinImgNormalizer.contained.ps1'
$legacyBytes = [IO.File]::ReadAllBytes($legacyPath)
$hasBom = $legacyBytes.Length -ge 3 -and $legacyBytes[0] -eq 239 -and
    $legacyBytes[1] -eq 187 -and $legacyBytes[2] -eq 191
$encoding = New-Object Text.UTF8Encoding($hasBom)
[IO.File]::WriteAllText($snapshotPath, $snapshotText, $encoding)
$snapshotHash = Get-Sha256 $snapshotPath
$powerShellExecutable = (Get-Process -Id $PID).Path
if ([string]::IsNullOrWhiteSpace($powerShellExecutable) -or
    -not (Test-Path -LiteralPath $powerShellExecutable -PathType Leaf)) {
    throw 'The current PowerShell executable could not be identified.'
}
$applicationArguments = @('-NoLogo', '-NoProfile', '-NonInteractive',
    '-ExecutionPolicy', 'Bypass', '-File', $snapshotPath, $sourceFull,
    $MaxBytes.ToString([Globalization.CultureInfo]::InvariantCulture))

$magickFull = $null
if (-not [string]::IsNullOrWhiteSpace($MagickPath)) {
    if (-not [IO.Path]::IsPathRooted($MagickPath)) { throw 'MagickPath must be absolute.' }
    $magickFull = [IO.Path]::GetFullPath($MagickPath)
    if (-not (Test-Path -LiteralPath $magickFull -PathType Leaf) -or
        [IO.Path]::GetFileName($magickFull) -notin @('magick.exe', 'magick.cmd')) {
        throw 'MagickPath must name an existing magick.exe or controlled magick.cmd.'
    }
    Assert-NoReparseAncestors $magickFull
    if (-not [string]::IsNullOrWhiteSpace($FakeMode) -and
        -not (Test-ContainsPath $scratchFull $magickFull)) {
        throw 'A controlled fake process must be owned by the marked scratch tree.'
    }
    $env:PATH = [IO.Path]::GetDirectoryName($magickFull) + [IO.Path]::PathSeparator + $env:PATH
    $resolvedMagick = Get-Command magick -ErrorAction Stop | Select-Object -First 1
    if ($resolvedMagick.CommandType -ne 'Application' -or
        -not $resolvedMagick.Source.Equals($magickFull, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'The explicitly selected ImageMagick command did not resolve exactly.'
    }
}
if (-not [string]::IsNullOrWhiteSpace($FakeMode) -and $null -eq $magickFull) {
    throw 'A fake-process scenario requires an explicitly selected native command.'
}
$env:MAGICK_TEMPORARY_PATH = $temporaryDirectory
$env:TEMP = $temporaryDirectory
$env:TMP = $temporaryDirectory
if (-not [string]::IsNullOrWhiteSpace($FakeMode)) { $env:WINIMG_FAKE_MODE = $FakeMode }
$metadata = [ordered]@{
    schema_version = 1
    recorded_utc = [DateTime]::UtcNow.ToString('o')
    execution_kind = 'instrumented-legacy-snapshot'
    seam = 'Replace only exact Pictures resolution block; invoke snapshot in a separate -File process, never dot-source.'
    legacy_sha256 = $legacyHash
    snapshot_sha256 = $snapshotHash
    legacy_path = $legacyPath
    snapshot_path = $snapshotPath
    scratch_root = $scratchFull
    source_root = $sourceFull
    pictures_root = $picturesFull
    temporary_root = $temporaryDirectory
    magick_path = $magickFull
    fake_mode = $FakeMode
    positional_arguments = @($sourceFull, $MaxBytes)
    powershell_version = $PSVersionTable.PSVersion.ToString()
    powershell_executable = $powerShellExecutable
    application_process_arguments = $applicationArguments
    original_application_executed = $false
    userprofile_redefined = $false
    nested_output_executed = $false
    parent_must_bound_process_tree = $true
    original_pictures_block = $picturesBlock
    replacement_pictures_block = $replacement
}
$metadataPath = Join-Path $snapshotDirectory 'containment.json'
[IO.File]::WriteAllText($metadataPath, ($metadata | ConvertTo-Json -Depth 5),
    (New-Object Text.UTF8Encoding($false)))
Write-Output ('CONTAINMENT_METADATA=' + $metadataPath)

# A separate -File host preserves actual application exit semantics, even when
# the last native ImageMagick call failed. The parent bounds this process tree.
& $powerShellExecutable @applicationArguments
$applicationExitCode = $LASTEXITCODE
exit $applicationExitCode
