# Development-only packaging. No install, tag, publication or runtime download.
[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidatePattern('^[0-9a-f]{40}$')][string]$Revision,
    [Parameter(Mandatory)][string]$OutputDirectory
)

$ErrorActionPreference = 'Stop'
$repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$scratch = [IO.Path]::GetFullPath((Join-Path $repository '.scratch'))
$utf8 = [Text.UTF8Encoding]::new($false, $true)
$git = (Get-Command git -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source

function Assert-ReleaseAncestors([string]$Path) {
    $current = [IO.DirectoryInfo]::new([IO.Path]::GetFullPath($Path))
    while ($null -ne $current) {
        if ($current.Exists -and ($current.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw 'Release directories must not traverse reparse points.'
        }
        $current = $current.Parent
    }
}

function Invoke-ReleaseGit([string[]]$Arguments, [switch]$AllowFailure) {
    # Preserve blob bytes (notably the CMD launcher's CRLF). A PowerShell native
    # text pipeline would decode/re-encode stdout and cannot do this safely.
    $start = New-Object Diagnostics.ProcessStartInfo
    $start.FileName = $git
    $tokens = @('--no-optional-locks', '--no-pager', '-C', $repository) + $Arguments
    $start.Arguments = ($tokens | ForEach-Object {
        '"' + [regex]::Replace([regex]::Replace([string]$_, '(\\*)"', '$1$1\"'), '(\\+)$', '$1$1') + '"'
    }) -join ' '
    $start.UseShellExecute = $false; $start.CreateNoWindow = $true
    $start.RedirectStandardOutput = $true; $start.RedirectStandardError = $true
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $start
    $buffer = New-Object IO.MemoryStream
    $started = $false
    try {
        if (-not $process.Start()) { throw 'Release Git process did not start.' }
        $started = $true
        $stdout = $process.StandardOutput.BaseStream.CopyToAsync($buffer)
        $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit(30000)) { $process.Kill(); throw 'Release Git process exceeded its bound.' }
        if (-not $stdout.Wait(5000) -or -not $stderr.Wait(5000)) { throw 'Release Git capture did not complete.' }
        if ($buffer.Length -gt 5MB) { throw 'Release Git output exceeds the packaging bound.' }
        if (-not $AllowFailure -and ($process.ExitCode -ne 0 -or $stderr.GetAwaiter().GetResult())) {
            throw 'Release Git command failed; no package accepted.'
        }
        return [pscustomobject]@{ Bytes = $buffer.ToArray(); ExitCode = $process.ExitCode }
    } finally {
        if ($started -and -not $process.HasExited) { $process.Kill(); $null = $process.WaitForExit(5000) }
        $process.Dispose(); $buffer.Dispose()
    }
}

function Get-ReleaseGitText([string[]]$Arguments) {
    $result = Invoke-ReleaseGit $Arguments
    return $utf8.GetString($result.Bytes).TrimEnd("`r", "`n")
}

function Assert-ReleaseRevision {
    if ((Get-ReleaseGitText @('rev-parse', '--show-toplevel')).Replace('/', '\') -ine $repository.Replace('/', '\')) {
        throw 'Builder must belong to the selected repository root.'
    }
    if ((Get-ReleaseGitText @('rev-parse', '--verify', 'HEAD')) -cne $Revision) {
        throw 'Release revision must be the exact current HEAD commit.'
    }
    if (Get-ReleaseGitText @('status', '--porcelain=v1', '--untracked-files=all', '--ignore-submodules=none')) {
        throw 'Release build requires a clean tracked and nonignored untracked worktree.'
    }
    # The builder being executed must itself come from that clean commit.
    $builderBlob = Get-ReleaseBlob 'tools/release/Build-Release.ps1'
    if ((Get-ReleaseHash $builderBlob) -cne (Get-ReleaseHash ([IO.File]::ReadAllBytes($PSCommandPath)))) {
        # A text checkout may differ only in line endings, which do not change
        # this development script. The ZIP still always uses exact Git bytes.
        $actual = $utf8.GetString([IO.File]::ReadAllBytes($PSCommandPath)).TrimStart([char]0xfeff).Replace("`r`n", "`n")
        $expected = $utf8.GetString($builderBlob).TrimStart([char]0xfeff).Replace("`r`n", "`n")
        if ($actual -cne $expected) { throw 'Executed release builder differs from the selected commit.' }
    }
}

function Get-ReleaseBlob([string]$Path) {
    $tree = Get-ReleaseGitText @('ls-tree', $Revision, '--', $Path)
    if ($tree -cnotmatch ('^100(?:644|755) blob [0-9a-f]{40}\t' + [regex]::Escape($Path) + '$')) {
        throw ('Release input is missing or is not a regular tracked file: ' + $Path)
    }
    $result = Invoke-ReleaseGit @('cat-file', 'blob', ($Revision + ':' + $Path))
    if ($result.Bytes.Length -eq 0) { throw 'Release inputs must not be empty.' }
    return ,$result.Bytes
}

function Get-ReleaseHash([byte[]]$Bytes) {
    $hasher = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($hasher.ComputeHash($Bytes)).Replace('-', '').ToLowerInvariant() }
    finally { $hasher.Dispose() }
}

Assert-ReleaseAncestors $repository
Assert-ReleaseRevision
$output = [IO.Path]::GetFullPath($OutputDirectory)
if (-not $output.StartsWith($scratch + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Release output must be a new child directory under repository .scratch.'
}
Assert-ReleaseAncestors $output
if (Test-Path -LiteralPath $output) { throw 'Release output already exists; no replacement is allowed.' }
$ignored = Invoke-ReleaseGit @('check-ignore', '--quiet', '--no-index', '--', $output) -AllowFailure
if ($ignored.ExitCode -ne 0) { throw 'Release output must already be ignored by Git.' }

# Reviewed content mapping. No directory recursion or user-supplied patterns.
# The fixture directory contributes only the original redistributed ICC license.
$allowlist = [ordered]@{
    'CHANGELOG.md' = 'CHANGELOG.md'
    'GETTING_STARTED.md' = 'docs/release/GETTING_STARTED.md'
    'LICENSE' = 'LICENSE'
    'LICENSE-CC0.txt' = 'tests/fixtures/colour/LICENSE-CC0.txt'
    'THIRD_PARTY_NOTICES.md' = 'docs/release/THIRD_PARTY_NOTICES.md'
    'WinImgNormalizer.bat' = 'WinImgNormalizer.bat'
    'WinImgNormalizer.ps1' = 'WinImgNormalizer.ps1'
}
$entries = @{}
$fileRecords = @()
foreach ($name in $allowlist.Keys) {
    [byte[]]$bytes = Get-ReleaseBlob $allowlist[$name]
    $entries[$name] = $bytes
    $fileRecords += [ordered]@{ path = $name; source_path = $allowlist[$name]; bytes = $bytes.Length; sha256 = Get-ReleaseHash $bytes }
}
$application = $utf8.GetString($entries['WinImgNormalizer.ps1']).TrimStart([char]0xfeff)
$versions = [regex]::Matches($application, '(?m)^function Get-WinImgVersion\s*\{\s*return ''((?:0|[1-9][0-9]*)\.(?:0|[1-9][0-9]*)\.(?:0|[1-9][0-9]*))''\s*\}')
if ($versions.Count -ne 1) { throw 'Source must expose exactly one literal semantic application version.' }
$version = $versions[0].Groups[1].Value
$metadata = $utf8.GetString((Get-ReleaseBlob 'release-metadata.json')) | ConvertFrom-Json
if ($metadata.version -cne $version -or $metadata.release_state -cne 'unreleased' -or $metadata.product -cne 'WinImgNormalizer') {
    throw 'Release metadata must agree with the source version and unreleased state.'
}
foreach ($field in @('tag', 'release_date', 'release_url', 'asset_filename', 'asset_bytes', 'asset_sha256', 'download_url')) {
    if ($null -ne $metadata.$field) { throw 'Preparation cannot claim published asset or release values.' }
}
$notes = $utf8.GetString((Get-ReleaseBlob 'docs/release/NOTES.md')).Replace("`r`n", "`n").TrimEnd()
$expectedChangelog = "# Changelog`n`n<!-- Generated by tools/release/Update-ReleaseMetadata.ps1; edit docs/release/NOTES.md. -->`n`n## $version (unreleased)`n`n$notes`n"
if ((Get-ReleaseHash $entries['CHANGELOG.md']) -cne (Get-ReleaseHash $utf8.GetBytes($expectedChangelog))) {
    throw 'Generated changelog is stale; regenerate before committing the build revision.'
}
$manifest = [ordered]@{
    schema_version = 1; product = 'WinImgNormalizer'; version = $version
    release_state = 'unreleased'; source_revision = $Revision; digital_signature = $null
    files = $fileRecords
}
$entries['package-manifest.json'] = $utf8.GetBytes(($manifest | ConvertTo-Json -Depth 6 -Compress) + "`n")
$asset = 'WinImgNormalizer-' + $version + '-portable.zip'
$staging = Join-Path $scratch ('release-staging-' + [Guid]::NewGuid().ToString('N'))
Assert-ReleaseAncestors $staging
if (Test-Path -LiteralPath $staging) { throw 'Release staging ownership collision.' }
$null = [IO.Directory]::CreateDirectory($staging)
# Retain a failed owned staging directory for inspection; never delete unknown
# files or partially replace an earlier build.
Add-Type -AssemblyName System.IO.Compression
$zipPath = Join-Path $staging $asset
$stream = [IO.File]::Open($zipPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
try {
    $zip = [IO.Compression.ZipArchive]::new($stream, [IO.Compression.ZipArchiveMode]::Create, $true)
    try {
        [string[]]$names = @($entries.Keys)
        [Array]::Sort($names, [StringComparer]::Ordinal)
        foreach ($name in $names) {
            $entry = $zip.CreateEntry($name, [IO.Compression.CompressionLevel]::NoCompression)
            $entry.LastWriteTime = [DateTimeOffset]::new(2000, 1, 1, 0, 0, 0, [TimeSpan]::Zero)
            $entry.ExternalAttributes = 0
            $entryStream = $entry.Open()
            try { $entryStream.Write($entries[$name], 0, $entries[$name].Length) }
            finally { $entryStream.Dispose() }
        }
    } finally { $zip.Dispose() }
} finally { $stream.Dispose() }
$assetHash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
$assetBytes = (Get-Item -LiteralPath $zipPath).Length
$provenance = [ordered]@{
    schema_version = 1; product = 'WinImgNormalizer'; version = $version
    release_state = 'unreleased'; source_revision = $Revision; digital_signature = $null
    asset_filename = $asset; asset_bytes = $assetBytes; asset_sha256 = $assetHash
    package_manifest_sha256 = Get-ReleaseHash $entries['package-manifest.json']
    builder_source_sha256 = Get-ReleaseHash (Get-ReleaseBlob 'tools/release/Build-Release.ps1')
    reproducibility = 'Exact Git content, ordinal entries, fixed ZIP time/attributes and the NoCompression setting. ZIP framing varies by runtime. Byte equality must be verified for the actual build hosts; no universal ZIP-writer guarantee.'
}
$provenancePath = Join-Path $staging 'build-provenance.json'
[IO.File]::WriteAllText($provenancePath, ($provenance | ConvertTo-Json -Depth 4 -Compress) + "`n", $utf8)
$provenanceHash = (Get-FileHash -LiteralPath $provenancePath -Algorithm SHA256).Hash.ToLowerInvariant()
[IO.File]::WriteAllText((Join-Path $staging 'SHA256SUMS.txt'), "$assetHash  $asset`n$provenanceHash  build-provenance.json`n", $utf8)
Assert-ReleaseRevision
Assert-ReleaseAncestors $output
Assert-ReleaseAncestors $staging
[IO.Directory]::Move($staging, $output)
[pscustomobject]@{ source_revision = $Revision; version = $version; asset_filename = $asset; asset_bytes = $assetBytes; asset_sha256 = $assetHash }
