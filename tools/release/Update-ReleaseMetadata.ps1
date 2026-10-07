# Development-only generation/checking; never creates a ZIP, tag or release.
[CmdletBinding()]
param([switch]$Check)

$ErrorActionPreference = 'Stop'
$repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
function Assert-ReleaseRegularFile([string]$Path, [switch]$AllowMissing) {
    $current = [IO.DirectoryInfo]::new([IO.Path]::GetDirectoryName($Path))
    while ($null -ne $current) {
        if ($current.Exists -and ($current.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Release metadata directory crosses a reparse point.' }
        $current = $current.Parent
    }
    if (Test-Path -LiteralPath $Path) {
        $item = Get-Item -LiteralPath $Path -Force
        if ($item -isnot [IO.FileInfo] -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Release metadata input/output must be a regular file.' }
    } elseif (-not $AllowMissing) { throw 'Required release metadata or publication proof is missing.' }
}
function Assert-ReleaseJsonObjectKeys {
    param([string]$Text, [hashtable]$Conversion)
    # JSON readers differ on first/last duplicate values. Track decoded names in
    # each individual object, so repeated asset field names remain valid while
    # duplicate or differently escaped spellings of one key cannot be accepted.
    $tokens = [regex]::Matches($Text, '"(?:[^"\\]|\\.)*"|[{}\[\]:,]|-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?|true|false|null')
    $scopes = New-Object 'Collections.Generic.Stack[object]'
    $end = 0
    for ($index = 0; $index -lt $tokens.Count; $index++) {
        $token = $tokens[$index]
        if ($Text.Substring($end, $token.Index - $end) -notmatch '^\s*$') { throw 'Release inputs must use ordinary JSON tokens.' }
        $end = $token.Index + $token.Length
        if ($token.Value -ceq '{') {
            $scopes.Push([pscustomobject]@{ kind='object'; keys=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase) })
        } elseif ($token.Value -ceq '[') {
            $scopes.Push([pscustomobject]@{ kind='array'; keys=$null })
        } elseif ($token.Value -cin @('}',']')) {
            if ($scopes.Count -eq 0) { throw 'Release JSON scope is invalid.' }
            $null = $scopes.Pop()
        } elseif ($token.Value.StartsWith('"', [StringComparison]::Ordinal) -and
                $index + 1 -lt $tokens.Count -and $tokens[$index + 1].Value -ceq ':') {
            if ($scopes.Count -eq 0 -or $scopes.Peek().kind -cne 'object') { throw 'Release JSON key scope is invalid.' }
            $name = ConvertFrom-Json -InputObject $token.Value @Conversion
            if (-not $scopes.Peek().keys.Add([string]$name)) { throw 'Release JSON object keys must be unique.' }
        }
    }
    if ($Text.Substring($end) -notmatch '^\s*$') { throw 'Release inputs must use ordinary JSON tokens.' }
}
function Read-ReleaseJson([string]$Path) {
    Assert-ReleaseRegularFile $Path
    if ((Get-Item -LiteralPath $Path).Length -gt 65536) { throw 'Release metadata or publication proof exceeds its bounded JSON size.' }
    try {
        $jsonArguments = @{ ErrorAction = 'Stop' }
        if ((Get-Command ConvertFrom-Json).Parameters.ContainsKey('DateKind')) { $jsonArguments.DateKind = 'String' }
        $text = [IO.File]::ReadAllText($Path)
        Assert-ReleaseJsonObjectKeys $text $jsonArguments
        return ($text | ConvertFrom-Json @jsonArguments)
    }
    catch { throw 'Release metadata or publication proof is stale or malformed JSON.' }
}
function Assert-ReleaseFields($Value, [string[]]$Names) {
    if ($Value -isnot [pscustomobject] -or
        (@($Value.PSObject.Properties.Name | Sort-Object) -join '|') -cne (@($Names | Sort-Object) -join '|')) { throw 'Release metadata or publication proof fields differ from the exact schema.' }
}
function Test-ReleaseInteger($Value) { return ($Value -is [int] -or $Value -is [long]) }
function Get-ReleaseUtcInstant($Value) {
    if ($Value -isnot [string] -or $Value -cnotmatch '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,7})?Z$') { throw 'Publication proof requires a UTC ISO timestamp.' }
    $instant = [DateTimeOffset]::MinValue
    [string[]]$formats = @("yyyy-MM-dd'T'HH:mm:ss'Z'", "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF'Z'")
    $style = [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal
    if (-not [DateTimeOffset]::TryParseExact($Value, $formats, [Globalization.CultureInfo]::InvariantCulture, $style, [ref]$instant)) { throw 'Publication proof UTC timestamp is invalid.' }
    return $instant
}
foreach ($relative in @('WinImgNormalizer.ps1', 'docs/release/NOTES.md')) { Assert-ReleaseRegularFile (Join-Path $repository $relative) }
. (Join-Path $repository 'WinImgNormalizer.ps1')
$applicationVersion = Get-WinImgVersion
if ($applicationVersion -notmatch '^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)$') {
    throw 'Application version must be a three-part semantic version.'
}
$notes = [IO.File]::ReadAllText((Join-Path $repository 'docs/release/NOTES.md')).Replace("`r`n", "`n").TrimEnd()
$metadata = [ordered]@{
    schema_version = 1
    product = 'WinImgNormalizer'
    version = $applicationVersion
    version_source = 'WinImgNormalizer.ps1:Get-WinImgVersion'
    release_state = 'unreleased'
    proposed_tag = 'v' + $applicationVersion
    tag = $null
    release_date = $null
    release_url = $null
    asset_filename = $null
    asset_bytes = $null
    asset_sha256 = $null
    download_url = $null
    license = 'MIT'
    author = 'Janne Vuorela'
    repository_url = 'https://github.com/PikkuJanne/WinImgNormalizer'
    release_notes_file = 'CHANGELOG.md'
    requirements_file = 'README.md'
}
$existingPath = Join-Path $repository 'release-metadata.json'
Assert-ReleaseRegularFile $existingPath -AllowMissing
$headingState = 'unreleased'
if ([IO.File]::Exists($existingPath)) {
    $existing = Read-ReleaseJson $existingPath
    if ($existing.release_state -cnotin @('unreleased', 'published')) { throw 'Release metadata state must be unreleased or published.' }
    if ($existing.release_state -ceq 'published') {
        Assert-ReleaseFields $existing @($metadata.Keys)
        # This offline witness describes the approved asset bytes, not a rebuilt
        # package or permission to publish. A proof alone never promotes a draft.
        $proofPath = Join-Path $repository ('docs/release/publication-v' + $applicationVersion + '.json')
        $proof = Read-ReleaseJson $proofPath
        Assert-ReleaseFields $proof @('schema_version', 'product', 'version', 'repository_url', 'visibility', 'tag', 'source_revision', 'release_id', 'release_url', 'published_at', 'observed_at', 'assets')
        foreach ($field in @('product', 'version', 'repository_url', 'visibility', 'tag', 'source_revision', 'release_url')) {
            if ($proof.$field -isnot [string]) { throw 'Publication proof identity must use scalar strings.' }
        }
        $releaseUrl = $metadata.repository_url + '/releases/tag/' + $metadata.proposed_tag
        if (-not (Test-ReleaseInteger $proof.schema_version) -or $proof.schema_version -ne 1 -or
            $proof.product -cne $metadata.product -or $proof.version -cne $applicationVersion -or $applicationVersion -cne '1.0.0' -or
            $proof.repository_url -cne $metadata.repository_url -or $proof.visibility -cne 'public' -or
            $proof.tag -cne $metadata.proposed_tag -or $proof.release_url -cne $releaseUrl -or
            $proof.source_revision -cne '8edbcbaeb3425ec3a52eeafde212c32553755af1' -or
            -not (Test-ReleaseInteger $proof.release_id) -or $proof.release_id -le 0) { throw 'Publication proof differs from the approved public release identity.' }
        $published = Get-ReleaseUtcInstant $proof.published_at
        $observed = Get-ReleaseUtcInstant $proof.observed_at
        if ($observed -lt $published) { throw 'Publication proof observation precedes publication.' }
        $approvedAssets = @(
            @{ name = 'WinImgNormalizer-1.0.0-portable.zip'; bytes = 184659; sha256 = '251828028e144759c919645f423fde08641cabf42e04db7024d7ad17da4ba14d' },
            @{ name = 'build-provenance.json'; bytes = 763; sha256 = '901598504d3bb60cdfa4a236873a07d3a7287cb8ff6b3c715b202b39865cefae' },
            @{ name = 'SHA256SUMS.txt'; bytes = 190; sha256 = '2257dbcd00707a93f3e811c9d3ff0eafb59a2e65bb756c440e152f1642540497' }
        )
        if ($proof.assets -isnot [Array] -or $proof.assets.Count -ne 3) { throw 'Publication proof must contain the exact three ordered release assets.' }
        $ids = New-Object 'Collections.Generic.HashSet[long]'
        for ($index = 0; $index -lt $approvedAssets.Count; $index++) {
            $asset = $proof.assets[$index]; $approved = $approvedAssets[$index]
            Assert-ReleaseFields $asset @('id', 'name', 'bytes', 'sha256', 'download_url', 'downloaded_bytes', 'downloaded_sha256')
            foreach ($field in @('name', 'sha256', 'download_url', 'downloaded_sha256')) {
                if ($asset.$field -isnot [string]) { throw 'Publication proof asset identity must use scalar strings.' }
            }
            if (-not (Test-ReleaseInteger $asset.id) -or $asset.id -le 0 -or -not $ids.Add([long]$asset.id) -or
                $asset.name -cne $approved.name -or -not (Test-ReleaseInteger $asset.bytes) -or $asset.bytes -ne $approved.bytes -or
                $asset.sha256 -cne $approved.sha256 -or -not (Test-ReleaseInteger $asset.downloaded_bytes) -or $asset.downloaded_bytes -ne $asset.bytes -or
                $asset.downloaded_sha256 -cne $asset.sha256 -or
                $asset.download_url -cne ($metadata.repository_url + '/releases/download/' + $metadata.proposed_tag + '/' + $approved.name)) { throw 'Publication proof asset differs from the approved downloaded bytes or identity.' }
        }
        $zip = $proof.assets[0]
        $metadata.release_state = 'published'
        $metadata.tag = $proof.tag
        $metadata.release_date = $published.UtcDateTime.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
        $metadata.release_url = $proof.release_url
        $metadata.asset_filename = $zip.name
        $metadata.asset_bytes = $zip.bytes
        $metadata.asset_sha256 = $zip.sha256
        $metadata.download_url = $zip.download_url
        foreach ($name in $metadata.Keys) {
            if (($metadata[$name] -is [string] -and $existing.$name -isnot [string]) -or $existing.$name -cne $metadata[$name]) { throw ('Published release metadata differs from its source/proof: ' + $name) }
        }
        if (-not (Test-ReleaseInteger $existing.schema_version) -or -not (Test-ReleaseInteger $existing.asset_bytes)) { throw 'Published release metadata integer types differ from its schema.' }
        $headingState = $metadata.release_date
    }
}
$changelog = "# Changelog`n`n<!-- Generated by tools/release/Update-ReleaseMetadata.ps1; edit docs/release/NOTES.md. -->`n`n## $applicationVersion ($headingState)`n`n$notes`n"
$metadataText = ($metadata | ConvertTo-Json -Depth 4 -Compress) + "`n"
$outputs = [ordered]@{ 'CHANGELOG.md' = $changelog; 'release-metadata.json' = $metadataText }
$encoding = [Text.UTF8Encoding]::new($false)
foreach ($name in $outputs.Keys) {
    $path = Join-Path $repository $name
    Assert-ReleaseRegularFile $path -AllowMissing
    if ($Check) {
        if (-not [IO.File]::Exists($path) -or
            [Convert]::ToBase64String([IO.File]::ReadAllBytes($path)) -cne [Convert]::ToBase64String($encoding.GetBytes($outputs[$name]))) {
            throw ('Generated release metadata is stale or missing: ' + $name + '. Run tools/release/Update-ReleaseMetadata.ps1.')
        }
    } elseif (-not [IO.File]::Exists($path) -or [Convert]::ToBase64String([IO.File]::ReadAllBytes($path)) -cne [Convert]::ToBase64String($encoding.GetBytes($outputs[$name]))) {
        [IO.File]::WriteAllText($path, $outputs[$name], $encoding)
    }
}
Write-Output ('Release metadata {0}: WinImgNormalizer {1} ({2})' -f $(if ($Check) { 'verified' } else { 'generated' }), $applicationVersion, $headingState)
