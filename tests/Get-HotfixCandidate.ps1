# Build-only acquisition. The workflow gives only this cache step a read token.
[CmdletBinding()]
param([Parameter(Mandatory)][ValidatePattern('^v\d+\.\d+\.\d+$')][string]$Tag)
$ErrorActionPreference = 'Stop'
$repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$output = Join-Path $repository '.scratch/ci-candidate'
if (Test-Path -LiteralPath $output) { throw 'Candidate cache already exists; no replacement.' }
$releaseText = @(& gh release view $Tag --repo PikkuJanne/WinImgNormalizer --json tagName,isDraft,isPrerelease,targetCommitish,assets,url)
if ($LASTEXITCODE -ne 0) { throw 'Cannot read the authorized draft candidate release.' }
$release = ($releaseText -join "`n") | ConvertFrom-Json
if ($release.isPrerelease -or $release.tagName -ne $Tag) { throw 'Only the selected stable release or its unpublished draft can be verified.' }
$latest = $false
if (-not $release.isDraft) {
    $latestText = @(& gh api repos/PikkuJanne/WinImgNormalizer/releases/latest)
    if ($LASTEXITCODE -ne 0) { throw 'Cannot read back public latest-release state.' }
    $latest = (($latestText -join "`n") | ConvertFrom-Json).tag_name -eq $Tag
    if (-not $latest) { throw 'Published candidate is not the public latest stable release.' }
}
$version = $Tag.Substring(1)
$name = 'WinImgNormalizer-' + $version + '-portable.zip'
$asset = @($release.assets | Where-Object name -eq $name)
if ($asset.Count -ne 1) { throw 'Draft candidate portable asset is absent or ambiguous.' }
[IO.Directory]::CreateDirectory($output) | Out-Null
foreach ($pattern in @($name, 'SHA256SUMS.txt', 'build-provenance.json')) {
    & gh release download $Tag --repo PikkuJanne/WinImgNormalizer --pattern $pattern --dir $output
    if ($LASTEXITCODE -ne 0) { throw 'Draft candidate asset download failed.' }
}
$zip = Join-Path $output $name
$digest = (Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash.ToLowerInvariant()
$provenance = Get-Content -LiteralPath (Join-Path $output 'build-provenance.json') -Raw | ConvertFrom-Json
$source = (& git -C $repository rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0 -or $source -notmatch '^[0-9a-f]{40}$') { throw 'Candidate checkout source could not be established.' }
$sums = [IO.File]::ReadAllText((Join-Path $output 'SHA256SUMS.txt'))
if ($provenance.source_revision -ne $source -or $provenance.version -ne $version -or $provenance.asset_filename -ne $name -or
    $provenance.asset_sha256 -ne $digest -or $asset[0].digest -ne ('sha256:' + $digest) -or
    $sums -notmatch ('(?m)^' + $digest + '  ' + [regex]::Escape($name) + '\r?$')) { throw 'Draft candidate bytes/checksum/provenance do not match the exact source checkout.' }
$provenanceDigest = (Get-FileHash -LiteralPath (Join-Path $output 'build-provenance.json')).Hash.ToLowerInvariant()
if ($sums -notmatch ('(?m)^' + $provenanceDigest + '  build-provenance\.json\r?$')) { throw 'Draft provenance checksum differs.' }
$record = [ordered]@{ Tag = $Tag; Version = $version; SourceRevision = $source; AssetFilename = $name; Sha256 = $digest; Draft = [bool]$release.isDraft; Stable = (-not $release.isDraft); Latest = $latest; ReleaseUrl = $release.url }
[IO.File]::WriteAllText((Join-Path $output 'candidate.json'), ($record | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
Write-Output ('Verified exact release bytes: ' + $name + ' SHA256 ' + $digest + '; draft=' + $release.isDraft + '; latest=' + $latest)
