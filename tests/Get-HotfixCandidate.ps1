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
if ($release.tagName -ne $Tag) { throw 'Only the explicitly selected release can be verified.' }
$latest = $false
$publicTagSource = $null
if (-not $release.isDraft) {
    # Public verification deliberately uses no token, authorization header,
    # default Windows credentials or authenticated gh download route.
    $anonymousRelease = Invoke-RestMethod -Uri ('https://api.github.com/repos/PikkuJanne/WinImgNormalizer/releases/tags/' + $Tag)
    $anonymousLatest = Invoke-RestMethod -Uri 'https://api.github.com/repos/PikkuJanne/WinImgNormalizer/releases/latest'
    if ($anonymousRelease.draft -or $anonymousRelease.tag_name -ne $Tag -or [bool]$anonymousRelease.prerelease -ne [bool]$release.isPrerelease) { throw 'Anonymous release readback differs from the explicitly selected public release state.' }
    $latest = $anonymousLatest.tag_name -eq $Tag
    if (-not $release.isPrerelease -and -not $latest) { throw 'Published stable candidate is not the public latest release.' }
    $release = [pscustomobject]@{ tagName = $anonymousRelease.tag_name; isDraft = $false; isPrerelease = [bool]$anonymousRelease.prerelease; assets = $anonymousRelease.assets; url = $anonymousRelease.html_url }
    $tagReference = Invoke-RestMethod -Uri ('https://api.github.com/repos/PikkuJanne/WinImgNormalizer/git/ref/tags/' + $Tag)
    $tagObject = $tagReference.object
    if ($tagObject.type -eq 'tag') { $tagObject = (Invoke-RestMethod -Uri $tagObject.url).object }
    if ($tagObject.type -ne 'commit') { throw 'Public release tag does not resolve to a source commit.' }
    $publicTagSource = $tagObject.sha
}
$version = $Tag.Substring(1)
$name = 'WinImgNormalizer-' + $version + '-portable.zip'
$asset = @($release.assets | Where-Object name -eq $name)
if ($asset.Count -ne 1) { throw 'Draft candidate portable asset is absent or ambiguous.' }
[IO.Directory]::CreateDirectory($output) | Out-Null
foreach ($pattern in @($name, 'SHA256SUMS.txt', 'build-provenance.json')) {
    if ($release.isDraft) {
        & gh release download $Tag --repo PikkuJanne/WinImgNormalizer --pattern $pattern --dir $output
        if ($LASTEXITCODE -ne 0) { throw 'Draft candidate asset download failed.' }
    } else {
        $anonymousAsset = @($release.assets | Where-Object name -eq $pattern)
        if ($anonymousAsset.Count -ne 1) { throw 'Anonymous public asset is absent or ambiguous.' }
        Invoke-WebRequest -UseBasicParsing -Uri $anonymousAsset[0].browser_download_url -OutFile (Join-Path $output $pattern)
    }
}
$zip = Join-Path $output $name
$digest = (Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash.ToLowerInvariant()
$provenance = Get-Content -LiteralPath (Join-Path $output 'build-provenance.json') -Raw | ConvertFrom-Json
$source = (& git -C $repository rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0 -or $source -notmatch '^[0-9a-f]{40}$') { throw 'Candidate checkout source could not be established.' }
if ($publicTagSource -and $publicTagSource -ne $source) { throw 'Anonymous public tag differs from the exact tested source checkout.' }
$sums = [IO.File]::ReadAllText((Join-Path $output 'SHA256SUMS.txt'))
if ($provenance.source_revision -ne $source -or $provenance.version -ne $version -or $provenance.asset_filename -ne $name -or
    $provenance.asset_sha256 -ne $digest -or $asset[0].digest -ne ('sha256:' + $digest) -or
    $sums -notmatch ('(?m)^' + $digest + '  ' + [regex]::Escape($name) + '\r?$')) { throw 'Draft candidate bytes/checksum/provenance do not match the exact source checkout.' }
$provenanceDigest = (Get-FileHash -LiteralPath (Join-Path $output 'build-provenance.json')).Hash.ToLowerInvariant()
if ($sums -notmatch ('(?m)^' + $provenanceDigest + '  build-provenance\.json\r?$')) { throw 'Draft provenance checksum differs.' }
$record = [ordered]@{ Tag = $Tag; Version = $version; SourceRevision = $source; AssetFilename = $name; Sha256 = $digest; Draft = [bool]$release.isDraft; Prerelease = [bool]$release.isPrerelease; Stable = (-not $release.isDraft -and -not $release.isPrerelease); Latest = $latest; ReleaseUrl = $release.url; AnonymousPublicDownload = (-not $release.isDraft); PublicTagSourceRevision = $publicTagSource }
[IO.File]::WriteAllText((Join-Path $output 'candidate.json'), ($record | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
Write-Output ('Verified exact release bytes: ' + $name + ' SHA256 ' + $digest + '; draft=' + $release.isDraft + '; prerelease=' + $release.isPrerelease + '; latest=' + $latest)
