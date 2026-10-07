# Read-only handoff consistency check; never publishes, downloads or renders a website.
[CmdletBinding()]
param(
    [string]$RepositoryRoot,
    [string]$MetadataPath
)

$ErrorActionPreference = 'Stop'

function Assert-WebsiteRegularPath {
    param([string]$Path, [string]$Root, [switch]$Directory)
    $resolved = [IO.Path]::GetFullPath($Path)
    if ($resolved -cne $Root -and
        -not $resolved.StartsWith($Root + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Website inputs must remain inside the selected repository.'
    }
    $item = Get-Item -LiteralPath $resolved -Force
    if ($Directory) {
        if ($item -isnot [IO.DirectoryInfo]) { throw 'The repository must be an ordinary directory.' }
    } elseif ($item -isnot [IO.FileInfo]) { throw 'Website inputs must be ordinary files.' }
    $ancestor = $item
    while ($null -ne $ancestor) {
        if (($ancestor.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw 'Website inputs must not traverse reparse points.'
        }
        if ($ancestor -is [IO.FileInfo]) { $ancestor = $ancestor.Directory }
        else { $ancestor = $ancestor.Parent }
    }
    return $resolved
}

function Assert-WebsiteKeys {
    param([object]$Value, [string[]]$Keys, [string]$Label)
    if ($Value -isnot [pscustomobject]) { throw ($Label + ' must be a JSON object.') }
    $actual = @($Value.PSObject.Properties.Name)
    if ($actual.Count -ne $Keys.Count) { throw ($Label + ' has missing or unknown fields.') }
    foreach ($key in $Keys) {
        if ($actual -cnotcontains $key) { throw ($Label + ' has missing or incorrectly cased fields.') }
    }
}

function Assert-WebsiteValue {
    param([object]$Value, [object]$Expected, [string]$Label)
    if ($null -eq $Expected) {
        if ($null -ne $Value) { throw ($Label + ' must be null in this handoff state.') }
    } elseif ($Expected -is [bool]) {
        if ($Value -isnot [bool] -or $Value -ne $Expected) { throw ($Label + ' must be the required JSON boolean.') }
    } elseif ($Expected -is [string]) {
        if ($Value -isnot [string] -or $Value -cne $Expected) { throw ($Label + ' does not match the reviewed value.') }
    } else {
        if (($Value -isnot [int] -and $Value -isnot [long]) -or $Value -ne $Expected) {
            throw ($Label + ' must be the required JSON integer.')
        }
    }
}

function Assert-WebsiteArray {
    param([object]$Value, [object[]]$Expected, [string]$Label)
    if ($Value -isnot [array] -or $Value.Count -ne $Expected.Count) { throw ($Label + ' must be the reviewed JSON array.') }
    for ($index = 0; $index -lt $Expected.Count; $index++) {
        Assert-WebsiteValue $Value[$index] $Expected[$index] $Label
    }
}

function Assert-WebsiteJsonObjectKeys {
    param([string]$Text, [hashtable]$Conversion)
    # JSON readers differ on first/last duplicate values. Track decoded names in
    # each individual object, so repeated asset field names remain valid while
    # duplicate or differently escaped spellings of one key cannot be accepted.
    $tokens = [regex]::Matches($Text, '"(?:[^"\\]|\\.)*"|[{}\[\]:,]|-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?|true|false|null')
    $scopes = New-Object 'Collections.Generic.Stack[object]'
    $end = 0
    for ($index = 0; $index -lt $tokens.Count; $index++) {
        $token = $tokens[$index]
        if ($Text.Substring($end, $token.Index - $end) -notmatch '^\s*$') { throw 'Website inputs must use ordinary JSON tokens.' }
        $end = $token.Index + $token.Length
        if ($token.Value -ceq '{') {
            $scopes.Push([pscustomobject]@{ kind='object'; keys=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase) })
        } elseif ($token.Value -ceq '[') {
            $scopes.Push([pscustomobject]@{ kind='array'; keys=$null })
        } elseif ($token.Value -cin @('}',']')) {
            if ($scopes.Count -eq 0) { throw 'Website JSON scope is invalid.' }
            $null = $scopes.Pop()
        } elseif ($token.Value.StartsWith('"', [StringComparison]::Ordinal) -and
                $index + 1 -lt $tokens.Count -and $tokens[$index + 1].Value -ceq ':') {
            if ($scopes.Count -eq 0 -or $scopes.Peek().kind -cne 'object') { throw 'Website JSON key scope is invalid.' }
            $name = ConvertFrom-Json -InputObject $token.Value @Conversion
            if (-not $scopes.Peek().keys.Add([string]$name)) { throw 'Website JSON object keys must be unique.' }
        }
    }
    if ($Text.Substring($end) -notmatch '^\s*$') { throw 'Website inputs must use ordinary JSON tokens.' }
}

function Read-WebsiteJson {
    param([string]$Path)
    if ((Get-Item -LiteralPath $Path).Length -gt 65536) { throw 'Website JSON inputs must be bounded to 64 KiB.' }
    # Newer PowerShell parses ISO date strings into DateTime by default. Preserve
    # the recorded JSON types so both maintained hosts validate the same schema.
    $conversion = @{ ErrorAction='Stop' }
    if ((Get-Command ConvertFrom-Json).Parameters.ContainsKey('DateKind')) { $conversion.DateKind = 'String' }
    $text = [IO.File]::ReadAllText($Path)
    Assert-WebsiteJsonObjectKeys $text $conversion
    return ($text | ConvertFrom-Json @conversion)
}

function Assert-WebsitePositiveInteger {
    param([object]$Value, [string]$Label)
    if (($Value -isnot [int] -and $Value -isnot [long]) -or $Value -le 0) {
        throw ($Label + ' must be a positive JSON integer.')
    }
}

function ConvertFrom-WebsiteUtcTimestamp {
    param([object]$Value, [string]$Label)
    if ($Value -isnot [string] -or $Value -cnotmatch '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,7})?Z$') {
        throw ($Label + ' must be an explicit UTC timestamp.')
    }
    $parsed = [DateTimeOffset]::MinValue
    if (-not [DateTimeOffset]::TryParse($Value, [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::None, [ref]$parsed)) {
        throw ($Label + ' must be a real UTC timestamp.')
    }
    return $parsed
}

function Assert-WebsitePublishedProof {
    param([object]$Release, [string]$Root)
    # This is offline consistency against the recorded publication observation.
    # It neither witnesses a fresh public download nor establishes a signature,
    # owner consent or website deployment approval.
    Assert-WebsiteValue $Release.version '1.0.0' 'Reviewed published version'
    $proofPath = Assert-WebsiteRegularPath (Join-Path $Root ('docs/release/publication-v' + $Release.version + '.json')) $Root
    $proof = Read-WebsiteJson $proofPath
    Assert-WebsiteKeys $proof @('schema_version','product','version','repository_url','visibility','tag','source_revision','release_id','release_url','published_at','observed_at','assets') 'Publication observation'
    Assert-WebsiteValue $proof.schema_version 1 'Publication schema version'
    Assert-WebsiteValue $proof.product 'WinImgNormalizer' 'Published product'
    Assert-WebsiteValue $proof.version '1.0.0' 'Published version'
    Assert-WebsiteValue $proof.repository_url 'https://github.com/PikkuJanne/WinImgNormalizer' 'Published repository'
    Assert-WebsiteValue $proof.visibility 'public' 'Release visibility'
    Assert-WebsiteValue $proof.tag 'v1.0.0' 'Published tag'
    Assert-WebsiteValue $proof.source_revision '8edbcbaeb3425ec3a52eeafde212c32553755af1' 'Published source revision'
    Assert-WebsitePositiveInteger $proof.release_id 'Release ID'
    Assert-WebsiteValue $proof.release_url ($proof.repository_url + '/releases/tag/' + $proof.tag) 'Published release URL'
    $publishedAt = ConvertFrom-WebsiteUtcTimestamp $proof.published_at 'Publication time'
    $observedAt = ConvertFrom-WebsiteUtcTimestamp $proof.observed_at 'Publication observation time'
    if ($observedAt -lt $publishedAt) { throw 'The publication observation cannot precede publication.' }
    $expectedAssets = @(
        @{ name='WinImgNormalizer-1.0.0-portable.zip'; bytes=184659; sha256='251828028e144759c919645f423fde08641cabf42e04db7024d7ad17da4ba14d' },
        @{ name='build-provenance.json'; bytes=763; sha256='901598504d3bb60cdfa4a236873a07d3a7287cb8ff6b3c715b202b39865cefae' },
        @{ name='SHA256SUMS.txt'; bytes=190; sha256='2257dbcd00707a93f3e811c9d3ff0eafb59a2e65bb756c440e152f1642540497' }
    )
    if ($proof.assets -isnot [array] -or $proof.assets.Count -ne $expectedAssets.Count) {
        throw 'The publication observation must bind exactly the three reviewed assets.'
    }
    $assetIds = @()
    for ($index = 0; $index -lt $expectedAssets.Count; $index++) {
        $asset = $proof.assets[$index]
        $expected = $expectedAssets[$index]
        Assert-WebsiteKeys $asset @('id','name','bytes','sha256','download_url','downloaded_bytes','downloaded_sha256') 'Published asset'
        Assert-WebsitePositiveInteger $asset.id 'Published asset ID'
        if ($assetIds -contains $asset.id) { throw 'Published asset IDs must be distinct.' }
        $assetIds += $asset.id
        foreach ($field in @('name','bytes','sha256')) { Assert-WebsiteValue $asset.$field $expected[$field] ('Published asset ' + $field) }
        Assert-WebsiteValue $asset.download_url ($proof.repository_url + '/releases/download/' + $proof.tag + '/' + $asset.name) 'Published download URL'
        Assert-WebsiteValue $asset.downloaded_bytes $expected.bytes 'Observed public download bytes'
        Assert-WebsiteValue $asset.downloaded_sha256 $expected.sha256 'Observed public download SHA-256'
    }
    Assert-WebsiteValue $Release.product $proof.product 'Canonical published product'
    Assert-WebsiteValue $Release.repository_url $proof.repository_url 'Canonical published repository'
    Assert-WebsiteValue $Release.tag $proof.tag 'Canonical published tag'
    Assert-WebsiteValue $Release.release_date $proof.published_at.Substring(0,10) 'Canonical publication date'
    Assert-WebsiteValue $Release.release_url $proof.release_url 'Canonical published release URL'
    Assert-WebsiteValue $Release.asset_filename $proof.assets[0].name 'Canonical download filename'
    Assert-WebsiteValue $Release.asset_bytes $proof.assets[0].bytes 'Canonical download bytes'
    Assert-WebsiteValue $Release.asset_sha256 $proof.assets[0].sha256 'Canonical download SHA-256'
    Assert-WebsiteValue $Release.download_url $proof.assets[0].download_url 'Canonical download URL'
    return $proof
}

function Assert-WebsiteDocumentLinks {
    param([string]$Path, [string]$Root, [string[]]$PublishedLinks = @())
    $text = [IO.File]::ReadAllText($Path)
    if ($text -match '(?m)^\s{0,3}\[[^\]]+\]:' -or $text -match '<[^>\r\n]+>') {
        throw 'Website documents require inline Markdown links without reference definitions, HTML or autolinks.'
    }
    if ($text -match '(?i)\bwww\.|[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}') {
        throw 'Website documents must not add bare-domain or email autolink destinations.'
    }
    $allowedExternal = @('https://github.com/PikkuJanne/WinImgNormalizer','https://imagemagick.org/download/#windows-binary-release') + $PublishedLinks
    # Some Markdown renderers turn bare URLs into links. Check every HTTP(S)
    # occurrence, not only conventional inline links, before page integration.
    foreach ($uriMatch in [regex]::Matches($text, 'https?://[^\s<>()"`]+', [Text.RegularExpressions.RegexOptions]::IgnoreCase)) {
        $uriText = $uriMatch.Value.TrimEnd('.', ',', ';')
        if ($uriText -cnotin $allowedExternal) { throw 'Website documents contain an unreviewed HTTP(S) destination.' }
    }
    foreach ($link in [regex]::Matches($text, '!?\[[^\]]+\]\(([^)]+)\)')) {
        $target = $link.Groups[1].Value
        if ($target -match '^https://') {
            # Published additions come only from the exact reviewed observation.
            if ($target -cnotin $allowedExternal) {
                throw 'Website documents contain an unreviewed external link.'
            }
            continue
        }
        if ($target -match '^[a-zA-Z][a-zA-Z0-9+.-]*:|^[/\\]') { throw 'Website links must use HTTPS or relative repository files.' }
        $parts = $target.Split('#', 2)
        $filePath = $Path
        if ($parts[0].Length -gt 0) {
            $relativePath = [Uri]::UnescapeDataString($parts[0])
            if ($relativePath -match '^[a-zA-Z][a-zA-Z0-9+.-]*:|^[/\\]') { throw 'Decoded website links must remain relative repository files.' }
            $filePath = Assert-WebsiteRegularPath (Join-Path (Split-Path -Parent $Path) $relativePath) $Root
        }
        if ($parts.Count -eq 2) {
            $headings = @([regex]::Matches([IO.File]::ReadAllText($filePath), '(?m)^#{1,6}\s+(.+?)\s*#*\s*$') | ForEach-Object {
                $_.Groups[1].Value.ToLowerInvariant() -replace '[^\w -]', '' -replace ' ', '-'
            })
            if ($headings -cnotcontains [Uri]::UnescapeDataString($parts[1])) { throw 'A website document heading link does not resolve.' }
        }
    }
}

try {
    # Windows PowerShell 5.1 can evaluate parameter defaults before PSScriptRoot
    # is populated. Resolve the default inside the script body in both hosts.
    if ([string]::IsNullOrWhiteSpace($RepositoryRoot)) { $RepositoryRoot = Join-Path $PSScriptRoot '../..' }
    $root = [IO.Path]::GetFullPath($RepositoryRoot).TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
    $root = Assert-WebsiteRegularPath $root $root -Directory
    $websiteDirectory = Join-Path $root 'docs/website'
    $defaultMetadata = [IO.Path]::GetFullPath((Join-Path $websiteDirectory 'metadata.json'))
    if ([string]::IsNullOrWhiteSpace($MetadataPath)) { $MetadataPath = $defaultMetadata }
    $MetadataPath = Assert-WebsiteRegularPath $MetadataPath $root
    if ($MetadataPath -cne $defaultMetadata -and
        -not $MetadataPath.StartsWith((Join-Path $root '.scratch') + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Alternate website metadata must be an owned repository .scratch file.'
    }
    $metadata = Read-WebsiteJson $MetadataPath
    Assert-WebsiteKeys $metadata @('schema_version','preparation_state','release_metadata_file','product','repository_url','license','platform','prerequisites','processing','deployment','download','documents','assets','screenshots','comparisons') 'Website metadata'
    Assert-WebsiteValue $metadata.schema_version 1 'Schema version'
    if ($metadata.preparation_state -isnot [string] -or $metadata.preparation_state -cnotin @('draft','published')) {
        throw 'Preparation state must be draft or published.'
    }
    Assert-WebsiteValue $metadata.release_metadata_file '../../release-metadata.json' 'Release metadata reference'
    $releasePath = Assert-WebsiteRegularPath (Join-Path $websiteDirectory $metadata.release_metadata_file) $root
    $release = Read-WebsiteJson $releasePath
    Assert-WebsiteKeys $release @('schema_version','product','version','version_source','release_state','proposed_tag','tag','release_date','release_url','asset_filename','asset_bytes','asset_sha256','download_url','license','author','repository_url','release_notes_file','requirements_file') 'Canonical release metadata'
    Assert-WebsiteValue $release.schema_version 1 'Release schema version'
    Assert-WebsiteValue $release.version_source 'WinImgNormalizer.ps1:Get-WinImgVersion' 'Version source'
    if ($release.version -isnot [string] -or $release.version -notmatch '^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)$') { throw 'Canonical application version must be semantic.' }
    Assert-WebsiteValue $release.proposed_tag ('v' + $release.version) 'Proposed tag'
    $publishedLinks = @()
    if ($metadata.preparation_state -ceq 'draft') {
        Assert-WebsiteValue $release.release_state 'unreleased' 'Release state'
        foreach ($field in @('tag','release_date','release_url','asset_filename','asset_bytes','asset_sha256','download_url')) {
            Assert-WebsiteValue $release.$field $null ('Canonical ' + $field)
        }
    } else {
        Assert-WebsiteValue $release.release_state 'published' 'Release state'
        $publication = Assert-WebsitePublishedProof $release $root
        $publishedLinks = @($publication.release_url) + @($publication.assets | ForEach-Object { $_.download_url })
    }
    # This reviewed helper checks the canonical projection against the real single
    # version source and release notes without replacing any bytes.
    $releaseChecker = Assert-WebsiteRegularPath (Join-Path $root 'tools/release/Update-ReleaseMetadata.ps1') $root
    & $releaseChecker -Check | Out-Null

    Assert-WebsiteKeys $metadata.product @('name','short_description') 'Product'
    Assert-WebsiteValue $metadata.product.name $release.product 'Product name'
    if ($metadata.product.short_description -isnot [string] -or [string]::IsNullOrWhiteSpace($metadata.product.short_description) -or
        $metadata.product.short_description.Length -gt 200 -or $metadata.product.short_description -match '[<>\r\n]') {
        throw 'Product description must be short plain text.'
    }
    Assert-WebsiteValue $metadata.repository_url $release.repository_url 'Repository URL'
    Assert-WebsiteValue $metadata.license $release.license 'License'
    Assert-WebsiteKeys $metadata.platform @('supported_os','powershell','tested_desktop','tested_ci','untested_os') 'Platform'
    Assert-WebsiteArray $metadata.platform.supported_os @('Windows 10 x64','Windows 11 x64') 'Supported platforms'
    Assert-WebsiteArray $metadata.platform.powershell @('Windows PowerShell 5.1','PowerShell 7') 'PowerShell hosts'
    Assert-WebsiteArray $metadata.platform.tested_desktop @('Windows 11 Pro build 26300') 'Tested desktop'
    Assert-WebsiteArray $metadata.platform.tested_ci @('Windows Server 2025') 'Tested CI'
    Assert-WebsiteArray $metadata.platform.untested_os @('Windows 10') 'Untested platform'
    Assert-WebsiteKeys $metadata.prerequisites @('imagemagick_minimum','supported_major','external_installation','format_support_depends_on_build','details_file') 'Prerequisites'
    Assert-WebsiteValue $metadata.prerequisites.imagemagick_minimum '7.1.2-32' 'ImageMagick minimum'
    Assert-WebsiteValue $metadata.prerequisites.supported_major 7 'ImageMagick major'
    Assert-WebsiteValue $metadata.prerequisites.external_installation $true 'External dependency installation'
    Assert-WebsiteValue $metadata.prerequisites.format_support_depends_on_build $true 'Codec qualification'
    Assert-WebsiteValue $metadata.prerequisites.details_file '../../README.md' 'Prerequisites reference'

    Assert-WebsiteKeys $metadata.processing @('location','uploads','processing_api','server_runtime','runtime_account_required','telemetry') 'Processing boundary'
    Assert-WebsiteValue $metadata.processing.location 'local_windows' 'Processing location'
    foreach ($field in @('uploads','processing_api','server_runtime','runtime_account_required','telemetry')) {
        Assert-WebsiteValue $metadata.processing.$field $false ('Processing ' + $field)
    }
    Assert-WebsiteKeys $metadata.deployment @('domain','framework','hosting_provider') 'Deployment context'
    foreach ($field in @('domain','framework','hosting_provider')) { Assert-WebsiteValue $metadata.deployment.$field $null ('Deployment ' + $field) }
    Assert-WebsiteKeys $metadata.download @('state','render_link','label') 'Download presentation'
    if ($metadata.preparation_state -ceq 'draft') {
        Assert-WebsiteValue $metadata.download.state 'unavailable' 'Download state'
        Assert-WebsiteValue $metadata.download.render_link $false 'Download link visibility'
        Assert-WebsiteValue $metadata.download.label 'Release download not available' 'Download label'
    } else {
        Assert-WebsiteValue $metadata.download.state 'available' 'Download state'
        Assert-WebsiteValue $metadata.download.render_link $true 'Download link visibility'
        Assert-WebsiteValue $metadata.download.label ('Download WinImgNormalizer ' + $release.version) 'Download label'
    }
    Assert-WebsiteArray $metadata.screenshots @() 'Screenshots'
    Assert-WebsiteArray $metadata.comparisons @() 'Comparisons'

    $documentPaths = [ordered]@{ product_copy='PRODUCT_COPY.md'; integration='INTEGRATION.md'; quick_start='../../README.md'; behavior='../BEHAVIOR.md'; security='../../SECURITY.md'; license='../../LICENSE'; release_notes='../../CHANGELOG.md' }
    Assert-WebsiteKeys $metadata.documents @($documentPaths.Keys) 'Documents'
    foreach ($name in $documentPaths.Keys) {
        Assert-WebsiteValue $metadata.documents.$name $documentPaths[$name] ('Document ' + $name)
        $document = Assert-WebsiteRegularPath (Join-Path $websiteDirectory $metadata.documents.$name) $root
        if ($name -in @('product_copy','integration')) { Assert-WebsiteDocumentLinks $document $root $publishedLinks }
    }

    $assets = @(
        @{ id='icon'; path='../../WinImgNormalizer_icon_variant.png'; media_type='image/png'; bytes=11141; sha256='479f819325df22004e38f42183ff1d1e2351f753abed935a59357fed7da750f9'; width=1024; height=1024; sizes=@(); usage='branding'; display=$true; alt='WinImgNormalizer image icon' },
        @{ id='favicon'; path='../../WinImgNormalizer_variant.ico'; media_type='image/vnd.microsoft.icon'; bytes=22917; sha256='78ef6c01d0b466e72e3bad4fe83bb12b57a74c2f245a1274a559857d35374e79'; width=$null; height=$null; sizes=@(16,32,48,64,128,256); usage='branding'; display=$true; alt='WinImgNormalizer icon' },
        @{ id='legacy_poster'; path='../../WinImgNormalizer_poster.png'; media_type='image/png'; bytes=2397241; sha256='99b1cd73450e725a5cce26a00ecdb6e660166ff8b5b9f75896882cfe7825d926'; width=1536; height=1024; sizes=@(); usage='historical_reference_only'; display=$false; alt='Historical WinImgNormalizer poster artwork; not a screenshot or current size guarantee' }
    )
    if ($metadata.assets -isnot [array] -or $metadata.assets.Count -ne $assets.Count) { throw 'Only the three reviewed original branding assets may be catalogued.' }
    for ($index = 0; $index -lt $assets.Count; $index++) {
        $asset = $metadata.assets[$index]
        $expected = $assets[$index]
        Assert-WebsiteKeys $asset @($expected.Keys) 'Branding asset'
        foreach ($field in $expected.Keys) {
            if ($field -eq 'sizes') { Assert-WebsiteArray $asset.$field $expected[$field] 'Icon sizes' }
            else { Assert-WebsiteValue $asset.$field $expected[$field] ('Asset ' + $field) }
        }
        $assetPath = Assert-WebsiteRegularPath (Join-Path $websiteDirectory $asset.path) $root
        if ((Get-Item -LiteralPath $assetPath).Length -ne $asset.bytes -or
            (Get-FileHash -LiteralPath $assetPath -Algorithm SHA256).Hash.ToLowerInvariant() -cne $asset.sha256) {
            throw 'Original branding bytes do not match the reviewed asset.'
        }
        $bytes = [IO.File]::ReadAllBytes($assetPath)
        if ($asset.media_type -eq 'image/png') {
            if ([BitConverter]::ToString($bytes[0..7]) -cne '89-50-4E-47-0D-0A-1A-0A' -or [Text.Encoding]::ASCII.GetString($bytes,12,4) -cne 'IHDR') { throw 'PNG branding header is invalid.' }
            $width = [uint32]$bytes[16] * 16777216 + [uint32]$bytes[17] * 65536 + [uint32]$bytes[18] * 256 + [uint32]$bytes[19]
            $height = [uint32]$bytes[20] * 16777216 + [uint32]$bytes[21] * 65536 + [uint32]$bytes[22] * 256 + [uint32]$bytes[23]
            if ($width -ne $asset.width -or $height -ne $asset.height) { throw 'PNG branding dimensions differ.' }
        } else {
            if ([BitConverter]::ToString($bytes[0..3]) -cne '00-00-01-00' -or [BitConverter]::ToUInt16($bytes,4) -ne $asset.sizes.Count) { throw 'ICO branding header is invalid.' }
            for ($entry = 0; $entry -lt $asset.sizes.Count; $entry++) {
                $width = [int]$bytes[6 + 16 * $entry]
                $height = [int]$bytes[7 + 16 * $entry]
                if ($width -eq 0) { $width = 256 }
                if ($height -eq 0) { $height = 256 }
                if ($width -ne $asset.sizes[$entry] -or $height -ne $asset.sizes[$entry]) { throw 'ICO branding sizes differ.' }
            }
        }
    }
    if ($metadata.preparation_state -ceq 'draft') {
        Write-Output ('WEBSITE HANDOFF VALID: draft content for {0} {1}; downloads unavailable; 3 original assets; no publication.' -f $release.product,$release.version)
    } else {
        Write-Output ('WEBSITE HANDOFF VALID: published record for {0} {1}; exact recorded download identity; 3 original assets; no deployment.' -f $release.product,$release.version)
    }
    exit 0
} catch {
    Write-Output ('WEBSITE HANDOFF INVALID: ' + $_.Exception.Message)
    exit 1
}
