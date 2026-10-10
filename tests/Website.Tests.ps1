BeforeAll {
    if ($env:OS -ne 'Windows_NT') { throw 'Website guard tests require actual Windows.' }
    $repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    $guard = Join-Path $repository 'tools/website/Test-WebsiteHandoff.ps1'
    $metadataPath = Join-Path $repository 'docs/website/metadata.json'
    $hostExecutable = (Get-Process -Id $PID).Path
    $scratch = Join-Path $repository '.scratch'
    $current = [IO.DirectoryInfo]::new($scratch)
    while ($current) {
        if ($current.Exists -and ($current.Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Website fixtures cross a reparse point.' }
        $current = $current.Parent
    }
    & git -C $repository check-ignore --quiet --no-index -- (Join-Path $scratch 'website-ignore-probe')
    if ($LASTEXITCODE -ne 0) { throw 'Website fixtures must already be ignored.' }
    $owned = Join-Path $scratch ('M4-T06-website-' + [guid]::NewGuid().ToString('N'))
    if (Test-Path -LiteralPath $owned) { throw 'Website fixture ownership collision.' }
    [IO.Directory]::CreateDirectory($owned) | Out-Null
    [IO.File]::WriteAllText((Join-Path $owned '.winimg-fixture-root'), 'M4-T06 owned synthetic website controls; no release, public download or owner approval is established by these fixtures.', [Text.UTF8Encoding]::new($false))
    $publicPaths = @('tools/website/Test-WebsiteHandoff.ps1', 'docs/website/metadata.json',
        'docs/website/PRODUCT_COPY.md', 'docs/website/INTEGRATION.md', 'WinImgNormalizer_icon_variant.png',
        'WinImgNormalizer_variant.ico', 'WinImgNormalizer.ps1',
        'WinImgNormalizer.bat', 'release-metadata.json', 'README.md', 'CHANGELOG.md', 'LICENSE',
        'SECURITY.md', 'docs/BEHAVIOR.md', 'docs/release/GETTING_STARTED.md', 'docs/release/NOTES.md',
        'docs/release/THIRD_PARTY_NOTICES.md', 'docs/release/PACKAGING.md',
        'tools/release/Update-ReleaseMetadata.ps1', 'tests/Website.Tests.ps1')
    $publicationRelative = 'docs/release/publication-v1.0.0.json'
    if ([IO.File]::Exists((Join-Path $repository $publicationRelative))) { $publicPaths += $publicationRelative }
    $before = @($publicPaths | ForEach-Object {
        [pscustomobject]@{ path = $_; sha256 = (Get-FileHash -LiteralPath (Join-Path $repository $_)).Hash.ToLowerInvariant() }
    }) | ConvertTo-Json -Compress
    $policiesBefore = @(Get-ExecutionPolicy -List | Where-Object Scope -ne Process | ForEach-Object { $_.Scope.ToString() + '=' + $_.ExecutionPolicy.ToString() }) -join '|'
    $observations = New-Object 'Collections.Generic.List[object]'
    $utf8 = [Text.UTF8Encoding]::new($false)
    function Read-WebsiteFixtureJson([string]$Path) {
        $conversion = @{ ErrorAction='Stop' }
        if ((Get-Command ConvertFrom-Json).Parameters.ContainsKey('DateKind')) { $conversion.DateKind = 'String' }
        return [IO.File]::ReadAllText($Path) | ConvertFrom-Json @conversion
    }
    function New-WebsiteMetadata {
        # Reparse the source for every control so one mutated object cannot
        # contaminate a later positive or negative fixture.
        $metadata = [IO.File]::ReadAllText($metadataPath) | ConvertFrom-Json
        $metadata.preparation_state = 'draft'
        $metadata.download.state = 'unavailable'
        $metadata.download.render_link = $false
        $metadata.download.label = 'Release download not available'
        return $metadata
    }
    function Write-WebsiteMetadata([object]$Metadata, [string]$Text) {
        $path = Join-Path (Join-Path $draftRepository '.scratch') ([guid]::NewGuid().ToString('N') + '.json')
        if (-not $PSBoundParameters.ContainsKey('Text')) { $Text = $Metadata | ConvertTo-Json -Depth 16 }
        [IO.File]::WriteAllText($path, $Text, $utf8)
        return $path
    }
    function Invoke-WebsiteGuard([string]$Metadata, [string]$Root = $repository, [switch]$DefaultPaths) {
        $capture = Join-Path $owned ([guid]::NewGuid().ToString('N'))
        $start = New-Object Diagnostics.ProcessStartInfo
        $start.FileName = $hostExecutable
        $tokens = @('-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $guard)
        if (-not $DefaultPaths) { $tokens += @('-RepositoryRoot', $Root) }
        if ($Metadata) { $tokens += @('-MetadataPath', $Metadata) }
        $start.Arguments = ($tokens | ForEach-Object {
            '"' + ([regex]::Replace([regex]::Replace([string]$_, '(\\*)"', '$1$1\"'), '(\\+)$', '$1$1')) + '"'
        }) -join ' '
        # Default resolution must use the script location, not this unrelated
        # working directory. No child executes normalization or personal media.
        $start.WorkingDirectory = $owned
        $start.UseShellExecute = $false; $start.CreateNoWindow = $true
        $start.RedirectStandardOutput = $true; $start.RedirectStandardError = $true
        foreach ($key in @($start.EnvironmentVariables.Keys)) {
            if ([string]$key -ieq 'PSModulePath') { $start.EnvironmentVariables.Remove([string]$key) }
        }
        $process = New-Object Diagnostics.Process
        $process.StartInfo = $start
        try {
            if (-not $process.Start()) { throw 'Website guard child did not start.' }
            $stdout = $process.StandardOutput.ReadToEndAsync(); $stderr = $process.StandardError.ReadToEndAsync()
            if (-not $process.WaitForExit(30000)) { $process.Kill(); throw 'Website guard child exceeded its deadline.' }
            if (-not $stdout.Wait(5000) -or -not $stderr.Wait(5000)) { throw 'Website guard capture incomplete.' }
            $outText = $stdout.GetAwaiter().GetResult(); $errText = $stderr.GetAwaiter().GetResult()
            [IO.File]::WriteAllText(($capture + '.stdout.txt'), $outText, $utf8)
            [IO.File]::WriteAllText(($capture + '.stderr.txt'), $errText, $utf8)
            $result = [pscustomobject]@{ Exit = $process.ExitCode; Text = $outText + $errText }
            $observations.Add([pscustomobject]@{ native_exit_code = $process.ExitCode;
                stdout_sha256 = (Get-FileHash -LiteralPath ($capture + '.stdout.txt')).Hash.ToLowerInvariant();
                stderr_sha256 = (Get-FileHash -LiteralPath ($capture + '.stderr.txt')).Hash.ToLowerInvariant() })
            return $result
        } finally { $process.Dispose() }
    }
    function New-WebsiteRepository {
        $root = Join-Path $owned ('repository-' + [guid]::NewGuid().ToString('N'))
        [IO.Directory]::CreateDirectory($root) | Out-Null
        $fixturePaths = @($publicPaths | Where-Object { $_ -notin @('tests/Website.Tests.ps1', $publicationRelative) }) + @('docs/codex-winimg/RELEASE_AND_WEBSITE.md')
        foreach ($relative in $fixturePaths) {
            $destination = Join-Path $root $relative
            [IO.Directory]::CreateDirectory((Split-Path -Parent $destination)) | Out-Null
            [IO.File]::Copy((Join-Path $repository $relative), $destination)
        }
        [IO.Directory]::CreateDirectory((Join-Path $root '.scratch')) | Out-Null
        # Always construct an explicit unpublished replica. Negative draft
        # controls must still reach their intended guard after the real checkout
        # becomes published; a state mismatch must not make them false positives.
        $releasePath = Join-Path $root 'release-metadata.json'
        $release = [IO.File]::ReadAllText($releasePath) | ConvertFrom-Json
        $release.release_state = 'unreleased'
        foreach ($field in @('tag','release_date','release_url','asset_filename','asset_bytes','asset_sha256','download_url')) { $release.$field = $null }
        [IO.File]::WriteAllText($releasePath, (($release | ConvertTo-Json -Depth 8 -Compress) + "`n"), $utf8)
        [IO.File]::WriteAllText((Join-Path $root 'docs/website/metadata.json'), ((New-WebsiteMetadata | ConvertTo-Json -Depth 16) + "`n"), $utf8)
        $notes = [IO.File]::ReadAllText((Join-Path $root 'docs/release/NOTES.md')).Replace("`r`n", "`n").TrimEnd()
        $changelog = "# Changelog`n`n<!-- Generated by tools/release/Update-ReleaseMetadata.ps1; edit docs/release/NOTES.md. -->`n`n## $($release.version) (unreleased)`n`n$notes`n"
        [IO.File]::WriteAllText((Join-Path $root 'CHANGELOG.md'), $changelog, $utf8)
        # These local-only fixture documents deliberately describe preparation.
        # The actual public documents are checked separately in the live guard.
        foreach ($name in @('PRODUCT_COPY.md','INTEGRATION.md')) {
            [IO.File]::WriteAllText((Join-Path $root ('docs/website/' + $name)), "# Draft website fixture`n`nDownloads unavailable.`n`n[Quick start](../../README.md)`n", $utf8)
        }
        return $root
    }
    function New-PublishedWebsiteRepository {
        $root = New-WebsiteRepository
        # These IDs/times are explicit synthetic offline controls. Only the root
        # publication operation and actual downloaded bytes can establish T075.
        $repositoryUrl = 'https://github.com/PikkuJanne/WinImgNormalizer'
        $assets = @(
            @{ id=101; name='WinImgNormalizer-1.0.0-portable.zip'; bytes=184659; sha256='251828028e144759c919645f423fde08641cabf42e04db7024d7ad17da4ba14d' },
            @{ id=102; name='build-provenance.json'; bytes=763; sha256='901598504d3bb60cdfa4a236873a07d3a7287cb8ff6b3c715b202b39865cefae' },
            @{ id=103; name='SHA256SUMS.txt'; bytes=190; sha256='2257dbcd00707a93f3e811c9d3ff0eafb59a2e65bb756c440e152f1642540497' }
        )
        $proof = [ordered]@{
            schema_version=1; product='WinImgNormalizer'; version='1.0.0'; repository_url=$repositoryUrl;
            visibility='public'; tag='v1.0.0'; source_revision='8edbcbaeb3425ec3a52eeafde212c32553755af1';
            release_id=100; release_url=($repositoryUrl + '/releases/tag/v1.0.0');
            published_at='2026-10-07T00:00:00Z'; observed_at='2026-10-07T00:00:01Z';
            assets=@($assets | ForEach-Object {
                [ordered]@{ id=$_.id; name=$_.name; bytes=$_.bytes; sha256=$_.sha256;
                    download_url=($repositoryUrl + '/releases/download/v1.0.0/' + $_.name);
                    downloaded_bytes=$_.bytes; downloaded_sha256=$_.sha256 }
            })
        }
        [IO.File]::WriteAllText((Join-Path $root $publicationRelative), (($proof | ConvertTo-Json -Depth 12) + "`n"), $utf8)
        $releasePath = Join-Path $root 'release-metadata.json'
        $release = [IO.File]::ReadAllText($releasePath) | ConvertFrom-Json
        $release.release_state = 'published'; $release.tag = $proof.tag; $release.release_date = '2026-10-07'
        $release.release_url = $proof.release_url; $release.asset_filename = $proof.assets[0].name
        $release.asset_bytes = $proof.assets[0].bytes; $release.asset_sha256 = $proof.assets[0].sha256
        $release.download_url = $proof.assets[0].download_url
        [IO.File]::WriteAllText($releasePath, (($release | ConvertTo-Json -Depth 8 -Compress) + "`n"), $utf8)
        $metadata = New-WebsiteMetadata
        $metadata.preparation_state = 'published'; $metadata.download.state = 'available'
        $metadata.download.render_link = $true; $metadata.download.label = 'Download WinImgNormalizer 1.0.0'
        [IO.File]::WriteAllText((Join-Path $root 'docs/website/metadata.json'), (($metadata | ConvertTo-Json -Depth 16) + "`n"), $utf8)
        $notes = [IO.File]::ReadAllText((Join-Path $root 'docs/release/NOTES.md')).Replace("`r`n", "`n").TrimEnd()
        $changelog = "# Changelog`n`n<!-- Generated by tools/release/Update-ReleaseMetadata.ps1; edit docs/release/NOTES.md. -->`n`n## 1.0.0 (2026-10-07)`n`n$notes`n"
        [IO.File]::WriteAllText((Join-Path $root 'CHANGELOG.md'), $changelog, $utf8)
        $links = @('[Release](' + $proof.release_url + ')') + @($proof.assets | ForEach-Object { '[' + $_.name + '](' + $_.download_url + ')' })
        [IO.File]::AppendAllText((Join-Path $root 'docs/website/PRODUCT_COPY.md'), ("`n" + ($links -join "`n") + "`n"), $utf8)
        return $root
    }
    function Assert-WebsiteRejected([string]$Metadata, [string]$Root = $draftRepository, [string]$Reason) {
        $observed = Invoke-WebsiteGuard -Metadata $Metadata -Root $Root
        $observed.Exit | Should -Be 1
        $observed.Text | Should -Match 'WEBSITE HANDOFF INVALID:'
        $observed.Text | Should -Not -Match 'ParserError|ParameterBindingException'
        if ($Reason) { $observed.Text | Should -Match $Reason }
    }
    function Get-WebsiteRepositoryState([string]$Root) {
        return @(Get-ChildItem -LiteralPath $Root -File -Recurse | Sort-Object FullName | ForEach-Object {
            [pscustomobject]@{ relative = $_.FullName.Substring($Root.Length + 1);
                bytes = $_.Length; modified = $_.LastWriteTimeUtc.Ticks;
                sha256 = (Get-FileHash -LiteralPath $_.FullName).Hash.ToLowerInvariant() }
        }) | ConvertTo-Json -Compress
    }
    function Assert-WebsiteDocumentLinks([string]$Path, [string]$Text) {
        foreach ($link in [regex]::Matches($Text, '!?\[[^\]]*\]\((?<target><[^>]+>|[^)\s]+)(?:\s+"[^"]*")?\)')) {
            $target = $link.Groups['target'].Value.Trim('<', '>')
            if ($target -match '^https://') {
                if (-not [Uri]::IsWellFormedUriString($target, [UriKind]::Absolute)) { throw 'Malformed website HTTPS link.' }
                continue
            }
            if ($target -match '^[A-Za-z][A-Za-z0-9+.-]*:|^[\\/]' -or $target -match '[\x00-\x1f]') { throw 'Unsafe website document link.' }
            $parts = $target.Split('#', 2)
            if (-not $parts[0]) { $destination = $Path }
            else { $destination = [IO.Path]::GetFullPath((Join-Path (Split-Path -Parent $Path) ([Uri]::UnescapeDataString($parts[0])))) }
            if (-not $destination.StartsWith($repository + '\', [StringComparison]::OrdinalIgnoreCase) -or
                -not [IO.File]::Exists($destination)) { throw 'Missing or escaped website document link.' }
            if ($parts.Count -eq 2) {
                $headings = @([regex]::Matches([IO.File]::ReadAllText($destination), '(?m)^#{1,6}\s+(.+?)\s*#*\s*$') | ForEach-Object {
                    [regex]::Replace($_.Groups[1].Value.Replace('`', '').Replace('*', '').ToLowerInvariant(), '[^\p{L}\p{N}_\- ]', '').Replace(' ', '-')
                })
                if ([Uri]::UnescapeDataString($parts[1]) -cnotin $headings) { throw 'Missing website document heading.' }
            }
        }
    }
    $draftRepository = New-WebsiteRepository
}

Describe 'T073 framework-neutral website preparation and metadata boundary' {
    It 'accepts the checked-in state and a deliberate draft through native CLI paths' {
        foreach ($result in @((Invoke-WebsiteGuard), (Invoke-WebsiteGuard -DefaultPaths),
            (Invoke-WebsiteGuard -Root $draftRepository -Metadata (Write-WebsiteMetadata (New-WebsiteMetadata))))) {
            $result.Exit | Should -Be 0
            $result.Text | Should -Match 'WEBSITE HANDOFF VALID:'
        }
    }
    It 'keeps the only release version and publication fields in the canonical release record' {
        $metadata = New-WebsiteMetadata
        $release = Get-Content -LiteralPath (Join-Path $draftRepository 'release-metadata.json') -Raw | ConvertFrom-Json
        $metadata.release_metadata_file | Should -Be '../../release-metadata.json'
        $metadata.PSObject.Properties.Name | Should -Not -Contain 'version'
        $metadata.product.PSObject.Properties.Name | Should -Not -Contain 'version'
        $metadata.download.PSObject.Properties.Name | Should -Not -Contain 'download_url'
        $release.version_source | Should -Be 'WinImgNormalizer.ps1:Get-WinImgVersion'
        $release.release_state | Should -Be 'unreleased'
        foreach ($field in @('tag', 'release_date', 'release_url', 'asset_filename', 'asset_bytes', 'asset_sha256', 'download_url')) {
            $release.$field | Should -BeNullOrEmpty
        }
        $metadata.download.state | Should -Be 'unavailable'
        $metadata.download.render_link | Should -BeFalse
        $metadata.screenshots.Count | Should -Be 0
        $metadata.comparisons.Count | Should -Be 0
        foreach ($asset in $metadata.assets) {
            $path = [IO.Path]::GetFullPath((Join-Path (Split-Path -Parent $metadataPath) $asset.path))
            (Get-FileHash -LiteralPath $path).Hash.ToLowerInvariant() | Should -Be $asset.sha256
            (Get-Item -LiteralPath $path).Length | Should -Be $asset.bytes
        }
    }
    It 'resolves actual local copy/integration links and refuses a missing or escaped link' {
        foreach ($relative in @('docs/website/PRODUCT_COPY.md', 'docs/website/INTEGRATION.md')) {
            $path = Join-Path $repository $relative
            { Assert-WebsiteDocumentLinks $path ([IO.File]::ReadAllText($path)) } | Should -Not -Throw
        }
        $path = Join-Path $repository 'docs/website/INTEGRATION.md'
        { Assert-WebsiteDocumentLinks $path '[broken](missing-private-file.md)' } | Should -Throw
        { Assert-WebsiteDocumentLinks $path '[escape](../../../outside-private.md)' } | Should -Throw
    }
    It 'rejects unpublished canonical field <Field> rather than constructing a live release link' -ForEach @(
        @{Field='tag';Value='v1.0.0'}, @{Field='release_date';Value='2026-10-06'},
        @{Field='release_url';Value='https://github.com/PikkuJanne/WinImgNormalizer/releases/tag/v1.0.0'},
        @{Field='asset_filename';Value='WinImgNormalizer-1.0.0.zip'}, @{Field='asset_bytes';Value=1234},
        @{Field='asset_bytes';Value=0}, @{Field='release_date';Value=''},
        @{Field='asset_sha256';Value=('a'*64)},
        @{Field='download_url';Value='https://github.com/PikkuJanne/WinImgNormalizer/releases/download/v1.0.0/WinImgNormalizer-1.0.0.zip'}
    ) {
        $root = New-WebsiteRepository
        $path = Join-Path $root 'release-metadata.json'
        $release = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
        $release.$Field = $Value
        [IO.File]::WriteAllText($path, ($release | ConvertTo-Json -Depth 8), $utf8)
        Assert-WebsiteRejected -Root $root
    }
    It 'rejects extra <Object> publication keys even when the canonical release is still draft' -ForEach @(
        @{Object='root';Field='version';Value='9.9.9'}, @{Object='product';Field='tag';Value='v1.0.0'},
        @{Object='download';Field='download_url';Value='https://github.com/PikkuJanne/WinImgNormalizer/releases/latest/download/WinImgNormalizer.zip'},
        @{Object='download';Field='asset_sha256';Value=('a'*64)}
    ) {
        $metadata = New-WebsiteMetadata
        $target = if ($Object -eq 'root') { $metadata } else { $metadata.$Object }
        $target | Add-Member -NotePropertyName $Field -NotePropertyValue $Value
        Assert-WebsiteRejected (Write-WebsiteMetadata $metadata)
    }
    It 'refuses a live download flag of <Value> with its actual JSON type' -ForEach @(
        @{Value=$true}, @{Value='false'}, @{Value=0}, @{Value=$null}
    ) {
        $metadata = New-WebsiteMetadata
        $metadata.download.render_link = $Value
        Assert-WebsiteRejected (Write-WebsiteMetadata $metadata)
    }
    It 'refuses processing boundary <Field> even with every release field still null' -ForEach @(
        @{Field='uploads'}, @{Field='processing_api'}, @{Field='server_runtime'},
        @{Field='runtime_account_required'}, @{Field='telemetry'}
    ) {
        $metadata = New-WebsiteMetadata
        $metadata.processing.$Field = $true
        Assert-WebsiteRejected (Write-WebsiteMetadata $metadata)
    }
    It 'rejects claimed deployment <Field> in an integration handoff without website context' -ForEach @(
        @{Field='domain';Value='https://invented.example.invalid'},
        @{Field='framework';Value='InventedFramework'}, @{Field='hosting_provider';Value='InventedProvider'}
    ) {
        $metadata = New-WebsiteMetadata
        $metadata.deployment.$Field = $Value
        Assert-WebsiteRejected (Write-WebsiteMetadata $metadata)
    }
    It 'rejects a fabricated <Field> capture or comparison while preserving the original branding' -ForEach @(
        @{Field='screenshots'}, @{Field='comparisons'}
    ) {
        $metadata = New-WebsiteMetadata
        $metadata.$Field = @([pscustomobject]@{ path='../../WinImgNormalizer_icon_variant.png'; caption='Invented application output'; sha256=$metadata.assets[0].sha256 })
        Assert-WebsiteRejected (Write-WebsiteMetadata $metadata)
    }
    It 'refuses metadata with stale <Field> identity or prerequisites' -ForEach @(
        @{Object='root';Field='repository_url';Value='https://github.com/another-owner/another-repository'},
        @{Object='root';Field='license';Value='Proprietary'},
        @{Object='root';Field='release_metadata_file';Value='../../README.md'},
        @{Object='prerequisites';Field='imagemagick_minimum';Value='7.1.1-47'},
        @{Object='prerequisites';Field='supported_major';Value=8},
        @{Object='product';Field='name';Value='Invented Product'}
    ) {
        $metadata = New-WebsiteMetadata
        $target = if ($Object -eq 'root') { $metadata } else { $metadata.$Object }
        $target.$Field = $Value
        Assert-WebsiteRejected (Write-WebsiteMetadata $metadata)
    }
    It 'rejects changed asset <Field> instead of treating it as an approved capture' -ForEach @(
        @{Index=0;Field='sha256';Value=('a'*64)}, @{Index=0;Field='bytes';Value=1},
        @{Index=0;Field='width';Value=1}, @{Index=0;Field='width';Value='1024'}, @{Index=0;Field='media_type';Value='image/jpeg'},
        @{Index=0;Field='path';Value='../../README.md'}, @{Index=0;Field='usage';Value='screenshot'},
        @{Index=1;Field='sizes';Value=@(16,32)}, @{Index=1;Field='display';Value=$false}
    ) {
        $metadata = New-WebsiteMetadata
        $metadata.assets[$Index].$Field = $Value
        Assert-WebsiteRejected (Write-WebsiteMetadata $metadata)
    }
    It 'rejects a missing local document and a path escape instead of accepting a broken handoff' {
        foreach ($value in @('missing-product-copy.md', '../../../outside-private-file.md')) {
            $metadata = New-WebsiteMetadata
            $metadata.documents.product_copy = $value
            Assert-WebsiteRejected (Write-WebsiteMetadata $metadata)
        }
        $root = New-WebsiteRepository
        $document = Join-Path $root 'docs/website/PRODUCT_COPY.md'
        Move-Item -LiteralPath $document -Destination ($document + '.owned-missing-control')
        Assert-WebsiteRejected -Root $root
    }
    It 'refuses actual broken links, upload markup and unpublished release links in owned visitor copy' {
        foreach ($text in @('[broken](missing-product-file.md)', '[escape](%2e%2e/%2e%2e/%2e%2e/outside-private.md)',
            'Download www.example.invalid/file.zip', 'Contact help@example.invalid',
            '<form action="https://invented.example.invalid/upload"><input type="file"></form>',
            '[download](https://github.com/PikkuJanne/WinImgNormalizer/releases/download/v1.0.0/WinImgNormalizer-1.0.0.zip)',
            '[download](https://github.com/PikkuJanne/WinImgNormalizer/releases/latest/download/WinImgNormalizer.zip)')) {
            $root = New-WebsiteRepository
            [IO.File]::AppendAllText((Join-Path $root 'docs/website/PRODUCT_COPY.md'), "`n" + $text + "`n", $utf8)
            Assert-WebsiteRejected -Root $root
        }
    }
    It 'refuses a <Kind> link that could expose an unpublished download in rendered copy' -ForEach @(
        @{Kind='reference definition';Text="[Download][release]`n`n[release]: https://invented.example.invalid/file.zip"},
        @{Kind='autolink';Text='<https://invented.example.invalid/file.zip>'},
        @{Kind='HTML anchor';Text='<a href="https://invented.example.invalid/file.zip">Download</a>'},
        @{Kind='bare URI';Text='Download https://invented.example.invalid/file.zip'}
    ) {
        $root = New-WebsiteRepository
        [IO.File]::AppendAllText((Join-Path $root 'docs/website/PRODUCT_COPY.md'), "`n" + $Text + "`n", $utf8)
        Assert-WebsiteRejected -Root $root
    }
    It 'refuses malformed, nonscalar and wrong-schema metadata as native guard failures' {
        foreach ($text in @('{"schema_version":', '[]', 'null')) {
            Assert-WebsiteRejected (Write-WebsiteMetadata -Text $text)
        }
        foreach ($value in @('1', 2)) {
            $metadata = New-WebsiteMetadata
            $metadata.schema_version = $value
            Assert-WebsiteRejected (Write-WebsiteMetadata $metadata)
        }
        $metadata = New-WebsiteMetadata
        $metadata.processing.uploads = 'false'
        Assert-WebsiteRejected (Write-WebsiteMetadata $metadata)
    }
    It 'refuses an unowned external or non-JSON metadata path before reading it as a draft' {
        Assert-WebsiteRejected -Metadata (Join-Path $repository 'README.md')
        $root = New-WebsiteRepository
        Assert-WebsiteRejected -Root $root -Metadata (Write-WebsiteMetadata (New-WebsiteMetadata))
        Assert-WebsiteRejected -Metadata (Join-Path (Join-Path $draftRepository '.scratch') 'missing.json')
        $path = Join-Path (Join-Path $draftRepository '.scratch') 'directory.json'
        [IO.Directory]::CreateDirectory($path) | Out-Null
        Assert-WebsiteRejected -Metadata $path
    }
    It 'accepts an ordinary owned replica and refuses stale canonical version and published state' {
        $root = New-WebsiteRepository
        $beforeReplica = Get-WebsiteRepositoryState $root
        (Invoke-WebsiteGuard -Root $root).Exit | Should -Be 0
        (Get-WebsiteRepositoryState $root) | Should -Be $beforeReplica
        $path = Join-Path $root 'release-metadata.json'
        $release = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
        $originalVersion = $release.version
        $originalTag = $release.proposed_tag
        $release.version = '9.9.9'; $release.proposed_tag = 'v9.9.9'
        [IO.File]::WriteAllText($path, ($release | ConvertTo-Json -Depth 8), $utf8)
        Assert-WebsiteRejected -Root $root
        $release.version = $originalVersion; $release.proposed_tag = $originalTag; $release.release_state = 'published'
        [IO.File]::WriteAllText($path, ($release | ConvertTo-Json -Depth 8), $utf8)
        Assert-WebsiteRejected -Root $root
    }
    It 'binds the actual application version body rather than trusting the projection alone' {
        $root = New-WebsiteRepository
        $application = Join-Path $root 'WinImgNormalizer.ps1'
        $text = [IO.File]::ReadAllText($application)
        $changed = [regex]::Replace($text, "(?m)(function Get-WinImgVersion \{\r?\n  return ')[^']+(')", '${1}9.9.9${2}')
        $changed | Should -Not -Be $text
        [IO.File]::WriteAllText($application, $changed, $utf8)
        Assert-WebsiteRejected -Root $root
    }
    It 'detects actual tampering with an owned branding file even when metadata remains untouched' {
        $root = New-WebsiteRepository
        $asset = Join-Path $root 'WinImgNormalizer_icon_variant.png'
        $stream = [IO.File]::Open($asset, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        try {
            $null = $stream.Seek(-1, [IO.SeekOrigin]::End)
            $value = $stream.ReadByte()
            $null = $stream.Seek(-1, [IO.SeekOrigin]::End)
            $stream.WriteByte([byte]($value -bxor 1))
        } finally { $stream.Dispose() }
        Assert-WebsiteRejected -Root $root
    }
}

Describe 'T073 recorded published website consistency using synthetic offline controls' {
    It 'accepts the exact recorded identity and approved document links without modifying its replica' {
        $root = New-PublishedWebsiteRepository
        $beforeReplica = Get-WebsiteRepositoryState $root
        $observed = Invoke-WebsiteGuard -Root $root
        $observed.Exit | Should -Be 0
        $observed.Text | Should -Match 'published record.*exact recorded download identity.*no deployment'
        (Get-WebsiteRepositoryState $root) | Should -Be $beforeReplica
    }
    It 'refuses a published handoff without a regular bounded publication observation' {
        $root = New-PublishedWebsiteRepository
        $path = Join-Path $root $publicationRelative
        Move-Item -LiteralPath $path -Destination ($path + '.owned-missing-control')
        Assert-WebsiteRejected -Root $root
        [IO.File]::WriteAllText($path, (' ' * 65537), $utf8)
        Assert-WebsiteRejected -Root $root -Reason 'bounded to 64 KiB'
    }
    It 'rejects publication observation <Field> rather than trusting a claimed release identity' -ForEach @(
        @{Field='schema_version';Value='1';Reason='Publication schema version'},
        @{Field='visibility';Value='private';Reason='Release visibility'},
        @{Field='tag';Value='v9.9.9';Reason='Published tag'},
        @{Field='source_revision';Value=('a'*40);Reason='Published source revision'},
        @{Field='release_id';Value=0;Reason='Release ID'},
        @{Field='release_id';Value='100';Reason='Release ID'},
        @{Field='repository_url';Value='https://github.com/another-owner/another-repository';Reason='Published repository'},
        @{Field='release_url';Value='https://github.com/PikkuJanne/WinImgNormalizer/releases/latest';Reason='Published release URL'},
        @{Field='published_at';Value='2026-02-30T00:00:00Z';Reason='Publication time'},
        @{Field='published_at';Value='2026-10-07T00:00:00+00:00';Reason='Publication time'},
        @{Field='observed_at';Value='2026-10-06T23:59:59Z';Reason='cannot precede publication'}
    ) {
        $root = New-PublishedWebsiteRepository
        $path = Join-Path $root $publicationRelative
        $proof = Read-WebsiteFixtureJson $path
        $proof.$Field = $Value
        [IO.File]::WriteAllText($path, ($proof | ConvertTo-Json -Depth 12), $utf8)
        Assert-WebsiteRejected -Root $root -Reason $Reason
    }
    It 'refuses missing, extra and incorrectly cased publication fields and asset sets' {
        foreach ($kind in @('missing field','extra field','cased field','missing asset','extra asset','reordered assets','duplicate ID')) {
            $root = New-PublishedWebsiteRepository
            $path = Join-Path $root $publicationRelative
            $proof = Read-WebsiteFixtureJson $path
            switch ($kind) {
                'missing field' { $proof.PSObject.Properties.Remove('source_revision') }
                'extra field' { $proof | Add-Member -NotePropertyName owner_approved -NotePropertyValue $true }
                'cased field' {
                    $value = $proof.source_revision; $proof.PSObject.Properties.Remove('source_revision')
                    $proof | Add-Member -NotePropertyName Source_Revision -NotePropertyValue $value
                }
                'missing asset' { $proof.assets = @($proof.assets[0],$proof.assets[1]) }
                'extra asset' { $proof.assets = @($proof.assets) + @($proof.assets[0]) }
                'reordered assets' { $proof.assets = @($proof.assets[1],$proof.assets[0],$proof.assets[2]) }
                'duplicate ID' { $proof.assets[1].id = $proof.assets[0].id }
            }
            [IO.File]::WriteAllText($path, ($proof | ConvertTo-Json -Depth 12), $utf8)
            Assert-WebsiteRejected -Root $root
        }
    }
    It 'rejects published asset <Index> <Field> when recorded bytes, hashes or URLs are changed' -ForEach @(
        @{Index=0;Field='name';Value='WinImgNormalizer.zip';Reason='Published asset name'},
        @{Index=0;Field='bytes';Value='184659';Reason='Published asset bytes'},
        @{Index=0;Field='sha256';Value=('a'*64);Reason='Published asset sha256'},
        @{Index=0;Field='downloaded_bytes';Value=184658;Reason='Observed public download bytes'},
        @{Index=0;Field='downloaded_sha256';Value=('a'*64);Reason='Observed public download SHA-256'},
        @{Index=0;Field='download_url';Value='https://github.com/PikkuJanne/WinImgNormalizer/releases/latest/download/WinImgNormalizer-1.0.0-portable.zip';Reason='Published download URL'},
        @{Index=1;Field='sha256';Value=('a'*64);Reason='Published asset sha256'},
        @{Index=2;Field='downloaded_sha256';Value=('a'*64);Reason='Observed public download SHA-256'}
    ) {
        $root = New-PublishedWebsiteRepository
        $path = Join-Path $root $publicationRelative
        $proof = Read-WebsiteFixtureJson $path
        $proof.assets[$Index].$Field = $Value
        [IO.File]::WriteAllText($path, ($proof | ConvertTo-Json -Depth 12), $utf8)
        Assert-WebsiteRejected -Root $root -Reason $Reason
    }
    It 'rejects canonical published <Field> that disagrees with the recorded public asset' -ForEach @(
        @{Field='tag';Value='v9.9.9';Reason='Canonical published tag'},
        @{Field='release_date';Value='2026-10-06';Reason='Canonical publication date'},
        @{Field='release_url';Value='https://github.com/PikkuJanne/WinImgNormalizer/releases/latest';Reason='Canonical published release URL'},
        @{Field='asset_filename';Value='WinImgNormalizer.zip';Reason='Canonical download filename'},
        @{Field='asset_bytes';Value=1;Reason='Canonical download bytes'},
        @{Field='asset_sha256';Value=('a'*64);Reason='Canonical download SHA-256'},
        @{Field='download_url';Value='https://invented.example.invalid/file.zip';Reason='Canonical download URL'}
    ) {
        $root = New-PublishedWebsiteRepository
        $path = Join-Path $root 'release-metadata.json'
        $release = [IO.File]::ReadAllText($path) | ConvertFrom-Json
        $release.$Field = $Value
        [IO.File]::WriteAllText($path, ($release | ConvertTo-Json -Depth 8), $utf8)
        Assert-WebsiteRejected -Root $root -Reason $Reason
    }
    It 'rejects published website projection <Field> rather than accepting a mismatched download state' -ForEach @(
        @{Object='root';Field='preparation_state';Value='draft';Reason='Release state'},
        @{Object='download';Field='state';Value='unavailable';Reason='Download state'},
        @{Object='download';Field='render_link';Value=$false;Reason='Download link visibility'},
        @{Object='download';Field='render_link';Value='true';Reason='Download link visibility'},
        @{Object='download';Field='label';Value='Download WinImgNormalizer 9.9.9';Reason='Download label'}
    ) {
        $root = New-PublishedWebsiteRepository
        $path = Join-Path $root 'docs/website/metadata.json'
        $metadata = [IO.File]::ReadAllText($path) | ConvertFrom-Json
        $target = if ($Object -eq 'root') { $metadata } else { $metadata.$Object }
        $target.$Field = $Value
        [IO.File]::WriteAllText($path, ($metadata | ConvertTo-Json -Depth 16), $utf8)
        Assert-WebsiteRejected -Root $root -Reason $Reason
    }
    It 'keeps publication from authorizing uploads, screenshots or deployment' {
        foreach ($kind in @('upload','screenshot','deployment')) {
            $root = New-PublishedWebsiteRepository
            $path = Join-Path $root 'docs/website/metadata.json'
            $metadata = [IO.File]::ReadAllText($path) | ConvertFrom-Json
            switch ($kind) {
                'upload' { $metadata.processing.uploads = $true }
                'screenshot' { $metadata.screenshots = @([pscustomobject]@{path='../../WinImgNormalizer_icon_variant.png'}) }
                'deployment' { $metadata.deployment.domain = 'https://invented.example.invalid' }
            }
            [IO.File]::WriteAllText($path, ($metadata | ConvertTo-Json -Depth 16), $utf8)
            Assert-WebsiteRejected -Root $root
        }
    }
    It 'refuses arbitrary and floating release destinations even in published visitor copy' {
        foreach ($uri in @('https://invented.example.invalid/file.zip',
            'https://github.com/PikkuJanne/WinImgNormalizer/releases/latest',
            'https://github.com/PikkuJanne/WinImgNormalizer/releases/download/v1.0.0/WinImgNormalizer.zip',
            'https://github.com/PikkuJanne/WinImgNormalizer/releases/download/v9.9.9/WinImgNormalizer-1.0.0-portable.zip',
            'https://github.com@invented.example.invalid/file.zip')) {
            $root = New-PublishedWebsiteRepository
            [IO.File]::AppendAllText((Join-Path $root 'docs/website/PRODUCT_COPY.md'), ("`n[Download](" + $uri + ")`n"), $utf8)
            Assert-WebsiteRejected -Root $root -Reason 'unreviewed HTTP\(S\) destination|bare-domain or email autolink destinations'
        }
    }
    It 'rejects ambiguous duplicate keys in <Target> instead of relying on parser overwrite order' -ForEach @(
        @{Target='canonical';Key='download_url';Alias='download_url'},
        @{Target='proof';Key='source_revision';Alias='source_revision'},
        @{Target='proof';Key='source_revision';Alias='\u0073ource_revision'},
        @{Target='proof';Key='source_revision';Alias='Source_Revision'},
        @{Target='proof';Key='sha256';Alias='sha256'},
        @{Target='website';Key='render_link';Alias='render_link'}
    ) {
        $root = New-PublishedWebsiteRepository
        $relative = switch ($Target) {
            'canonical' { 'release-metadata.json' }
            'proof' { $publicationRelative }
            'website' { 'docs/website/metadata.json' }
        }
        $path = Join-Path $root $relative
        $text = [IO.File]::ReadAllText($path)
        # Put a duplicate before the unchanged original. Readers that keep the
        # last value would otherwise preserve the original passing projection.
        $pattern = '"' + [regex]::Escape($Key) + '"\s*:'
        $first = [regex]::Match($text, $pattern)
        $first.Success | Should -BeTrue
        $duplicate = '"' + $Alias + '":null,' + $first.Value
        $changed = $text.Substring(0, $first.Index) + $duplicate + $text.Substring($first.Index + $first.Length)
        [IO.File]::WriteAllText($path, $changed, $utf8)
        Assert-WebsiteRejected -Root $root -Reason 'JSON object keys must be unique'
    }
}

AfterAll {
    $after = @($publicPaths | ForEach-Object {
        [pscustomobject]@{ path = $_; sha256 = (Get-FileHash -LiteralPath (Join-Path $repository $_)).Hash.ToLowerInvariant() }
    }) | ConvertTo-Json -Compress
    $after | Should -Be $before
    $policiesAfter = @(Get-ExecutionPolicy -List | Where-Object Scope -ne Process | ForEach-Object { $_.Scope.ToString() + '=' + $_.ExecutionPolicy.ToString() }) -join '|'
    $policiesAfter | Should -Be $policiesBefore
    [IO.File]::WriteAllText((Join-Path $owned 'native-observations.json'), ($observations | ConvertTo-Json -Depth 6), $utf8)
}
