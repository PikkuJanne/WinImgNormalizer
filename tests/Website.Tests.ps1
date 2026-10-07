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
    $owned = Join-Path $scratch ('M4-T04-website-' + [guid]::NewGuid().ToString('N'))
    if (Test-Path -LiteralPath $owned) { throw 'Website fixture ownership collision.' }
    [IO.Directory]::CreateDirectory($owned) | Out-Null
    [IO.File]::WriteAllText((Join-Path $owned '.winimg-fixture-root'), 'M4-T04 owned synthetic website draft controls; retained for inspection.', [Text.UTF8Encoding]::new($false))
    $publicPaths = @('tools/website/Test-WebsiteHandoff.ps1', 'docs/website/metadata.json',
        'docs/website/PRODUCT_COPY.md', 'docs/website/INTEGRATION.md', 'WinImgNormalizer_icon_variant.png',
        'WinImgNormalizer_variant.ico', 'WinImgNormalizer_poster.png', 'WinImgNormalizer.ps1',
        'WinImgNormalizer.bat', 'release-metadata.json', 'README.md', 'CHANGELOG.md', 'LICENSE',
        'SECURITY.md', 'docs/BEHAVIOR.md', 'docs/release/GETTING_STARTED.md', 'docs/release/NOTES.md',
        'docs/release/THIRD_PARTY_NOTICES.md', 'docs/release/PACKAGING.md',
        'tools/release/Update-ReleaseMetadata.ps1', 'tests/Website.Tests.ps1')
    $before = @($publicPaths | ForEach-Object {
        [pscustomobject]@{ path = $_; sha256 = (Get-FileHash -LiteralPath (Join-Path $repository $_)).Hash.ToLowerInvariant() }
    }) | ConvertTo-Json -Compress
    $policiesBefore = @(Get-ExecutionPolicy -List | Where-Object Scope -ne Process | ForEach-Object { $_.Scope.ToString() + '=' + $_.ExecutionPolicy.ToString() }) -join '|'
    $observations = New-Object 'Collections.Generic.List[object]'
    $utf8 = [Text.UTF8Encoding]::new($false)
    function New-WebsiteMetadata {
        # Reparse the source for every control so one mutated object cannot
        # contaminate a later positive or negative fixture.
        return [IO.File]::ReadAllText($metadataPath) | ConvertFrom-Json
    }
    function Write-WebsiteMetadata([object]$Metadata, [string]$Text) {
        $path = Join-Path $owned ([guid]::NewGuid().ToString('N') + '.json')
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
        $fixturePaths = @($publicPaths | Where-Object { $_ -ne 'tests/Website.Tests.ps1' }) + @('docs/codex-winimg/RELEASE_AND_WEBSITE.md')
        foreach ($relative in $fixturePaths) {
            $destination = Join-Path $root $relative
            [IO.Directory]::CreateDirectory((Split-Path -Parent $destination)) | Out-Null
            [IO.File]::Copy((Join-Path $repository $relative), $destination)
        }
        return $root
    }
    function Assert-WebsiteRejected([string]$Metadata, [string]$Root = $repository) {
        $observed = Invoke-WebsiteGuard -Metadata $Metadata -Root $Root
        $observed.Exit | Should -Be 1
        $observed.Text | Should -Match 'WEBSITE HANDOFF INVALID:'
        $observed.Text | Should -Not -Match 'ParserError|ParameterBindingException'
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
}

Describe 'T073 framework-neutral website preparation and metadata boundary' {
    It 'accepts the checked-in draft through explicit and script-relative native CLI paths' {
        foreach ($result in @((Invoke-WebsiteGuard), (Invoke-WebsiteGuard -DefaultPaths),
            (Invoke-WebsiteGuard -Metadata (Write-WebsiteMetadata (New-WebsiteMetadata))))) {
            $result.Exit | Should -Be 0
            $result.Text | Should -Match 'WEBSITE HANDOFF VALID:'
        }
    }
    It 'keeps the only release version and publication fields in the canonical release record' {
        $metadata = New-WebsiteMetadata
        $release = Get-Content -LiteralPath (Join-Path $repository 'release-metadata.json') -Raw | ConvertFrom-Json
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
        @{Index=1;Field='sizes';Value=@(16,32)}, @{Index=2;Field='display';Value=$true}
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
        Assert-WebsiteRejected -Metadata (Join-Path $owned 'missing.json')
        $path = Join-Path $owned 'directory.json'
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

AfterAll {
    $after = @($publicPaths | ForEach-Object {
        [pscustomobject]@{ path = $_; sha256 = (Get-FileHash -LiteralPath (Join-Path $repository $_)).Hash.ToLowerInvariant() }
    }) | ConvertTo-Json -Compress
    $after | Should -Be $before
    $policiesAfter = @(Get-ExecutionPolicy -List | Where-Object Scope -ne Process | ForEach-Object { $_.Scope.ToString() + '=' + $_.ExecutionPolicy.ToString() }) -join '|'
    $policiesAfter | Should -Be $policiesBefore
    [IO.File]::WriteAllText((Join-Path $owned 'native-observations.json'), ($observations | ConvertTo-Json -Depth 6), $utf8)
}
