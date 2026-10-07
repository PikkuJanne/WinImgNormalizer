BeforeAll {
    $repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    $application = Join-Path $repository 'WinImgNormalizer.ps1'
    $generator = Join-Path $repository 'tools/release/Update-ReleaseMetadata.ps1'
    $scratch = Join-Path $repository '.scratch'
    function Assert-VersionAncestors([string]$Path) {
        $current = [IO.DirectoryInfo]::new([IO.Path]::GetFullPath($Path))
        while ($current) {
            if ($current.Exists -and ($current.Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Versioning fixture crosses a reparse point.' }
            $current = $current.Parent
        }
    }
    Assert-VersionAncestors $scratch
    foreach ($pictures in @([Environment]::GetFolderPath('MyPictures'), (Join-Path $env:USERPROFILE 'Pictures'))) {
        if (-not $pictures) { continue }
        $protected = [IO.Path]::GetFullPath($pictures).TrimEnd('\', '/')
        if ($scratch.Equals($protected, [StringComparison]::OrdinalIgnoreCase) -or
            $scratch.StartsWith($protected + '\', [StringComparison]::OrdinalIgnoreCase) -or
            $protected.StartsWith($scratch + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Versioning fixtures overlap real Pictures.' }
    }
    & git -C $repository check-ignore --quiet --no-index -- (Join-Path $scratch 'versioning-ignore-probe')
    if ($LASTEXITCODE -ne 0) { throw 'Versioning fixtures must already be ignored.' }
    $ownedRoot = Join-Path $scratch ('M4-T01-versioning-' + [Guid]::NewGuid().ToString('N'))
    if ([IO.Directory]::Exists($ownedRoot)) { throw 'Versioning ownership collision.' }
    $null = [IO.Directory]::CreateDirectory($ownedRoot)
    [IO.File]::WriteAllText((Join-Path $ownedRoot '.winimg-fixture-root'), 'M4-T01 owned synthetic versioning fixtures; retained for inspection.', [Text.UTF8Encoding]::new($false))
    if (-not $env:WINIMG_TEST_MAGICK -or -not [IO.Path]::IsPathRooted($env:WINIMG_TEST_MAGICK)) { throw 'Explicit verified ImageMagick required.' }
    $magick = [IO.Path]::GetFullPath($env:WINIMG_TEST_MAGICK)
    Assert-VersionAncestors $magick
    $imagePin = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'legacy/toolchain.json') -Raw | ConvertFrom-Json
    (Get-FileHash -LiteralPath $magick).Hash.ToLowerInvariant() | Should -Be $imagePin.portable_imagemagick.executable_sha256
    $artifactPaths = @('WinImgNormalizer.ps1', 'tools/release/Update-ReleaseMetadata.ps1', 'docs/release/NOTES.md', 'CHANGELOG.md', 'release-metadata.json')
    $proofRelative = 'docs/release/publication-v1.0.0.json'
    if ([IO.File]::Exists((Join-Path $repository $proofRelative))) { $artifactPaths += $proofRelative }
    function Get-VersionBindings {
        foreach ($relative in $artifactPaths + @('tests/Versioning.Tests.ps1', 'LICENSE')) {
            [pscustomobject]@{ Path = $relative; Sha256 = (Get-FileHash -LiteralPath (Join-Path $repository $relative)).Hash.ToLowerInvariant() }
        }
    }
    $bindingsBefore = @(Get-VersionBindings)
    $observations = New-Object 'Collections.Generic.List[object]'
    function New-VersionDirectory([string]$Label) {
        $path = [IO.Path]::GetFullPath((Join-Path $ownedRoot ($Label + '-' + [Guid]::NewGuid().ToString('N'))))
        if (-not $path.StartsWith($ownedRoot + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Versioning fixture escaped ownership.' }
        Assert-VersionAncestors $path
        if ([IO.Directory]::Exists($path)) { throw 'Versioning fixture directory collision.' }
        $null = [IO.Directory]::CreateDirectory($path)
        return $path
    }
    function ConvertTo-VersionWindowsArgument([string]$Value) {
        $escaped = [regex]::Replace($Value, '(\\*)"', '$1$1\"')
        $escaped = [regex]::Replace($escaped, '(\\+)$', '$1$1')
        return '"' + $escaped + '"'
    }
    function Invoke-VersionChild([string]$Script, [string[]]$Arguments = @(), [switch]$WithoutMagick) {
        Assert-VersionAncestors (Split-Path -Parent $Script)
        $work = New-VersionDirectory 'child work'
        $temporary = Join-Path $work 'temporary'
        $null = [IO.Directory]::CreateDirectory($temporary)
        $start = New-Object Diagnostics.ProcessStartInfo
        $start.FileName = (Get-Process -Id $PID).Path
        $tokens = @('-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $Script)
        if ($Arguments.Count -gt 0) { $tokens += $Arguments }
        $start.Arguments = ($tokens | ForEach-Object { ConvertTo-VersionWindowsArgument ([string]$_) }) -join ' '
        $start.WorkingDirectory = $work; $start.UseShellExecute = $false; $start.CreateNoWindow = $true
        $start.RedirectStandardOutput = $true; $start.RedirectStandardError = $true
        foreach ($key in @($start.EnvironmentVariables.Keys)) {
            if ([string]::Equals([string]$key, 'PSModulePath', [StringComparison]::OrdinalIgnoreCase)) { $start.EnvironmentVariables.Remove([string]$key) }
        }
        $start.EnvironmentVariables['PATH'] = if ($WithoutMagick) { '' } else { [IO.Path]::GetDirectoryName($magick) + ';' + $env:PATH }
        $start.EnvironmentVariables['TEMP'] = $temporary; $start.EnvironmentVariables['TMP'] = $temporary
        $start.EnvironmentVariables['MAGICK_TEMPORARY_PATH'] = $temporary
        $process = New-Object Diagnostics.Process
        $process.StartInfo = $start
        try {
            if (-not $process.Start()) { throw 'Owned versioning child did not start.' }
            $stdout = $process.StandardOutput.ReadToEndAsync(); $stderr = $process.StandardError.ReadToEndAsync()
            if (-not $process.WaitForExit(30000)) {
                & (Join-Path $env:SystemRoot 'System32/taskkill.exe') /PID $process.Id /T /F 2>&1 | Out-Null
                if (-not $process.WaitForExit(5000)) { $process.Kill() }
                throw 'Owned versioning child exceeded its 30-second bound.'
            }
            if (-not $stdout.Wait(5000) -or -not $stderr.Wait(5000)) { throw 'Owned versioning child streams did not finish.' }
            $result = [pscustomobject]@{ ExitCode = $process.ExitCode; StdOut = $stdout.GetAwaiter().GetResult(); StdErr = $stderr.GetAwaiter().GetResult(); Work = $work }
            $observations.Add($result)
            return $result
        } finally { $process.Dispose() }
    }
    function Get-VersionFileState([string]$Path) {
        $item = Get-Item -LiteralPath $Path -Force
        [pscustomobject]@{ Sha256 = (Get-FileHash -LiteralPath $Path).Hash; Length = $item.Length; CreationTicks = $item.CreationTimeUtc.Ticks; ModifiedTicks = $item.LastWriteTimeUtc.Ticks; Attributes = [int]$item.Attributes } | ConvertTo-Json -Compress
    }
    function Read-VersionJson([string]$Path) {
        $jsonArguments = @{}
        if ((Get-Command ConvertFrom-Json).Parameters.ContainsKey('DateKind')) { $jsonArguments.DateKind = 'String' }
        return (Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json @jsonArguments)
    }
    function Get-VersionGeneratedState([string]$Root) {
        foreach ($relative in @('CHANGELOG.md', 'release-metadata.json')) { Get-VersionFileState (Join-Path $Root $relative) }
        if ([IO.File]::Exists((Join-Path $Root $proofRelative))) { Get-VersionFileState (Join-Path $Root $proofRelative) }
    }
    function New-VersionReplica {
        $root = New-VersionDirectory 'disposable project with spaces'
        foreach ($relative in $artifactPaths) {
            if ($relative -ceq $proofRelative) { continue }
            $target = Join-Path $root $relative
            $null = [IO.Directory]::CreateDirectory((Split-Path -Parent $target))
            [IO.File]::WriteAllBytes($target, [IO.File]::ReadAllBytes((Join-Path $repository $relative)))
        }
        # Draft controls are owned fixtures, regardless of the live release state.
        $metadataPath = Join-Path $root 'release-metadata.json'
        $draft = Read-VersionJson $metadataPath
        $draft.release_state = 'unreleased'
        foreach ($field in @('tag', 'release_date', 'release_url', 'asset_filename', 'asset_bytes', 'asset_sha256', 'download_url')) { $draft.$field = $null }
        [IO.File]::WriteAllText($metadataPath, (($draft | ConvertTo-Json -Depth 4 -Compress) + "`n"), [Text.UTF8Encoding]::new($false))
        $draftNotes = [IO.File]::ReadAllText((Join-Path $root 'docs/release/NOTES.md')).Replace("`r`n", "`n").TrimEnd()
        [IO.File]::WriteAllText((Join-Path $root 'CHANGELOG.md'), "# Changelog`n`n<!-- Generated by tools/release/Update-ReleaseMetadata.ps1; edit docs/release/NOTES.md. -->`n`n## $($draft.version) (unreleased)`n`n$draftNotes`n", [Text.UTF8Encoding]::new($false))
        return $root
    }
    function New-VersionPublicationProof {
        return [ordered]@{
            schema_version = 1; product = 'WinImgNormalizer'; version = '1.0.0'; repository_url = 'https://github.com/PikkuJanne/WinImgNormalizer'
            visibility = 'public'; tag = 'v1.0.0'; source_revision = '8edbcbaeb3425ec3a52eeafde212c32553755af1'; release_id = 100
            release_url = 'https://github.com/PikkuJanne/WinImgNormalizer/releases/tag/v1.0.0'; published_at = '2026-10-07T00:00:00Z'; observed_at = '2026-10-07T00:00:01Z'
            assets = @(
                [ordered]@{ id = 101; name = 'WinImgNormalizer-1.0.0-portable.zip'; bytes = 184659; sha256 = '251828028e144759c919645f423fde08641cabf42e04db7024d7ad17da4ba14d'; download_url = 'https://github.com/PikkuJanne/WinImgNormalizer/releases/download/v1.0.0/WinImgNormalizer-1.0.0-portable.zip'; downloaded_bytes = 184659; downloaded_sha256 = '251828028e144759c919645f423fde08641cabf42e04db7024d7ad17da4ba14d' },
                [ordered]@{ id = 102; name = 'build-provenance.json'; bytes = 763; sha256 = '901598504d3bb60cdfa4a236873a07d3a7287cb8ff6b3c715b202b39865cefae'; download_url = 'https://github.com/PikkuJanne/WinImgNormalizer/releases/download/v1.0.0/build-provenance.json'; downloaded_bytes = 763; downloaded_sha256 = '901598504d3bb60cdfa4a236873a07d3a7287cb8ff6b3c715b202b39865cefae' },
                [ordered]@{ id = 103; name = 'SHA256SUMS.txt'; bytes = 190; sha256 = '2257dbcd00707a93f3e811c9d3ff0eafb59a2e65bb756c440e152f1642540497'; download_url = 'https://github.com/PikkuJanne/WinImgNormalizer/releases/download/v1.0.0/SHA256SUMS.txt'; downloaded_bytes = 190; downloaded_sha256 = '2257dbcd00707a93f3e811c9d3ff0eafb59a2e65bb756c440e152f1642540497' }
            )
        }
    }
    function New-VersionPublishedReplica {
        # Explicit synthetic consistency fixture; this does not witness T075.
        $root = New-VersionReplica
        $proof = New-VersionPublicationProof
        [IO.File]::WriteAllText((Join-Path $root $proofRelative), (($proof | ConvertTo-Json -Depth 8 -Compress) + "`n"), [Text.UTF8Encoding]::new($false))
        $metadataPath = Join-Path $root 'release-metadata.json'
        $published = Read-VersionJson $metadataPath
        $published.release_state = 'published'; $published.tag = $proof.tag; $published.release_date = '2026-10-07'; $published.release_url = $proof.release_url
        $published.asset_filename = $proof.assets[0].name; $published.asset_bytes = $proof.assets[0].bytes; $published.asset_sha256 = $proof.assets[0].sha256; $published.download_url = $proof.assets[0].download_url
        [IO.File]::WriteAllText($metadataPath, (($published | ConvertTo-Json -Depth 4 -Compress) + "`n"), [Text.UTF8Encoding]::new($false))
        $publishedNotes = [IO.File]::ReadAllText((Join-Path $root 'docs/release/NOTES.md')).Replace("`r`n", "`n").TrimEnd()
        [IO.File]::WriteAllText((Join-Path $root 'CHANGELOG.md'), "# Changelog`n`n<!-- Generated by tools/release/Update-ReleaseMetadata.ps1; edit docs/release/NOTES.md. -->`n`n## 1.0.0 (2026-10-07)`n`n$publishedNotes`n", [Text.UTF8Encoding]::new($false))
        return $root
    }
    . $application
    $version = Get-WinImgVersion
}

Describe 'M4-T01 one application version and witnessed release artifacts' {
    It 'T068 prints the source version through a real positional command with no ImageMagick or output writes' {
        $result = Invoke-VersionChild -Script $application -WithoutMagick
        $result.ExitCode | Should -Be 1
        $result.StdErr | Should -BeNullOrEmpty
        $result.StdOut | Should -Match ([regex]::Escape($version))
        $result.StdOut | Should -Match 'Usage: WinImgNormalizer\.ps1 <sourceFolder> \[maxBytes\]'
        @(Get-ChildItem -LiteralPath $result.Work -Recurse -File -Force).Count | Should -Be 0
    }

    It 'T068 returns versioned usage before native resolution or normalization' {
        Mock Resolve-WinImgMagickApplication { throw 'Usage must not resolve ImageMagick.' }
        Mock Invoke-WinImgNormalizer { throw 'Usage must not start normalization.' }
        $output = @(& Invoke-WinImgNormalizerCommand -Arguments @() 6>&1)
        @($output | Where-Object { $_ -is [int] }).Count | Should -Be 1
        @($output | Where-Object { $_ -is [int] })[0] | Should -Be 1
        ($output -join "`n") | Should -Match ([regex]::Escape($version))
        Should -Invoke Resolve-WinImgMagickApplication -Times 0 -Exactly
        Should -Invoke Invoke-WinImgNormalizer -Times 0 -Exactly
    }

    It 'T068 records the same version in an actual run log while copying synthetic video bytes unchanged' {
        $source = New-VersionDirectory 'synthetic video source'
        $parent = New-VersionDirectory 'synthetic video output'
        $video = Join-Path $source 'synthetic.mp4'
        $bytes = New-Object byte[] 4096
        for ($i = 0; $i -lt $bytes.Length; $i++) { $bytes[$i] = [byte](($i * 29 + 17) % 256) }
        [IO.File]::WriteAllBytes($video, $bytes)
        $fixed = [DateTime]::Parse('2019-04-05T06:07:08Z').ToUniversalTime()
        [IO.File]::SetCreationTimeUtc($video, $fixed); [IO.File]::SetLastWriteTimeUtc($video, $fixed)
        $before = Get-VersionFileState $video
        $output = @(& Invoke-WinImgNormalizer -Source $source -OutputParent $parent -MagickPath $magick -NativeTemporaryRoot $ownedRoot 6>&1)
        @($output | Where-Object { $_ -is [int] }).Count | Should -Be 1
        @($output | Where-Object { $_ -is [int] })[0] | Should -Be 0
        $runs = @(Get-ChildItem -LiteralPath $parent -Directory); $runs.Count | Should -Be 1
        $logs = @(Get-ChildItem -LiteralPath $runs[0].FullName -Recurse -File -Filter '*.log'); $logs.Count | Should -Be 1
        $log = [IO.File]::ReadAllText($logs[0].FullName)
        $log | Should -Match ([regex]::Escape('WinImgNormalizer ' + $version))
        $log | Should -Match 'SUMMARY ConvertedImages=0 CopiedVideos=1 Duplicates=0 Unsupported=0 Errors=0'
        (Get-FileHash -LiteralPath (Join-Path $runs[0].FullName 'synthetic.mp4')).Hash | Should -Be (Get-FileHash -LiteralPath $video).Hash
        Get-VersionFileState $video | Should -Be $before
        # A second actual run rejects a hard-coded log version that would happen
        # to match today's source. Only the version function is substituted.
        Mock Get-WinImgVersion { '9.8.7' }
        $changedParent = New-VersionDirectory 'changed version log output'
        $changedOutput = @(& Invoke-WinImgNormalizer -Source $source -OutputParent $changedParent -MagickPath $magick -NativeTemporaryRoot $ownedRoot 6>&1)
        @($changedOutput | Where-Object { $_ -is [int] })[0] | Should -Be 0
        $changedLogs = @(Get-ChildItem -LiteralPath $changedParent -Recurse -File -Filter '*.log')
        $changedLogs.Count | Should -Be 1
        $changedLog = [IO.File]::ReadAllText($changedLogs[0].FullName)
        $changedLog | Should -Match 'WinImgNormalizer 9\.8\.7 started'
        $changedLog | Should -Not -Match ([regex]::Escape('WinImgNormalizer ' + $version + ' started'))
        Get-VersionFileState $video | Should -Be $before
    }

    It 'T068 checks exact generated artifacts without changing their bytes or timestamps' {
        $before = @(Get-VersionGeneratedState $repository)
        $result = Invoke-VersionChild -Script $generator -Arguments @('-Check') -WithoutMagick
        $result.ExitCode | Should -Be 0
        $result.StdErr | Should -BeNullOrEmpty
        @(Get-VersionGeneratedState $repository) -join '|' | Should -Be ($before -join '|')
        $metadata = Get-Content -LiteralPath (Join-Path $repository 'release-metadata.json') -Raw | ConvertFrom-Json
        $metadata.version | Should -Be $version
        $metadata.release_state | Should -BeIn @('unreleased', 'published')
        if ($metadata.release_state -ceq 'unreleased') {
            foreach ($field in @('tag', 'release_date', 'release_url', 'asset_filename', 'asset_bytes', 'asset_sha256', 'download_url')) { $metadata.$field | Should -BeNullOrEmpty }
        } else {
            $metadata.tag | Should -BeExactly ('v' + $version)
            $metadata.asset_sha256 | Should -BeExactly '251828028e144759c919645f423fde08641cabf42e04db7024d7ad17da4ba14d'
            [IO.File]::Exists((Join-Path $repository $proofRelative)) | Should -BeTrue
        }
        [IO.File]::ReadAllText((Join-Path $repository 'CHANGELOG.md')) | Should -Match ([regex]::Escape($version))
    }

    It 'T068 rejects stale generated metadata and preserves the stale evidence' {
        $replica = New-VersionReplica
        $metadata = Join-Path $replica 'release-metadata.json'
        [IO.File]::AppendAllText($metadata, "`nSTALE VERSION CONTROL`n", [Text.UTF8Encoding]::new($false))
        $before = @(Get-VersionGeneratedState $replica)
        $result = Invoke-VersionChild -Script (Join-Path $replica 'tools/release/Update-ReleaseMetadata.ps1') -Arguments @('-Check') -WithoutMagick
        $result.ExitCode | Should -Not -Be 0
        @($result.StdOut, $result.StdErr) -join "`n" | Should -Match '(?i)(stale|differ|mismatch|regenerat|update)'
        @(Get-VersionGeneratedState $replica) -join '|' | Should -Be ($before -join '|')
    }

    It 'T068 derives usage and both generated artifact versions from a changed source in a disposable project' {
        $replica = New-VersionReplica
        $replicaApplication = Join-Path $replica 'WinImgNormalizer.ps1'
        $text = [IO.File]::ReadAllText($replicaApplication)
        $pattern = '(?s)(function\s+Get-WinImgVersion\s*\{\s*return\s*)''([^'']+)''(\s*\})'
        $matches = [regex]::Matches($text, $pattern)
        $matches.Count | Should -Be 1
        $matches[0].Groups[2].Value | Should -Be $version
        $changedVersion = '9.8.7'
        $literal = $matches[0].Groups[2]
        $changed = $text.Substring(0, $literal.Index) + $changedVersion + $text.Substring($literal.Index + $literal.Length)
        [IO.File]::WriteAllText($replicaApplication, $changed, [Text.UTF8Encoding]::new($true))
        $replicaGenerator = Join-Path $replica 'tools/release/Update-ReleaseMetadata.ps1'
        $stale = Invoke-VersionChild -Script $replicaGenerator -Arguments @('-Check') -WithoutMagick
        $stale.ExitCode | Should -Not -Be 0
        $result = Invoke-VersionChild -Script $replicaGenerator -WithoutMagick
        $result.ExitCode | Should -Be 0
        $result.StdErr | Should -BeNullOrEmpty
        $metadata = Get-Content -LiteralPath (Join-Path $replica 'release-metadata.json') -Raw | ConvertFrom-Json
        $metadata.version | Should -Be $changedVersion
        $changelog = [IO.File]::ReadAllText((Join-Path $replica 'CHANGELOG.md'))
        $changelog | Should -Match ([regex]::Escape($changedVersion))
        $changelog | Should -Not -Match ('(?m)^##.*' + [regex]::Escape($version))
        $usage = Invoke-VersionChild -Script $replicaApplication -WithoutMagick
        $usage.ExitCode | Should -Be 1
        $usage.StdOut | Should -Match ([regex]::Escape($changedVersion))
        $before = @('CHANGELOG.md', 'release-metadata.json' | ForEach-Object { (Get-FileHash -LiteralPath (Join-Path $replica $_)).Hash })
        $repeat = Invoke-VersionChild -Script $replicaGenerator -WithoutMagick
        $repeat.ExitCode | Should -Be 0
        @('CHANGELOG.md', 'release-metadata.json' | ForEach-Object { (Get-FileHash -LiteralPath (Join-Path $replica $_)).Hash }) -join '|' | Should -Be ($before -join '|')
        $check = Invoke-VersionChild -Script $replicaGenerator -Arguments @('-Check') -WithoutMagick
        $check.ExitCode | Should -Be 0
    }

    It 'T068 checks and regenerates a published synthetic witness without resetting its metadata or proof' {
        $root = New-VersionPublishedReplica
        $before = @(Get-VersionGeneratedState $root)
        $script = Join-Path $root 'tools/release/Update-ReleaseMetadata.ps1'
        foreach ($checkMode in @($true, $false)) {
            $arguments = if ($checkMode) { @('-Check') } else { @() }
            $result = Invoke-VersionChild -Script $script -Arguments $arguments -WithoutMagick
            $result.ExitCode | Should -Be 0
            $result.StdErr | Should -BeNullOrEmpty
            $result.StdOut | Should -Match '1\.0\.0 \(2026-10-07\)'
            @(Get-VersionGeneratedState $root) -join '|' | Should -BeExactly ($before -join '|')
        }
    }

    It 'T068 repairs a stale published changelog while preserving publication bytes and the dated heading' {
        $root = New-VersionPublishedReplica
        $before = @('release-metadata.json', $proofRelative | ForEach-Object { Get-VersionFileState (Join-Path $root $_) })
        [IO.File]::WriteAllText((Join-Path $root 'CHANGELOG.md'), 'STALE PUBLISHED CHANGELOG', [Text.UTF8Encoding]::new($false))
        $result = Invoke-VersionChild -Script (Join-Path $root 'tools/release/Update-ReleaseMetadata.ps1') -WithoutMagick
        $result.ExitCode | Should -Be 0
        @('release-metadata.json', $proofRelative | ForEach-Object { Get-VersionFileState (Join-Path $root $_) }) -join '|' | Should -BeExactly ($before -join '|')
        [IO.File]::ReadAllText((Join-Path $root 'CHANGELOG.md')) | Should -Match '(?m)^## 1\.0\.0 \(2026-10-07\)$'
    }

    It 'T068 never promotes an explicit draft merely because a publication proof exists' {
        $root = New-VersionReplica
        [IO.File]::WriteAllText((Join-Path $root $proofRelative), ((New-VersionPublicationProof | ConvertTo-Json -Depth 8 -Compress) + "`n"), [Text.UTF8Encoding]::new($false))
        $result = Invoke-VersionChild -Script (Join-Path $root 'tools/release/Update-ReleaseMetadata.ps1') -Arguments @('-Check') -WithoutMagick
        $result.ExitCode | Should -Be 0
        (Get-Content -LiteralPath (Join-Path $root 'release-metadata.json') -Raw | ConvertFrom-Json).release_state | Should -BeExactly 'unreleased'
    }

    It 'T068 rejects <Mode> in published synthetic records without writing either output or the proof' -ForEach @(
        @{ Mode = 'missing proof' }, @{ Mode = 'malformed proof' }, @{ Mode = 'extra proof field' }, @{ Mode = 'fractional schema' },
        @{ Mode = 'wrong version' }, @{ Mode = 'array version' }, @{ Mode = 'private visibility' }, @{ Mode = 'wrong source' }, @{ Mode = 'boolean release ID' },
        @{ Mode = 'invalid UTC date' }, @{ Mode = 'non UTC date' }, @{ Mode = 'observation before publication' },
        @{ Mode = 'missing asset' }, @{ Mode = 'reordered assets' }, @{ Mode = 'duplicate asset ID' }, @{ Mode = 'extra asset field' },
        @{ Mode = 'string asset bytes' }, @{ Mode = 'wrong asset hash' }, @{ Mode = 'wrong downloaded bytes' }, @{ Mode = 'wrong downloaded hash' }, @{ Mode = 'wrong download URL' },
        @{ Mode = 'canonical hash drift' }, @{ Mode = 'canonical date drift' }, @{ Mode = 'array canonical author' }, @{ Mode = 'extra canonical field' }, @{ Mode = 'invalid canonical state' },
        @{ Mode = 'duplicate proof field' }, @{ Mode = 'unicode alias proof field' }, @{ Mode = 'case alias proof field' }, @{ Mode = 'duplicate canonical field' }
    ) {
        $root = New-VersionPublishedReplica
        $proofPath = Join-Path $root $proofRelative
        $proof = Read-VersionJson $proofPath
        $metadataPath = Join-Path $root 'release-metadata.json'
        $metadata = Read-VersionJson $metadataPath
        switch ($Mode) {
            'extra proof field' { $proof | Add-Member -NotePropertyName extra -NotePropertyValue 'refused' }
            'fractional schema' { $proof.schema_version = 1.5 }
            'wrong version' { $proof.version = '9.8.7' }
            'array version' { $proof.version = @('1.0.0') }
            'private visibility' { $proof.visibility = 'private' }
            'wrong source' { $proof.source_revision = '0' * 40 }
            'boolean release ID' { $proof.release_id = $true }
            'invalid UTC date' { $proof.published_at = '2026-02-30T00:00:00Z' }
            'non UTC date' { $proof.published_at = '2026-10-07T00:00:00+00:00' }
            'observation before publication' { $proof.observed_at = '2026-10-06T23:59:59Z' }
            'missing asset' { $proof.assets = @($proof.assets[0], $proof.assets[1]) }
            'reordered assets' { $proof.assets = @($proof.assets[1], $proof.assets[0], $proof.assets[2]) }
            'duplicate asset ID' { $proof.assets[1].id = $proof.assets[0].id }
            'extra asset field' { $proof.assets[0] | Add-Member -NotePropertyName extra -NotePropertyValue 'refused' }
            'string asset bytes' { $proof.assets[0].bytes = '184659' }
            'wrong asset hash' { $proof.assets[0].sha256 = '0' * 64 }
            'wrong downloaded bytes' { $proof.assets[0].downloaded_bytes = 184660 }
            'wrong downloaded hash' { $proof.assets[0].downloaded_sha256 = '0' * 64 }
            'wrong download URL' { $proof.assets[0].download_url = 'https://example.invalid/unapproved.zip' }
            'canonical hash drift' { $metadata.asset_sha256 = '0' * 64 }
            'canonical date drift' { $metadata.release_date = '2026-10-06' }
            'array canonical author' { $metadata.author = @('Janne Vuorela') }
            'extra canonical field' { $metadata | Add-Member -NotePropertyName extra -NotePropertyValue 'refused' }
            'invalid canonical state' { $metadata.release_state = 'released' }
        }
        [IO.File]::WriteAllText($proofPath, (($proof | ConvertTo-Json -Depth 8 -Compress) + "`n"), [Text.UTF8Encoding]::new($false))
        [IO.File]::WriteAllText($metadataPath, (($metadata | ConvertTo-Json -Depth 8 -Compress) + "`n"), [Text.UTF8Encoding]::new($false))
        $rawProof = [IO.File]::ReadAllText($proofPath)
        switch ($Mode) {
            'duplicate proof field' { [IO.File]::WriteAllText($proofPath, $rawProof.Replace('"source_revision":', '"source_revision":"8edbcbaeb3425ec3a52eeafde212c32553755af1","source_revision":'), [Text.UTF8Encoding]::new($false)) }
            'unicode alias proof field' { [IO.File]::WriteAllText($proofPath, $rawProof.Replace('"source_revision":', '"source_revis\u0069on":"8edbcbaeb3425ec3a52eeafde212c32553755af1","source_revision":'), [Text.UTF8Encoding]::new($false)) }
            'case alias proof field' { [IO.File]::WriteAllText($proofPath, $rawProof.Replace('"source_revision":', '"SOURCE_REVISION":"8edbcbaeb3425ec3a52eeafde212c32553755af1","source_revision":'), [Text.UTF8Encoding]::new($false)) }
            'duplicate canonical field' { [IO.File]::WriteAllText($metadataPath, ([IO.File]::ReadAllText($metadataPath)).Replace('"author":', '"author":"Janne Vuorela","author":'), [Text.UTF8Encoding]::new($false)) }
        }
        if ($Mode -ceq 'malformed proof') { [IO.File]::WriteAllText($proofPath, '{"schema_version":', [Text.UTF8Encoding]::new($false)) }
        if ($Mode -ceq 'missing proof') { [IO.File]::Move($proofPath, (Join-Path $root 'retained-missing-proof.json')) }
        $before = @(Get-VersionGeneratedState $root)
        foreach ($checkMode in @($true, $false)) {
            $arguments = if ($checkMode) { @('-Check') } else { @() }
            $result = Invoke-VersionChild -Script (Join-Path $root 'tools/release/Update-ReleaseMetadata.ps1') -Arguments $arguments -WithoutMagick
            $result.ExitCode | Should -Not -Be 0
            @(Get-VersionGeneratedState $root) -join '|' | Should -BeExactly ($before -join '|')
        }
    }
}

AfterAll {
    if ($bindingsBefore) {
        @(Get-VersionBindings) | ConvertTo-Json -Depth 4 -Compress | Should -Be ($bindingsBefore | ConvertTo-Json -Depth 4 -Compress)
    }
    if ($ownedRoot -and $observations) {
        [IO.File]::WriteAllText((Join-Path $ownedRoot 'observations.json'), ($observations | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))
    }
}
