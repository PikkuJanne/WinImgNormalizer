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
        $tokens = @('-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $Script) + $Arguments
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
    function Get-VersionGeneratedState([string]$Root) {
        foreach ($relative in @('CHANGELOG.md', 'release-metadata.json')) { Get-VersionFileState (Join-Path $Root $relative) }
    }
    function New-VersionReplica {
        $root = New-VersionDirectory 'disposable project with spaces'
        foreach ($relative in $artifactPaths) {
            $target = Join-Path $root $relative
            $null = [IO.Directory]::CreateDirectory((Split-Path -Parent $target))
            [IO.File]::WriteAllBytes($target, [IO.File]::ReadAllBytes((Join-Path $repository $relative)))
        }
        return $root
    }
    . $application
    $version = Get-WinImgVersion
}

Describe 'M4-T01 one application version and unpublished release artifacts' {
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
        $metadata.release_state | Should -Be 'unreleased'
        foreach ($field in @('tag', 'release_date', 'release_url', 'asset_filename', 'asset_bytes', 'asset_sha256', 'download_url')) {
            $metadata.$field | Should -BeNullOrEmpty
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
}

AfterAll {
    if ($bindingsBefore) {
        @(Get-VersionBindings) | ConvertTo-Json -Depth 4 -Compress | Should -Be ($bindingsBefore | ConvertTo-Json -Depth 4 -Compress)
    }
    if ($ownedRoot -and $observations) {
        [IO.File]::WriteAllText((Join-Path $ownedRoot 'observations.json'), ($observations | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))
    }
}
