BeforeAll {
    if ($env:OS -ne 'Windows_NT') { throw 'Packaging smoke tests require actual Windows.' }
    $repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    $scratch = Join-Path $repository '.scratch'
    function Assert-PackageAncestors([string]$Path) {
        $current = [IO.DirectoryInfo]::new([IO.Path]::GetFullPath($Path))
        while ($current) {
            if ($current.Exists -and ($current.Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Packaging fixture crosses a reparse point.' }
            $current = $current.Parent
        }
    }
    Assert-PackageAncestors $scratch
    foreach ($pictures in @([Environment]::GetFolderPath('MyPictures'), (Join-Path $env:USERPROFILE 'Pictures'))) {
        if (-not $pictures) { continue }
        $protected = [IO.Path]::GetFullPath($pictures).TrimEnd('\', '/')
        if ($scratch.Equals($protected, [StringComparison]::OrdinalIgnoreCase) -or
            $scratch.StartsWith($protected + '\', [StringComparison]::OrdinalIgnoreCase) -or
            $protected.StartsWith($scratch + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Packaging fixtures overlap real Pictures.' }
    }
    $git = (Get-Command git -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
    & $git -C $repository check-ignore --quiet --no-index -- (Join-Path $scratch 'packaging-ignore-probe')
    if ($LASTEXITCODE -ne 0) { throw 'Packaging fixtures must already be ignored.' }
    $ownedRoot = Join-Path $scratch ('M4-T02-packaging-' + [Guid]::NewGuid().ToString('N'))
    if (Test-Path -LiteralPath $ownedRoot) { throw 'Packaging ownership collision.' }
    $null = [IO.Directory]::CreateDirectory($ownedRoot)
    [IO.File]::WriteAllText((Join-Path $ownedRoot '.winimg-fixture-root'), 'M4-T02 owned synthetic packaging fixtures; retained for inspection; no real Pictures or recursive cleanup.', [Text.UTF8Encoding]::new($false))
    $gitConfiguration = Join-Path $ownedRoot 'empty-git-config'
    [IO.File]::WriteAllText($gitConfiguration, '', [Text.UTF8Encoding]::new($false))
    $observations = New-Object 'Collections.Generic.List[object]'
    $allowlist = @(
        [pscustomobject]@{ path = 'CHANGELOG.md'; source_path = 'CHANGELOG.md' },
        [pscustomobject]@{ path = 'GETTING_STARTED.md'; source_path = 'docs/release/GETTING_STARTED.md' },
        [pscustomobject]@{ path = 'LICENSE'; source_path = 'LICENSE' },
        [pscustomobject]@{ path = 'LICENSE-CC0.txt'; source_path = 'tests/fixtures/colour/LICENSE-CC0.txt' },
        [pscustomobject]@{ path = 'THIRD_PARTY_NOTICES.md'; source_path = 'docs/release/THIRD_PARTY_NOTICES.md' },
        [pscustomobject]@{ path = 'WinImgNormalizer.bat'; source_path = 'WinImgNormalizer.bat' },
        [pscustomobject]@{ path = 'WinImgNormalizer.ps1'; source_path = 'WinImgNormalizer.ps1' }
    )
    $replicaPaths = @($allowlist | ForEach-Object { $_.source_path }) + @('tools/release/Build-Release.ps1', 'tools/release/Update-ReleaseMetadata.ps1', 'docs/release/NOTES.md', 'release-metadata.json', '.gitattributes', '.gitignore')
    $bindingsBefore = @($replicaPaths + @('tests/Packaging.Tests.ps1', 'tests/fixtures/LauncherConsoleFixture.cs', 'tests/legacy/toolchain.json') | Select-Object -Unique | ForEach-Object {
        [pscustomobject]@{ Path = $_; Sha256 = (Get-FileHash -LiteralPath (Join-Path $repository $_)).Hash.ToLowerInvariant() }
    })
    if (-not $env:WINIMG_TEST_MAGICK -or -not [IO.Path]::IsPathRooted($env:WINIMG_TEST_MAGICK)) { throw 'Explicit verified ImageMagick required.' }
    $magick = [IO.Path]::GetFullPath($env:WINIMG_TEST_MAGICK)
    Assert-PackageAncestors $magick
    $imagePin = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'legacy/toolchain.json') -Raw | ConvertFrom-Json
    (Get-FileHash -LiteralPath $magick).Hash.ToLowerInvariant() | Should -Be $imagePin.portable_imagemagick.executable_sha256
    $hostExecutable = (Get-Process -Id $PID).Path
    $legacyHost = Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe'
    $cmd = Join-Path $env:SystemRoot 'System32/cmd.exe'
    function New-PackageDirectory([string]$Label) {
        $path = Join-Path $ownedRoot ($Label + '-' + [Guid]::NewGuid().ToString('N'))
        Assert-PackageAncestors $path
        if (Test-Path -LiteralPath $path) { throw 'Packaging fixture collision.' }
        $null = [IO.Directory]::CreateDirectory($path)
        return $path
    }
    function ConvertTo-PackageArgument([string]$Value) {
        $escaped = [regex]::Replace($Value, '(\\*)"', '$1$1\"')
        $escaped = [regex]::Replace($escaped, '(\\+)$', '$1$1')
        return '"' + $escaped + '"'
    }
    function New-PackageProcessStart([string]$Executable, [string[]]$Arguments, [string]$WorkingDirectory, [switch]$PreserveGitConfiguration) {
        Assert-PackageAncestors $WorkingDirectory
        $start = New-Object Diagnostics.ProcessStartInfo
        $start.FileName = $Executable
        $start.Arguments = ($Arguments | ForEach-Object { ConvertTo-PackageArgument ([string]$_) }) -join ' '
        $start.WorkingDirectory = $WorkingDirectory
        $start.UseShellExecute = $false; $start.CreateNoWindow = $true
        $start.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
        $start.RedirectStandardOutput = $true; $start.RedirectStandardError = $true
        foreach ($key in @($start.EnvironmentVariables.Keys)) {
            if ((-not $PreserveGitConfiguration -and [string]$key -like 'GIT_*') -or [string]::Equals([string]$key, 'PSModulePath', [StringComparison]::OrdinalIgnoreCase)) { $start.EnvironmentVariables.Remove([string]$key) }
        }
        if (-not $PreserveGitConfiguration) {
            $start.EnvironmentVariables['GIT_CONFIG_NOSYSTEM'] = '1'
            $start.EnvironmentVariables['GIT_CONFIG_GLOBAL'] = $gitConfiguration
        }
        $start.EnvironmentVariables['PATH'] = [IO.Path]::GetDirectoryName($legacyHost) + ';' + [IO.Path]::GetDirectoryName($magick) + ';' + $env:PATH
        $start.EnvironmentVariables['TEMP'] = $ownedRoot; $start.EnvironmentVariables['TMP'] = $ownedRoot
        $start.EnvironmentVariables['MAGICK_TEMPORARY_PATH'] = $ownedRoot
        return $start
    }
    function Invoke-PackageProcess($Start, [switch]$Binary) {
        $process = New-Object Diagnostics.Process
        $process.StartInfo = $Start
        $buffer = New-Object IO.MemoryStream
        $started = $false
        try {
            if (-not $process.Start()) { throw 'Owned packaging process did not start.' }
            $started = $true
            $stdout = if ($Binary) { $process.StandardOutput.BaseStream.CopyToAsync($buffer) } else { $process.StandardOutput.ReadToEndAsync() }
            $stderr = $process.StandardError.ReadToEndAsync()
            if (-not $process.WaitForExit(30000)) { throw 'Owned packaging child exceeded its 30-second bound.' }
            if (-not $stdout.Wait(5000) -or -not $stderr.Wait(5000)) { throw 'Owned packaging streams did not finish.' }
            $result = [pscustomobject]@{ ExitCode = $process.ExitCode; StdOut = $(if ($Binary) { $null } else { $stdout.GetAwaiter().GetResult() }); StdErr = $stderr.GetAwaiter().GetResult(); Bytes = $(if ($Binary) { $buffer.ToArray() } else { $null }) }
            if (-not $Binary) { $observations.Add([pscustomobject]@{ Executable = $Start.FileName; Work = $Start.WorkingDirectory; Result = $result }) }
            return $result
        } finally {
            if ($started -and -not $process.HasExited) {
                & (Join-Path $env:SystemRoot 'System32/taskkill.exe') /PID $process.Id /T /F 2>&1 | Out-Null
                if (-not $process.WaitForExit(5000)) { $process.Kill(); $null = $process.WaitForExit(5000) }
            }
            $process.Dispose(); $buffer.Dispose()
        }
    }
    function Invoke-PackageGit([string]$Root, [string[]]$Arguments, [switch]$Binary) {
        $actualRepository = [string]::Equals($Root, $repository, [StringComparison]::OrdinalIgnoreCase)
        $result = Invoke-PackageProcess (New-PackageProcessStart $git (@('-C', $Root) + $Arguments) $ownedRoot -PreserveGitConfiguration:$actualRepository) -Binary:$Binary
        if ($result.ExitCode -ne 0) { throw ('Owned packaging Git failed: ' + $result.StdErr) }
        return $result
    }
    function Commit-PackageReplica([string]$Root) {
        $null = Invoke-PackageGit $Root @('-c', 'user.name=WinImg Synthetic Test', '-c', 'user.email=synthetic@example.invalid', '-c', 'commit.gpgsign=false', 'commit', '-m', 'Owned synthetic packaging control')
        return (Invoke-PackageGit $Root @('rev-parse', 'HEAD')).StdOut.Trim()
    }
    function New-PackageReplica {
        $root = New-PackageDirectory 'disposable clean project with spaces'
        foreach ($relative in $replicaPaths) {
            $target = Join-Path $root $relative
            $null = [IO.Directory]::CreateDirectory((Split-Path -Parent $target))
            [IO.File]::WriteAllBytes($target, [IO.File]::ReadAllBytes((Join-Path $repository $relative)))
        }
        [IO.File]::WriteAllText((Join-Path $root 'unrelated-tracked.txt'), 'Synthetic tracked exclusion control.', [Text.UTF8Encoding]::new($false))
        $null = Invoke-PackageGit $root @('init', '-b', 'packaging-regression')
        $null = Invoke-PackageGit $root @('config', '--local', 'core.symlinks', 'false')
        $null = Invoke-PackageGit $root (@('-c', 'core.autocrlf=false', 'add', '--') + $replicaPaths + @('unrelated-tracked.txt'))
        $revision = Commit-PackageReplica $root
        $null = [IO.Directory]::CreateDirectory((Join-Path $root '.scratch/private build debris'))
        [IO.File]::WriteAllText((Join-Path $root '.scratch/private build debris/credentials.txt'), 'synthetic-private-control-secret', [Text.UTF8Encoding]::new($false))
        $status = Invoke-PackageGit $root @('status', '--porcelain=v1', '--untracked-files=all')
        $status.StdOut | Should -BeNullOrEmpty
        return [pscustomobject]@{ Root = $root; Revision = $revision }
    }
    $buildDriver = Join-Path $ownedRoot 'invoke-owned-build.ps1'
    [IO.File]::WriteAllText($buildDriver, @'
param([string]$Root, [string]$Revision, [string]$OutputDirectory, [string]$Record)
$ErrorActionPreference = 'Stop'
try {
    $result = @(& (Join-Path $Root 'tools/release/Build-Release.ps1') -Revision $Revision -OutputDirectory $OutputDirectory)
    if ($result.Count -ne 1) { throw 'Builder must return exactly one result object.' }
    [IO.File]::WriteAllText($Record, ($result[0] | ConvertTo-Json -Depth 8), [Text.UTF8Encoding]::new($false))
    exit 0
} catch { Write-Error $_; exit 1 }
'@, [Text.UTF8Encoding]::new($false))
    function Invoke-PackageBuild($Replica, [string]$OutputDirectory, [string]$Revision) {
        if (-not $OutputDirectory) { $OutputDirectory = Join-Path $Replica.Root ('.scratch/package-' + [Guid]::NewGuid().ToString('N')) }
        if (-not $Revision) { $Revision = $Replica.Revision }
        $work = New-PackageDirectory 'build child work'
        $record = Join-Path $work 'builder-result.json'
        $tokens = @('-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $buildDriver, $Replica.Root, $Revision, $OutputDirectory, $record)
        $actualRepository = [string]::Equals($Replica.Root, $repository, [StringComparison]::OrdinalIgnoreCase)
        $result = Invoke-PackageProcess (New-PackageProcessStart $hostExecutable $tokens $work -PreserveGitConfiguration:$actualRepository)
        return [pscustomobject]@{ ExitCode = $result.ExitCode; StdOut = $result.StdOut; StdErr = $result.StdErr; OutputDirectory = $OutputDirectory; Result = $(if ([IO.File]::Exists($record)) { Get-Content -LiteralPath $record -Raw | ConvertFrom-Json } else { $null }) }
    }
    function Get-PackageSha256([byte[]]$Bytes) {
        $hash = [Security.Cryptography.SHA256]::Create()
        try { return ([BitConverter]::ToString($hash.ComputeHash($Bytes))).Replace('-', '').ToLowerInvariant() } finally { $hash.Dispose() }
    }
    function Get-PackageFileState([string]$Path) {
        $item = Get-Item -LiteralPath $Path -Force
        return ([ordered]@{ hash = (Get-FileHash -LiteralPath $Path).Hash; bytes = $item.Length; created = $item.CreationTimeUtc.Ticks; modified = $item.LastWriteTimeUtc.Ticks; attributes = [int]$item.Attributes } | ConvertTo-Json -Compress)
    }
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    function Read-PackageArchive([string]$Path) {
        $archive = [IO.Compression.ZipFile]::OpenRead($Path)
        try {
            return @($archive.Entries | ForEach-Object {
                $entry = $_; $stream = $entry.Open(); $buffer = New-Object IO.MemoryStream
                try { $stream.CopyTo($buffer); $bytes = $buffer.ToArray() } finally { $stream.Dispose(); $buffer.Dispose() }
                [pscustomobject]@{ Path = $entry.FullName; Bytes = $bytes; Length = $entry.Length; CompressedLength = $entry.CompressedLength; LastWriteTime = $entry.LastWriteTime; Attributes = $entry.ExternalAttributes }
            })
        } finally { $archive.Dispose() }
    }
    function Get-PackageSecurityState {
        return ([ordered]@{
            policies = @(Get-ExecutionPolicy -List | Where-Object { $_.Scope -ne 'Process' } | ForEach-Object { [ordered]@{ scope = $_.Scope.ToString(); policy = $_.ExecutionPolicy.ToString() } })
            user_path = [Environment]::GetEnvironmentVariable('PATH', 'User')
            machine_path = [Environment]::GetEnvironmentVariable('PATH', 'Machine')
            process_path = $env:PATH
            user_profile = $env:USERPROFILE
            process_policy = (Get-ExecutionPolicy -Scope Process).ToString()
        } | ConvertTo-Json -Depth 6 -Compress)
    }
    $securityBefore = Get-PackageSecurityState
    $replica = New-PackageReplica
    $built = Invoke-PackageBuild $replica
    if ($built.ExitCode -ne 0) { throw ('Initial synthetic portable build failed: ' + $built.StdOut + $built.StdErr) }
    $assetFilename = 'WinImgNormalizer-1.0.0-portable.zip'
    $zipPath = Join-Path $built.OutputDirectory $assetFilename
    $entries = @(Read-PackageArchive $zipPath)
    $manifestEntry = @($entries | Where-Object { $_.Path -ceq 'package-manifest.json' })
    $manifest = [Text.Encoding]::UTF8.GetString($manifestEntry[0].Bytes) | ConvertFrom-Json
    $provenancePath = Join-Path $built.OutputDirectory 'build-provenance.json'
    $provenance = Get-Content -LiteralPath $provenancePath -Raw | ConvertFrom-Json
    $compiler = Join-Path $env:SystemRoot 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
    $consoleFixture = Join-Path $ownedRoot 'LauncherConsoleFixture.exe'
    $fixtureSource = Join-Path $PSScriptRoot 'fixtures/LauncherConsoleFixture.cs'
    $compile = Invoke-PackageProcess (New-PackageProcessStart $compiler @('/nologo', '/target:winexe', '/reference:System.Web.Extensions.dll', ('/out:' + $consoleFixture), $fixtureSource) $ownedRoot)
    if ($compile.ExitCode -ne 0 -or -not [IO.File]::Exists($consoleFixture)) { throw 'Packaging private-console fixture did not compile.' }
    $fixtureBinding = [ordered]@{ SourceSha256 = (Get-FileHash -LiteralPath $fixtureSource).Hash.ToLowerInvariant(); CompilerSha256 = (Get-FileHash -LiteralPath $compiler).Hash.ToLowerInvariant(); ExecutableSha256 = (Get-FileHash -LiteralPath $consoleFixture).Hash.ToLowerInvariant() }
    # Final clean gates exercise the actual implementation commit. Disposable
    # replicas keep dirty developer-focused runs useful without claiming that
    # those synthetic revision hashes are the distribution provenance.
    $actualStatus = Invoke-PackageGit $repository @('status', '--porcelain=v1', '--untracked-files=all')
    $smokeIsActualRevision = [string]::IsNullOrWhiteSpace($actualStatus.StdOut)
    if (-not $smokeIsActualRevision -and ($env:GITHUB_ACTIONS -eq 'true' -or $env:WINIMG_TEST_REQUIRE_ACTUAL_PACKAGE -eq '1')) {
        throw 'Acceptance packaging smoke requires the actual clean implementation revision; a synthetic replica fallback is not allowed.'
    }
    $smokeBuild = $built
    if ($smokeIsActualRevision) {
        $actualRevision = (Invoke-PackageGit $repository @('rev-parse', 'HEAD')).StdOut.Trim()
        $actualReplica = [pscustomobject]@{ Root = $repository; Revision = $actualRevision }
        $smokeBuild = Invoke-PackageBuild $actualReplica -OutputDirectory (Join-Path $ownedRoot 'actual clean revision package')
        if ($smokeBuild.ExitCode -ne 0) { throw ('Actual clean implementation portable build failed: ' + $smokeBuild.StdOut + $smokeBuild.StdErr) }
    }
    $smokeZipPath = Join-Path $smokeBuild.OutputDirectory $assetFilename
    $smokeEntries = @(Read-PackageArchive $smokeZipPath)
    $smokeProvenance = Get-Content -LiteralPath (Join-Path $smokeBuild.OutputDirectory 'build-provenance.json') -Raw | ConvertFrom-Json
    $observations.Add([pscustomobject]@{
        Kind = 'extraction source binding'; ActualCleanImplementationRevision = $smokeIsActualRevision
        SourceRevision = $smokeProvenance.source_revision; AssetSha256 = $smokeProvenance.asset_sha256
        EntryHashes = @($smokeEntries | ForEach-Object { [ordered]@{ Path = $_.Path; Bytes = $_.Length; Sha256 = Get-PackageSha256 $_.Bytes } })
        DeveloperReplicaFallback = (-not $smokeIsActualRevision)
    })
    $extractParent = New-PackageDirectory 'fresh extraction with spaces'
    $extracted = Join-Path $extractParent 'portable application with spaces'
    Assert-PackageAncestors $extracted
    [IO.Compression.ZipFile]::ExtractToDirectory($smokeZipPath, $extracted)
    function Invoke-PackagedBat([string[]]$Arguments = @()) {
        $work = New-PackageDirectory 'real packaged BAT child'
        $record = Join-Path $work 'console-observation.json'
        $start = New-PackageProcessStart $consoleFixture @() $work
        $start.RedirectStandardInput = $true
        $start.EnvironmentVariables['WINIMG_TEST_PACKAGED_BAT'] = Join-Path $extracted 'WinImgNormalizer.bat'
        $command = '"%WINIMG_TEST_PACKAGED_BAT%"'
        for ($i = 0; $i -lt $Arguments.Count; $i++) {
            $key = 'WINIMG_TEST_PACKAGED_ARG' + $i
            $start.EnvironmentVariables[$key] = $Arguments[$i]
            $command += ' "%' + $key + '%"'
        }
        $start.EnvironmentVariables['WINIMG_TEST_PRIVATE_CONSOLE_CMD'] = $cmd
        $start.EnvironmentVariables['WINIMG_TEST_PRIVATE_CONSOLE_ARGUMENTS'] = '/d /v:off /s /c "' + $command + '"'
        $start.EnvironmentVariables['WINIMG_TEST_PRIVATE_CONSOLE_RECORD'] = $record
        $process = New-Object Diagnostics.Process; $process.StartInfo = $start
        $started = $false; $text = New-Object Text.StringBuilder
        try {
            if (-not $process.Start()) { throw 'Packaged BAT console process did not start.' }
            $started = $true; $stderr = $process.StandardError.ReadToEndAsync()
            $clock = [Diagnostics.Stopwatch]::StartNew()
            $expected = if ($Arguments.Count -eq 0) { 'ERROR:.*exactly one FOLDER' } else { 'ERROR: Processing did not complete reliably' }
            do {
                $lineTask = $process.StandardOutput.ReadLineAsync()
                if (-not $lineTask.Wait([Math]::Max(1, 20000 - [int]$clock.ElapsedMilliseconds))) { throw 'Packaged BAT did not reach its setup result.' }
                $line = $lineTask.GetAwaiter().GetResult()
                if ($null -eq $line) { throw 'Packaged BAT exited before its setup result.' }
                $null = $text.AppendLine($line)
            } while ($line -notmatch $expected -and $clock.ElapsedMilliseconds -lt 20000)
            if ($line -notmatch $expected) { throw 'Packaged BAT setup result was not observed.' }
            $pauseObserved = -not $process.WaitForExit(200)
            $stdout = $process.StandardOutput.ReadToEndAsync()
            $process.StandardInput.WriteLine('x'); $process.StandardInput.Flush(); $process.StandardInput.Close()
            if (-not $process.WaitForExit(15000)) { throw 'Packaged BAT exceeded its completion bound.' }
            if (-not $stdout.Wait(5000) -or -not $stderr.Wait(5000)) { throw 'Packaged BAT streams did not finish.' }
            $null = $text.Append($stdout.GetAwaiter().GetResult())
            $console = Get-Content -LiteralPath $record -Raw | ConvertFrom-Json
            $console.HarnessError | Should -BeNullOrEmpty
            $console.PrivateConsoleAllocated | Should -BeTrue
            $console.PrivateJobAssignedBeforeCmd | Should -BeTrue
            $console.KeyRecordsWritten | Should -Be 2
            @($console.MembersBeforeKey).Count | Should -Be 2
            $console.ExitCode | Should -Be $process.ExitCode
            $result = [pscustomobject]@{ ExitCode = $process.ExitCode; StdOut = $text.ToString(); StdErr = $stderr.GetAwaiter().GetResult(); PauseObserved = $pauseObserved; Console = $console }
            $observations.Add([pscustomobject]@{ Kind = 'exact extracted production BAT setup only'; Arguments = $Arguments; Result = $result })
            return $result
        } finally {
            if ($started -and -not $process.HasExited) { $process.Kill(); $null = $process.WaitForExit(5000) }
            $process.Dispose()
        }
    }
}

Describe 'M4-T02 portable release content and revision refusal' {
    It 'T069 includes exactly the seven approved Git inputs and generated manifest with canonical bytes' {
        $expectedNames = @($allowlist.path + @('package-manifest.json'))
        [Array]::Sort($expectedNames, [StringComparer]::Ordinal)
        @($entries.Path) -join '|' | Should -BeExactly ($expectedNames -join '|')
        @($manifest.files).Count | Should -Be 7
        @($manifest.files.path) -join '|' | Should -BeExactly ($allowlist.path -join '|')
        $manifest.schema_version | Should -Be 1
        $manifest.product | Should -BeExactly 'WinImgNormalizer'
        $manifest.version | Should -BeExactly '1.0.0'
        $manifest.release_state | Should -BeExactly 'unreleased'
        $manifest.source_revision | Should -BeExactly $replica.Revision
        $manifest.digital_signature | Should -BeNullOrEmpty
        @($manifest.PSObject.Properties.Name) -join '|' | Should -BeExactly 'schema_version|product|version|release_state|source_revision|digital_signature|files'
        foreach ($approved in $allowlist) {
            $entry = @($entries | Where-Object { $_.Path -ceq $approved.path })[0]
            $blob = (Invoke-PackageGit $replica.Root @('show', ($replica.Revision + ':' + $approved.source_path)) -Binary).Bytes
            (Get-PackageSha256 $entry.Bytes) | Should -BeExactly (Get-PackageSha256 $blob)
            $entry.Length | Should -Be $blob.Length
            $file = @($manifest.files | Where-Object { $_.path -ceq $approved.path })[0]
            @($file.PSObject.Properties.Name) -join '|' | Should -BeExactly 'path|source_path|bytes|sha256'
            $file.source_path | Should -BeExactly $approved.source_path
            $file.bytes | Should -Be $blob.Length
            $file.sha256 | Should -BeExactly (Get-PackageSha256 $blob)
        }
        $batText = [Text.Encoding]::ASCII.GetString(@($entries | Where-Object { $_.Path -ceq 'WinImgNormalizer.bat' })[0].Bytes)
        $batText | Should -Match "`r`n"
        $batText.Replace("`r`n", '') | Should -Not -Match "`n"
        foreach ($entry in $entries) { [Text.Encoding]::UTF8.GetString($entry.Bytes) | Should -Not -Match 'synthetic-private-control-secret|Synthetic tracked exclusion control' }
        @($entries.Path) -join '|' | Should -Not -Match '(?i)(\.git|tests/|tools/|credentials|\.scratch|AGENTS|codex-winimg|magick\.exe)'
    }

    It 'T069 refuses dirty <Label> without creating an output directory' -ForEach @(
        @{ Label = 'tracked file'; Mode = 'tracked' }, @{ Label = 'untracked file'; Mode = 'untracked' }, @{ Label = 'staged tracked file'; Mode = 'staged' }
    ) {
        $case = New-PackageReplica
        $path = Join-Path $case.Root $(if ($Mode -eq 'untracked') { 'untracked-control.txt' } else { 'unrelated-tracked.txt' })
        [IO.File]::AppendAllText($path, ' Synthetic dirty control.', [Text.UTF8Encoding]::new($false))
        if ($Mode -eq 'staged') { $null = Invoke-PackageGit $case.Root @('add', '--', 'unrelated-tracked.txt') }
        $before = Get-PackageFileState $path
        $result = Invoke-PackageBuild $case
        $result.ExitCode | Should -Not -Be 0
        @($result.StdOut, $result.StdErr) -join ' ' | Should -Match '(?i)(clean|dirty|untracked|working tree|worktree)'
        Test-Path -LiteralPath $result.OutputDirectory | Should -BeFalse
        Get-PackageFileState $path | Should -BeExactly $before
    }

    It 'T069 refuses an old clean revision instead of silently packaging current HEAD' {
        $case = New-PackageReplica
        [IO.File]::AppendAllText((Join-Path $case.Root 'unrelated-tracked.txt'), ' New synthetic commit.', [Text.UTF8Encoding]::new($false))
        $null = Invoke-PackageGit $case.Root @('add', '--', 'unrelated-tracked.txt')
        $newRevision = Commit-PackageReplica $case.Root
        $newRevision | Should -Not -BeExactly $case.Revision
        $result = Invoke-PackageBuild $case
        $result.ExitCode | Should -Not -Be 0
        Test-Path -LiteralPath $result.OutputDirectory | Should -BeFalse
    }

    It 'T069 refuses noncanonical revision <Revision>' -ForEach @(
        @{ Revision = 'HEAD' }, @{ Revision = '0123456' }, @{ Revision = 'ABCDEF0123456789ABCDEF0123456789ABCDEF0123' }
    ) {
        $result = Invoke-PackageBuild $replica -Revision $Revision
        $result.ExitCode | Should -Not -Be 0
        Test-Path -LiteralPath $result.OutputDirectory | Should -BeFalse
    }

    It 'T069 refuses a committed missing mandatory input without publishing a partial ZIP' {
        $case = New-PackageReplica
        $null = Invoke-PackageGit $case.Root @('rm', '--', 'LICENSE')
        $case.Revision = Commit-PackageReplica $case.Root
        $result = Invoke-PackageBuild $case
        $result.ExitCode | Should -Not -Be 0
        Test-Path -LiteralPath $result.OutputDirectory | Should -BeFalse
    }

    It 'T069 refuses committed <Label> drift instead of packaging stale release claims' -ForEach @(
        @{ Label = 'generated changelog'; Mode = 'changelog' }, @{ Label = 'version metadata'; Mode = 'metadata' }
    ) {
        $case = New-PackageReplica
        if ($Mode -eq 'changelog') {
            $relative = 'CHANGELOG.md'
            [IO.File]::AppendAllText((Join-Path $case.Root $relative), "`nSynthetic stale changelog control.`n", [Text.UTF8Encoding]::new($false))
        } else {
            $relative = 'release-metadata.json'
            $path = Join-Path $case.Root $relative
            $text = [IO.File]::ReadAllText($path).Replace('"version":"1.0.0"', '"version":"9.8.7"')
            [IO.File]::WriteAllText($path, $text, [Text.UTF8Encoding]::new($false))
        }
        $null = Invoke-PackageGit $case.Root @('add', '--', $relative)
        $case.Revision = Commit-PackageReplica $case.Root
        $before = Get-PackageFileState (Join-Path $case.Root $relative)
        $result = Invoke-PackageBuild $case
        $result.ExitCode | Should -Not -Be 0
        Test-Path -LiteralPath $result.OutputDirectory | Should -BeFalse
        Get-PackageFileState (Join-Path $case.Root $relative) | Should -BeExactly $before
    }

    It 'T069 derives the portable filename and both manifests from a changed committed source version' {
        $case = New-PackageReplica
        $path = Join-Path $case.Root 'WinImgNormalizer.ps1'
        $text = [IO.File]::ReadAllText($path)
        $matches = [regex]::Matches($text, '(?s)(function\s+Get-WinImgVersion\s*\{\s*return\s*)''([^'']+)''(\s*\})')
        $matches.Count | Should -Be 1
        $literal = $matches[0].Groups[2]
        $literal.Value | Should -BeExactly '1.0.0'
        $changed = $text.Substring(0, $literal.Index) + '9.8.7' + $text.Substring($literal.Index + $literal.Length)
        [IO.File]::WriteAllText($path, $changed, [Text.UTF8Encoding]::new($true))
        $tokens = @('-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', (Join-Path $case.Root 'tools/release/Update-ReleaseMetadata.ps1'))
        $generated = Invoke-PackageProcess (New-PackageProcessStart $hostExecutable $tokens $ownedRoot)
        $generated.ExitCode | Should -Be 0
        $generated.StdErr | Should -BeNullOrEmpty
        $null = Invoke-PackageGit $case.Root @('add', '--', 'WinImgNormalizer.ps1', 'CHANGELOG.md', 'release-metadata.json')
        $case.Revision = Commit-PackageReplica $case.Root
        $result = Invoke-PackageBuild $case
        $result.ExitCode | Should -Be 0
        $result.Result.version | Should -BeExactly '9.8.7'
        $result.Result.asset_filename | Should -BeExactly 'WinImgNormalizer-9.8.7-portable.zip'
        $changedEntries = @(Read-PackageArchive (Join-Path $result.OutputDirectory $result.Result.asset_filename))
        $changedManifest = [Text.Encoding]::UTF8.GetString(@($changedEntries | Where-Object { $_.Path -ceq 'package-manifest.json' })[0].Bytes) | ConvertFrom-Json
        $changedManifest.version | Should -BeExactly '9.8.7'
        $changedManifest.source_revision | Should -BeExactly $case.Revision
        $changedProvenance = Get-Content -LiteralPath (Join-Path $result.OutputDirectory 'build-provenance.json') -Raw | ConvertFrom-Json
        $changedProvenance.version | Should -BeExactly '9.8.7'
        $changedProvenance.source_revision | Should -BeExactly $case.Revision
        $changedProvenance.asset_filename | Should -BeExactly $result.Result.asset_filename
    }

    It 'T069 refuses a committed Git symbolic-link mode even when checkout bytes look like a file' {
        $case = New-PackageReplica
        $blob = (Invoke-PackageGit $case.Root @('rev-parse', ($case.Revision + ':LICENSE'))).StdOut.Trim()
        $null = Invoke-PackageGit $case.Root @('update-index', '--cacheinfo', ('120000,' + $blob + ',LICENSE'))
        $case.Revision = Commit-PackageReplica $case.Root
        # Git core.symlinks=false materializes this committed mode as ordinary
        # bytes on Windows, proving rejection relies on reviewed tree modes.
        $null = Invoke-PackageGit $case.Root @('-c', 'core.symlinks=false', 'checkout-index', '-f', '--', 'LICENSE')
        (Invoke-PackageGit $case.Root @('status', '--porcelain=v1', '--untracked-files=all')).StdOut | Should -BeNullOrEmpty
        $result = Invoke-PackageBuild $case
        $result.ExitCode | Should -Not -Be 0
        Test-Path -LiteralPath $result.OutputDirectory | Should -BeFalse
    }

    It 'T069 refuses an existing output and preserves every built byte and timestamp' {
        $before = @(Get-ChildItem -LiteralPath $built.OutputDirectory -File | Sort-Object Name | ForEach-Object { Get-PackageFileState $_.FullName })
        $result = Invoke-PackageBuild $replica -OutputDirectory $built.OutputDirectory
        $result.ExitCode | Should -Not -Be 0
        @(Get-ChildItem -LiteralPath $built.OutputDirectory -File | Sort-Object Name | ForEach-Object { Get-PackageFileState $_.FullName }) -join '|' | Should -BeExactly ($before -join '|')
    }

    It 'T069 refuses output <Label> outside a new ignored scratch child' -ForEach @(
        @{ Label = 'at the scratch root'; Mode = 'root' }, @{ Label = 'in nonignored repository content'; Mode = 'nonignored' }, @{ Label = 'outside the replica repository'; Mode = 'outside' }
    ) {
        $path = switch ($Mode) {
            'root' { Join-Path $replica.Root '.scratch' }
            'nonignored' { Join-Path $replica.Root 'unexpected-output' }
            'outside' { Join-Path $ownedRoot ('outside-output-' + [Guid]::NewGuid().ToString('N')) }
        }
        $existed = Test-Path -LiteralPath $path
        $result = Invoke-PackageBuild $replica -OutputDirectory $path
        $result.ExitCode | Should -Not -Be 0
        Test-Path -LiteralPath $path | Should -Be $existed
    }

    It 'T069 refuses a junction output ancestor and leaves its owned target empty' {
        $case = New-PackageReplica
        $target = New-PackageDirectory 'owned junction target'
        $junction = Join-Path $case.Root '.scratch/output-junction'
        $null = New-Item -ItemType Junction -Path $junction -Value $target -ErrorAction Stop
        (Get-Item -LiteralPath $junction -Force).Attributes -band [IO.FileAttributes]::ReparsePoint | Should -Not -Be 0
        $result = Invoke-PackageBuild $case -OutputDirectory (Join-Path $junction 'refused-release')
        $result.ExitCode | Should -Not -Be 0
        @(Get-ChildItem -LiteralPath $target -Force).Count | Should -Be 0
    }
}

Describe 'M4-T02 checksum and reproducibility evidence' {
    It 'T071 binds the exact unsigned asset and builder to the clean recorded revision' {
        @(Get-ChildItem -LiteralPath $built.OutputDirectory -File | Select-Object -ExpandProperty Name | Sort-Object) -join '|' | Should -BeExactly (($assetFilename, 'build-provenance.json', 'SHA256SUMS.txt' | Sort-Object) -join '|')
        $provenance.schema_version | Should -Be 1
        $provenance.product | Should -BeExactly 'WinImgNormalizer'
        $provenance.version | Should -BeExactly $manifest.version
        $provenance.release_state | Should -BeExactly 'unreleased'
        $provenance.source_revision | Should -BeExactly $replica.Revision
        $provenance.digital_signature | Should -BeNullOrEmpty
        $provenance.asset_filename | Should -BeExactly $assetFilename
        $provenance.asset_bytes | Should -Be (Get-Item -LiteralPath $zipPath).Length
        $provenance.asset_sha256 | Should -BeExactly (Get-FileHash -LiteralPath $zipPath).Hash.ToLowerInvariant()
        $provenance.package_manifest_sha256 | Should -BeExactly (Get-PackageSha256 $manifestEntry[0].Bytes)
        $builderBlob = (Invoke-PackageGit $replica.Root @('show', ($replica.Revision + ':tools/release/Build-Release.ps1')) -Binary).Bytes
        $provenance.builder_source_sha256 | Should -BeExactly (Get-PackageSha256 $builderBlob)
        @($provenance.PSObject.Properties.Name) -join '|' | Should -BeExactly 'schema_version|product|version|release_state|source_revision|digital_signature|asset_filename|asset_bytes|asset_sha256|package_manifest_sha256|builder_source_sha256|reproducibility'
        foreach ($property in $provenance.PSObject.Properties) { $property.Name | Should -Not -Match '(?i)(path|timestamp|built_at|user|host)' }
        $text = [IO.File]::ReadAllText($provenancePath)
        $text | Should -Not -Match ([regex]::Escape($ownedRoot))
        $text | Should -Not -Match '[A-Za-z]:[\\/]|\\\\|synthetic-private-control-secret'
        $built.Result.source_revision | Should -BeExactly $replica.Revision
        $built.Result.version | Should -BeExactly $manifest.version
        $built.Result.asset_filename | Should -BeExactly $assetFilename
        $built.Result.asset_bytes | Should -Be $provenance.asset_bytes
        $built.Result.asset_sha256 | Should -BeExactly $provenance.asset_sha256
    }

    It 'T071 recomputes both exact SHA256SUMS assets independently' {
        $lines = @([IO.File]::ReadAllLines((Join-Path $built.OutputDirectory 'SHA256SUMS.txt')))
        $lines.Count | Should -Be 2
        $names = New-Object 'Collections.Generic.List[string]'
        foreach ($line in $lines) {
            $line | Should -Match '^([0-9a-f]{64})  ([A-Za-z0-9._-]+)$'
            $match = [regex]::Match($line, '^([0-9a-f]{64})  ([A-Za-z0-9._-]+)$')
            $hash = $match.Groups[1].Value; $name = $match.Groups[2].Value
            $names.Add($name)
            $hash | Should -BeExactly (Get-FileHash -LiteralPath (Join-Path $built.OutputDirectory $name)).Hash.ToLowerInvariant()
        }
        @($names | Sort-Object) -join '|' | Should -BeExactly (($assetFilename, 'build-provenance.json' | Sort-Object) -join '|')
    }

    It 'T071 fixes entry ordering, timestamps and attributes and repeats identical bytes on this host' {
        foreach ($entry in $entries) {
            $entry.LastWriteTime.ToString('yyyy-MM-ddTHH:mm:ss', [Globalization.CultureInfo]::InvariantCulture) | Should -BeExactly '2000-01-01T00:00:00'
            $entry.Attributes | Should -Be 0
            # NoCompression uses Stored entries on current .NET and DEFLATE
            # stored-block framing on .NET Framework. Exact extraction hashes
            # above and repeat bytes below are the reproducibility contract.
            $entry.CompressedLength | Should -BeGreaterOrEqual $entry.Length
        }
        $repeat = Invoke-PackageBuild $replica
        $repeat.ExitCode | Should -Be 0
        foreach ($name in @($assetFilename, 'build-provenance.json', 'SHA256SUMS.txt')) {
            (Get-FileHash -LiteralPath (Join-Path $repeat.OutputDirectory $name)).Hash | Should -BeExactly (Get-FileHash -LiteralPath (Join-Path $built.OutputDirectory $name)).Hash
        }
        (Invoke-PackageGit $replica.Root @('status', '--porcelain=v1', '--untracked-files=all')).StdOut | Should -BeNullOrEmpty
        $observations.Add([pscustomobject]@{ Kind = 'same host same clean revision byte identical repeat'; SourceRevision = $replica.Revision; First = $built.OutputDirectory; Second = $repeat.OutputDirectory; AssetSha256 = $provenance.asset_sha256; HostVersion = $PSVersionTable.PSVersion.ToString() })
    }
}

Describe 'M4-T02 exact clean extracted portable entry points' {
    It 'T070 extracts only approved root files into a fresh directory with spaces' {
        $extracted | Should -Match ' '
        @(Get-ChildItem -LiteralPath $extracted -Force).Count | Should -Be 8
        @(Get-ChildItem -LiteralPath $extracted -Directory -Force).Count | Should -Be 0
        foreach ($entry in $smokeEntries) {
            (Get-FileHash -LiteralPath (Join-Path $extracted $entry.Path)).Hash.ToLowerInvariant() | Should -BeExactly (Get-PackageSha256 $entry.Bytes)
        }
        $smokeManifestEntry = @($smokeEntries | Where-Object { $_.Path -ceq 'package-manifest.json' })[0]
        $smokeManifest = [Text.Encoding]::UTF8.GetString($smokeManifestEntry.Bytes) | ConvertFrom-Json
        $smokeManifest.source_revision | Should -BeExactly $smokeProvenance.source_revision
        $smokeManifest.version | Should -BeExactly $smokeProvenance.version
        $smokeProvenance.package_manifest_sha256 | Should -BeExactly (Get-PackageSha256 $smokeManifestEntry.Bytes)
        $smokeProvenance.asset_sha256 | Should -BeExactly (Get-FileHash -LiteralPath $smokeZipPath).Hash.ToLowerInvariant()
        if ($smokeIsActualRevision) {
            $smokeProvenance.source_revision | Should -BeExactly $actualRevision
            $smokeRoot = $repository
        } else { $smokeRoot = $replica.Root }
        foreach ($approved in $allowlist) {
            $blob = (Invoke-PackageGit $smokeRoot @('show', ($smokeProvenance.source_revision + ':' + $approved.source_path)) -Binary).Bytes
            (Get-FileHash -LiteralPath (Join-Path $extracted $approved.path)).Hash.ToLowerInvariant() | Should -BeExactly (Get-PackageSha256 $blob)
        }
    }

    It 'T070 invokes the exact extracted BAT usage and retained pause without starting normalization' {
        $result = Invoke-PackagedBat
        $result.ExitCode | Should -Be 1
        $result.StdOut | Should -Match 'ERROR:.*exactly one FOLDER'
        $result.StdErr | Should -BeNullOrEmpty
        $result.PauseObserved | Should -BeTrue
        @(Get-ChildItem -LiteralPath $extracted -File -Force).Count | Should -Be 8
    }

    It 'T070 invokes the matching extracted BAT and PS1 through real Windows PowerShell for invalid source setup' {
        $missing = Join-Path $ownedRoot 'absent synthetic source'
        Test-Path -LiteralPath $missing | Should -BeFalse
        $result = Invoke-PackagedBat -Arguments @($missing)
        $result.ExitCode | Should -Be 1
        $result.StdOut | Should -Match '(?i)(source|folder).*(exist|found)|not.*(exist|found)'
        $result.StdOut | Should -Match 'ERROR: Processing did not complete reliably \(exit code 1\)'
        $result.StdErr | Should -BeNullOrEmpty
        $result.PauseObserved | Should -BeTrue
        Test-Path -LiteralPath $missing | Should -BeFalse
    }

    It 'T070 runs extracted production PS1 with actual verified ImageMagick on synthetic media and owned output' {
        $source = New-PackageDirectory 'synthetic packaged source with spaces'
        $parent = New-PackageDirectory 'owned packaged output with spaces'
        $nested = Join-Path $source 'nested media'; $null = [IO.Directory]::CreateDirectory($nested)
        $png = Join-Path $nested 'synthetic alpha.png'
        $generated = Invoke-PackageProcess (New-PackageProcessStart $magick @('-size', '32x24', 'xc:rgba(30,100,180,0.5)', ('PNG:' + $png)) $ownedRoot)
        $generated.ExitCode | Should -Be 0
        $video = Join-Path $nested 'synthetic video.mp4'
        $videoBytes = New-Object byte[] 4096
        for ($i = 0; $i -lt $videoBytes.Length; $i++) { $videoBytes[$i] = [byte](($i * 29 + 17) % 256) }
        [IO.File]::WriteAllBytes($video, $videoBytes)
        $before = @($png, $video | ForEach-Object { Get-PackageFileState $_ })
        $driver = Join-Path $ownedRoot 'invoke-extracted-media.ps1'
        [IO.File]::WriteAllText($driver, @'
param([string]$Application, [string]$Source, [string]$OutputParent, [string]$Magick, [string]$TemporaryRoot, [string]$Record)
$ErrorActionPreference = 'Stop'
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$context = [ordered]@{
    HostExecutable = (Get-Process -Id $PID).Path
    UserSid = $identity.User.Value
    IsAdministrator = [Security.Principal.WindowsPrincipal]::new($identity).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    PersistentPolicies = @(Get-ExecutionPolicy -List | Where-Object { $_.Scope -ne 'Process' } | ForEach-Object { [ordered]@{ scope = $_.Scope.ToString(); policy = $_.ExecutionPolicy.ToString() } })
    ProcessPolicy = (Get-ExecutionPolicy -Scope Process).ToString()
}
. $Application
$result = @(Invoke-WinImgNormalizer -Source $Source -OutputParent $OutputParent -MagickPath $Magick -NativeTemporaryRoot $TemporaryRoot)
if ($result.Count -ne 1 -or $result[0] -isnot [int]) { throw 'Extracted normalizer must return one integer result.' }
$context.ExitCode = $result[0]
[IO.File]::WriteAllText($Record, ($context | ConvertTo-Json -Depth 8), [Text.UTF8Encoding]::new($false))
exit $result[0]
'@, [Text.UTF8Encoding]::new($false))
        $record = Join-Path $parent 'ordinary-child-security.json'
        $tokens = @('-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $driver, (Join-Path $extracted 'WinImgNormalizer.ps1'), $source, $parent, $magick, $ownedRoot, $record)
        $result = Invoke-PackageProcess (New-PackageProcessStart $hostExecutable $tokens $ownedRoot)
        $result.ExitCode | Should -Be 0
        $result.StdErr | Should -BeNullOrEmpty
        $context = Get-Content -LiteralPath $record -Raw | ConvertFrom-Json
        $context.UserSid | Should -BeExactly ([Security.Principal.WindowsIdentity]::GetCurrent().User.Value)
        $context.IsAdministrator | Should -Be ([Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
        $context.HostExecutable | Should -BeExactly $hostExecutable
        $context.ProcessPolicy | Should -BeExactly 'Bypass'
        ($context.PersistentPolicies | ConvertTo-Json -Depth 6 -Compress) | Should -BeExactly ((($securityBefore | ConvertFrom-Json).policies) | ConvertTo-Json -Depth 6 -Compress)
        $runs = @(Get-ChildItem -LiteralPath $parent -Directory); $runs.Count | Should -Be 1
        $jpeg = Join-Path $runs[0].FullName 'nested media/synthetic alpha.jpeg'
        Test-Path -LiteralPath $jpeg | Should -BeTrue
        $identified = Invoke-PackageProcess (New-PackageProcessStart $magick @('identify', '-ping', '-format', '%m|%wx%h', ('JPEG:' + $jpeg)) $ownedRoot)
        $identified.ExitCode | Should -Be 0
        $identified.StdOut | Should -BeExactly 'JPEG|32x24'
        (Get-FileHash -LiteralPath (Join-Path $runs[0].FullName 'nested media/synthetic video.mp4')).Hash | Should -BeExactly (Get-FileHash -LiteralPath $video).Hash
        @($png, $video | ForEach-Object { Get-PackageFileState $_ }) -join '|' | Should -BeExactly ($before -join '|')
        $logs = @(Get-ChildItem -LiteralPath $runs[0].FullName -Recurse -File -Filter '*.log'); $logs.Count | Should -Be 1
        $log = [IO.File]::ReadAllText($logs[0].FullName)
        $log | Should -Match 'WinImgNormalizer 1\.0\.0 started'
        $log | Should -Match 'SUMMARY ConvertedImages=1 CopiedVideos=1 Duplicates=0 Unsupported=0 Errors=0'
        Get-PackageSecurityState | Should -BeExactly $securityBefore
        $observations.Add([pscustomobject]@{ Kind = 'exact extracted PS1 real media using owned OutputParent seam'; ActualCleanImplementationRevision = $smokeIsActualRevision; SourceRevision = $smokeProvenance.source_revision; AssetSha256 = $smokeProvenance.asset_sha256; ImageMagickSha256 = $imagePin.portable_imagemagick.executable_sha256; Context = $context; JpegSha256 = (Get-FileHash -LiteralPath $jpeg).Hash.ToLowerInvariant(); VideoSha256 = (Get-FileHash -LiteralPath $video).Hash.ToLowerInvariant(); ManualDragAndDrop = $false; RealPicturesNormalization = $false })
    }
}

AfterAll {
    if ($ownedRoot -and $observations) {
        [IO.File]::WriteAllText((Join-Path $ownedRoot 'packaging-observations.json'), (@($observations.ToArray()) | ConvertTo-Json -Depth 12), [Text.UTF8Encoding]::new($false))
    }
    if ($securityBefore) {
        $securityAfter = Get-PackageSecurityState
        [IO.File]::WriteAllText((Join-Path $ownedRoot 'security-observation.json'), ([ordered]@{ Before = ($securityBefore | ConvertFrom-Json); After = ($securityAfter | ConvertFrom-Json); Unchanged = ($securityBefore -eq $securityAfter); SourceBindings = $bindingsBefore; FixtureBinding = $fixtureBinding } | ConvertTo-Json -Depth 8), [Text.UTF8Encoding]::new($false))
        $securityAfter | Should -BeExactly $securityBefore
    }
    foreach ($binding in $bindingsBefore) { (Get-FileHash -LiteralPath (Join-Path $repository $binding.Path)).Hash.ToLowerInvariant() | Should -BeExactly $binding.Sha256 }
    if ($fixtureBinding) {
        (Get-FileHash -LiteralPath $consoleFixture).Hash.ToLowerInvariant() | Should -BeExactly $fixtureBinding.ExecutableSha256
        (Get-FileHash -LiteralPath $compiler).Hash.ToLowerInvariant() | Should -BeExactly $fixtureBinding.CompilerSha256
    }
}
