BeforeAll {
    $repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    $application = Join-Path $repository 'WinImgNormalizer.ps1'

    function Assert-NoReparseAncestors {
        param([string]$Path)
        $current = [IO.Path]::GetFullPath($Path)
        while ($current) {
            if (Test-Path -LiteralPath $current) {
                $entry = Get-Item -LiteralPath $current -Force
                if (($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                    throw 'Test containment path includes a reparse point.'
                }
            }
            $parent = [IO.Directory]::GetParent($current)
            if ($null -eq $parent) { break }
            $current = $parent.FullName
        }
    }

    $scratchParent = Join-Path $repository '.scratch'
    Assert-NoReparseAncestors $scratchParent
    $protectedPictures = @([Environment]::GetFolderPath('MyPictures'),
        (Join-Path $env:USERPROFILE 'Pictures'))
    foreach ($pictures in $protectedPictures) {
        if ([string]::IsNullOrWhiteSpace($pictures)) { continue }
        $picturesFull = [IO.Path]::GetFullPath($pictures).TrimEnd('\', '/')
        if ($scratchParent.Equals($picturesFull, [StringComparison]::OrdinalIgnoreCase) -or
            $scratchParent.StartsWith($picturesFull + '\', [StringComparison]::OrdinalIgnoreCase) -or
            $picturesFull.StartsWith($scratchParent + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Test scratch overlaps a real Pictures location.'
        }
    }
    $git = Get-Command git -CommandType Application -ErrorAction Stop | Select-Object -First 1
    & $git.Source -C $repository check-ignore --quiet --no-index -- (Join-Path $scratchParent 'test-ignore-probe')
    if ($LASTEXITCODE -ne 0) { throw 'Synthetic test scratch must already be ignored.' }
    $ownedRoot = Join-Path $scratchParent ('M0-T03-tests-' + [Guid]::NewGuid().ToString('N'))
    if (Test-Path -LiteralPath $ownedRoot) { throw 'Test ownership directory already exists.' }
    $null = [IO.Directory]::CreateDirectory($ownedRoot)
    [IO.File]::WriteAllText((Join-Path $ownedRoot '.winimg-fixture-root'), 'M0-T03 synthetic test ownership')

    function New-OwnedDirectory {
        param([string]$RelativePath)
        $path = [IO.Path]::GetFullPath((Join-Path $ownedRoot $RelativePath))
        if (-not $path.StartsWith($ownedRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Test path escapes its owned scratch directory.'
        }
        Assert-NoReparseAncestors $path
        if (Test-Path -LiteralPath $path) { throw 'Test directory already exists.' }
        $null = [IO.Directory]::CreateDirectory($path)
        return $path
    }

    # Quote actual Windows argv tokens; shared tests do not require ArgumentList.
    function ConvertTo-WindowsArgument {
        param([string]$Value)
        $escaped = [regex]::Replace($Value, '(\\*)"', '$1$1\"')
        $escaped = [regex]::Replace($escaped, '(\\+)$', '$1$1')
        return '"' + $escaped + '"'
    }

    function Invoke-OwnedPowerShell {
        param([string]$Script, [string[]]$Arguments, [string]$WorkDirectory)
        Assert-NoReparseAncestors $Script
        Assert-NoReparseAncestors $WorkDirectory
        $temporary = Join-Path $WorkDirectory 'temporary'
        $null = [IO.Directory]::CreateDirectory($temporary)
        $start = New-Object Diagnostics.ProcessStartInfo
        $start.FileName = (Get-Process -Id $PID).Path
        $tokens = @('-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $Script) + $Arguments
        $start.Arguments = ($tokens | ForEach-Object { ConvertTo-WindowsArgument ([string]$_) }) -join ' '
        $start.WorkingDirectory = $WorkDirectory
        $start.UseShellExecute = $false
        $start.CreateNoWindow = $true
        $start.RedirectStandardOutput = $true
        $start.RedirectStandardError = $true
        # Cross-edition inherited module paths can break an otherwise real -File run.
        foreach ($key in @($start.EnvironmentVariables.Keys)) {
            if ([string]::Equals([string]$key, 'PSModulePath', [StringComparison]::OrdinalIgnoreCase)) {
                $start.EnvironmentVariables.Remove([string]$key)
            }
        }
        $start.EnvironmentVariables['PATH'] = [IO.Path]::GetDirectoryName($magick) + ';' + $env:PATH
        $start.EnvironmentVariables['TEMP'] = $temporary
        $start.EnvironmentVariables['TMP'] = $temporary
        $start.EnvironmentVariables['MAGICK_TEMPORARY_PATH'] = $temporary
        $process = New-Object Diagnostics.Process
        $process.StartInfo = $start
        try {
            if (-not $process.Start()) { throw 'Owned child PowerShell failed to start.' }
            $stdout = $process.StandardOutput.ReadToEndAsync()
            $stderr = $process.StandardError.ReadToEndAsync()
            if (-not $process.WaitForExit(30000)) {
                # Only the PID just created by this helper and its descendants are terminated.
                & (Join-Path $env:SystemRoot 'System32\taskkill.exe') /PID $process.Id /T /F 2>&1 | Out-Null
                if (-not $process.WaitForExit(5000)) { $process.Kill() }
                throw 'Owned child PowerShell exceeded the 30-second test bound.'
            }
            if (-not $stdout.Wait(5000) -or -not $stderr.Wait(5000)) {
                throw 'Owned child output streams did not finish.'
            }
            return [pscustomobject]@{
                ExitCode = $process.ExitCode
                StdOut = $stdout.GetAwaiter().GetResult()
                StdErr = $stderr.GetAwaiter().GetResult()
            }
        } finally { $process.Dispose() }
    }

    if ([string]::IsNullOrWhiteSpace($env:WINIMG_TEST_MAGICK) -or
        -not [IO.Path]::IsPathRooted($env:WINIMG_TEST_MAGICK)) {
        throw 'WINIMG_TEST_MAGICK must explicitly select the verified test magick.exe.'
    }
    $magick = [IO.Path]::GetFullPath($env:WINIMG_TEST_MAGICK)
    Assert-NoReparseAncestors $magick
    if (-not (Test-Path -LiteralPath $magick -PathType Leaf) -or
        [IO.Path]::GetFileName($magick) -ne 'magick.exe') {
        throw 'The explicitly selected test magick.exe is missing.'
    }

    function Invoke-TestMagick {
        param([string[]]$Arguments)
        $result = & $magick @Arguments 2>&1
        $code = $LASTEXITCODE
        if ($code -ne 0) { throw ('Test ImageMagick exited {0}: {1}' -f $code, ($result -join "`n")) }
        return ($result -join "`n")
    }

    $null = Invoke-TestMagick @('-version')
    $fixtureSource = New-OwnedDirectory 'ordinary-source'
    $null = [IO.Directory]::CreateDirectory((Join-Path $fixtureSource 'sub'))
    $null = [IO.Directory]::CreateDirectory((Join-Path $fixtureSource 'empty-directory'))
    # Small synthetic gradients cover the same ordinary formats/tree as M0-T02.
    # These are independent recipes, not cross-build JPEG byte expectations.
    $imageCases = @(
        @{ Source = 'landscape.jpg'; Output = 'landscape.jpeg'; Width = 64; Height = 48; Coder = 'JPEG' },
        @{ Source = 'sub\portrait.png'; Output = 'sub\portrait.jpeg'; Width = 48; Height = 64; Coder = 'PNG' },
        @{ Source = 'sub\lines.bmp'; Output = 'sub\lines.jpeg'; Width = 64; Height = 48; Coder = 'BMP' },
        @{ Source = 'seeded-noise.tiff'; Output = 'seeded-noise.jpeg'; Width = 80; Height = 56; Coder = 'TIFF' }
    )
    foreach ($image in $imageCases) {
        $target = Join-Path $fixtureSource $image.Source
        $null = Invoke-TestMagick @('-quiet', '-size', ($image.Width.ToString() + 'x' + $image.Height),
            'gradient:#205080-#E0A070', '-depth', '8', ($image.Coder + ':' + $target))
    }
    $videoCases = @('sub\opaque-video.mp4', 'sub\opaque-video.mov', 'sub\opaque-video.mkv')
    for ($index = 0; $index -lt $videoCases.Count; $index++) {
        $bytes = New-Object byte[] 4096
        for ($offset = 0; $offset -lt $bytes.Length; $offset++) {
            $bytes[$offset] = [byte](($offset * ($index + 1) + 13) % 256)
        }
        [IO.File]::WriteAllBytes((Join-Path $fixtureSource $videoCases[$index]), $bytes)
    }
    [IO.File]::WriteAllText((Join-Path $fixtureSource 'ignored.txt'), 'Synthetic unsupported-file control.')
    $fixedUtc = [DateTime]::Parse('2020-02-03T04:05:06Z').ToUniversalTime()
    foreach ($file in Get-ChildItem -LiteralPath $fixtureSource -Recurse -File) {
        [IO.File]::SetCreationTimeUtc($file.FullName, $fixedUtc)
        [IO.File]::SetLastWriteTimeUtc($file.FullName, $fixedUtc)
    }

    function Get-SourceState {
        param([string]$Root)
        return @(
            foreach ($file in Get-ChildItem -LiteralPath $Root -Recurse -File | Sort-Object FullName) {
                [pscustomobject]@{
                    Path = $file.FullName.Substring($Root.Length + 1)
                    Hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
                    Length = $file.Length
                    CreationTicks = $file.CreationTimeUtc.Ticks
                    ModifiedTicks = $file.LastWriteTimeUtc.Ticks
                }
            }
        ) | ConvertTo-Json -Depth 4 -Compress
    }

    # Probe import in an owned child BEFORE importing into Pester. A regression
    # that exits zero must fail setup, not terminate the runner with a false pass.
    $importWork = New-OwnedDirectory 'pure-import'
    $control = Join-Path $importWork 'Import-Control.ps1'
    $encodedApplication = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($application))
    $scriptText = @'
$ErrorActionPreference = 'Continue'
$script:LogPath = 'import-log-sentinel'
function New-Item { throw 'Import attempted New-Item.' }
function Set-Content { throw 'Import attempted Set-Content.' }
function Add-Content { throw 'Import attempted Add-Content.' }
function Copy-Item { throw 'Import attempted Copy-Item.' }
function Get-Command { throw 'Import attempted native command discovery.' }
function Write-Host { throw 'Import attempted application output.' }
function magick { throw 'Import attempted a native conversion.' }
$path = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('__APPLICATION__'))
. $path 'unused-source-sentinel' 1048576
if ($ErrorActionPreference -ne 'Continue') { throw 'Import mutated caller error preference.' }
if ($script:LogPath -ne 'import-log-sentinel') { throw 'Import mutated caller log path.' }
if (-not (Test-Path Function:\Invoke-WinImgNormalizer) -or
    -not (Test-Path Function:\Invoke-WinImgNormalizerCommand)) { throw 'Callable definitions are absent.' }
[Console]::WriteLine('WINIMG_IMPORT_COMPLETED')
'@
    $scriptText = $scriptText.Replace('__APPLICATION__', $encodedApplication)
    [IO.File]::WriteAllText($control, $scriptText, (New-Object Text.UTF8Encoding($false)))
    $importResult = Invoke-OwnedPowerShell -Script $control -WorkDirectory $importWork
    if ($importResult.ExitCode -ne 0 -or $importResult.StdErr -or
        $importResult.StdOut.Trim() -ne 'WINIMG_IMPORT_COMPLETED' -or
        @(Get-ChildItem -LiteralPath $importWork -Recurse -File).Count -ne 1) {
        throw ('Isolated import probe failed before Pester import: exit={0}; stdout={1}; stderr={2}' -f
            $importResult.ExitCode, $importResult.StdOut, $importResult.StdErr)
    }
    . $application
}

Describe 'M0-T03 callable boundary and positional compatibility' {
    It 'T005 imports definitions without writes, native execution, caller mutation or host exit' {
        $importResult.ExitCode | Should -Be 0
        $importResult.StdErr | Should -BeNullOrEmpty
        $importResult.StdOut.Trim() | Should -Be 'WINIMG_IMPORT_COMPLETED'
        @(Get-ChildItem -LiteralPath $importWork -Recurse -File).Count | Should -Be 1
    }

    It 'T005 returns usage status without executing or terminating the test host' {
        $parent = Join-Path $ownedRoot 'unused-usage-output'
        $result = Invoke-WinImgNormalizerCommand -Arguments @() -OutputParent $parent -ProcessRunner {
            throw 'Usage attempted process execution.'
        }
        $result | Should -Be 1
        Test-Path -LiteralPath $parent | Should -BeFalse
    }

    It 'T005 returns zero for an empty tree through the injected destination and process runner' {
        $source = New-OwnedDirectory 'empty-source'
        $parent = New-OwnedDirectory 'empty-output'
        $trace = New-Object 'Collections.Generic.List[object]'
        $runner = {
            param([string]$Executable, [string[]]$Arguments)
            $trace.Add(@($Arguments))
            return 0
        }
        $result = Invoke-WinImgNormalizer -Source $source -OutputParent $parent -MagickPath $magick -ProcessRunner $runner
        $result | Should -Be 0
        $trace.Count | Should -Be 0
        $runs = @(Get-ChildItem -LiteralPath $parent -Directory)
        $runs.Count | Should -Be 1
        $logs = @(Get-ChildItem -LiteralPath $runs[0].FullName -Recurse -File -Filter '*.log')
        $logs.Count | Should -Be 1
        Get-Content -LiteralPath $logs[0].FullName -Raw | Should -Match 'No images or videos found\.'
    }

    It 'T005 routes conversion arguments and a diagnosed sharing retry through the injected runner' {
        $source = New-OwnedDirectory 'runner-source'
        $parent = New-OwnedDirectory 'runner-output'
        $sourceImage = Join-Path $source 'single.png'
        [IO.File]::WriteAllBytes($sourceImage,
            [IO.File]::ReadAllBytes((Join-Path $fixtureSource 'sub\portrait.png')))
        $sourceBefore = Get-SourceState $source
        $jpegBytes = [IO.File]::ReadAllBytes((Join-Path $fixtureSource 'landscape.jpg'))
        $trace = New-Object 'Collections.Generic.List[object]'
        $runner = {
            param([string]$Executable, [string[]]$Arguments)
            $trace.Add([pscustomobject]@{ Executable = $Executable; Arguments = @($Arguments) })
            if ($trace.Count -eq 1) { return [pscustomobject]@{ ExitCode = 9; StdErr = "magick.exe: sharing violation @ error/blob.c/OpenBlob/3590." } }
            $target = $Arguments[-1].Substring('JPEG:'.Length)
            if (-not $target.StartsWith($parent + '\', [StringComparison]::OrdinalIgnoreCase)) {
                throw 'Injected conversion escaped its owned output parent.'
            }
            [IO.File]::WriteAllBytes($target, $jpegBytes)
            return 0
        }
        $result = Invoke-WinImgNormalizer -Source $source -OutputParent $parent -MagickPath $magick -ProcessRunner $runner
        $result | Should -Be 0
        $trace.Count | Should -Be 2
        $first = $trace[0].Arguments
        $retry = $trace[1].Arguments
        $snapshotPath = $first[4]
        if ($snapshotPath.StartsWith('\\?\', [StringComparison]::Ordinal)) { $snapshotPath = $snapshotPath.Substring(4) }
        $run = @(Get-ChildItem -LiteralPath $parent -Directory)
        $run.Count | Should -Be 1
        $snapshotPath.StartsWith((Join-Path $run[0].FullName '.WinImgNormalizer\work\'), [StringComparison]::OrdinalIgnoreCase) | Should -BeTrue
        [IO.Path]::GetFileName($snapshotPath) | Should -Be 'source.png'
        $first[4] | Should -Not -Be $sourceImage
        $retry[4] | Should -Be $first[4]
        Test-Path -LiteralPath $snapshotPath | Should -BeFalse
        ($first -join '|') | Should -Be (
            @('-quiet', '-regard-warnings', '-define', 'registry:filename:literal=true', $first[4], '-auto-orient', '-colorspace', 'sRGB',
                '-background', 'white', '-alpha', 'remove', '-alpha', 'off', '-strip',
                '-sampling-factor', '4:2:0', '-interlace', 'Line', '-resize', '100%', '-define',
                'jpeg:extent=1048576B', $first[-1]) -join '|')
        ($retry -join '|') | Should -Be (
            @('-quiet', '-regard-warnings', '-define', 'registry:filename:literal=true', $first[4], '-auto-orient', '-colorspace', 'sRGB',
                '-background', 'white', '-alpha', 'remove', '-alpha', 'off', '-strip',
                '-sampling-factor', '4:2:0', '-interlace', 'Line', '-resize', '100%', '-define',
                'jpeg:extent=1048576B', $retry[-1]) -join '|')
        $retry[-1] | Should -Not -Be $first[-1]
        $trace[0].Executable | Should -Be $magick
        $trace[1].Executable | Should -Be $magick
        Get-SourceState $source | Should -Be $sourceBefore
        $logs = @(Get-ChildItem -LiteralPath $parent -Recurse -File -Filter '*.log')
        Get-Content -LiteralPath $logs[0].FullName -Raw |
            Should -Match 'SUMMARY ConvertedImages=1 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=0'
    }

    It 'T006 preserves the <Form> positional -File form and ordinary image/video behavior' -ForEach @(
        @{ Form = 'source-only'; ExplicitCap = $false },
        @{ Form = 'source-plus-1048576'; ExplicitCap = $true }
    ) {
        $work = New-OwnedDirectory ('invocation-' + $Form)
        $parent = Join-Path $work 'injected-pictures'
        $snapshot = Join-Path $work 'WinImgNormalizer.contained.ps1'
        $text = [IO.File]::ReadAllText($application)
        $entry = 'exit (Invoke-WinImgNormalizerCommand -Arguments $args)'
        $first = $text.IndexOf($entry, [StringComparison]::Ordinal)
        $first | Should -BeGreaterOrEqual 0
        $text.IndexOf($entry, $first + $entry.Length, [StringComparison]::Ordinal) | Should -Be -1
        $encodedParent = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($parent))
        $injection = "exit (Invoke-WinImgNormalizerCommand -Arguments `$args -OutputParent ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('$encodedParent'))))"
        $contained = $text.Substring(0, $first) + $injection + $text.Substring($first + $entry.Length)
        # Only the final call receives the internal destination seam. Argument binding,
        # native runner, conversion, discovery and exit execute unchanged in a -File host.
        # This intentionally does not claim coverage of real Known Folder resolution.
        [IO.File]::WriteAllText($snapshot, $contained, (New-Object Text.UTF8Encoding($false)))
        $before = Get-SourceState $fixtureSource
        $arguments = @($fixtureSource)
        if ($ExplicitCap) { $arguments += '1048576' }
        $result = Invoke-OwnedPowerShell -Script $snapshot -Arguments $arguments -WorkDirectory $work
        $result.ExitCode | Should -Be 0 -Because $result.StdOut
        $result.StdErr | Should -BeNullOrEmpty
        Get-SourceState $fixtureSource | Should -Be $before
        $runs = @(Get-ChildItem -LiteralPath $parent -Directory)
        $runs.Count | Should -Be 1
        $runs[0].Name | Should -Match '^ordinary-source_WinImgNormalized_\d{8}_\d{6}_[0-9a-f]{32}$'
        $output = $runs[0].FullName
        @(Get-ChildItem -LiteralPath $output -Recurse -File).Count | Should -Be 8
        @(Get-ChildItem -LiteralPath $output -Recurse -Directory).Count | Should -Be 5
        Test-Path -LiteralPath (Join-Path $output 'empty-directory') -PathType Container | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $output 'ignored.txt') | Should -BeFalse
        foreach ($image in $imageCases) {
            $path = Join-Path $output $image.Output
            $item = Get-Item -LiteralPath $path
            $item.Length | Should -BeGreaterThan 0
            $item.Length | Should -BeLessOrEqual 1048576
            Invoke-TestMagick @('identify', '-format', '%m|%w|%h|%n', $path) |
                Should -Be ('JPEG|{0}|{1}|1' -f $image.Width, $image.Height)
            $null = Invoke-TestMagick @('-quiet', $path, 'null:')
            $original = Get-Item -LiteralPath (Join-Path $fixtureSource $image.Source)
            $item.CreationTimeUtc.Ticks | Should -Be $original.CreationTimeUtc.Ticks
            $item.LastWriteTimeUtc.Ticks | Should -Be $original.LastWriteTimeUtc.Ticks
        }
        foreach ($relative in $videoCases) {
            $original = Get-Item -LiteralPath (Join-Path $fixtureSource $relative)
            $copy = Get-Item -LiteralPath (Join-Path $output $relative)
            (Get-FileHash -LiteralPath $copy.FullName -Algorithm SHA256).Hash |
                Should -Be (Get-FileHash -LiteralPath $original.FullName -Algorithm SHA256).Hash
            $copy.CreationTimeUtc.Ticks | Should -Be $original.CreationTimeUtc.Ticks
            $copy.LastWriteTimeUtc.Ticks | Should -Be $original.LastWriteTimeUtc.Ticks
        }
        $logs = @(Get-ChildItem -LiteralPath $output -Recurse -File -Filter '*.log')
        $logs.Count | Should -Be 1
        $log = Get-Content -LiteralPath $logs[0].FullName -Raw
        $log | Should -Match 'SUMMARY ConvertedImages=4 CopiedVideos=3 Duplicates=0 Unsupported=0 Errors=0'
        $log | Should -Match 'MaxBytes: 1048576 bytes \(1 MiB; best-effort target\)'
        $log | Should -Match 'Completed '
    }
}
