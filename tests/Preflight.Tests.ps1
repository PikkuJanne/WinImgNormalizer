BeforeAll {
    $repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    $application = Join-Path $repository 'WinImgNormalizer.ps1'

    function ConvertFrom-PreflightNativePath {
        param([string]$Path)
        # Generated scratch can require an extended Windows prefix. Strip only
        # that native spelling before applying the existing owned-path guard.
        if ($Path.StartsWith('\\?\UNC\', [StringComparison]::OrdinalIgnoreCase)) { return '\\' + $Path.Substring(8) }
        if ($Path.StartsWith('\\?\', [StringComparison]::Ordinal)) { return $Path.Substring(4) }
        return $Path
    }

    function Assert-PreflightNoReparseAncestors {
        param([string]$Path)
        $current = [IO.Path]::GetFullPath($Path)
        while ($current) {
            if (Test-Path -LiteralPath $current) {
                if (((Get-Item -LiteralPath $current -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                    throw 'Preflight test containment path includes a reparse point.'
                }
            }
            $parent = [IO.Directory]::GetParent($current)
            if ($null -eq $parent) { break }
            $current = $parent.FullName
        }
    }

    $scratchParent = [IO.Path]::GetFullPath((Join-Path $repository '.scratch'))
    Assert-PreflightNoReparseAncestors $scratchParent
    foreach ($pictures in @([Environment]::GetFolderPath('MyPictures'), (Join-Path $env:USERPROFILE 'Pictures'))) {
        if ([string]::IsNullOrWhiteSpace($pictures)) { continue }
        $picturesFull = [IO.Path]::GetFullPath($pictures).TrimEnd('\', '/')
        if ($scratchParent.Equals($picturesFull, [StringComparison]::OrdinalIgnoreCase) -or
            $scratchParent.StartsWith($picturesFull + '\', [StringComparison]::OrdinalIgnoreCase) -or
            $picturesFull.StartsWith($scratchParent + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Preflight test scratch overlaps a real Pictures location.'
        }
    }
    $git = Get-Command git -CommandType Application -ErrorAction Stop | Select-Object -First 1
    & $git.Source -C $repository check-ignore --quiet --no-index -- (Join-Path $scratchParent 'preflight-ignore-probe')
    if ($LASTEXITCODE -ne 0) { throw 'Preflight synthetic scratch must already be ignored.' }
    $ownedRoot = Join-Path $scratchParent ('M1-T01-tests-' + [Guid]::NewGuid().ToString('N'))
    if (Test-Path -LiteralPath $ownedRoot) { throw 'Preflight ownership directory already exists.' }
    $null = [IO.Directory]::CreateDirectory($ownedRoot)
    [IO.File]::WriteAllText((Join-Path $ownedRoot '.winimg-fixture-root'), 'M1-T01 synthetic test ownership')

    function New-PreflightDirectory {
        param([string]$Label)
        $path = [IO.Path]::GetFullPath((Join-Path $ownedRoot ($Label + '-' + [Guid]::NewGuid().ToString('N'))))
        if (-not $path.StartsWith($ownedRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Preflight test path escapes its owned scratch directory.'
        }
        Assert-PreflightNoReparseAncestors $path
        if (Test-Path -LiteralPath $path) { throw 'Preflight test directory already exists.' }
        $null = [IO.Directory]::CreateDirectory($path)
        return $path
    }

    if ([string]::IsNullOrWhiteSpace($env:WINIMG_TEST_MAGICK) -or
        -not [IO.Path]::IsPathRooted($env:WINIMG_TEST_MAGICK)) {
        throw 'WINIMG_TEST_MAGICK must explicitly select the verified test magick.exe.'
    }
    $magick = [IO.Path]::GetFullPath($env:WINIMG_TEST_MAGICK)
    Assert-PreflightNoReparseAncestors $magick
    if (-not (Test-Path -LiteralPath $magick -PathType Leaf) -or [IO.Path]::GetFileName($magick) -ne 'magick.exe') {
        throw 'The explicitly selected preflight test magick.exe is missing.'
    }

    function Invoke-PreflightTestMagick {
        param([string[]]$Arguments)
        $text = @(& $magick @Arguments 2>&1)
        if ($LASTEXITCODE -ne 0) { throw ('Fixture ImageMagick failed: ' + ($text -join "`n")) }
        return ($text -join "`n")
    }

    function Get-PreflightSourceState {
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

    function ConvertTo-PreflightWindowsArgument {
        param([string]$Value)
        $escaped = [regex]::Replace($Value, '(\\*)"', '$1$1\"')
        $escaped = [regex]::Replace($escaped, '(\\+)$', '$1$1')
        return '"' + $escaped + '"'
    }

    function Invoke-PreflightOwnedPowerShell {
        param([string]$Script, [string[]]$Arguments, [string]$WorkDirectory)
        Assert-PreflightNoReparseAncestors $Script
        Assert-PreflightNoReparseAncestors $WorkDirectory
        $temporary = Join-Path $WorkDirectory 'temporary'
        $null = [IO.Directory]::CreateDirectory($temporary)
        $start = New-Object Diagnostics.ProcessStartInfo
        $start.FileName = (Get-Process -Id $PID).Path
        $tokens = @('-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $Script) + $Arguments
        $start.Arguments = ($tokens | ForEach-Object { ConvertTo-PreflightWindowsArgument ([string]$_) }) -join ' '
        $start.WorkingDirectory = $WorkDirectory
        $start.UseShellExecute = $false
        $start.CreateNoWindow = $true
        $start.RedirectStandardOutput = $true
        $start.RedirectStandardError = $true
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
            if (-not $process.Start()) { throw 'Owned preflight PowerShell failed to start.' }
            $stdout = $process.StandardOutput.ReadToEndAsync()
            $stderr = $process.StandardError.ReadToEndAsync()
            if (-not $process.WaitForExit(30000)) {
                & (Join-Path $env:SystemRoot 'System32\taskkill.exe') /PID $process.Id /T /F 2>&1 | Out-Null
                if (-not $process.WaitForExit(5000)) { $process.Kill() }
                throw 'Owned preflight child exceeded the 30-second test bound.'
            }
            if (-not $stdout.Wait(5000) -or -not $stderr.Wait(5000)) { throw 'Owned preflight child streams did not finish.' }
            return [pscustomobject]@{ ExitCode = $process.ExitCode; StdOut = $stdout.GetAwaiter().GetResult(); StdErr = $stderr.GetAwaiter().GetResult() }
        } finally { $process.Dispose() }
    }

    $supportedVersionText = @'
Version: ImageMagick 7.1.2-32 Q16-HDRI x64 synthetic-test-build
Copyright: (C) ImageMagick Studio LLC
Features: Cipher DPC HDRI OpenMP
Delegates (built-in): bzlib jpeg lcms png tiff webp zlib
'@
    $supportedFormatText = @'
   Format  Module    Mode  Description
-----------------------------------------------------
     JPEG* JPEG      rw-   Joint Photographic Experts Group JFIF format
      PNG* PNG       rw-   Portable Network Graphics
      BMP* BMP       rw-   Microsoft Windows bitmap image
     TIFF* TIFF      rw+   Tagged Image File Format
      GIF* GIF       rw+   CompuServe graphics interchange format
     HEIC* HEIC      rw+   High Efficiency Image Format
     HEIF* HEIC      rw+   High Efficiency Image Format
     WEBP* WEBP      rw+   WebP Image Format
'@

    function New-ControlledPreflightRunner {
        param(
            [string]$VersionText = $supportedVersionText,
            [string]$FormatText = $supportedFormatText,
            [int]$VersionExitCode = 0,
            [int]$FormatExitCode = 0,
            [string]$VersionError = '',
            [string]$FormatError = '',
            [bool]$TimedOut = $false,
            [object]$Trace
        )
        $runner = {
            param([string]$Executable, [string[]]$Arguments)
            if ($null -ne $Trace) {
                $Trace.Add([pscustomobject]@{ Executable = $Executable; Arguments = @($Arguments) })
            }
            if (($Arguments -join '|') -eq '-version') {
                return [pscustomobject]@{ ExitCode = $VersionExitCode; StdOut = $VersionText; StdErr = $VersionError; TimedOut = $TimedOut }
            }
            if (($Arguments -join '|') -eq '-list|format') {
                return [pscustomobject]@{ ExitCode = $FormatExitCode; StdOut = $FormatText; StdErr = $FormatError; TimedOut = $TimedOut }
            }
            throw ('Unexpected preflight arguments: ' + ($Arguments -join '|'))
        }
        return $runner.GetNewClosure()
    }

    function Invoke-ContainedPreflight {
        param(
            [object[]]$InputArguments,
            [string]$OutputParent,
            [string]$MagickPath = $magick,
            [scriptblock]$PreflightRunner,
            [scriptblock]$ProcessRunner
        )
        $invoke = @{ Arguments = $InputArguments; OutputParent = $OutputParent; MagickPath = $MagickPath }
        if ($PreflightRunner) { $invoke.PreflightRunner = $PreflightRunner }
        if ($ProcessRunner) { $invoke.ProcessRunner = $ProcessRunner }
        $observed = @(& Invoke-WinImgNormalizerCommand @invoke 6>&1 2>&1)
        $codes = @($observed | Where-Object { $_ -is [int] -or $_ -is [long] })
        if ($codes.Count -ne 1) { throw ('Expected one callable exit status; observed: ' + ($observed -join "`n")) }
        return [pscustomobject]@{
            Code = $codes[0]
            Text = (@($observed | Where-Object { $_ -isnot [int] -and $_ -isnot [long] }) -join "`n")
        }
    }

    function Assert-PreflightNoOutput {
        param([string]$Path)
        Test-Path -LiteralPath $Path | Should -BeFalse
    }

    # The M0-T03 suite performs its child-host import probe before this file.
    # This direct import is also callable-only when this suite is selected alone.
    . $application
}

Describe 'M1-T01 friendly argument and source validation (T008)' {
    It 'returns a friendly nonzero real -File exit for <Label> before run output' -ForEach @(
        @{ Label = 'zero arguments'; Case = 'zero' },
        @{ Label = 'extra argument'; Case = 'extra' },
        @{ Label = 'invalid cap'; Case = 'cap' },
        @{ Label = 'missing dependency'; Case = 'dependency' }
    ) {
        $work = New-PreflightDirectory 'file-setup-failure'
        $source = Join-Path $work 'source'
        $null = [IO.Directory]::CreateDirectory($source)
        [IO.File]::WriteAllText((Join-Path $source 'source-sentinel.txt'), 'Synthetic source preservation sentinel.')
        $before = Get-PreflightSourceState $source
        $parent = Join-Path $work 'output'
        $selectedMagick = $magick
        if ($Case -eq 'dependency') { $selectedMagick = Join-Path $work 'missing-magick.exe' }
        $script = Join-Path $work 'WinImgNormalizer.contained.ps1'
        $text = [IO.File]::ReadAllText($application)
        $entry = 'exit (Invoke-WinImgNormalizerCommand -Arguments $args)'
        $first = $text.IndexOf($entry, [StringComparison]::Ordinal)
        $first | Should -BeGreaterOrEqual 0
        $text.IndexOf($entry, $first + $entry.Length, [StringComparison]::Ordinal) | Should -Be -1
        $encodedParent = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($parent))
        $encodedMagick = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($selectedMagick))
        $injection = "exit (Invoke-WinImgNormalizerCommand -Arguments `$args -OutputParent ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('$encodedParent'))) -MagickPath ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('$encodedMagick'))))"
        $contained = $text.Substring(0, $first) + $injection + $text.Substring($first + $entry.Length)
        [IO.File]::WriteAllText($script, $contained, (New-Object Text.UTF8Encoding($false)))
        $arguments = @($source)
        if ($Case -eq 'zero') { $arguments = @() }
        if ($Case -eq 'extra') { $arguments += @('1048576', 'unexpected') }
        if ($Case -eq 'cap') { $arguments += '1.5' }
        $result = Invoke-PreflightOwnedPowerShell -Script $script -Arguments $arguments -WorkDirectory $work
        $result.ExitCode | Should -Be 1
        $result.StdErr | Should -BeNullOrEmpty
        $result.StdOut | Should -Match '(?i)(Usage:|maxBytes|ImageMagick)'
        Assert-PreflightNoOutput $parent
        Get-PreflightSourceState $source | Should -Be $before
    }

    It 'rejects <Label> argument count before dependency execution or output creation' -ForEach @(
        @{ Label = 'zero'; Count = 0 },
        @{ Label = 'three'; Count = 3 }
    ) {
        $work = New-PreflightDirectory 'argument-count'
        $source = Join-Path $work 'source'
        $null = [IO.Directory]::CreateDirectory($source)
        $parent = Join-Path $work 'output'
        $arguments = @()
        if ($Count -eq 3) { $arguments = @($source, '1048576', 'unexpected') }
        $result = Invoke-ContainedPreflight -InputArguments $arguments -OutputParent $parent -PreflightRunner {
            throw 'Argument failure attempted dependency execution.'
        } -ProcessRunner { throw 'Argument failure attempted conversion.' }
        $result.Code | Should -Be 1
        $result.Text | Should -Match 'Usage:.*<sourceFolder>.*\[maxBytes\]'
        Assert-PreflightNoOutput $parent
    }

    It 'rejects <Label> byte cap without a casting exception or output' -ForEach @(
        @{ Label = 'zero'; Value = '0' },
        @{ Label = 'negative'; Value = '-1' },
        @{ Label = 'decimal'; Value = '1.5' },
        @{ Label = 'nonnumeric'; Value = 'invalid' },
        @{ Label = 'overflow'; Value = '9223372036854775808' },
        @{ Label = 'blank'; Value = ' ' },
        @{ Label = 'exponent'; Value = '1e6' },
        @{ Label = 'grouped'; Value = '1,024' }
    ) {
        $work = New-PreflightDirectory 'invalid-cap'
        $source = Join-Path $work 'source'
        $null = [IO.Directory]::CreateDirectory($source)
        $parent = Join-Path $work 'output'
        $result = Invoke-ContainedPreflight -InputArguments @($source, $Value) -OutputParent $parent -PreflightRunner {
            throw 'Invalid byte cap attempted dependency execution.'
        } -ProcessRunner { throw 'Invalid byte cap attempted conversion.' }
        $result.Code | Should -Be 1
        $result.Text | Should -Match '(?i)(maxBytes|byte cap|byte limit).*(positive|Int64|integer)'
        $result.Text | Should -Not -Match 'Cannot convert|InvalidCastException|ParameterBindingException'
        Assert-PreflightNoOutput $parent
    }

    It 'rejects <Label> source before dependency execution or output' -ForEach @(
        @{ Label = 'missing'; Kind = 'missing' },
        @{ Label = 'file'; Kind = 'file' },
        @{ Label = 'non-FileSystem provider'; Kind = 'provider' },
        @{ Label = 'blank'; Kind = 'blank' }
    ) {
        $work = New-PreflightDirectory 'invalid-source'
        $parent = Join-Path $work 'output'
        $source = Join-Path $work 'missing'
        if ($Kind -eq 'file') { [IO.File]::WriteAllText($source, 'Synthetic regular file.') }
        if ($Kind -eq 'provider') { $source = 'Env:\' }
        if ($Kind -eq 'blank') { $source = ' ' }
        $result = Invoke-ContainedPreflight -InputArguments @($source) -OutputParent $parent -PreflightRunner {
            throw 'Invalid source attempted dependency execution.'
        } -ProcessRunner { throw 'Invalid source attempted conversion.' }
        $result.Code | Should -Be 1
        $result.Text | Should -Match '(?i)(source|Usage:)'
        Assert-PreflightNoOutput $parent
    }

    It 'passes <Label> positive Int64 cap unchanged through the positional adapter' -ForEach @(
        @{ Label = 'default'; Explicit = $false; Value = '1048576'; Expected = [long]1048576 },
        @{ Label = 'one byte'; Explicit = $true; Value = '1'; Expected = [long]1 },
        @{ Label = 'non-KiB multiple'; Explicit = $true; Value = '1048577'; Expected = [long]1048577 },
        @{ Label = 'Int64 maximum'; Explicit = $true; Value = '9223372036854775807'; Expected = [long]::MaxValue }
    ) {
        $work = New-PreflightDirectory 'valid-adapter'
        $parent = Join-Path $work 'output'
        Mock Invoke-WinImgNormalizer { return 0 }
        $arguments = @($work)
        if ($Explicit) { $arguments += $Value }
        Invoke-WinImgNormalizerCommand -Arguments $arguments -OutputParent $parent -MagickPath $magick | Should -Be 0
        Should -Invoke Invoke-WinImgNormalizer -Times 1 -Exactly -ParameterFilter {
            $Source -eq $work -and [string]$MaxBytes -eq [string]$Expected -and $OutputParent -eq $parent
        }
        $parsed = ConvertTo-WinImgByteCap -Value $Value
        $parsed | Should -Be $Expected
        $parsed | Should -BeOfType ([long])
        Assert-PreflightNoOutput $parent
    }
}

Describe 'M1-T01 root-preserving paths (T009)' {
    It 'preserves root and absolute semantics for <Label> without accessing or processing that root' -ForEach @(
        @{ Label = 'drive root'; Path = 'C:\'; Expected = 'C:\' },
        @{ Label = 'drive root with extra separator'; Path = 'C:\\'; Expected = 'C:\' },
        @{ Label = 'directory trailing separators'; Path = 'C:\synthetic\photos\\'; Expected = 'C:\synthetic\photos' },
        @{ Label = 'UNC share root'; Path = '\\synthetic-server\synthetic-share'; Expected = '\\synthetic-server\synthetic-share\' },
        @{ Label = 'UNC share root with separator'; Path = '\\synthetic-server\synthetic-share\'; Expected = '\\synthetic-server\synthetic-share\' },
        @{ Label = 'UNC nested directory'; Path = '\\synthetic-server\synthetic-share\photos\\'; Expected = '\\synthetic-server\synthetic-share\photos' }
    ) {
        $actual = Normalize-WinImgRootPath -Path $Path
        $actual | Should -Be $Expected
        [IO.Path]::IsPathRooted($actual) | Should -BeTrue
        $actual | Should -Not -Match '^[A-Za-z]:$'
    }

    It 'resolves an owned relative source and retains its real DirectoryInfo root' {
        $work = New-PreflightDirectory 'relative-source'
        $source = Join-Path $work 'source'
        $null = [IO.Directory]::CreateDirectory($source)
        Push-Location -LiteralPath $work
        try { $actual = Resolve-WinImgSourcePath -Path '.\source\' }
        finally { Pop-Location }
        $actual | Should -Be $source
    }
}

Describe 'M1-T01 reviewed ImageMagick identity and preflight (T010)' {
    It 'resolves the verified real executable despite shadowing aliases and functions in a child host' {
        $work = New-PreflightDirectory 'shadowed-magick'
        $script = Join-Path $work 'Check-Application-Identity.ps1'
        $encodedApplication = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($application))
        $text = @'
$ErrorActionPreference = 'Stop'
function magick { throw 'A function named magick was invoked.' }
function magick.exe { throw 'A function named magick.exe was invoked.' }
function Invoke-SyntheticMagickShadow { throw 'A magick alias was invoked.' }
Set-Alias -Name magick -Value Invoke-SyntheticMagickShadow
Set-Alias -Name magick.exe -Value Invoke-SyntheticMagickShadow
. ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('__APPLICATION__')))
$info = Get-WinImgMagickInfo
[Console]::WriteLine($info.Path)
[Console]::WriteLine($info.Version)
'@
        [IO.File]::WriteAllText($script, $text.Replace('__APPLICATION__', $encodedApplication), (New-Object Text.UTF8Encoding($false)))
        $result = Invoke-PreflightOwnedPowerShell -Script $script -WorkDirectory $work
        $result.ExitCode | Should -Be 0
        $result.StdErr | Should -BeNullOrEmpty
        $lines = @($result.StdOut.Trim() -split '\r?\n')
        $lines.Count | Should -Be 2
        $lines[0] | Should -Be $magick
        $lines[1] | Should -Be '7.1.2.32'
    }

    It 'ignores function and alias names and resolves only an Application command' {
        $applicationCommand = Microsoft.PowerShell.Core\Get-Command -Name $magick -CommandType Application -ErrorAction Stop
        Mock Get-Command {
            if ($CommandType -ne 'Application') { throw 'Dependency lookup accepted non-Application commands.' }
            return $applicationCommand
        }
        $actual = Resolve-WinImgMagickApplication
        $actual | Should -Be $magick
        Should -Invoke Get-Command -Times 1 -Exactly -ParameterFilter { $CommandType -eq 'Application' }
    }

    It 'fails missing application without invoking a same-name function or creating output' {
        $source = New-PreflightDirectory 'missing-magick-source'
        $parent = Join-Path $ownedRoot ('unused-output-' + [Guid]::NewGuid().ToString('N'))
        Mock Get-Command { return $null }
        $result = Invoke-ContainedPreflight -InputArguments @($source) -OutputParent $parent -MagickPath '' -PreflightRunner {
            throw 'Missing application was invoked.'
        } -ProcessRunner { throw 'Missing application attempted conversion.' }
        $result.Code | Should -Be 1
        $result.Text | Should -Match '(?i)ImageMagick.*(not found|missing|application|executable)'
        Assert-PreflightNoOutput $parent
    }

    It 'rejects an explicit non-executable path before launching or creating output' {
        $work = New-PreflightDirectory 'invalid-executable'
        $source = Join-Path $work 'source'
        $null = [IO.Directory]::CreateDirectory($source)
        $fake = Join-Path $work 'magick.ps1'
        [IO.File]::WriteAllText($fake, 'throw "This file must never execute."')
        $parent = Join-Path $work 'output'
        $result = Invoke-ContainedPreflight -InputArguments @($source) -OutputParent $parent -MagickPath $fake -PreflightRunner {
            throw 'Non-executable dependency was launched.'
        }
        $result.Code | Should -Be 1
        $result.Text | Should -Match '(?i)(ImageMagick|magick|executable|application)'
        Assert-PreflightNoOutput $parent
    }

    It 'rejects <Label> dependency response before conversion or output' -ForEach @(
        @{ Label = 'nonzero version exit'; VersionExit = 7; FormatExit = 0; Version = 'current'; Formats = 'valid'; Timeout = $false; Pattern = '(?i)(version|exit|ImageMagick)' },
        @{ Label = 'timeout'; VersionExit = 0; FormatExit = 0; Version = 'current'; Formats = 'valid'; Timeout = $true; Pattern = '(?i)(timed out|timeout)' },
        @{ Label = 'unparseable version'; VersionExit = 0; FormatExit = 0; Version = 'not ImageMagick'; Formats = 'valid'; Timeout = $false; Pattern = '(?i)(version|ImageMagick)' },
        @{ Label = 'older than reviewed supported build'; VersionExit = 0; FormatExit = 0; Version = 'Version: ImageMagick 7.1.2-14 Q16 x64'; Formats = 'valid'; Timeout = $false; Pattern = '(?i)(7\.1\.2-32|supported|patched|upgrade)' },
        @{ Label = 'ImageMagick 6'; VersionExit = 0; FormatExit = 0; Version = 'Version: ImageMagick 6.9.13-0 Q16 x64'; Formats = 'valid'; Timeout = $false; Pattern = '(?i)(7\.1\.2-32|supported|patched|upgrade)' },
        @{ Label = 'nonzero codec command'; VersionExit = 0; FormatExit = 9; Version = 'current'; Formats = 'valid'; Timeout = $false; Pattern = '(?i)(format|codec|exit)' },
        @{ Label = 'unparseable codec table'; VersionExit = 0; FormatExit = 0; Version = 'current'; Formats = 'invalid table'; Timeout = $false; Pattern = '(?i)(format|codec)' }
    ) {
        $work = New-PreflightDirectory 'dependency-response'
        $source = Join-Path $work 'source'
        $null = [IO.Directory]::CreateDirectory($source)
        $parent = Join-Path $work 'output'
        $versionText = $Version
        if ($Version -eq 'current') { $versionText = $supportedVersionText }
        $formatText = $Formats
        if ($Formats -eq 'valid') { $formatText = $supportedFormatText }
        $runner = New-ControlledPreflightRunner -VersionText $versionText -FormatText $formatText -VersionExitCode $VersionExit -FormatExitCode $FormatExit -TimedOut $Timeout
        $result = Invoke-ContainedPreflight -InputArguments @($source) -OutputParent $parent -PreflightRunner $runner -ProcessRunner {
            throw 'Rejected dependency attempted conversion.'
        }
        $result.Code | Should -Be 1
        $result.Text | Should -Match $Pattern
        Assert-PreflightNoOutput $parent
    }

    It 'parses and records the exact real pinned executable, supported version, delegates and relevant codecs' {
        $info = Get-WinImgMagickInfo -MagickPath $magick
        $info.Path | Should -Be $magick
        $info.Version.ToString() | Should -Match '^7\.1\.2[-\.]32$'
        $info.VersionText | Should -Match 'ImageMagick 7\.1\.2-32'
        $info.Delegates | Should -Match 'jpeg'
        $info.Formats.JPEG.Read | Should -BeTrue
        $info.Formats.JPEG.Write | Should -BeTrue
        $info.Formats.PNG.Read | Should -BeTrue
    }

    It 'parses <Label> capability rows including non-starred read-only HEIC' -ForEach @(
        @{ Label = 'module-column'; FormatText = "JPEG* JPEG rw- JPEG description`nHEIC HEIC r-- HEIC description`nPNG* PNG rw- PNG description" },
        @{ Label = 'static portable'; FormatText = "JPEG* rw- JPEG description`nHEIC r-- HEIC description`nPNG* rw- PNG description" }
    ) {
        $info = Get-WinImgMagickInfo -MagickPath $magick -PreflightRunner (New-ControlledPreflightRunner -FormatText $FormatText)
        $info.Formats.JPEG.Read | Should -BeTrue
        $info.Formats.JPEG.Write | Should -BeTrue
        $info.Formats.HEIC.Read | Should -BeTrue
        $info.Formats.HEIC.Write | Should -BeFalse
        $info.Formats.PNG.Read | Should -BeTrue
    }
}

Describe 'M1-T01 per-format capability handling (T011)' {
    It 'accepts a large Int64 cap and emits exact bytes <ExpectedExtent> without narrowing the full Int64 value' -ForEach @(
        @{ Cap = '2251799813685248'; ExpectedExtent = '2251799813685248B' },
        @{ Cap = '2251799813686272'; ExpectedExtent = '2251799813686272B' }
    ) {
        $work = New-PreflightDirectory 'large-cap-conversion'
        $source = Join-Path $work 'source'
        $null = [IO.Directory]::CreateDirectory($source)
        $image = Join-Path $source 'small.png'
        $null = Invoke-PreflightTestMagick @('-size', '3x2', 'xc:#309070', ('PNG:' + $image))
        $sourceBefore = Get-PreflightSourceState $source
        $jpegFixture = Join-Path $work 'synthetic.jpg'
        $null = Invoke-PreflightTestMagick @('-size', '3x2', 'xc:#309070', ('JPEG:' + $jpegFixture))
        $jpegBytes = [IO.File]::ReadAllBytes($jpegFixture)
        $parent = Join-Path $work 'output'
        $trace = New-Object 'Collections.Generic.List[object]'
        $runner = {
            param([string]$Executable, [string[]]$Arguments)
            $trace.Add([pscustomobject]@{ Executable = $Executable; Arguments = @($Arguments) })
            $target = $Arguments[-1].Substring('JPEG:'.Length)
            if (-not (ConvertFrom-PreflightNativePath $target).StartsWith($parent + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Large-cap control escaped owned output.' }
            [IO.File]::WriteAllBytes($target, $jpegBytes)
            return 0
        }
        $parsed = ConvertTo-WinImgByteCap -Value $Cap
        $parsed | Should -BeOfType ([long])
        $parsed | Should -Be ([long]$Cap)
        $result = Invoke-ContainedPreflight -InputArguments @($source, $Cap) -OutputParent $parent -PreflightRunner (New-ControlledPreflightRunner) -ProcessRunner $runner
        $result.Code | Should -Be 0 -Because $result.Text
        $trace.Count | Should -Be 1
        $trace[0].Arguments | Should -Contain ('jpeg:extent=' + $ExpectedExtent)
        $run = @(Get-ChildItem -LiteralPath $parent -Directory)
        $run.Count | Should -Be 1
        Invoke-PreflightTestMagick @('identify', '-format', '%m|%w|%h|%n', (Join-Path $run[0].FullName 'small.jpeg')) | Should -Be 'JPEG|3|2|1'
        Get-PreflightSourceState $source | Should -Be $sourceBefore
    }

    It 'uses the HEIF decoder for .heif when HEIC decoding is unavailable' {
        $work = New-PreflightDirectory 'heif-only-decoder'
        $source = Join-Path $work 'source'
        $null = [IO.Directory]::CreateDirectory($source)
        $image = Join-Path $source 'supported.heif'
        [IO.File]::WriteAllBytes($image, [byte[]](1, 2, 3, 4))
        $sourceBefore = Get-PreflightSourceState $source
        $jpegFixture = Join-Path $work 'synthetic.jpg'
        $null = Invoke-PreflightTestMagick @('-size', '3x2', 'xc:#408090', ('JPEG:' + $jpegFixture))
        $jpegBytes = [IO.File]::ReadAllBytes($jpegFixture)
        $parent = Join-Path $work 'output'
        $formats = $supportedFormatText -replace '(?m)^\s*HEIC.*(?:\r?\n|$)', ''
        # This tiny controlled input tests decoder routing; real HEIC/HEIF
        # collection decoding is required separately in Frames.Tests.ps1.
        Mock Get-WinImgSourceImageInfo {
            return [pscustomobject]@{ SourceCount = 1; Omitted = 0; Unit = 'Images'; Policy = 'DecoderPrimaryOrFirstImage'; Decoder = 'HEIF' }
        }
        Mock Get-WinImgColourInfo {
            return [pscustomobject]@{ ColourSpace = 'sRGB'; HasIcc = $false; Policy = 'AssumeSrgb' }
        }
        $trace = New-Object 'Collections.Generic.List[object]'
        $runner = {
            param([string]$Executable, [string[]]$Arguments)
            $trace.Add([pscustomobject]@{ Executable = $Executable; Arguments = @($Arguments) })
            $target = $Arguments[-1].Substring('JPEG:'.Length)
            if (-not (ConvertFrom-PreflightNativePath $target).StartsWith($parent + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'HEIF control escaped owned output.' }
            [IO.File]::WriteAllBytes($target, $jpegBytes)
            return 0
        }
        $result = Invoke-ContainedPreflight -InputArguments @($source) -OutputParent $parent -PreflightRunner (New-ControlledPreflightRunner -FormatText $formats) -ProcessRunner $runner
        $result.Code | Should -Be 0 -Because $result.Text
        $trace.Count | Should -Be 1
        $trace[0].Arguments[0] | Should -Be '-quiet'
        $trace[0].Arguments[1] | Should -Be '-regard-warnings'
        $trace[0].Arguments[2] | Should -Be '-define'
        $trace[0].Arguments[3] | Should -Be 'registry:filename:literal=true'
        $trace[0].Arguments[4] | Should -Be '-define'
        $trace[0].Arguments[5] | Should -Be 'image:frames=0'
        $snapshotPath = ConvertFrom-PreflightNativePath $trace[0].Arguments[6]
        $snapshotPath.StartsWith($parent + '\', [StringComparison]::OrdinalIgnoreCase) | Should -BeTrue
        [IO.Path]::GetFileName($snapshotPath) | Should -Be 'source.heif'
        Test-Path -LiteralPath $snapshotPath | Should -BeFalse
        $run = @(Get-ChildItem -LiteralPath $parent -Directory)
        $run.Count | Should -Be 1
        Invoke-PreflightTestMagick @('identify', '-format', '%m|%w|%h|%n', (Join-Path $run[0].FullName 'supported.jpeg')) | Should -Be 'JPEG|3|2|1'
        Get-PreflightSourceState $source | Should -Be $sourceBefore
    }

    It 'skips an unavailable HEIC decoder once, converts a supported image and copies video bytes' {
        $work = New-PreflightDirectory 'mixed-codecs'
        $source = Join-Path $work 'source'
        $null = [IO.Directory]::CreateDirectory($source)
        $parent = Join-Path $work 'output'
        $png = Join-Path $source 'supported.png'
        $null = Invoke-PreflightTestMagick @('-size', '24x16', 'gradient:#206090-#C0D080', ('PNG:' + $png))
        [IO.File]::WriteAllBytes((Join-Path $source 'missing-decoder.heic'), [byte[]](0, 1, 2, 3, 4, 5))
        $video = Join-Path $source 'opaque-video.mp4'
        [IO.File]::WriteAllBytes($video, [byte[]](17, 18, 19, 20, 255, 0, 99))
        $before = (Get-FileHash -LiteralPath $video -Algorithm SHA256).Hash
        $sourceBefore = Get-PreflightSourceState $source
        $formatText = $supportedFormatText -replace '(?m)^.*HEIC.*(?:\r?\n|$)', ''
        $preflightTrace = New-Object 'Collections.Generic.List[object]'
        $preflight = New-ControlledPreflightRunner -FormatText $formatText -Trace $preflightTrace
        $conversionTrace = New-Object 'Collections.Generic.List[object]'
        $runner = {
            param([string]$Executable, [string[]]$Arguments)
            $conversionTrace.Add([pscustomobject]@{ Executable = $Executable; Arguments = @($Arguments) })
            & $Executable @Arguments 1>$null 2>$null
            return $LASTEXITCODE
        }
        $result = Invoke-ContainedPreflight -InputArguments @($source) -OutputParent $parent -PreflightRunner $preflight -ProcessRunner $runner
        $result.Code | Should -Be 2 -Because $result.Text
        $conversionTrace.Count | Should -Be 1
        $conversionTrace[0].Executable | Should -Be $magick
        $conversionTrace[0].Arguments[0] | Should -Be '-quiet'
        $conversionTrace[0].Arguments[1] | Should -Be '-regard-warnings'
        $snapshotPath = ConvertFrom-PreflightNativePath $conversionTrace[0].Arguments[4]
        $snapshotPath.StartsWith($parent + '\', [StringComparison]::OrdinalIgnoreCase) | Should -BeTrue
        [IO.Path]::GetFileName($snapshotPath) | Should -Be 'source.png'
        $snapshotPath | Should -Not -Be $png
        Test-Path -LiteralPath $snapshotPath | Should -BeFalse
        $preflightTrace.Count | Should -Be 2
        foreach ($call in $preflightTrace) { $call.Executable | Should -Be $magick }
        $run = @(Get-ChildItem -LiteralPath $parent -Directory)
        $run.Count | Should -Be 1
        $jpeg = Join-Path $run[0].FullName 'supported.jpeg'
        Invoke-PreflightTestMagick @('identify', '-format', '%m|%w|%h|%n', $jpeg) | Should -Be 'JPEG|24|16|1'
        $null = Invoke-PreflightTestMagick @('-quiet', $jpeg, 'null:')
        Test-Path -LiteralPath (Join-Path $run[0].FullName 'missing-decoder.jpeg') | Should -BeFalse
        (Get-FileHash -LiteralPath (Join-Path $run[0].FullName 'opaque-video.mp4') -Algorithm SHA256).Hash | Should -Be $before
        (Get-FileHash -LiteralPath $video -Algorithm SHA256).Hash | Should -Be $before
        Get-PreflightSourceState $source | Should -Be $sourceBefore
        $log = Get-Content -LiteralPath (Get-ChildItem -LiteralPath $run[0].FullName -Recurse -File -Filter '*.log').FullName -Raw
        $log | Should -Match '(?i)missing-decoder\.heic.*(decoder|read|codec|HEIC)'
        $log | Should -Match 'SUMMARY ConvertedImages=1 CopiedVideos=1 Duplicates=0 Unsupported=0 Errors=1'
        $log | Should -Match ([regex]::Escape($magick))
        $log | Should -Match 'ImageMagick 7\.1\.2-32'
        $log | Should -Match '(?i)Delegates.*jpeg'
    }

    It 'rejects a missing JPEG writer for readable images before run output creation' {
        $work = New-PreflightDirectory 'missing-jpeg-writer'
        $source = Join-Path $work 'source'
        $null = [IO.Directory]::CreateDirectory($source)
        [IO.File]::WriteAllBytes((Join-Path $source 'image.png'), [byte[]](1, 2, 3))
        $parent = Join-Path $work 'output'
        $formats = $supportedFormatText -replace 'JPEG\* JPEG      rw-', 'JPEG* JPEG      r--'
        $result = Invoke-ContainedPreflight -InputArguments @($source) -OutputParent $parent -PreflightRunner (New-ControlledPreflightRunner -FormatText $formats) -ProcessRunner {
            throw 'Missing JPEG writer attempted conversion.'
        }
        $result.Code | Should -Be 1
        $result.Text | Should -Match '(?i)JPEG.*(write|encoder|encoding)'
        Assert-PreflightNoOutput $parent
    }
}

Describe 'M1-T01 destination and video-only policy (T012)' {
    It 'budgets a small image practically under an Int64 maximum cap' {
        $parent = New-PreflightDirectory 'space-large-cap'
        $files = @([pscustomobject]@{ Extension = '.png'; Length = [long]4 })
        Mock Get-WinImgAvailableBytes { return [long]90MB }
        $info = Get-WinImgDestinationInfo -Path $parent -Files $files -MaxBytes ([long]::MaxValue)
        $info.RequiredBytes | Should -Be ([decimal]64MB + 1MB + 2 * 4)
        @(Get-ChildItem -LiteralPath $parent -Force).Count | Should -Be 0
    }

    It 'excludes unavailable decoder files from the destination space budget' {
        $work = New-PreflightDirectory 'missing-codec-space'
        $source = Join-Path $work 'source'
        $null = [IO.Directory]::CreateDirectory($source)
        [IO.File]::WriteAllBytes((Join-Path $source 'unsupported.heic'), [byte[]](1, 2, 3, 4))
        $video = Join-Path $source 'video.mp4'
        [IO.File]::WriteAllBytes($video, [byte[]](4, 3, 2, 1))
        $sourceBefore = Get-PreflightSourceState $source
        $parent = Join-Path $work 'output'
        $formats = $supportedFormatText -replace '(?m)^\s*HEIC.*(?:\r?\n|$)', ''
        Mock Get-WinImgAvailableBytes { return [long](64MB + 4) }
        $result = Invoke-ContainedPreflight -InputArguments @($source) -OutputParent $parent -PreflightRunner (New-ControlledPreflightRunner -FormatText $formats) -ProcessRunner {
            throw 'Unavailable decoder file attempted conversion.'
        }
        $result.Code | Should -Be 2 -Because $result.Text
        $run = @(Get-ChildItem -LiteralPath $parent -Directory)
        $run.Count | Should -Be 1
        (Get-FileHash -LiteralPath (Join-Path $run[0].FullName 'video.mp4') -Algorithm SHA256).Hash |
            Should -Be (Get-FileHash -LiteralPath $video -Algorithm SHA256).Hash
        Get-PreflightSourceState $source | Should -Be $sourceBefore
    }

    It 'reports an actual run-directory creation failure as friendly status without a run or log' {
        $source = New-PreflightDirectory 'mkdir-denied-source'
        $parent = New-PreflightDirectory 'mkdir-denied-output'
        Mock New-WinImgExclusiveDirectory { throw [UnauthorizedAccessException]::new('Synthetic run directory creation denied.') }
        $result = Invoke-ContainedPreflight -InputArguments @($source) -OutputParent $parent -PreflightRunner (New-ControlledPreflightRunner)
        $result.Code | Should -Be 1
        $result.Text | Should -Match '(?i)(destination|directory|creation|denied)'
        @(Get-ChildItem -LiteralPath $parent -Force).Count | Should -Be 0
        @(Get-ChildItem -LiteralPath $source -Force).Count | Should -Be 0
    }

    It 'includes copied video bytes, per-image cap budgets and largest-image scratch in the space estimate' {
        $parent = New-PreflightDirectory 'space-estimate'
        $files = @(
            [pscustomobject]@{ Extension = '.mp4'; Length = [long]4 },
            [pscustomobject]@{ Extension = '.png'; Length = [long]3 },
            [pscustomobject]@{ Extension = '.jpg'; Length = [long]4 }
        )
        Mock Get-WinImgAvailableBytes { return [long]::MaxValue }
        $info = Get-WinImgDestinationInfo -Path $parent -Files $files -MaxBytes 1048576
        $info.RequiredBytes | Should -Be ([decimal]64MB + 4 + 2 * 1048576 + 2 * 4)
        $info.AvailableBytes | Should -Be ([long]::MaxValue)
        @(Get-ChildItem -LiteralPath $parent -Force).Count | Should -Be 0
    }

    It 'rejects an estimate exceeding Int64 capacity without overflow wrapping or large fixture allocation' {
        $parent = New-PreflightDirectory 'space-overflow'
        $files = @([pscustomobject]@{ Extension = '.mp4'; Length = [long]::MaxValue })
        Mock Get-WinImgAvailableBytes { return [long]::MaxValue }
        { Get-WinImgDestinationInfo -Path $parent -Files $files -MaxBytes 1048576 } | Should -Throw '*Insufficient*space*'
        @(Get-ChildItem -LiteralPath $parent -Force).Count | Should -Be 0
    }

    It 'reports a denied write probe before output, leaving an unrelated sentinel untouched' {
        $work = New-PreflightDirectory 'denied-parent'
        $source = Join-Path $work 'source'
        $null = [IO.Directory]::CreateDirectory($source)
        $parent = Join-Path $work 'output'
        $sentinel = Join-Path $work 'unrelated.txt'
        [IO.File]::WriteAllText($sentinel, 'Unrelated synthetic sentinel.')
        $hash = (Get-FileHash -LiteralPath $sentinel -Algorithm SHA256).Hash
        Mock Test-WinImgDestinationWritable { throw [UnauthorizedAccessException]::new('Synthetic destination write denied.') }
        $result = Invoke-ContainedPreflight -InputArguments @($source) -OutputParent $parent -PreflightRunner (New-ControlledPreflightRunner)
        $result.Code | Should -Be 1
        $result.Text | Should -Match '(?i)(destination|write|writable|denied)'
        Assert-PreflightNoOutput $parent
        (Get-FileHash -LiteralPath $sentinel -Algorithm SHA256).Hash | Should -Be $hash
        @(Get-ChildItem -LiteralPath $work -Force).Count | Should -Be 2
    }

    It 'rejects a regular file destination parent before output creation' {
        $work = New-PreflightDirectory 'file-parent'
        $source = Join-Path $work 'source'
        $null = [IO.Directory]::CreateDirectory($source)
        $parent = Join-Path $work 'not-a-directory'
        [IO.File]::WriteAllText($parent, 'Synthetic destination path collision.')
        $hash = (Get-FileHash -LiteralPath $parent -Algorithm SHA256).Hash
        $result = Invoke-ContainedPreflight -InputArguments @($source) -OutputParent $parent -PreflightRunner (New-ControlledPreflightRunner)
        $result.Code | Should -Be 1
        $result.Text | Should -Match '(?i)(destination|directory|parent)'
        (Get-FileHash -LiteralPath $parent -Algorithm SHA256).Hash | Should -Be $hash
        @(Get-ChildItem -LiteralPath $work -Force).Count | Should -Be 2
    }

    It 'rejects known insufficient space without run output or leftover probes' {
        $work = New-PreflightDirectory 'low-space'
        $source = Join-Path $work 'source'
        $null = [IO.Directory]::CreateDirectory($source)
        [IO.File]::WriteAllBytes((Join-Path $source 'video.mp4'), [byte[]](1, 2, 3, 4))
        $parent = Join-Path $work 'output'
        Mock Get-WinImgAvailableBytes { return [long]1 }
        $result = Invoke-ContainedPreflight -InputArguments @($source) -OutputParent $parent -PreflightRunner (New-ControlledPreflightRunner)
        $result.Code | Should -Be 1
        $result.Text | Should -Match '(?i)(space|available|capacity)'
        Assert-PreflightNoOutput $parent
        @(Get-ChildItem -LiteralPath $work -Force).Count | Should -Be 1
    }

    It 'continues video-only input when available space cannot be determined, with an explicit advisory' {
        $work = New-PreflightDirectory 'unknown-space'
        $source = Join-Path $work 'source'
        $null = [IO.Directory]::CreateDirectory($source)
        $video = Join-Path $source 'video.mp4'
        [IO.File]::WriteAllBytes($video, [byte[]](30, 31, 32, 33))
        $parent = Join-Path $work 'output'
        Mock Get-WinImgAvailableBytes { return $null }
        $result = Invoke-ContainedPreflight -InputArguments @($source) -OutputParent $parent -PreflightRunner (New-ControlledPreflightRunner) -ProcessRunner {
            throw 'Video-only input attempted image conversion.'
        }
        $result.Code | Should -Be 0 -Because $result.Text
        $result.Text | Should -Match '(?i)(space|capacity).*(unknown|unavailable|unable|could not|cannot|best.effort)'
        $run = @(Get-ChildItem -LiteralPath $parent -Directory)
        $run.Count | Should -Be 1
        (Get-FileHash -LiteralPath (Join-Path $run[0].FullName 'video.mp4') -Algorithm SHA256).Hash |
            Should -Be (Get-FileHash -LiteralPath $video -Algorithm SHA256).Hash
    }

    It 'requires reviewed ImageMagick for video-only input and does not copy on dependency failure' {
        $work = New-PreflightDirectory 'video-dependency-failure'
        $source = Join-Path $work 'source'
        $null = [IO.Directory]::CreateDirectory($source)
        $video = Join-Path $source 'video.mp4'
        [IO.File]::WriteAllBytes($video, [byte[]](50, 51, 52, 53))
        $before = (Get-FileHash -LiteralPath $video -Algorithm SHA256).Hash
        $parent = Join-Path $work 'output'
        $result = Invoke-ContainedPreflight -InputArguments @($source) -OutputParent $parent -PreflightRunner (New-ControlledPreflightRunner -VersionExitCode 3)
        $result.Code | Should -Be 1
        $result.Text | Should -Match '(?i)(ImageMagick|version)'
        Assert-PreflightNoOutput $parent
        (Get-FileHash -LiteralPath $video -Algorithm SHA256).Hash | Should -Be $before
    }

    It 'does not require JPEG encoding for video-only input after dependency verification' {
        $work = New-PreflightDirectory 'video-without-jpeg'
        $source = Join-Path $work 'source'
        $null = [IO.Directory]::CreateDirectory($source)
        $video = Join-Path $source 'video.mp4'
        [IO.File]::WriteAllBytes($video, [byte[]](80, 81, 82, 83, 84))
        $parent = Join-Path $work 'output'
        $formats = $supportedFormatText -replace '(?m)^.*JPEG.*(?:\r?\n|$)', ''
        $result = Invoke-ContainedPreflight -InputArguments @($source) -OutputParent $parent -PreflightRunner (New-ControlledPreflightRunner -FormatText $formats) -ProcessRunner {
            throw 'Video-only input attempted image conversion.'
        }
        $result.Code | Should -Be 0 -Because $result.Text
        $run = @(Get-ChildItem -LiteralPath $parent -Directory)
        $run.Count | Should -Be 1
        (Get-FileHash -LiteralPath (Join-Path $run[0].FullName 'video.mp4') -Algorithm SHA256).Hash |
            Should -Be (Get-FileHash -LiteralPath $video -Algorithm SHA256).Hash
    }

    It 'cleans the real exclusive write probe without changing an existing parent' {
        $parent = New-PreflightDirectory 'real-write-probe'
        $sentinel = Join-Path $parent 'sentinel.txt'
        [IO.File]::WriteAllText($sentinel, 'Synthetic write-probe sentinel.')
        $hash = (Get-FileHash -LiteralPath $sentinel -Algorithm SHA256).Hash
        $null = Test-WinImgDestinationWritable -Path $parent
        @(Get-ChildItem -LiteralPath $parent -Force).Count | Should -Be 1
        (Get-FileHash -LiteralPath $sentinel -Algorithm SHA256).Hash | Should -Be $hash
    }
}
