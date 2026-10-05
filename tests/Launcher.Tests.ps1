BeforeAll {
    if ($env:OS -ne 'Windows_NT') { throw 'Launcher tests require actual Windows CMD and Windows PowerShell.' }
    $repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    $launcher = Join-Path $repository 'WinImgNormalizer.bat'
    $suitePath = Join-Path $PSScriptRoot 'Launcher.Tests.ps1'
    $scratch = [IO.Path]::GetFullPath((Join-Path $repository '.scratch'))

    function Assert-LauncherNoReparseAncestors([string]$Path) {
        $current = [IO.DirectoryInfo]::new([IO.Path]::GetFullPath($Path))
        while ($null -ne $current) {
            if ($current.Exists -and (($current.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0)) {
                throw 'Launcher fixture path includes a reparse point.'
            }
            $current = $current.Parent
        }
    }
    Assert-LauncherNoReparseAncestors $scratch
    foreach ($pictures in @([Environment]::GetFolderPath('MyPictures'), (Join-Path $env:USERPROFILE 'Pictures'))) {
        if ([string]::IsNullOrWhiteSpace($pictures)) { continue }
        $picturesFull = [IO.Path]::GetFullPath($pictures).TrimEnd('\', '/')
        if ($scratch.Equals($picturesFull, [StringComparison]::OrdinalIgnoreCase) -or
            $scratch.StartsWith($picturesFull + '\', [StringComparison]::OrdinalIgnoreCase) -or
            $picturesFull.StartsWith($scratch + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Launcher synthetic scratch overlaps a real Pictures location.'
        }
    }
    $git = Get-Command git -CommandType Application -ErrorAction Stop | Select-Object -First 1
    & $git.Source -C $repository check-ignore --quiet --no-index -- (Join-Path $scratch 'launcher-ignore-probe')
    if ($LASTEXITCODE -ne 0) { throw 'Launcher synthetic scratch must already be ignored.' }
    $ownedRoot = Join-Path $scratch ('M3-T05-launcher-' + [Guid]::NewGuid().ToString('N'))
    if (Test-Path -LiteralPath $ownedRoot) { throw 'Launcher fixture ownership directory already exists.' }
    [IO.Directory]::CreateDirectory($ownedRoot) | Out-Null
    [IO.File]::WriteAllText((Join-Path $ownedRoot 'fixture-owner.txt'), 'Owned synthetic Windows launcher fixtures; no real Pictures, persistent environment/policy change or recursive cleanup.', [Text.UTF8Encoding]::new($false))
    $observations = New-Object 'Collections.Generic.List[object]'
    $productionBytes = [IO.File]::ReadAllBytes($launcher)
    $bindingsBefore = @($launcher, $suitePath) | ForEach-Object {
        [pscustomobject]@{ Path = $_.Substring($repository.Length + 1).Replace('\', '/'); Sha256 = (Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash.ToLowerInvariant() }
    }
    $cmd = Join-Path $env:SystemRoot 'System32/cmd.exe'
    $legacyHost = Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe'
    if (-not [IO.File]::Exists($cmd) -or -not [IO.File]::Exists($legacyHost)) { throw 'Required Windows launcher hosts are missing.' }

    function Get-LauncherSecurityState {
        [pscustomobject]@{
            Policies = @(Get-ExecutionPolicy -List | Where-Object { $_.Scope -ne 'Process' } | ForEach-Object {
                [pscustomobject]@{ Scope = $_.Scope.ToString(); Policy = $_.ExecutionPolicy.ToString() }
            })
            UserPath = [Environment]::GetEnvironmentVariable('PATH', 'User')
            MachinePath = [Environment]::GetEnvironmentVariable('PATH', 'Machine')
            ProcessPath = [Environment]::GetEnvironmentVariable('PATH', 'Process')
            UserProfile = [Environment]::GetEnvironmentVariable('USERPROFILE', 'Process')
            ProcessPolicy = (Get-ExecutionPolicy -Scope Process).ToString()
        } | ConvertTo-Json -Depth 5 -Compress
    }
    $securityBefore = Get-LauncherSecurityState
    function Get-LauncherLegacyPolicies {
        $start = New-Object Diagnostics.ProcessStartInfo
        $start.FileName = $legacyHost
        $command = '@(Get-ExecutionPolicy -List | Where-Object { $_.Scope -ne ''Process'' } | ForEach-Object { [pscustomobject]@{ Scope = $_.Scope.ToString(); Policy = $_.ExecutionPolicy.ToString() } }) | ConvertTo-Json -Depth 4 -Compress'
        $start.Arguments = '-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command "' + $command + '"'
        $start.WorkingDirectory = $ownedRoot
        $start.UseShellExecute = $false
        $start.CreateNoWindow = $true
        $start.RedirectStandardOutput = $true
        $start.RedirectStandardError = $true
        foreach ($key in @($start.EnvironmentVariables.Keys)) {
            if ([string]::Equals([string]$key, 'PSModulePath', [StringComparison]::OrdinalIgnoreCase)) { $start.EnvironmentVariables.Remove([string]$key) }
        }
        $start.EnvironmentVariables['TEMP'] = $ownedRoot
        $start.EnvironmentVariables['TMP'] = $ownedRoot
        $process = New-Object Diagnostics.Process
        $process.StartInfo = $start
        $started = $false
        try {
            if (-not $process.Start()) { throw 'Legacy-host policy probe failed to start.' }
            $started = $true
            $stdout = $process.StandardOutput.ReadToEndAsync()
            $stderr = $process.StandardError.ReadToEndAsync()
            if (-not $process.WaitForExit(15000)) { throw 'Legacy-host policy probe exceeded its 15-second bound.' }
            if (-not $stdout.Wait(5000) -or -not $stderr.Wait(5000)) { throw 'Legacy-host policy probe streams did not finish.' }
            if ($process.ExitCode -ne 0 -or $stderr.GetAwaiter().GetResult()) { throw 'Legacy-host policy probe failed.' }
            return (($stdout.GetAwaiter().GetResult() | ConvertFrom-Json) | ConvertTo-Json -Depth 5 -Compress)
        } finally {
            if ($started -and -not $process.HasExited) { $process.Kill(); $null = $process.WaitForExit(5000) }
            $process.Dispose()
        }
    }
    # PowerShell editions can use different persistent-policy registry views;
    # compare each host with its own observation before launcher execution.
    $legacyPoliciesBefore = Get-LauncherLegacyPolicies
    $driverText = @'
$ErrorActionPreference = 'Stop'
$record = [ordered]@{
    Arguments = @($args)
    HostExecutable = (Get-Process -Id $PID).Path
    Version = $PSVersionTable.PSVersion.ToString()
    Edition = $PSVersionTable.PSEdition
    ProcessPolicy = (Get-ExecutionPolicy -Scope Process).ToString()
    PersistentPolicies = @(Get-ExecutionPolicy -List | Where-Object { $_.Scope -ne 'Process' } | ForEach-Object {
        [pscustomobject]@{ Scope = $_.Scope.ToString(); Policy = $_.ExecutionPolicy.ToString() }
    })
    UserSid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    IsAdministrator = ([Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    ProcessId = $PID
    StartedTicks = (Get-Process -Id $PID).StartTime.ToUniversalTime().Ticks
    ExitCode = [int]$env:WINIMG_LAUNCHER_CODE
}
[IO.File]::WriteAllText($env:WINIMG_LAUNCHER_RECORD, ($record | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))
[IO.File]::WriteAllText(($env:WINIMG_LAUNCHER_RECORD + '.ready'), 'Driver observation is complete.', [Text.UTF8Encoding]::new($false))
exit ([int]$env:WINIMG_LAUNCHER_CODE)
'@

    function New-LauncherCase([string]$Label, [switch]$MissingDriver) {
        $root = [IO.Path]::GetFullPath((Join-Path $ownedRoot ($Label + '-' + [Guid]::NewGuid().ToString('N'))))
        if (-not $root.StartsWith($ownedRoot + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Launcher case escapes its owned directory.' }
        Assert-LauncherNoReparseAncestors $root
        [IO.Directory]::CreateDirectory($root) | Out-Null
        $bat = Join-Path $root 'WinImgNormalizer.bat'
        [IO.File]::WriteAllBytes($bat, $productionBytes)
        if ([Convert]::ToBase64String([IO.File]::ReadAllBytes($bat)) -ne [Convert]::ToBase64String($productionBytes)) { throw 'Copied BAT differs from exact production bytes.' }
        if (-not $MissingDriver) { [IO.File]::WriteAllText((Join-Path $root 'WinImgNormalizer.ps1'), $driverText, [Text.UTF8Encoding]::new($true)) }
        $source = Join-Path $root 'source folder'
        [IO.Directory]::CreateDirectory($source) | Out-Null
        $temporary = Join-Path $root 'temporary'
        [IO.Directory]::CreateDirectory($temporary) | Out-Null
        [pscustomobject]@{ Root = $root; Bat = $bat; Source = $source; Temporary = $temporary; Record = (Join-Path $root 'driver-observation.json') }
    }

    function Invoke-LauncherChild {
        param($Case, [string[]]$Arguments, [int]$Code = 0, [switch]$ExpectDriver, [switch]$VerifyPause, [hashtable]$Environment)
        $start = New-Object Diagnostics.ProcessStartInfo
        $start.FileName = $cmd
        # CMD expands each environment reference once; literal percent signs in
        # the resulting value are not re-expanded. CALL would add another parse.
        $command = '"%WINIMG_LAUNCHER_BAT%"'
        for ($i = 0; $i -lt $Arguments.Count; $i++) {
            $key = 'WINIMG_LAUNCHER_ARG' + $i
            if ([string]$Arguments[$i] -eq '') {
                # An undefined CMD environment reference remains literal;
                # use actual empty quoted argv for this count regression.
                $command += ' ""'
            } else {
                $start.EnvironmentVariables[$key] = [string]$Arguments[$i]
                $command += ' "%' + $key + '%"'
            }
        }
        $start.Arguments = '/d /v:off /s /c "' + $command + '"'
        $start.WorkingDirectory = $Case.Root
        $start.UseShellExecute = $false
        $start.CreateNoWindow = $true
        $start.RedirectStandardInput = $true
        $start.RedirectStandardOutput = $true
        $start.RedirectStandardError = $true
        foreach ($key in @($start.EnvironmentVariables.Keys)) {
            if ([string]::Equals([string]$key, 'PSModulePath', [StringComparison]::OrdinalIgnoreCase)) { $start.EnvironmentVariables.Remove([string]$key) }
        }
        # Select the existing legacy host with only a child PATH change. The BAT
        # itself remains byte-for-byte production, including host selection.
        $start.EnvironmentVariables['PATH'] = [IO.Path]::GetDirectoryName($legacyHost) + ';' + $env:PATH
        $start.EnvironmentVariables['TEMP'] = $Case.Temporary
        $start.EnvironmentVariables['TMP'] = $Case.Temporary
        $start.EnvironmentVariables['WINIMG_LAUNCHER_BAT'] = $Case.Bat
        $start.EnvironmentVariables['WINIMG_LAUNCHER_RECORD'] = $Case.Record
        $start.EnvironmentVariables['WINIMG_LAUNCHER_CODE'] = [string]$Code
        $start.EnvironmentVariables['WINIMG_TEST_PERCENT'] = 'MUST-NOT-EXPAND-PERCENT'
        $start.EnvironmentVariables['WINIMG_TEST_BANG'] = 'MUST-NOT-EXPAND-BANG'
        if ($Environment) {
            foreach ($key in $Environment.Keys) { $start.EnvironmentVariables[[string]$key] = [string]$Environment[$key] }
        }
        $process = New-Object Diagnostics.Process
        $process.StartInfo = $start
        $pauseObserved = $false
        $driver = $null
        $started = $false
        try {
            if (-not $process.Start()) { throw 'Owned CMD launcher child did not start.' }
            $started = $true
            $stdout = $process.StandardOutput.ReadToEndAsync()
            $stderr = $process.StandardError.ReadToEndAsync()
            if ($ExpectDriver) {
                $clock = [Diagnostics.Stopwatch]::StartNew()
                while (-not [IO.File]::Exists($Case.Record + '.ready') -and -not $process.HasExited -and $clock.ElapsedMilliseconds -lt 15000) { Start-Sleep -Milliseconds 20 }
                if ([IO.File]::Exists($Case.Record + '.ready')) {
                    $driver = [IO.File]::ReadAllText($Case.Record) | ConvertFrom-Json
                    # Confirm the real PowerShell process has exited before
                    # claiming the BAT is waiting at its retained pause.
                    while ($clock.ElapsedMilliseconds -lt 15000) {
                        $sameDriver = $false
                        try {
                            $child = [Diagnostics.Process]::GetProcessById([int]$driver.ProcessId)
                            try { $sameDriver = (-not $child.HasExited -and $child.StartTime.ToUniversalTime().Ticks -eq [long]$driver.StartedTicks) } finally { $child.Dispose() }
                        } catch [ArgumentException] { $sameDriver = $false }
                        if (-not $sameDriver) { break }
                        Start-Sleep -Milliseconds 20
                    }
                    if ($sameDriver) { throw 'Synthetic PowerShell driver did not exit within its readiness bound.' }
                }
            }
            if ($VerifyPause) { $pauseObserved = -not $process.WaitForExit(300) }
            if (-not $process.HasExited) {
                $process.StandardInput.WriteLine('x')
                $process.StandardInput.Flush()
                $process.StandardInput.Close()
            }
            if (-not $process.WaitForExit(15000)) { throw 'Owned CMD launcher child exceeded its 15-second completion bound.' }
            if (-not $stdout.Wait(5000) -or -not $stderr.Wait(5000)) { throw 'Owned CMD launcher streams did not finish.' }
            $result = [pscustomobject]@{ ExitCode = $process.ExitCode; StdOut = $stdout.GetAwaiter().GetResult(); StdErr = $stderr.GetAwaiter().GetResult(); PauseObserved = $pauseObserved; Driver = $driver }
            $observations.Add([pscustomobject]@{ Case = $Case.Root; Arguments = @($Arguments); RequestedCode = $Code; Result = $result; ProductionBatSha256 = $bindingsBefore[0].Sha256 })
            [IO.File]::WriteAllText((Join-Path $Case.Root 'launcher-result.json'), ($result | ConvertTo-Json -Depth 8), [Text.UTF8Encoding]::new($false))
            return $result
        } finally {
            if ($started -and -not $process.HasExited) {
                & (Join-Path $env:SystemRoot 'System32/taskkill.exe') /PID $process.Id /T /F 2>&1 | Out-Null
                if (-not $process.WaitForExit(5000)) { $process.Kill() }
            }
            $process.Dispose()
        }
    }
}

AfterAll {
    if ($ownedRoot) {
        $securityAfter = Get-LauncherSecurityState
        $legacyPoliciesAfter = Get-LauncherLegacyPolicies
        [IO.File]::WriteAllText((Join-Path $ownedRoot 'launcher-observations.json'), (@($observations.ToArray()) | ConvertTo-Json -Depth 10), [Text.UTF8Encoding]::new($false))
        [IO.File]::WriteAllText((Join-Path $ownedRoot 'security-observation.json'), ([ordered]@{ Before = ($securityBefore | ConvertFrom-Json); After = ($securityAfter | ConvertFrom-Json); LegacyPoliciesBefore = ($legacyPoliciesBefore | ConvertFrom-Json); LegacyPoliciesAfter = ($legacyPoliciesAfter | ConvertFrom-Json); Unchanged = ($securityBefore -eq $securityAfter -and $legacyPoliciesBefore -eq $legacyPoliciesAfter); SourceBindings = $bindingsBefore } | ConvertTo-Json -Depth 8), [Text.UTF8Encoding]::new($false))
        $securityAfter | Should -BeExactly $securityBefore
        $legacyPoliciesAfter | Should -BeExactly $legacyPoliciesBefore
        foreach ($binding in $bindingsBefore) {
            (Get-FileHash -LiteralPath (Join-Path $repository $binding.Path) -Algorithm SHA256).Hash.ToLowerInvariant() | Should -BeExactly $binding.Sha256
        }
    }
}

Describe 'M3-T05 actual Windows CMD launcher contract' {
    It 'T061 preserves native exit <Code> after real child exit, message and blocked pause' -ForEach @(
        @{ Code = 0; Message = 'Completed successfully.' },
        @{ Code = 1; Message = 'ERROR: Processing did not complete reliably (exit code 1).' },
        @{ Code = 2; Message = 'Completed with warnings or partial failures.' },
        @{ Code = 130; Message = 'Cancelled. Completed outputs were retained.' },
        @{ Code = 37; Message = 'ERROR: Processing did not complete reliably (exit code 37).' }
    ) {
        $case = New-LauncherCase ('code-' + $Code)
        $result = Invoke-LauncherChild -Case $case -Arguments @($case.Source) -Code $Code -ExpectDriver -VerifyPause
        $result.Driver | Should -Not -BeNullOrEmpty
        $result.Driver.ExitCode | Should -Be $Code
        $result.ExitCode | Should -Be $Code
        $result.PauseObserved | Should -BeTrue
        $result.StdOut | Should -Match ([regex]::Escape($Message))
        $result.StdOut | Should -Match 'Press any key to close\.'
        $result.StdErr | Should -BeNullOrEmpty
        if ($Code -ne 0) { $result.StdOut | Should -Not -Match 'Completed successfully\.|\bDone\.' }
    }

    It 'T061 ignores an inherited ERRORLEVEL variable when the actual child returns <Code>' -ForEach @(
        @{ Code = 0; Message = 'Completed successfully.' },
        @{ Code = 2; Message = 'Completed with warnings or partial failures.' },
        @{ Code = 130; Message = 'Cancelled. Completed outputs were retained.' }
    ) {
        $case = New-LauncherCase ('inherited-errorlevel-' + $Code)
        $result = Invoke-LauncherChild -Case $case -Arguments @($case.Source) -Code $Code -ExpectDriver -VerifyPause -Environment @{ ERRORLEVEL = '47' }
        $result.Driver.ExitCode | Should -Be $Code
        $result.ExitCode | Should -Be $Code
        $result.PauseObserved | Should -BeTrue
        $result.StdOut | Should -Match ([regex]::Escape($Message))
        $result.StdOut | Should -Not -Match 'exit code 47'
        $result.StdErr | Should -BeNullOrEmpty
    }

    It 'T062 rejects <Label> with usage code before starting PowerShell' -ForEach @(
        @{ Label = 'zero arguments'; Mode = 'zero' },
        @{ Label = 'an empty first argument'; Mode = 'empty' },
        @{ Label = 'two folders'; Mode = 'two' },
        @{ Label = 'a hostile literal second folder'; Mode = 'hostile-two' },
        @{ Label = 'an empty extra argument'; Mode = 'extra-empty' },
        @{ Label = 'an empty second argument followed by a third'; Mode = 'third' },
        @{ Label = 'two empty extra arguments'; Mode = 'two-empty' }
    ) {
        $case = New-LauncherCase ('count-' + $Mode)
        $arguments = switch ($Mode) {
            'zero' { @() }
            'empty' { @('') }
            'two' { @($case.Source, $case.Source) }
            'hostile-two' {
                $extra = Join-Path $case.Root ('dropped ' + [char]0x00E5 + [char]0x96EA + ' %WINIMG_TEST_PERCENT% !WINIMG_TEST_BANG! &(part)[part]')
                [IO.Directory]::CreateDirectory($extra) | Out-Null
                @($case.Source, $extra)
            }
            'extra-empty' { @($case.Source, '') }
            'third' { @($case.Source, '', $case.Source) }
            'two-empty' { @($case.Source, '', '') }
        }
        $result = Invoke-LauncherChild -Case $case -Arguments $arguments -VerifyPause
        $result.ExitCode | Should -Be 1
        $result.PauseObserved | Should -BeTrue
        $result.StdOut | Should -Match 'ERROR:.*exactly one FOLDER'
        $result.StdOut | Should -Not -Match 'Completed successfully\.|\bDone\.'
        $result.StdErr | Should -BeNullOrEmpty
        [IO.File]::Exists($case.Record) | Should -BeFalse
    }

    It 'T062 reports a missing companion script with code1 after the pause' {
        $case = New-LauncherCase 'missing-driver' -MissingDriver
        $result = Invoke-LauncherChild -Case $case -Arguments @($case.Source) -VerifyPause
        $result.ExitCode | Should -Be 1
        $result.PauseObserved | Should -BeTrue
        $result.StdOut | Should -Match 'ERROR: Missing PowerShell script:'
        $result.StdOut | Should -Not -Match 'Completed successfully\.|\bDone\.'
        $result.StdErr | Should -BeNullOrEmpty
        [IO.File]::Exists($case.Record) | Should -BeFalse
    }

    It 'T061 retains native9009 and a failure message when the unchanged host is unavailable in child PATH' {
        $case = New-LauncherCase 'missing-host'
        $result = Invoke-LauncherChild -Case $case -Arguments @($case.Source) -VerifyPause -Environment @{ PATH = '' }
        $result.ExitCode | Should -Be 9009
        $result.PauseObserved | Should -BeTrue
        $result.StdOut | Should -Match 'ERROR: Processing did not complete reliably \(exit code 9009\)\.'
        $result.StdOut | Should -Not -Match 'Completed successfully\.|\bDone\.'
        $result.StdErr | Should -Match 'powershell'
        [IO.File]::Exists($case.Record) | Should -BeFalse
    }

    It 'T062 preserves a single literal source and companion path with <Label>' -ForEach @(
        @{ Label = 'spaces'; Directory = 'folder with spaces'; Separator = '' },
        @{ Label = 'Unicode'; Directory = 'Unicode-' + [char]0x00E5 + [char]0x96EA; Separator = '' },
        @{ Label = 'percent signs'; Directory = 'percent-%WINIMG_TEST_PERCENT%'; Separator = '' },
        @{ Label = 'exclamation marks'; Directory = 'bang-!WINIMG_TEST_BANG!'; Separator = '' },
        @{ Label = 'ampersand'; Directory = 'folder&part'; Separator = '' },
        @{ Label = 'parentheses'; Directory = 'folder(part)'; Separator = '' },
        @{ Label = 'brackets'; Directory = 'folder[part]'; Separator = '' },
        @{ Label = 'all characters'; Directory = 'spaced ' + [char]0x00E5 + [char]0x96EA + ' %WINIMG_TEST_PERCENT% !WINIMG_TEST_BANG! &(part)[part]'; Separator = '' },
        @{ Label = 'all characters and trailing backslash'; Directory = 'trailing ' + [char]0x00E5 + [char]0x96EA + ' %WINIMG_TEST_PERCENT% !WINIMG_TEST_BANG! &(part)[part]'; Separator = '\' },
        @{ Label = 'all characters and trailing slash'; Directory = 'slash ' + [char]0x00E5 + [char]0x96EA + ' %WINIMG_TEST_PERCENT% !WINIMG_TEST_BANG! &(part)[part]'; Separator = '/' }
    ) {
        $case = New-LauncherCase $Directory
        $inputArgument = $case.Source + $Separator
        $expectedArgument = $inputArgument
        if ($Separator) { $expectedArgument += '.' }
        $result = Invoke-LauncherChild -Case $case -Arguments @($inputArgument) -ExpectDriver
        $result.ExitCode | Should -Be 0
        $result.StdOut | Should -Match 'Completed successfully\.'
        $result.StdErr | Should -BeNullOrEmpty
        @($result.Driver.Arguments).Count | Should -Be 1
        $result.Driver.Arguments[0] | Should -BeExactly $expectedArgument
        [IO.Path]::GetFullPath([string]$result.Driver.Arguments[0]).TrimEnd('\', '/') | Should -BeExactly $case.Source
    }

    It 'T063 retains the legacy host, process-only Bypass, user identity and persistent security context' {
        $case = New-LauncherCase 'security-context'
        $before = Get-LauncherSecurityState
        $result = Invoke-LauncherChild -Case $case -Arguments @($case.Source) -ExpectDriver -VerifyPause
        $result.ExitCode | Should -Be 0
        $result.Driver.HostExecutable | Should -BeExactly $legacyHost
        $result.Driver.Edition | Should -Be 'Desktop'
        $result.Driver.Version | Should -Match '^5\.1\.'
        $result.Driver.ProcessPolicy | Should -Be 'Bypass'
        $result.Driver.UserSid | Should -BeExactly ([Security.Principal.WindowsIdentity]::GetCurrent().User.Value)
        $result.Driver.IsAdministrator | Should -Be ([Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
        ($result.Driver.PersistentPolicies | ConvertTo-Json -Depth 5 -Compress) | Should -BeExactly $legacyPoliciesBefore
        Get-LauncherLegacyPolicies | Should -BeExactly $legacyPoliciesBefore
        Get-LauncherSecurityState | Should -BeExactly $before
        Get-LauncherSecurityState | Should -BeExactly $securityBefore
    }
}
