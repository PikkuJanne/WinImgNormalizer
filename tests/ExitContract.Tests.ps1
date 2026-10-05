BeforeAll {
    $repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    $application = Join-Path $repository 'WinImgNormalizer.ps1'
    $launcher = Join-Path $repository 'WinImgNormalizer.bat'
    $scratch = Join-Path $repository '.scratch'
    function Assert-ExitAncestors([string]$Path) {
        $current = [IO.DirectoryInfo]::new([IO.Path]::GetFullPath($Path))
        while ($current) {
            if ($current.Exists -and ($current.Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Exit fixture crosses a reparse point.' }
            $current = $current.Parent
        }
    }
    Assert-ExitAncestors $scratch
    foreach ($pictures in @([Environment]::GetFolderPath('MyPictures'), (Join-Path $env:USERPROFILE 'Pictures'))) {
        if (-not $pictures) { continue }
        $protected = [IO.Path]::GetFullPath($pictures).TrimEnd('\', '/')
        if ($scratch.Equals($protected, [StringComparison]::OrdinalIgnoreCase) -or
            $scratch.StartsWith($protected + '\', [StringComparison]::OrdinalIgnoreCase) -or
            $protected.StartsWith($scratch + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Exit fixtures overlap real Pictures.' }
    }
    & git -C $repository check-ignore --quiet --no-index -- (Join-Path $scratch 'exit-ignore-probe')
    if ($LASTEXITCODE -ne 0) { throw 'Exit fixtures must already be ignored.' }
    $ownedRoot = Join-Path $scratch ('M3-T05-exit-' + [Guid]::NewGuid().ToString('N'))
    if ([IO.Directory]::Exists($ownedRoot)) { throw 'Exit fixture ownership collision.' }
    $null = [IO.Directory]::CreateDirectory($ownedRoot)
    [IO.File]::WriteAllText((Join-Path $ownedRoot '.winimg-fixture-root'), 'M3-T05 owned synthetic exit fixtures; retained for inspection.', [Text.UTF8Encoding]::new($false))
    if (-not $env:WINIMG_TEST_MAGICK -or -not [IO.Path]::IsPathRooted($env:WINIMG_TEST_MAGICK)) { throw 'Explicit verified ImageMagick required.' }
    $magick = [IO.Path]::GetFullPath($env:WINIMG_TEST_MAGICK)
    Assert-ExitAncestors $magick
    $imagePin = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'legacy/toolchain.json') -Raw | ConvertFrom-Json
    (Get-FileHash -LiteralPath $magick).Hash.ToLowerInvariant() | Should -Be $imagePin.portable_imagemagick.executable_sha256
    $hostExecutable = (Get-Process -Id $PID).Path
    $consoleFixtureSource = Join-Path $PSScriptRoot 'fixtures/LauncherConsoleFixture.cs'
    $consoleFixtureCompiler = Join-Path $env:SystemRoot 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
    $consoleFixtureExecutable = Join-Path $ownedRoot 'LauncherConsoleFixture.exe'
    Assert-ExitAncestors $consoleFixtureSource; Assert-ExitAncestors $consoleFixtureCompiler
    if (-not [IO.File]::Exists($consoleFixtureSource) -or -not [IO.File]::Exists($consoleFixtureCompiler)) { throw 'Shared console fixture source and existing Windows Framework compiler required.' }
    $fixed = [DateTime]::Parse('2019-04-05T06:07:08Z').ToUniversalTime()
    $observations = New-Object 'Collections.Generic.List[object]'
    function Get-ExitBindings {
        foreach ($path in @($application, $launcher, (Join-Path $PSScriptRoot 'ExitContract.Tests.ps1'), $consoleFixtureSource, $consoleFixtureCompiler, $magick)) {
            [pscustomobject]@{ Path = $path; Sha256 = (Get-FileHash -LiteralPath $path).Hash.ToLowerInvariant() }
        }
    }
    $bindingsBefore = @(Get-ExitBindings)
    & $consoleFixtureCompiler /nologo /target:winexe /reference:System.Web.Extensions.dll ('/out:' + $consoleFixtureExecutable) $consoleFixtureSource *> (Join-Path $ownedRoot 'console-fixture-compiler.log')
    if ($LASTEXITCODE -ne 0) { throw 'Shared hidden console fixture compilation failed.' }
    $consoleFixtureExecutableHash = (Get-FileHash -LiteralPath $consoleFixtureExecutable).Hash.ToLowerInvariant()
    $policyBefore = @(Get-ExecutionPolicy -List | Where-Object Scope -ne Process | ForEach-Object { $_.Scope.ToString() + '=' + $_.ExecutionPolicy.ToString() }) -join '|'
    function New-ExitDirectory([string]$Relative) {
        $path = [IO.Path]::GetFullPath((Join-Path $ownedRoot $Relative))
        if (-not $path.StartsWith($ownedRoot + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Exit fixture escaped ownership.' }
        Assert-ExitAncestors $path
        $null = [IO.Directory]::CreateDirectory($path)
        return $path
    }
    function Write-ExitBytes([string]$Path, [byte[]]$Bytes) {
        if (-not ([IO.Path]::GetFullPath($Path)).StartsWith($ownedRoot + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Exit bytes escaped ownership.' }
        Assert-ExitAncestors ([IO.Path]::GetDirectoryName($Path))
        $stream = [IO.FileStream]::new($Path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        try { $stream.Write($Bytes, 0, $Bytes.Length) } finally { $stream.Dispose() }
        [IO.File]::SetCreationTimeUtc($Path, $fixed); [IO.File]::SetLastWriteTimeUtc($Path, $fixed)
    }
    function New-ExitVideo([string]$Path, [int]$Length = 4096) {
        # Opaque synthetic video bytes exercise the application's byte-copy contract;
        # they are deliberately not advertised as playable/codec coverage.
        $bytes = New-Object byte[] $Length
        for ($i = 0; $i -lt $bytes.Length; $i++) { $bytes[$i] = [byte](($i * 29 + 17) % 256) }
        Write-ExitBytes $Path $bytes
    }
    function Get-ExitSourceState([string]$Root) {
        $rows = @(@([IO.DirectoryInfo]::new($Root)) + @(Get-ChildItem -LiteralPath $Root -Recurse -Force) | Sort-Object FullName | ForEach-Object {
            $_.Refresh(); $directory = $_ -is [IO.DirectoryInfo]
            [pscustomobject]@{ Relative = $_.FullName.Substring($Root.Length); Directory = $directory;
                Length = if ($directory) { $null } else { $_.Length };
                Hash = if ($directory) { $null } else { (Get-FileHash -LiteralPath $_.FullName).Hash };
                CreationTicks = $_.CreationTimeUtc.Ticks; ModifiedTicks = $_.LastWriteTimeUtc.Ticks; Attributes = [int]$_.Attributes }
        })
        return ConvertTo-Json -InputObject $rows -Depth 6 -Compress
    }
    function ConvertTo-ExitWindowsArgument([string]$Value) {
        $escaped = [regex]::Replace($Value, '(\\*)"', '$1$1\"')
        $escaped = [regex]::Replace($escaped, '(\\+)$', '$1$1')
        return '"' + $escaped + '"'
    }
    function New-ExitCase([string]$Label, [string]$Mode) {
        $work = New-ExitDirectory ($Label + '-' + [Guid]::NewGuid().ToString('N'))
        $source = New-ExitDirectory ($work.Substring($ownedRoot.Length + 1) + '/source')
        if ($Mode -ne 'empty') { Write-ExitBytes (Join-Path $source 'ignored.txt') ([Text.Encoding]::UTF8.GetBytes('Owned ignored sentinel.')) }
        if ($Mode -eq 'partial') {
            Write-ExitBytes (Join-Path $source '01-corrupt.png') ([byte[]]@(137,80,78,71,13,10,26,10,0,0,0,0))
            New-ExitVideo (Join-Path $source '02-kept.mp4')
        }
        if ($Mode -eq 'global') {
            New-ExitVideo (Join-Path $source '01-kept.mp4')
            New-ExitVideo (Join-Path $source '02-before-throw.mp4')
        }
        if ($Mode -eq 'cancel') {
            New-ExitVideo (Join-Path $source '01-kept.mp4')
            New-ExitVideo (Join-Path $source '02-current.mp4') (1024 * 1024)
            New-ExitVideo (Join-Path $source '03-unstarted.mp4')
        }
        foreach ($directory in @(@(Get-ChildItem -LiteralPath $source -Recurse -Directory) | Sort-Object FullName -Descending) + @([IO.DirectoryInfo]::new($source))) {
            [IO.Directory]::SetCreationTimeUtc($directory.FullName, $fixed); [IO.Directory]::SetLastWriteTimeUtc($directory.FullName, $fixed)
        }
        $parent = Join-Path $work 'output'
        $statePath = Join-Path $work 'child-state.json'
        $driver = Join-Path $work 'WinImgNormalizer.ps1'
        $selectedMagick = if ($Mode -eq 'setup') { Join-Path $work 'missing-magick.exe' } else { $magick }
        $driverText = @'
$ErrorActionPreference = 'Stop'
function Decode-ExitFixture([string]$Value) { [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Value)) }
$app = Decode-ExitFixture '__APP__'
$parent = Decode-ExitFixture '__PARENT__'
$magick = Decode-ExitFixture '__MAGICK__'
$statePath = Decode-ExitFixture '__STATE__'
$mode = '__MODE__'
function Get-ExitChildPolicy { return (@(Get-ExecutionPolicy -List | Where-Object Scope -ne Process | ForEach-Object { $_.Scope.ToString() + '=' + $_.ExecutionPolicy.ToString() }) -join '|') }
function Get-ExitChildProfiles {
    return (@(@($PROFILE.AllUsersAllHosts, $PROFILE.AllUsersCurrentHost, $PROFILE.CurrentUserAllHosts, $PROFILE.CurrentUserCurrentHost) | Select-Object -Unique | ForEach-Object {
        $exists = [IO.File]::Exists($_)
        [pscustomobject]@{ Path = $_; Exists = $exists; Hash = if ($exists) { (Get-FileHash -LiteralPath $_).Hash } else { $null } }
    }) | ConvertTo-Json -Depth 4 -Compress)
}
$policyBefore = Get-ExitChildPolicy
$profilesBefore = Get-ExitChildProfiles
. $app
$script:exitReport = $null
$script:exitCopied = 0
$script:exitStage = $null
$script:exitStoppedPropagated = $false
$script:exitStoppedState = $null
$script:exitStoppedReasonType = $null
$invoke = @{ Arguments = $args; OutputParent = $parent; MagickPath = $magick;
    NativeTemporaryRoot = [IO.Path]::GetDirectoryName($statePath);
    ReportObserver = { param($State) $script:exitReport = $State } }
if ($mode -eq 'global') {
    $invoke.RunStageObserver = { param($Stage, $State, $RelativePath, $Path)
        if ($Stage -eq 'BeforeItem' -and $RelativePath -eq '02-before-throw.mp4') {
            $script:exitStage = $Stage
            throw [IO.IOException]::new("Owned global failure`r`n" + ('X' * 4096))
        }
    }
}
if ($mode -eq 'cancel') {
    $invoke.CopyProgressObserver = { param($SourcePath, $CandidatePath, $Copied, $State)
        if ($SourcePath.EndsWith('\02-current.mp4') -and $CandidatePath.EndsWith('\video.partial')) {
            $script:exitCopied = $Copied
            $State.Request()
        }
    }
}
if ($mode -eq 'stopped') {
    # This fixture deliberately replaces the core, solely to inspect the outer
    # Command exception boundary. A stopped engine pipeline cannot finish this
    # driver's observation in the same pipeline, so inspect a nested runspace.
    # This is not a simulated console event or application cancellation result.
    $nested = [Management.Automation.PowerShell]::Create()
    try {
        $nestedText = '. ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String(''__APP__''))); function Invoke-WinImgNormalizer { throw [Management.Automation.PipelineStoppedException]::new(''Owned stopped-pipeline control'') }; Invoke-WinImgNormalizerCommand -Arguments @(''owned-controlled-argument'') -OutputParent ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String(''__PARENT__'')))'
        $null = $nested.AddScript($nestedText)
        $nestedOutput = @($nested.Invoke())
        $script:exitStoppedState = $nested.InvocationStateInfo.State.ToString()
        $script:exitStoppedReasonType = $nested.InvocationStateInfo.Reason.GetType().FullName
        $script:exitStoppedPropagated = $nestedOutput.Count -eq 0 -and $script:exitStoppedState -eq 'Stopped' -and
            $nested.InvocationStateInfo.Reason -is [Management.Automation.PipelineStoppedException]
    } finally { $nested.Dispose() }
    if (-not $script:exitStoppedPropagated) { throw 'Command swallowed the stopped pipeline.' }
    $code = 0 # Driver assertion succeeded; no application cancellation code is claimed.
} else {
    if ($mode -eq 'outofsession') {
        function Invoke-WinImgNormalizer { throw [OperationCanceledException]::new('Owned cancellation exception outside any active run') }
    }
    $observed = @(Invoke-WinImgNormalizerCommand @invoke)
    if ($observed.Count -ne 1 -or $observed[0] -isnot [int]) { throw 'Command did not return exactly one Int32 exit status.' }
    $code = $observed[0]
}
$record = [ordered]@{ Code = $code; HostVersion = $PSVersionTable.PSVersion.ToString(); HostEdition = $PSVersionTable.PSEdition;
    HostExecutable = (Get-Process -Id $PID).Path; CommandLine = [Environment]::CommandLine;
    Pid = $PID; StartTicks = (Get-Process -Id $PID).StartTime.ToUniversalTime().Ticks;
    PolicyBefore = $policyBefore; PolicyAfter = Get-ExitChildPolicy; ProfilesBefore = $profilesBefore; ProfilesAfter = Get-ExitChildProfiles;
    CopiedBeforeControlledRequest = $script:exitCopied; ControlledStage = $script:exitStage;
    StoppedPropagated = $script:exitStoppedPropagated; ControlledCoreFixture = $mode -in @('stopped', 'outofsession');
    StoppedState = $script:exitStoppedState; StoppedReasonType = $script:exitStoppedReasonType;
    Report = if ($script:exitReport) { [ordered]@{ RunState = $script:exitReport.RunState; DestinationRoot = $script:exitReport.DestinationRoot;
        ReportComplete = $script:exitReport.ReportComplete; CsvPath = $script:exitReport.CsvPath; Rows = $script:exitReport.Rows.ToArray(); Summary = $script:exitReport.Summary } } else { $null } }
[IO.File]::WriteAllText($statePath, (ConvertTo-Json -InputObject $record -Depth 12), [Text.UTF8Encoding]::new($false))
exit $code
'@
        foreach ($replacement in @{ APP = $application; PARENT = $parent; MAGICK = $selectedMagick; STATE = $statePath }.GetEnumerator()) {
            $driverText = $driverText.Replace('__' + $replacement.Key + '__', [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($replacement.Value)))
        }
        $driverText = $driverText.Replace('__MODE__', $Mode)
        [IO.File]::WriteAllText($driver, $driverText, [Text.UTF8Encoding]::new($false))
        $arguments = @($source)
        if ($Mode -eq 'usage') { $arguments = @() }
        if ($Mode -eq 'cap') { $arguments += '1.5' }
        return [pscustomobject]@{ Work = $work; Source = $source; OutputParent = $parent; Driver = $driver; Arguments = $arguments;
            StatePath = $statePath; Mode = $Mode; Before = Get-ExitSourceState $source }
    }
    function Invoke-ExitChild([object]$Case, [switch]$Batch) {
        $temporary = Join-Path $Case.Work 'temporary'
        $null = [IO.Directory]::CreateDirectory($temporary)
        $start = New-Object Diagnostics.ProcessStartInfo
        if ($Batch) {
            $batchPath = Join-Path $Case.Work 'WinImgNormalizer.bat'
            [IO.File]::Copy($launcher, $batchPath, $false)
            (Get-FileHash -LiteralPath $batchPath).Hash | Should -Be (Get-FileHash -LiteralPath $launcher).Hash
            $start.FileName = Join-Path $env:SystemRoot 'System32/cmd.exe'
            $start.Arguments = '/d /s /c ""' + $batchPath + '" "' + $Case.Source + '""'
            $start.RedirectStandardInput = $true
            # The fixture owns a hidden private console and gives the unchanged
            # CMD/BAT actual console stdin. Its main process waits only for CMD;
            # the background key reader cannot manufacture a blocked pause.
            $start.EnvironmentVariables['WINIMG_TEST_PRIVATE_CONSOLE_CMD'] = $start.FileName
            $start.EnvironmentVariables['WINIMG_TEST_PRIVATE_CONSOLE_ARGUMENTS'] = $start.Arguments
            $start.EnvironmentVariables['WINIMG_TEST_PRIVATE_CONSOLE_RECORD'] = Join-Path $Case.Work 'private-console.json'
            $start.FileName = $consoleFixtureExecutable
            $start.Arguments = ''
        } else {
            $start.FileName = $hostExecutable
            $tokens = @('-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $Case.Driver) + $Case.Arguments
            $start.Arguments = ($tokens | ForEach-Object { ConvertTo-ExitWindowsArgument ([string]$_) }) -join ' '
        }
        $start.WorkingDirectory = $Case.Work; $start.UseShellExecute = $false; $start.CreateNoWindow = $true
        $start.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
        $start.RedirectStandardOutput = $true; $start.RedirectStandardError = $true
        foreach ($key in @($start.EnvironmentVariables.Keys)) {
            if ([string]::Equals([string]$key, 'PSModulePath', [StringComparison]::OrdinalIgnoreCase)) { $start.EnvironmentVariables.Remove([string]$key) }
        }
        # The production BAT's chosen powershell command must resolve to actual
        # Windows PowerShell 5.1, independently of the Pester host edition.
        $start.EnvironmentVariables['PATH'] = (Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0') + ';' + $env:PATH
        $start.EnvironmentVariables['TEMP'] = $temporary; $start.EnvironmentVariables['TMP'] = $temporary
        $start.EnvironmentVariables['MAGICK_TEMPORARY_PATH'] = $temporary
        $process = New-Object Diagnostics.Process
        $process.StartInfo = $start
        $pauseWasPending = $false
        $driverExitedBeforePauseRelease = $false
        try {
            if (-not $process.Start()) { throw 'Owned exit-contract child failed to start.' }
            $stdout = $process.StandardOutput.ReadToEndAsync(); $stderr = $process.StandardError.ReadToEndAsync()
            if ($Batch) {
                $wait = [Diagnostics.Stopwatch]::StartNew()
                while (-not [IO.File]::Exists($Case.StatePath) -and -not $process.WaitForExit(20) -and $wait.ElapsedMilliseconds -lt 35000) { }
                if (-not [IO.File]::Exists($Case.StatePath)) { throw 'Safe application driver did not record its result before BAT pause.' }
                $childIdentity = Get-Content -LiteralPath $Case.StatePath -Raw | ConvertFrom-Json
                $child = $null
                try {
                    try { $child = [Diagnostics.Process]::GetProcessById($childIdentity.Pid) } catch [ArgumentException] { $driverExitedBeforePauseRelease = $true }
                    if ($child) {
                        if ($child.StartTime.ToUniversalTime().Ticks -ne $childIdentity.StartTicks) { throw 'Safe driver process identity changed before pause inspection.' }
                        if (-not $child.WaitForExit(5000)) { throw 'Safe driver did not exit after recording its code.' }
                        $driverExitedBeforePauseRelease = $true
                    }
                } finally { if ($child) { $child.Dispose() } }
                $pauseWasPending = -not $process.WaitForExit(300)
                $pauseWasPending | Should -BeTrue
                $process.StandardInput.WriteLine('x'); $process.StandardInput.Flush(); $process.StandardInput.Close()
            }
            if (-not $process.WaitForExit(45000)) {
                if ($Batch) {
                    # Terminate this exact fixture handle. Its private job owns
                    # CMD and descendants and closes when the fixture exits.
                    $process.Kill(); $null = $process.WaitForExit(5000)
                } else {
                    & (Join-Path $env:SystemRoot 'System32/taskkill.exe') /PID $process.Id /T /F 2>&1 | Out-Null
                    if (-not $process.WaitForExit(5000)) { $process.Kill() }
                }
                throw 'Owned exit-contract child exceeded its 45-second bound.'
            }
            if (-not $stdout.Wait(5000) -or -not $stderr.Wait(5000)) { throw 'Exit-contract child streams did not finish.' }
            $out = $stdout.GetAwaiter().GetResult(); $err = $stderr.GetAwaiter().GetResult()
            [IO.File]::WriteAllText((Join-Path $Case.Work 'stdout.log'), $out, [Text.UTF8Encoding]::new($false))
            [IO.File]::WriteAllText((Join-Path $Case.Work 'stderr.log'), $err, [Text.UTF8Encoding]::new($false))
            $record = [pscustomobject]@{ Batch = [bool]$Batch; Mode = $Case.Mode; ExitCode = $process.ExitCode; StdOut = $out; StdErr = $err;
                ChildState = if ([IO.File]::Exists($Case.StatePath)) { Get-Content -LiteralPath $Case.StatePath -Raw | ConvertFrom-Json } else { $null };
                Executable = $start.FileName; Arguments = $start.Arguments; Work = $Case.Work;
                ConsoleFixtureRecord = if ($Batch -and [IO.File]::Exists((Join-Path $Case.Work 'private-console.json'))) { Get-Content -LiteralPath (Join-Path $Case.Work 'private-console.json') -Raw | ConvertFrom-Json } else { $null };
                PauseWasPending = $pauseWasPending; DriverExitedBeforePauseRelease = $driverExitedBeforePauseRelease;
                DriverSha256 = (Get-FileHash -LiteralPath $Case.Driver).Hash; SourceBefore = $Case.Before; SourceAfter = Get-ExitSourceState $Case.Source }
            $observations.Add($record)
            return $record
        } finally {
            if ($process.Id -and -not $process.HasExited) {
                if ($Batch) { $process.Kill(); $null = $process.WaitForExit(5000) }
                else {
                    & (Join-Path $env:SystemRoot 'System32/taskkill.exe') /PID $process.Id /T /F 2>&1 | Out-Null
                    $null = $process.WaitForExit(5000)
                }
            }
            $process.Dispose()
        }
    }
    function Assert-ExitResult([object]$Case, [object]$Result, [int]$ExpectedCode) {
        $Result.ExitCode | Should -Be $ExpectedCode
        $Result.StdErr | Should -BeNullOrEmpty
        $Result.ChildState | Should -Not -BeNullOrEmpty
        $Result.ChildState.Code | Should -Be $ExpectedCode
        $Result.ChildState.CommandLine | Should -Match '(?i)-NoProfile\b'
        $Result.ChildState.PolicyAfter | Should -BeExactly $Result.ChildState.PolicyBefore
        $Result.ChildState.ProfilesAfter | Should -BeExactly $Result.ChildState.ProfilesBefore
        $Result.SourceAfter | Should -BeExactly $Result.SourceBefore
        if ($Result.Batch) {
            $Result.ChildState.HostEdition | Should -Be 'Desktop'
            $Result.ChildState.HostVersion | Should -Match '^5\.1\.'
            $Result.PauseWasPending | Should -BeTrue
            $Result.DriverExitedBeforePauseRelease | Should -BeTrue
            $Result.ConsoleFixtureRecord | Should -Not -BeNullOrEmpty
            $Result.ConsoleFixtureRecord.PrivateConsoleAllocated | Should -BeTrue
            $Result.ConsoleFixtureRecord.PrivateJobAssignedBeforeCmd | Should -BeTrue
            $Result.ConsoleFixtureRecord.HarnessError | Should -BeNullOrEmpty
            $Result.ConsoleFixtureRecord.ExitCode | Should -Be $ExpectedCode
            $Result.ConsoleFixtureRecord.ControllerExitCode | Should -Be $ExpectedCode
            $Result.ConsoleFixtureRecord.KeyRecordsWritten | Should -Be 2
            @($Result.ConsoleFixtureRecord.MembersBeforeKey).Count | Should -Be 2
            $Result.ConsoleFixtureRecord.MembersBeforeKey | Should -Contain $Result.ConsoleFixtureRecord.CmdId
            $Result.ConsoleFixtureRecord.MembersBeforeKey | Should -Contain $Result.ConsoleFixtureRecord.ControllerId
        } else {
            $Result.ChildState.HostEdition | Should -Be $PSVersionTable.PSEdition
            $Result.ChildState.HostVersion | Should -Be $PSVersionTable.PSVersion.ToString()
        }
        if ($Case.Mode -in @('usage', 'setup', 'cap')) {
            Test-Path -LiteralPath $Case.OutputParent | Should -BeFalse
            if ($Case.Mode -eq 'usage') { $Result.StdOut | Should -Match 'Usage:' }
            else { $Result.StdOut | Should -Match 'Setup error:' }
            return
        }
        $report = $Result.ChildState.Report
        $report | Should -Not -BeNullOrEmpty
        $runs = @(Get-ChildItem -LiteralPath $Case.OutputParent -Directory)
        $runs.Count | Should -Be 1
        $report.DestinationRoot | Should -Be $runs[0].FullName
        $files = @(Get-ChildItem -LiteralPath $report.DestinationRoot -Recurse -File)
        @($files | Where-Object { $_.Extension -eq '.jpeg' -or $_.Name -eq 'video.partial' }).Count | Should -Be 0
        if ($Case.Mode -in @('ignored', 'empty')) {
            $report.RunState | Should -Be 'Completed'; $report.ReportComplete | Should -BeTrue
            @($files | Where-Object Extension -eq '.mp4').Count | Should -Be 0
            $report.Summary.Counts.Converted | Should -Be 0; $report.Summary.Counts.CopiedVideo | Should -Be 0
            $report.Summary.Counts.Ignored | Should -Be $(if ($Case.Mode -eq 'ignored') { 1 } else { 0 })
            $Result.StdOut | Should -Match 'No images or videos found\.'
        }
        if ($Case.Mode -eq 'global') {
            $Result.ChildState.ControlledStage | Should -Be 'BeforeItem'
            $line = @($Result.StdOut -split '\r?\n' | Where-Object { $_.StartsWith('Run error:') })
            $line.Count | Should -Be 1; $line[0].Length | Should -BeLessOrEqual 1035
            $line[0] | Should -Match '^Run error: Owned global failure\\u000D\\u000A'
            $Result.StdOut | Should -Not -Match 'IOException|FullyQualifiedErrorId|CategoryInfo'
            $report.RunState | Should -Be 'Partial'; $report.ReportComplete | Should -BeTrue
            $report.Summary.Counts.CopiedVideo | Should -Be 1; $report.Summary.Counts.NotStarted | Should -Be 1
            $kept = Join-Path $report.DestinationRoot '01-kept.mp4'
            (Get-FileHash -LiteralPath $kept).Hash | Should -Be (Get-FileHash -LiteralPath (Join-Path $Case.Source '01-kept.mp4')).Hash
            Test-Path -LiteralPath (Join-Path $report.DestinationRoot '02-before-throw.mp4') | Should -BeFalse
            @($files | Where-Object Extension -eq '.mp4').Count | Should -Be 1
        }
        if ($Case.Mode -eq 'partial') {
            $report.RunState | Should -Be 'Partial'; $report.ReportComplete | Should -BeTrue
            $report.Summary.Counts.Error | Should -Be 1; $report.Summary.Counts.CopiedVideo | Should -Be 1
            @($report.Rows | Where-Object { $_.SourceRelativePath -eq '01-corrupt.png' -and $_.Status -eq 'Error' }).Count | Should -Be 1
            $kept = Join-Path $report.DestinationRoot '02-kept.mp4'
            (Get-FileHash -LiteralPath $kept).Hash | Should -Be (Get-FileHash -LiteralPath (Join-Path $Case.Source '02-kept.mp4')).Hash
            @($files | Where-Object Extension -eq '.mp4').Count | Should -Be 1
        }
        if ($Case.Mode -eq 'cancel') {
            $Result.ChildState.CopiedBeforeControlledRequest | Should -BeGreaterThan 0
            $Result.ChildState.CopiedBeforeControlledRequest | Should -BeLessThan (1024 * 1024)
            $report.RunState | Should -Be 'Interrupted'; $report.ReportComplete | Should -BeTrue
            $report.Summary.Counts.CopiedVideo | Should -Be 1; $report.Summary.Counts.Cancelled | Should -Be 1
            $report.Summary.Counts.NotStarted | Should -Be 1
            $kept = Join-Path $report.DestinationRoot '01-kept.mp4'
            (Get-FileHash -LiteralPath $kept).Hash | Should -Be (Get-FileHash -LiteralPath (Join-Path $Case.Source '01-kept.mp4')).Hash
            Test-Path -LiteralPath (Join-Path $report.DestinationRoot '02-current.mp4') | Should -BeFalse
            Test-Path -LiteralPath (Join-Path $report.DestinationRoot '03-unstarted.mp4') | Should -BeFalse
            @($files | Where-Object Extension -eq '.mp4').Count | Should -Be 1
            $Result.StdOut | Should -Match 'INTERRUPTED Mode=Cooperative ExitCode=130'
        }
    }
}

Describe 'M3-T05 actual script command exit contract (T061/T063)' {
    It 'returns native <Code> for <Label> through the unchanged Command on the actual host' -ForEach @(
        @{ Label = 'ignored input'; Mode = 'ignored'; Code = 0 },
        @{ Label = 'empty input'; Mode = 'empty'; Code = 0 },
        @{ Label = 'usage failure'; Mode = 'usage'; Code = 1 },
        @{ Label = 'missing dependency setup failure'; Mode = 'setup'; Code = 1 },
        @{ Label = 'invalid positional cap setup failure'; Mode = 'cap'; Code = 1 },
        @{ Label = 'controlled unexpected run-level failure'; Mode = 'global'; Code = 1 },
        @{ Label = 'actual corrupt image with retained video'; Mode = 'partial'; Code = 2 },
        @{ Label = 'controlled cancellation during actual copy'; Mode = 'cancel'; Code = 130 }
    ) {
        $case = New-ExitCase ('script-' + $Mode) $Mode
        $result = Invoke-ExitChild $case
        Assert-ExitResult $case $result $Code
    }
}

Describe 'M3-T05 complete application flow through the exact production BAT (T061/T063)' {
    It 'preserves application <Code> after the actual BAT echo and pause for <Label>' -ForEach @(
        @{ Label = 'completed ignored input'; Mode = 'ignored'; Code = 0; Message = 'Completed successfully' },
        @{ Label = 'dependency setup failure'; Mode = 'setup'; Code = 1; Message = 'Processing did not complete reliably \(exit code 1\)' },
        @{ Label = 'controlled global failure'; Mode = 'global'; Code = 1; Message = 'Processing did not complete reliably \(exit code 1\)' },
        @{ Label = 'actual file failure with kept output'; Mode = 'partial'; Code = 2; Message = 'Completed with warnings|Completed with.*partial' },
        @{ Label = 'controlled cooperative interruption'; Mode = 'cancel'; Code = 130; Message = 'cancelled|canceled|interrupted' }
    ) {
        $case = New-ExitCase ('batch-' + $Mode) $Mode
        $result = Invoke-ExitChild $case -Batch
        Assert-ExitResult $case $result $Code
        $result.StdOut | Should -Match ('(?i)' + $Message)
        $result.StdOut | Should -Match '(?i)Press any key'
        $result.StdOut | Should -Not -Match '(?m)^Done\.'
    }
}

Describe 'M3-T05 outer Command exception boundary controls (T061)' {
    It 'propagates a controlled stopped pipeline without inventing application completion or cancellation' {
        $case = New-ExitCase 'script-stopped-control' 'stopped'
        $result = Invoke-ExitChild $case
        $result.ExitCode | Should -Be 0 # The driver caught the propagated exception.
        $result.StdErr | Should -BeNullOrEmpty
        $result.ChildState.ControlledCoreFixture | Should -BeTrue
        $result.ChildState.StoppedPropagated | Should -BeTrue
        $result.ChildState.StoppedState | Should -Be 'Stopped'
        $result.ChildState.StoppedReasonType | Should -Be 'System.Management.Automation.PipelineStoppedException'
        $result.StdOut | Should -Not -Match 'Run error:|INTERRUPTED|ExitCode=130'
        Test-Path -LiteralPath $case.OutputParent | Should -BeFalse
        $result.SourceAfter | Should -BeExactly $result.SourceBefore
    }

    It 'returns 1 for a controlled cancellation exception outside any active core session' {
        $case = New-ExitCase 'script-outofsession-control' 'outofsession'
        $result = Invoke-ExitChild $case
        $result.ExitCode | Should -Be 1
        $result.StdErr | Should -BeNullOrEmpty
        $result.ChildState.ControlledCoreFixture | Should -BeTrue
        $result.ChildState.Code | Should -Be 1
        $result.StdOut | Should -Match 'Run error: Owned cancellation exception outside any active run'
        $result.StdOut | Should -Not -Match 'INTERRUPTED|ExitCode=130'
        Test-Path -LiteralPath $case.OutputParent | Should -BeFalse
        $result.SourceAfter | Should -BeExactly $result.SourceBefore
    }
}

AfterAll {
    $bindingsAfter = @(Get-ExitBindings)
    $policyAfter = @(Get-ExecutionPolicy -List | Where-Object Scope -ne Process | ForEach-Object { $_.Scope.ToString() + '=' + $_.ExecutionPolicy.ToString() }) -join '|'
    $record = [ordered]@{ SchemaVersion = 1; Task = 'M3-T05'; ObservedAtUtc = [DateTime]::UtcNow.ToString('o');
        HostExecutable = $hostExecutable; HostVersion = $PSVersionTable.PSVersion.ToString(); HostEdition = $PSVersionTable.PSEdition;
        BindingsBefore = $bindingsBefore; BindingsAfter = $bindingsAfter; PolicyBefore = $policyBefore; PolicyAfter = $policyAfter;
        ConsoleFixtureExecutableSha256 = $consoleFixtureExecutableHash;
        Observations = @($observations.ToArray());
        Limitations = @('Driver calls unchanged application Command with an owned OutputParent; production entry with real Pictures is not executed.',
            'Cancellation and unexpected run-level throw are controlled callbacks; actual native host exit, video copy, corrupt-image failure, retained outputs and cleanup are asserted.',
            'BAT is copied byte-for-byte beside the safe application driver and invokes its existing actual PS5.1 host. A shared test-only hidden private console supplies two key records only to its owned CMD after the application host exits and the real pause is observed. No physical desktop drag event is claimed.',
            'Profile inventory and persistent policy equality are observations. No existing profile or global policy is changed.') }
    [IO.File]::WriteAllText((Join-Path $ownedRoot 'observations.json'), (ConvertTo-Json -InputObject $record -Depth 18), [Text.UTF8Encoding]::new($false))
    Write-Host ('Owned M3-T05 exit observations: ' + (Join-Path $ownedRoot 'observations.json'))
    (ConvertTo-Json -InputObject $bindingsAfter -Depth 4 -Compress) | Should -BeExactly (ConvertTo-Json -InputObject $bindingsBefore -Depth 4 -Compress)
    (Get-FileHash -LiteralPath $consoleFixtureExecutable).Hash.ToLowerInvariant() | Should -BeExactly $consoleFixtureExecutableHash
    $policyAfter | Should -BeExactly $policyBefore
}
