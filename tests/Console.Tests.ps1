BeforeAll {
    $repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    . (Join-Path $repository 'WinImgNormalizer.ps1')
    $fullOutput = 'C:\Synthetic media\Tour photos\output with a deliberately long and complete path'
    $fullReport = $fullOutput + '\reports\WinImgNormalizer.csv'
    $resultState = [pscustomobject]@{
        RunState='Completed'; DestinationRoot=$fullOutput; CsvPath=$fullReport
        Summary=[pscustomobject]@{ Counts=[pscustomobject]@{ Discovered=12; Converted=11; CopiedVideo=1; Error=0 } }
    }
    $consoleScratch = Join-Path $repository ('.scratch/console-bytes-' + [guid]::NewGuid().ToString('N'))
    Assert-WinImgNoReparseAncestors ([IO.Path]::GetDirectoryName($consoleScratch))
    $null = [IO.Directory]::CreateDirectory($consoleScratch)
    $consoleChildHost = (Get-Process -Id $PID).Path
    function Invoke-ConsoleByteCapture([string]$Mode) {
        $root = Join-Path $consoleScratch ($Mode + '-' + [guid]::NewGuid().ToString('N'))
        $null = [IO.Directory]::CreateDirectory($root)
        $driver = Join-Path $root 'unicode-driver.ps1'
        $record = Join-Path $root 'encoding.json'
        $unicodePath = Join-Path $root ('synthetic ' + [char]0x00e5 + [char]0x96ea + '\complete output.jpeg')
        $encodedApplication = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes((Join-Path $repository 'WinImgNormalizer.ps1')))
        $encodedRecord = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($record))
        $driverText = @'
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::GetEncoding(437)
$before = [Console]::OutputEncoding.CodePage
$errorBefore = [Console]::Error.Encoding.CodePage
. ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('__APPLICATION__')))
$script:unicodePath = '__PATH_LITERAL__'
$script:mode = '__MODE__'
$script:during = $null
$script:errorDuring = $null
function Write-UnicodeFixture {
    $script:during = [Console]::OutputEncoding.CodePage
    $script:errorDuring = [Console]::Error.Encoding.CodePage
    Write-WinImgConsoleLine ('Output: ' + $script:unicodePath)
    [Console]::Error.WriteLine('Diagnostic: ' + $script:unicodePath)
}
function Get-WinImgUsage { Write-UnicodeFixture; return ('Usage: WinImgNormalizer.ps1 <sourceFolder> [maxBytes] ' + $script:unicodePath) }
function Invoke-WinImgNormalizer {
    Write-UnicodeFixture
    switch ($script:mode) {
        'run-error' { throw [IO.IOException]::new('Synthetic failure: ' + $script:unicodePath) }
        'setup' { Write-WinImgConsoleLine ('Setup error: ' + $script:unicodePath) -Kind Error; return 1 }
        'partial' { return 2 }
        'cancel' { return 130 }
        default { return 0 }
    }
}
$arguments = @()
if ($script:mode -ne 'usage') { $arguments = @($script:unicodePath) }
$code = Invoke-WinImgNormalizerCommand -Arguments $arguments -CaptureConsoleControl $false
$result = [ordered]@{ Before=$before; During=$script:during; After=[Console]::OutputEncoding.CodePage;
    ErrorBefore=$errorBefore; ErrorDuring=$script:errorDuring; ErrorAfter=[Console]::Error.Encoding.CodePage;
    Edition=$PSVersionTable.PSEdition; Version=$PSVersionTable.PSVersion.ToString(); Code=$code;
    OutputRedirected=[Console]::IsOutputRedirected; ErrorRedirected=[Console]::IsErrorRedirected }
$recordPath = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('__RECORD__'))
[IO.File]::WriteAllText($recordPath, ($result | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
exit $code
'@
        $driverText = $driverText.Replace('__APPLICATION__',$encodedApplication).Replace('__PATH_LITERAL__',$unicodePath.Replace("'","''")).Replace('__RECORD__',$encodedRecord).Replace('__MODE__',$Mode)
        # Windows PowerShell must decode the driver itself before the app can
        # select its pipe encoding. Keep a real Unicode literal in BOM-marked UTF-8.
        [IO.File]::WriteAllText($driver, $driverText, [Text.UTF8Encoding]::new($true))
        $start = [Diagnostics.ProcessStartInfo]::new()
        $start.FileName = $consoleChildHost
        $start.Arguments = '-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + $driver + '"'
        $start.UseShellExecute=$false; $start.CreateNoWindow=$true
        $start.RedirectStandardOutput=$true; $start.RedirectStandardError=$true
        $process = [Diagnostics.Process]::new(); $process.StartInfo=$start
        $stdout = [IO.MemoryStream]::new(); $stderr = [IO.MemoryStream]::new()
        try {
            if (-not $process.Start()) { throw 'Native Unicode fixture did not start.' }
            $outTask = $process.StandardOutput.BaseStream.CopyToAsync($stdout)
            $errTask = $process.StandardError.BaseStream.CopyToAsync($stderr)
            if (-not $process.WaitForExit(15000)) { $process.Kill(); throw 'Native Unicode fixture exceeded its deadline.' }
            if (-not $outTask.Wait(5000) -or -not $errTask.Wait(5000)) { throw 'Native Unicode fixture streams did not complete.' }
            $outBytes=$stdout.ToArray(); $errBytes=$stderr.ToArray()
            [IO.File]::WriteAllBytes((Join-Path $root 'stdout.bin'), $outBytes)
            [IO.File]::WriteAllBytes((Join-Path $root 'stderr.bin'), $errBytes)
            $utf8 = [Text.UTF8Encoding]::new($false,$true)
            if (-not [IO.File]::Exists($record)) { throw ('Native Unicode fixture did not record its result: ' + $utf8.GetString($errBytes)) }
            return [pscustomobject]@{ ExitCode=$process.ExitCode; StdOut=$utf8.GetString($outBytes); StdErr=$utf8.GetString($errBytes);
                UnicodePath=$unicodePath; State=(Get-Content -LiteralPath $record -Raw -Encoding UTF8 | ConvertFrom-Json) }
        } finally {
            if ($process.Id -and -not $process.HasExited) { $process.Kill(); $null=$process.WaitForExit(5000) }
            $process.Dispose(); $stdout.Dispose(); $stderr.Dispose()
        }
    }
}

Describe 'Hotfix field console deterministic presentation' {
    It 'shows the engine section and actual app directory before a private resolver failure without inventing an engine' {
        $source = Join-Path $consoleScratch ('resolver-source-' + [guid]::NewGuid().ToString('N'))
        [IO.Directory]::CreateDirectory($source) | Out-Null
        [IO.File]::WriteAllBytes((Join-Path $source 'synthetic.png'), [byte[]](1,2,3))
        $output = Join-Path $consoleScratch ('unused-output-' + [guid]::NewGuid().ToString('N'))
        # A previous run's selection must never leak into this failure report.
        $script:WinImgEngineDiagnostic = [pscustomobject]@{Path='C:\stale\magick.exe';Source='external';DisplayVersion='7.1.2-5 Q16-HDRI x64'}
        Mock Get-WinImgPrivateRuntime { throw 'Incomplete WinImgNormalizer app copy; WinImgNormalizer-1.0.2-portable.zip is the complete portable asset.' }
        Mock Invoke-WinImgNativeProcess { throw 'No runtime bytes must mean no native execution.' }
        $observed = @(& Invoke-WinImgNormalizerCommand -Arguments @($source) -OutputParent $output -CaptureConsoleControl $false 6>&1 2>&1)
        @($observed | Where-Object { $_ -is [int] }) | Should -Be @(1)
        $text = (@($observed | Where-Object { $_ -isnot [int] }) -join "`n")
        $text.IndexOf('[02 / ENGINE]') | Should -BeGreaterThan $text.IndexOf('[01 / INPUT]')
        $text.IndexOf('[04 / RESULT]') | Should -BeGreaterThan $text.IndexOf('[02 / ENGINE]')
        $text | Should -Match ('Application:\s+' + [regex]::Escape($repository))
        $text | Should -Match 'Selected engine: none'
        $text | Should -Match 'Incomplete WinImgNormalizer app copy'
        $text | Should -Not -Match 'external|stale|7\.1\.2|Not needed'
        Test-Path -LiteralPath $output | Should -BeFalse
        Should -Invoke Invoke-WinImgNativeProcess -Times 0 -Exactly
    }

    It 'shows only the selected path with an unavailable version when policy isolation fails before a query' {
        $source = Join-Path $consoleScratch ('policy-source-' + [guid]::NewGuid().ToString('N'))
        [IO.Directory]::CreateDirectory($source) | Out-Null
        [IO.File]::WriteAllBytes((Join-Path $source 'synthetic.png'), [byte[]](1,2,3))
        $output = Join-Path $consoleScratch ('unused-policy-output-' + [guid]::NewGuid().ToString('N'))
        Mock Get-WinImgPrivateRuntime { return [pscustomobject]@{Path='C:\synthetic-private\magick.exe';Source='cache';Version='7.1.2-35';Environment=@{};PolicyFiles=@();PolicyDirectories=@()} }
        Mock Initialize-WinImgBundledEnvironment { throw 'Explicit policy denied before any engine query.' }
        Mock Invoke-WinImgNativeProcess { throw 'Policy denial must not execute an engine.' }
        $observed = @(& Invoke-WinImgNormalizerCommand -Arguments @($source) -OutputParent $output -CaptureConsoleControl $false 6>&1 2>&1)
        @($observed | Where-Object { $_ -is [int] }) | Should -Be @(1)
        $text = (@($observed | Where-Object { $_ -isnot [int] }) -join "`n")
        $text.IndexOf('[04 / RESULT]') | Should -BeGreaterThan $text.IndexOf('[02 / ENGINE]')
        $text | Should -Match 'Selected engine: cache \| version unavailable'
        $text | Should -Match ([regex]::Escape('C:\synthetic-private\magick.exe'))
        $text | Should -Match 'Explicit policy denied before any engine query'
        $text | Should -Not -Match 'external|7\.1\.2-35|Selected engine: none'
        Test-Path -LiteralPath $output | Should -BeFalse
        Should -Invoke Invoke-WinImgNativeProcess -Times 0 -Exactly
    }

    It 'captures complete UTF-8 native stdout/error for <Mode> and restores the original OEM process encoding' -ForEach @(
        @{Mode='success';Code=0}, @{Mode='partial';Code=2}, @{Mode='cancel';Code=130},
        @{Mode='setup';Code=1}, @{Mode='run-error';Code=1}, @{Mode='usage';Code=1}
    ) {
        $actual = Invoke-ConsoleByteCapture $Mode
        $actual.ExitCode | Should -Be $Code
        $actual.State.Code | Should -Be $Code
        $actual.State.Edition | Should -Be $PSVersionTable.PSEdition
        $actual.State.Version | Should -Be $PSVersionTable.PSVersion.ToString()
        $actual.State.OutputRedirected | Should -BeTrue
        $actual.State.ErrorRedirected | Should -BeTrue
        $actual.State.Before | Should -Be 437
        $actual.State.During | Should -Be 65001
        $actual.State.After | Should -Be 437
        $actual.State.ErrorBefore | Should -Be 437
        $actual.State.ErrorDuring | Should -Be 65001
        $actual.State.ErrorAfter | Should -Be 437
        $actual.StdOut | Should -Match ([regex]::Escape($actual.UnicodePath))
        $actual.StdErr | Should -BeExactly ('Diagnostic: ' + $actual.UnicodePath + [Environment]::NewLine)
        $actual.StdOut | Should -Not -Match '[\x1b\x00-\x08]'
    }

    It 'leaves the encoding untouched for an interactive standard console' {
        $before = [Console]::OutputEncoding
        $scope = Initialize-WinImgConsoleEncoding -OutputRedirected $false -ErrorRedirected $false
        try {
            $scope.Changed | Should -BeFalse
            [object]::ReferenceEquals([Console]::OutputEncoding,$before) | Should -BeTrue
        } finally { Restore-WinImgConsoleEncoding $scope }
    }

    It 'renders actual state at <Width> columns and keeps final paths complete' -ForEach @(@{Width=40},@{Width=80},@{Width=120}) {
        $lines = @(
            Get-WinImgConsoleLayout Start ([pscustomobject]@{Source='Tour photos'}) $Width
            Get-WinImgConsoleLayout Engine ([pscustomobject]@{Path='runtime\imagemagick\magick.exe';Version=[version]'7.1.2.35';Source='bundled'}) $Width
            Get-WinImgConsoleLayout Progress ([pscustomobject]@{Percent=67;Status='8 / 12 : stage-left-with-a-deliberately-long-synthetic-name.png'}) $Width
            Get-WinImgConsoleLayout Result $resultState $Width
        )
        $expected = Get-Content -LiteralPath (Join-Path $PSScriptRoot ('fixtures/console/' + $Width + '.txt'))
        ($lines -join "`n") | Should -BeExactly ($expected -join "`n")
        $lines | Should -Contain $fullOutput
        $lines | Should -Contain $fullReport
        @($lines | Where-Object { $_ -match '^\[03' })[0].Length | Should -BeLessOrEqual $Width
        ($lines -join "`n") | Should -Not -Match '[\x1b\x00-\x08]'
    }

    It 'names partial completion and cancellation without false success' -ForEach @(@{State='Partial';Word='PARTIAL'},@{State='Interrupted';Word='CANCELLED'}) {
        $resultState.RunState=$State
        try {
            $lines = @(Get-WinImgConsoleLayout Result $resultState 40)
            ($lines -join "`n") | Should -Match $Word
            $lines | Should -Contain 'Completed outputs retained.'
            $lines[0] | Should -Not -Match '\bCOMPLETE\b'
        } finally { $resultState.RunState='Completed' }
    }

    It 'leads a setup failure with the specific cause and states no output' {
        $lines = @(Get-WinImgConsoleLayout Failure ([pscustomobject]@{Cause='Runtime integrity failed. Re-extract the official ZIP.'}) 40)
        $lines[1] | Should -Be 'Runtime integrity failed. Re-extract the official ZIP.'
        $lines[2] | Should -Be 'No output created.'
    }

    It 'labels no-engine work and an explicit empty no-op' {
        @(Get-WinImgConsoleLayout Engine ([pscustomobject]@{Path=''}) 40)[0] | Should -Match 'Not needed'
        $empty = [pscustomobject]@{RunState='Completed';DestinationRoot='';CsvPath='';Summary=[pscustomobject]@{Counts=[pscustomobject]@{Discovered=0;Converted=0;CopiedVideo=0;Error=0}}}
        @(Get-WinImgConsoleLayout Result $empty 40) | Should -Contain 'No eligible media found; nothing to process.'
    }

    It 'uses word labels with colour and never forces a console background' {
        Mock Get-WinImgConsoleCapabilities { [pscustomobject]@{Interactive=$true;Color=$true;Width=80} }
        Mock Write-Host {}
        Write-WinImgConsoleLine 'WARNING: optional decoder unavailable' Warning
        Write-WinImgConsoleLine 'ERROR: runtime integrity failed' Error
        Write-WinImgConsoleLine '[01 / INPUT]' Label
        Should -Invoke Write-Host -Times 1 -Exactly -ParameterFilter { $ForegroundColor -eq 'Yellow' -and -not $BackgroundColor }
        Should -Invoke Write-Host -Times 1 -Exactly -ParameterFilter { $ForegroundColor -eq 'Red' -and -not $BackgroundColor }
        Should -Invoke Write-Host -Times 1 -Exactly -ParameterFilter { $ForegroundColor -eq 'Cyan' -and -not $BackgroundColor }
    }

    It 'uses plain lines without RawUI, colour or cursor rewrites in redirected output' {
        Mock Get-WinImgConsoleCapabilities { [pscustomobject]@{Interactive=$false;Color=$false;Width=80} }
        Mock Write-Host {}
        Mock Write-Progress { throw 'A redirected run must not draw host progress.' }
        $state = New-WinImgReportState
        Write-WinImgReportProgress $state '1 / 1 : synthetic.png' 100
        Write-WinImgReportProgress $state -Completed
        Should -Invoke Write-Progress -Times 0 -Exactly
        Should -Invoke Write-Host -Times 1 -Exactly -ParameterFilter { [string]$Object -match '^\[03 / PROCESS\]' -and -not $ForegroundColor }
    }

    It 'keeps word labels legible on a <Background> background without forcing it' -ForEach @(
        @{Background='dark';Light=$false;Label='Cyan';Warning='Yellow';Failure='Red'},
        @{Background='light';Light=$true;Label='DarkCyan';Warning='DarkYellow';Failure='DarkRed'}
    ) {
        $script:consoleFixtureLight = $Light
        Mock Get-WinImgConsoleCapabilities { [pscustomobject]@{Interactive=$true;Color=$true;Width=80;Light=$script:consoleFixtureLight} }
        Mock Write-Host {}
        Write-WinImgConsoleLine '[01 / INPUT]' Label
        Write-WinImgConsoleLine 'WARNING: optional decoder unavailable' Warning
        Write-WinImgConsoleLine 'ERROR: runtime integrity failed' Error
        Should -Invoke Write-Host -Times 1 -Exactly -ParameterFilter { $Object -eq '[01 / INPUT]' -and [string]$ForegroundColor -eq $Label -and -not $BackgroundColor }
        Should -Invoke Write-Host -Times 1 -Exactly -ParameterFilter { [string]$Object -like 'WARNING:*' -and [string]$ForegroundColor -eq $Warning -and -not $BackgroundColor }
        Should -Invoke Write-Host -Times 1 -Exactly -ParameterFilter { [string]$Object -like 'ERROR:*' -and [string]$ForegroundColor -eq $Failure -and -not $BackgroundColor }
    }

    It 'keeps NO_COLOR labels in the host foreground on a <Background> background' -ForEach @(@{Background='dark';Light=$false},@{Background='light';Light=$true}) {
        $script:consoleFixtureLight = $Light
        Mock Get-WinImgConsoleCapabilities { [pscustomobject]@{Interactive=$true;Color=$false;Width=80;Light=$script:consoleFixtureLight} }
        Mock Write-Host {}
        Write-WinImgConsoleLine '[04 / RESULT]   PARTIAL - review warnings' Warning
        Should -Invoke Write-Host -Times 1 -Exactly -ParameterFilter { [string]$Object -match 'PARTIAL' -and -not $ForegroundColor -and -not $BackgroundColor }
    }

    It 'draws bounded actual progress through the interactive host branch' {
        Mock Get-WinImgConsoleCapabilities { [pscustomobject]@{Interactive=$true;Color=$true;Width=40} }
        Mock Write-Progress {}
        $state=New-WinImgReportState
        Write-WinImgReportProgress $state '1 / 2 : synthetic.png' 50
        Write-WinImgReportProgress $state '1 / 2 : synthetic.png' 50
        Write-WinImgReportProgress $state '2 / 2 : synthetic.png' 100
        Write-WinImgReportProgress $state -Completed
        Should -Invoke Write-Progress -Times 3 -Exactly
        Should -Invoke Write-Progress -Times 1 -Exactly -ParameterFilter { $Completed }
    }

    It 'honours NO_COLOR without modifying persistent settings' {
        $before=$env:NO_COLOR
        try { $env:NO_COLOR='1'; (Get-WinImgConsoleCapabilities).Color | Should -BeFalse }
        finally { $env:NO_COLOR=$before }
    }
}
