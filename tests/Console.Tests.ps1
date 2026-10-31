BeforeAll {
    $repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    . (Join-Path $repository 'WinImgNormalizer.ps1')
    $fullOutput = 'C:\Synthetic media\Tour photos\output with a deliberately long and complete path'
    $fullReport = $fullOutput + '\reports\WinImgNormalizer.csv'
    $resultState = [pscustomobject]@{
        RunState='Completed'; DestinationRoot=$fullOutput; CsvPath=$fullReport
        Summary=[pscustomobject]@{ Counts=[pscustomobject]@{ Discovered=12; Converted=11; CopiedVideo=1; Error=0 } }
    }
}

Describe 'Hotfix field console deterministic presentation' {
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
