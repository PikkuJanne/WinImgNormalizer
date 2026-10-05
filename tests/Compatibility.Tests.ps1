BeforeAll {
    $repository = Split-Path -Parent $PSScriptRoot
    . (Join-Path $PSScriptRoot 'TestGate.ps1')
    $scratch = Join-Path $repository '.scratch'
    $owned = Join-Path $scratch ('M3-T06-gates-' + [guid]::NewGuid().ToString('N'))
    $current = [IO.DirectoryInfo]::new($scratch)
    while ($current) {
        if ($current.Exists -and ($current.Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Gate fixture ancestor is linked.' }
        $current = $current.Parent
    }
    [IO.Directory]::CreateDirectory($owned) | Out-Null
    function New-GateResult {
        param([hashtable]$Changes = @{})
        $values = @{ TotalCount=1; PassedCount=1; FailedCount=0; SkippedCount=0; NotRunCount=0; InconclusiveCount=0;
            FailedBlocksCount=0; FailedContainersCount=0; Result='Passed'; Tests=@([pscustomobject]@{Result='Passed';ExpandedPath=@('T064 fixture','passes')});
            Containers=@([pscustomobject]@{Item='Example.Tests.ps1';TotalCount=1;Result='Passed'}) }
        foreach ($key in $Changes.Keys) { $values[$key] = $Changes[$key] }
        return [pscustomobject]$values
    }
    function Invoke-GateChild {
        param([string]$TestFile, [string]$ResultDirectory)
        $start = New-Object Diagnostics.ProcessStartInfo
        $start.FileName = (Get-Process -Id $PID).Path
        $tokens = @('-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',(Join-Path $PSScriptRoot 'Invoke-Tests.ps1'),'-Path',$TestFile,'-ResultDirectory',$ResultDirectory)
        $start.Arguments = ($tokens | ForEach-Object {
            '"' + ([regex]::Replace([regex]::Replace([string]$_, '(\\*)"', '$1$1\"'), '(\\+)$', '$1$1')) + '"'
        }) -join ' '
        $start.UseShellExecute=$false; $start.CreateNoWindow=$true
        $start.RedirectStandardOutput=$true; $start.RedirectStandardError=$true
        foreach ($key in @($start.EnvironmentVariables.Keys)) {
            if ([string]$key -ieq 'PSModulePath') { $start.EnvironmentVariables.Remove([string]$key) }
        }
        $process = New-Object Diagnostics.Process
        $process.StartInfo=$start
        try {
            $process.Start() | Should -BeTrue
            $stdout=$process.StandardOutput.ReadToEndAsync(); $stderr=$process.StandardError.ReadToEndAsync()
            if (-not $process.WaitForExit(90000)) { $process.Kill(); throw 'Owned gate fixture exceeded deadline.' }
            if (-not $stdout.Wait(5000) -or -not $stderr.Wait(5000)) { throw 'Gate fixture capture incomplete.' }
            $summary = Get-Content -LiteralPath (Join-Path $ResultDirectory 'summary.json') -Raw -Encoding UTF8 | ConvertFrom-Json
            return [pscustomobject]@{ Exit=$process.ExitCode; Summary=$summary }
        } finally { $process.Dispose() }
    }
}

Describe 'T064 full-corpus discovery and completion gates' {
    It 'accepts a complete passing result with its required case and nonempty suite' {
        @(Get-WinImgTestGateFailures -Result (New-GateResult) -RequiredCases T064 -RequiredSuites Example.Tests.ps1).Count | Should -Be 0
    }
    It 'rejects <Label> even if other counts look successful' -ForEach @(
        @{Label='zero discovery';Change=@{TotalCount=0}}, @{Label='a skip';Change=@{SkippedCount=1}},
        @{Label='not run';Change=@{NotRunCount=1}}, @{Label='inconclusive';Change=@{InconclusiveCount=1}},
        @{Label='assertion failure';Change=@{FailedCount=1}}, @{Label='container failure';Change=@{FailedContainersCount=1}},
        @{Label='block failure';Change=@{FailedBlocksCount=1}}, @{Label='failing overall result';Change=@{Result='Failed'}}
    ) {
        @(Get-WinImgTestGateFailures -Result (New-GateResult $Change)).Count | Should -BeGreaterThan 0
    }
    It 'rejects a missing case, absent suite and an empty required suite' {
        @(Get-WinImgTestGateFailures -Result (New-GateResult) -RequiredCases T065).Count | Should -BeGreaterThan 0
        @(Get-WinImgTestGateFailures -Result (New-GateResult) -RequiredSuites Missing.Tests.ps1).Count | Should -BeGreaterThan 0
        $result = New-GateResult @{Containers=@([pscustomobject]@{Item='Example.Tests.ps1';TotalCount=0;Result='Passed'})}
        @(Get-WinImgTestGateFailures -Result $result -RequiredSuites Example.Tests.ps1).Count | Should -BeGreaterThan 0
    }
    It 'does not count skipped test names or range text as actual passed individual cases' {
        $result = New-GateResult @{Tests=@([pscustomobject]@{Result='Skipped';ExpandedPath=@('T065 absent')}, [pscustomobject]@{Result='Passed';ExpandedPath=@('T029-T031 range')})}
        @(Get-WinImgPassedTestCases $result) | Should -Not -Contain T065
        @(Get-WinImgPassedTestCases $result) | Should -Not -Contain T030
    }
    It 'requires executed real codec tests rather than advertised capability or a generic case label' {
        $result = New-GateResult @{Tests=@([pscustomobject]@{Result='Passed';ExpandedPath=@('T065 fully decodes an actual .png source and its JPEG derivative')})}
        $coverage = @(Get-WinImgExecutedCodecCoverage $result)
        @($coverage | Where-Object executed).Count | Should -Be 1
        @($coverage | Where-Object { $_.extension -eq 'png' -and $_.executed }).Count | Should -Be 1
        @($coverage | Where-Object { $_.extension -eq 'heic' -and $_.executed }).Count | Should -Be 0
        $result.Tests[0].Result='Skipped'
        @(Get-WinImgExecutedCodecCoverage $result | Where-Object executed).Count | Should -Be 0
    }
    It 'rejects deletion and unregistered addition in the suite inventory' {
        $directory = Join-Path $owned 'inventory'; [IO.Directory]::CreateDirectory($directory) | Out-Null
        $file=Join-Path $directory 'Example.Tests.ps1'; [IO.File]::WriteAllText($file, 'Describe fixture {}')
        $manifest=[pscustomobject]@{suites=@('Example.Tests.ps1')}
        { Assert-WinImgMandatoryTestManifest $directory $manifest } | Should -Not -Throw
        { Assert-WinImgMandatoryTestManifest $directory ([pscustomobject]@{suites=@('Example.Tests.ps1','Missing.Tests.ps1')}) } | Should -Throw
        [IO.File]::WriteAllText((Join-Path $directory 'Unregistered.Tests.ps1'), 'Describe fixture {}')
        { Assert-WinImgMandatoryTestManifest $directory $manifest } | Should -Throw
    }
    It 'returns native failure for actual Pester <Kind> discovery without a broken assertion' -ForEach @(
        @{Kind='empty';Text="Describe 'empty control' {}";Count=0;Skipped=0},
        @{Kind='skipped';Text="Describe 'skip control' { It 'unavailable mandatory fixture' -Skip { 1 | Should -Be 1 } }";Count=1;Skipped=1}
    ) {
        $file = Join-Path $owned ($Kind + '.Tests.ps1'); [IO.File]::WriteAllText($file, $Text)
        $result = Invoke-GateChild $file (Join-Path $owned ($Kind + '-result'))
        $result.Exit | Should -Be 1
        $result.Summary.exit_code | Should -Be 1
        $result.Summary.total_count | Should -Be $Count
        $result.Summary.skipped_count | Should -Be $Skipped
        $result.Summary.failed_count | Should -Be 0
        $result.Summary.suite_scope | Should -Be 'focused'
        $result.Summary.infrastructure_error | Should -BeNullOrEmpty
        @($result.Summary.gate_failures).Count | Should -BeGreaterThan 0
    }
}

Describe 'T066 synthetic evidence export privacy and containment' {
    It 'exports counts and hashes while excluding diagnostic text, test names and private absolute paths' {
        $directory=Join-Path $owned 'export-input'; [IO.Directory]::CreateDirectory($directory) | Out-Null
        $privateText='DO-NOT-UPLOAD-C:\Users\private-owner\media'; $head=(git -C $repository rev-parse HEAD).Trim()
        $summary=@{ tested_commit=$head; source_worktree_dirty=$true; powershell_version=$PSVersionTable.PSVersion.ToString(); powershell_edition=$PSVersionTable.PSEdition;
            os_caption='Microsoft Windows synthetic fixture'; source_sha256=@([pscustomobject]@{relative_path='WinImgNormalizer.ps1';sha256=('a'*64);path=$privateText});
            total_count=1;passed_count=1;exit_code=0;test_paths=@($privateText);failed_tests=@($privateText);infrastructure_error=$privateText;
            gate_failures=@($privateText);host_executable=$privateText;imagemagick_version_output=@($privateText);passed_case_ids=@('T064');required_case_ids=@('T064') }
        $summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'summary.json') -Encoding UTF8
        $output=Join-Path $owned 'export-output'
        & (Join-Path $PSScriptRoot 'Export-TestEvidence.ps1') -ResultDirectories @($directory) -OutputDirectory $output
        $text=Get-Content -LiteralPath (Join-Path $output 'test-evidence.json') -Raw -Encoding UTF8
        $text | Should -Not -Match 'DO-NOT-UPLOAD|private-owner|host_executable|test_paths|failed_tests|infrastructure_error|imagemagick_version_output'
        $data=$text | ConvertFrom-Json
        $data.runs[0].tested_commit | Should -Be $head
        $data.runs[0].total_count | Should -Be 1
        $data.runs[0].status | Should -Be 'infrastructure_failed'
        $data.runs[0].source_sha256[0].path | Should -Be 'WinImgNormalizer.ps1'
        @(Get-ChildItem -LiteralPath $output -File).Count | Should -Be 2
        { & (Join-Path $PSScriptRoot 'Export-TestEvidence.ps1') -ResultDirectories @($directory) -OutputDirectory $output } | Should -Throw
    }
    It 'refuses source bindings containing an absolute private path' {
        $directory=Join-Path $owned 'unsafe-input'; [IO.Directory]::CreateDirectory($directory) | Out-Null
        @{tested_commit=('a'*40);source_sha256=@(@{path='C:\Users\private-owner\image.jpeg';sha256=('a'*64)})} | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $directory 'summary.json') -Encoding UTF8
        { & (Join-Path $PSScriptRoot 'Export-TestEvidence.ps1') -ResultDirectories @($directory) -OutputDirectory (Join-Path $owned 'unsafe-output') } | Should -Throw
    }
    It 'rejects a nested environment object instead of retaining private text behind its string conversion' {
        $directory=Join-Path $owned 'nested-input'; [IO.Directory]::CreateDirectory($directory) | Out-Null
        @{tested_commit=('a'*40);runner_image=@{nested=@{private_path='C:/Users/private-owner/media'}};source_sha256=@()} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'summary.json') -Encoding UTF8
        $output=Join-Path $owned 'nested-output'
        { & (Join-Path $PSScriptRoot 'Export-TestEvidence.ps1') -ResultDirectories @($directory) -OutputDirectory $output } | Should -Throw '*must be scalar*'
        Test-Path -LiteralPath $output | Should -BeFalse
    }
}
