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
        param([string]$TestFile, [string]$ResultDirectory,
            [string]$ScriptPath = (Join-Path $PSScriptRoot 'Invoke-Tests.ps1'), [string[]]$AdditionalArguments = @())
        $start = New-Object Diagnostics.ProcessStartInfo
        $start.FileName = (Get-Process -Id $PID).Path
        $tokens = @('-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',$ScriptPath,'-Path',$TestFile,'-ResultDirectory',$ResultDirectory) + $AdditionalArguments
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
    It 'matches actual file-object containers and treats a scriptblock as text rather than a Windows filename' {
        $result = New-GateResult @{Containers=@([pscustomobject]@{Item=[IO.FileInfo]::new((Join-Path $owned 'Example.Tests.ps1'));TotalCount=1;Result='Passed'})}
        @(Get-WinImgTestGateFailures -Result $result -RequiredCases T064 -RequiredSuites Example.Tests.ps1).Count | Should -Be 0
        Get-WinImgTestContainerName ([scriptblock]::Create("Describe 'control' { It 'x' { 1 | Should -Be 2 } }")) | Should -Be 'scriptblock-control'
    }
    It 'accepts real pinned Pester file and scriptblock containers in a fresh host through the required-suite gate' {
        $directory = Join-Path $owned 'actual-containers'
        [IO.Directory]::CreateDirectory($directory) | Out-Null
        $fixture = Join-Path $directory 'ContainerFixture.Tests.ps1'
        [IO.File]::WriteAllText($fixture, "Describe 'T064 actual file fixture' { It 'passes' { 1 | Should -Be 1 } }", [Text.UTF8Encoding]::new($false))
        $probe = Join-Path $directory 'Probe-ActualContainers.ps1'
        $probeText = @'
param([string[]]$Path, [string]$ResultDirectory, [string]$PesterManifest, [string]$TestGatePath,
    [string]$ExpectedManifestHash, [string]$ExpectedPesterVersion)
$ErrorActionPreference = 'Stop'
if (Test-Path -LiteralPath $ResultDirectory) { throw 'Actual container result directory already exists.' }
$manifestHash = (Get-FileHash -LiteralPath $PesterManifest -Algorithm SHA256).Hash.ToLowerInvariant()
if ($manifestHash -ne $ExpectedManifestHash) { throw 'Actual container probe Pester manifest differs from the reviewed pin.' }
$pester = Import-Module $PesterManifest -Force -PassThru
if ($pester.Version.ToString() -ne $ExpectedPesterVersion) { throw 'Actual container probe Pester version differs from the reviewed pin.' }
. $TestGatePath
$configuration = New-PesterConfiguration
$configuration.Run.Path = $Path
$configuration.Run.ScriptBlock = @({ Describe 'T064 actual scriptblock fixture' { It 'passes' { 1 | Should -Be 1 } } })
$configuration.Run.PassThru = $true
$configuration.Run.Exit = $false
$configuration.Run.Throw = $false
$configuration.Output.Verbosity = 'None'
$result = Invoke-Pester -Configuration $configuration
$requiredSuite = [IO.Path]::GetFileName($Path[0])
$failures = @(Get-WinImgTestGateFailures -Result $result -RequiredCases T064 -RequiredSuites $requiredSuite)
$exitCode = if ($failures.Count -eq 0) { 0 } else { 1 }
$summary = [ordered]@{
    exit_code = $exitCode; pester_version = $pester.Version.ToString(); pester_manifest_sha256 = $manifestHash
    total_count = $result.TotalCount; passed_count = $result.PassedCount; failed_count = $result.FailedCount
    skipped_count = $result.SkippedCount; not_run_count = $result.NotRunCount; gate_failures = $failures
    containers = @($result.Containers | ForEach-Object {
        [pscustomobject]@{ item_type = $_.Item.GetType().FullName; name = Get-WinImgTestContainerName $_.Item; result = [string]$_.Result; total = $_.TotalCount }
    })
}
[IO.Directory]::CreateDirectory($ResultDirectory) | Out-Null
$summary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $ResultDirectory 'summary.json') -Encoding UTF8
exit $exitCode
'@
        [IO.File]::WriteAllText($probe, $probeText, [Text.UTF8Encoding]::new($false))
        $loadedPester = Get-Module Pester | Select-Object -First 1
        $manifest = Join-Path $loadedPester.ModuleBase 'Pester.psd1'
        $pin = (Get-Content -LiteralPath (Join-Path $PSScriptRoot 'dependencies.json') -Raw | ConvertFrom-Json).pester
        $arguments = @('-PesterManifest', $manifest, '-TestGatePath', (Join-Path $PSScriptRoot 'TestGate.ps1'),
            '-ExpectedManifestHash', $pin.manifest_sha256, '-ExpectedPesterVersion', $pin.version)
        $observed = Invoke-GateChild -TestFile $fixture -ResultDirectory (Join-Path $directory 'result') -ScriptPath $probe -AdditionalArguments $arguments
        $observed.Exit | Should -Be 0
        $observed.Summary.exit_code | Should -Be 0
        $observed.Summary.total_count | Should -Be 2
        $observed.Summary.passed_count | Should -Be 2
        $observed.Summary.failed_count | Should -Be 0
        $observed.Summary.skipped_count | Should -Be 0
        $observed.Summary.not_run_count | Should -Be 0
        @($observed.Summary.gate_failures).Count | Should -Be 0
        $observed.Summary.pester_version | Should -Be $pin.version
        $observed.Summary.pester_manifest_sha256 | Should -Be $pin.manifest_sha256
        @($observed.Summary.containers | Where-Object { $_.item_type -eq 'System.IO.FileInfo' -and $_.name -eq 'ContainerFixture.Tests.ps1' -and $_.result -eq 'Passed' -and $_.total -eq 1 }).Count | Should -Be 1
        @($observed.Summary.containers | Where-Object { $_.item_type -eq 'System.Management.Automation.ScriptBlock' -and $_.name -eq 'scriptblock-control' -and $_.result -eq 'Passed' -and $_.total -eq 1 }).Count | Should -Be 1
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
    It 'exports both maintained gate types with exact public release bindings and excludes their private source paths' {
        $directory=Join-Path $owned 'release-bindings-input'; [IO.Directory]::CreateDirectory($directory) | Out-Null
        $privateText='DO-NOT-UPLOAD-C:\Users\private-owner\checkout'
        $publicPaths=@('tools/release/Update-ReleaseMetadata.ps1','docs/release/NOTES.md','CHANGELOG.md',
            'release-metadata.json','README.md','LICENSE','.gitattributes','.gitignore',
            'tools/release/Build-Release.ps1','docs/release/GETTING_STARTED.md',
            'docs/release/THIRD_PARTY_NOTICES.md','docs/release/PACKAGING.md','docs/BEHAVIOR.md','SECURITY.md')
        $bindings=@(foreach ($relative in $publicPaths) {
            [pscustomobject]@{relative_path=$relative;path=$privateText;sha256=(Get-FileHash -LiteralPath (Join-Path $repository $relative)).Hash.ToLowerInvariant()}
        })
        $summary=@{tested_commit=('a'*40);source_sha256=$bindings;exit_code=0;result='passed'}
        foreach ($name in @('summary.json','static-analysis-summary.json')) {
            $summary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $directory $name) -Encoding UTF8
        }
        $output=Join-Path $owned 'release-bindings-output'
        & (Join-Path $PSScriptRoot 'Export-TestEvidence.ps1') -ResultDirectories @($directory) -StaticAnalysisDirectories @($directory) -OutputDirectory $output
        foreach ($name in @('test-evidence.json','static-analysis-evidence.json')) {
            $text=Get-Content -LiteralPath (Join-Path $output $name) -Raw -Encoding UTF8
            $text | Should -Not -Match 'DO-NOT-UPLOAD|private-owner'
            $data=$text | ConvertFrom-Json
            $data.runs.Count | Should -Be 1
            $data.runs[0].status | Should -Be 'passed'
            $data.runs[0].source_sha256.Count | Should -Be $publicPaths.Count
            foreach ($binding in $bindings) {
                $exported=@($data.runs[0].source_sha256 | Where-Object { $_.path -ceq $binding.relative_path })
                $exported.Count | Should -Be 1
                $exported[0].sha256 | Should -Be $binding.sha256
                @($exported[0].PSObject.Properties.Name | Sort-Object) -join '|' | Should -Be 'path|sha256'
            }
        }
    }
    It 'refuses nearby nonallowlisted release source <Relative>' -ForEach @(
        @{Relative='docs/release/private-notes.md'}, @{Relative='docs/codex-winimg/SESSION_LOG.md'},
        @{Relative='tools/release/private-command.ps1'}, @{Relative='README.md.private'},
        @{Relative='CHANGELOG.md.bak'}, @{Relative='docs/release/../release/NOTES.md'},
        @{Relative='tools/release/Build-Release.ps1.private'}, @{Relative='docs/release/GETTING_STARTED.md.bak'},
        @{Relative='docs/release/THIRD_PARTY_NOTICES.md.private'}, @{Relative='docs/release/PACKAGING.md.bak'},
        @{Relative='.gitignore.private'}, @{Relative='docs/BEHAVIOR.md.private'},
        @{Relative='SECURITY.md.bak'}, @{Relative='docs/../SECURITY.md'}
    ) {
        $directory=Join-Path $owned ('nonallowlisted-source-' + [guid]::NewGuid().ToString('N'))
        [IO.Directory]::CreateDirectory($directory) | Out-Null
        @{tested_commit=('a'*40);source_sha256=@(@{relative_path=$Relative;sha256=('a'*64)})} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'summary.json') -Encoding UTF8
        $output=Join-Path $directory 'output'
        { & (Join-Path $PSScriptRoot 'Export-TestEvidence.ps1') -ResultDirectories @($directory) -OutputDirectory $output } | Should -Throw '*Unsafe evidence source binding*'
        Test-Path -LiteralPath $output | Should -BeFalse
    }
    It 'rejects a nested <Field> list entry containing an otherwise allowed identifier and private text' -ForEach @(
        @{Field='passed_case_ids';Allowed='T064';Analysis=$false},
        @{Field='required_case_ids';Allowed='T064';Analysis=$false},
        @{Field='imagemagick_delegates';Allowed='jpeg';Analysis=$false},
        @{Field='rules';Allowed='PSAvoidUsingInvokeExpression';Analysis=$true}
    ) {
        $directory=Join-Path $owned ('nested-list-' + $Field); [IO.Directory]::CreateDirectory($directory) | Out-Null
        $summary=[ordered]@{tested_commit=('a'*40);source_sha256=@()}
        # Preserve one inner array across JSON parsing. PowerShell array -match
        # accepts its allowed member; exporting the original array leaks its peer.
        $summary[$Field]=,@($Allowed, 'DO-NOT-UPLOAD-C:\Users\private-owner\media')
        $name=if($Analysis){'static-analysis-summary.json'}else{'summary.json'}
        $summary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $directory $name) -Encoding UTF8
        $parameters=@{OutputDirectory=(Join-Path $owned ('nested-list-output-' + $Field))}
        if($Analysis){$parameters.StaticAnalysisDirectories=@($directory)}else{$parameters.ResultDirectories=@($directory)}
        { & (Join-Path $PSScriptRoot 'Export-TestEvidence.ps1') @parameters } | Should -Throw '*must be scalar strings*'
        Test-Path -LiteralPath $parameters.OutputDirectory | Should -BeFalse
    }
    It 'rejects structured source <Field> before coercion can hide additional values' -ForEach @(
        @{Field='relative_path';Values=@('tests/synthetic.ps1','DO-NOT-UPLOAD-private-owner')},
        @{Field='sha256';Values=@(('a'*64),('b'*64))}
    ) {
        $directory=Join-Path $owned ('structured-source-' + $Field); [IO.Directory]::CreateDirectory($directory) | Out-Null
        $binding=[ordered]@{relative_path='tests/synthetic.ps1';sha256=('a'*64)}
        $binding[$Field]=$Values
        @{tested_commit=('a'*40);source_sha256=@($binding)} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $directory 'summary.json') -Encoding UTF8
        $output=Join-Path $owned ('structured-source-output-' + $Field)
        { & (Join-Path $PSScriptRoot 'Export-TestEvidence.ps1') -ResultDirectories @($directory) -OutputDirectory $output } | Should -Throw '*must be scalar strings*'
        Test-Path -LiteralPath $output | Should -BeFalse
    }
    It 'rejects a structured codec format even when every array member is an allowed codec' {
        $directory=Join-Path $owned 'structured-codec'; [IO.Directory]::CreateDirectory($directory) | Out-Null
        @{tested_commit=('a'*40);source_sha256=@();codec_capabilities=@(@{format=@('JPEG','HEIC');read=$true;write=$true})} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $directory 'summary.json') -Encoding UTF8
        $output=Join-Path $owned 'structured-codec-output'
        { & (Join-Path $PSScriptRoot 'Export-TestEvidence.ps1') -ResultDirectories @($directory) -OutputDirectory $output } | Should -Throw '*must be a scalar string*'
        Test-Path -LiteralPath $output | Should -BeFalse
    }
    It 'exports counts and hashes while excluding diagnostic text, test names and private absolute paths' {
        $directory=Join-Path $owned 'export-input'; [IO.Directory]::CreateDirectory($directory) | Out-Null
        $privateText='DO-NOT-UPLOAD-C:\Users\private-owner\media'; $head=(git -C $repository rev-parse HEAD).Trim()
        $timestamp='2026-10-05T15:12:34.2440815Z'
        $summary=@{ tested_commit=$head; source_worktree_dirty=$true; powershell_version=$PSVersionTable.PSVersion.ToString(); powershell_edition=$PSVersionTable.PSEdition;
            observed_at_utc=$timestamp; os_caption='Microsoft Windows synthetic fixture'; source_sha256=@([pscustomobject]@{relative_path='WinImgNormalizer.ps1';sha256=('a'*64);path=$privateText});
            total_count=1;passed_count=1;exit_code=0;test_paths=@($privateText);failed_tests=@($privateText);infrastructure_error=$privateText;
            gate_failures=@($privateText);host_executable=$privateText;imagemagick_version_output=@($privateText);passed_case_ids=@('T064');required_case_ids=@('T064') }
        $summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'summary.json') -Encoding UTF8
        $output=Join-Path $owned 'export-output'
        & (Join-Path $PSScriptRoot 'Export-TestEvidence.ps1') -ResultDirectories @($directory) -OutputDirectory $output
        $text=Get-Content -LiteralPath (Join-Path $output 'test-evidence.json') -Raw -Encoding UTF8
        $text | Should -Not -Match 'DO-NOT-UPLOAD|private-owner|host_executable|test_paths|failed_tests|infrastructure_error|imagemagick_version_output'
        $text | Should -Match ('"observed_at_utc"\s*:\s*"' + [regex]::Escape($timestamp) + '"')
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
