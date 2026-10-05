# Run in a fresh -NoProfile Windows PowerShell 5.1 or PowerShell 7 process.
[CmdletBinding()]
param(
    [string[]]$Path,
    [string]$MagickPath,
    [switch]$DownloadDependencies,
    [switch]$DeliberateFailure,
    [string]$ResultDirectory
)

$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path -Parent $PSScriptRoot
$mandatory = -not $PSBoundParameters.ContainsKey('Path')
$manifestPath = Join-Path $PSScriptRoot 'mandatory-tests.json'
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
. (Join-Path $PSScriptRoot 'TestGate.ps1')
if ($mandatory) { $Path = @($manifest.suites | ForEach-Object { Join-Path $PSScriptRoot $_ }) }
function Get-TestSourceBindings {
    param([string[]]$SelectedPaths)
    return @(
        $files = @((Join-Path $repositoryRoot 'WinImgNormalizer.ps1'), (Join-Path $repositoryRoot 'WinImgNormalizer.bat'), $PSCommandPath, (Join-Path $PSScriptRoot 'Initialize-TestDependencies.ps1'), (Join-Path $PSScriptRoot 'dependencies.json'), $manifestPath, (Join-Path $PSScriptRoot 'TestGate.ps1'), (Join-Path $PSScriptRoot 'Export-TestEvidence.ps1'), (Join-Path $PSScriptRoot 'Invoke-StaticAnalysis.ps1'), (Join-Path $PSScriptRoot 'PSScriptAnalyzerSettings.psd1'), (Join-Path $repositoryRoot '.github/workflows/windows-tests.yml'), (Join-Path $PSScriptRoot 'legacy/toolchain.json'))
        # Colour assertions use immutable profiles and an independent reference;
        # bind those exact bytes in targeted and full-suite run summaries too.
        $colourFixtureRoot = Join-Path $PSScriptRoot 'fixtures/colour'
        $files += @(Get-ChildItem -LiteralPath $colourFixtureRoot -Recurse -File -Force | ForEach-Object { $_.FullName })
        # Bind the procedural size/quality measurement recipe with the gate.
        $files += Join-Path $PSScriptRoot 'Measure-SizeQuality.ps1'
        $files += Join-Path $PSScriptRoot 'fixtures/LauncherConsoleFixture.cs'
        foreach ($testPath in $SelectedPaths) {
            if (Test-Path -LiteralPath $testPath -PathType Container) {
                $files += @(Get-ChildItem -LiteralPath $testPath -Recurse -Filter '*.Tests.ps1' | ForEach-Object { $_.FullName })
            }
            elseif (Test-Path -LiteralPath $testPath -PathType Leaf) { $files += $testPath }
        }
        foreach ($file in ($files | Select-Object -Unique)) {
            $hash = Get-FileHash -LiteralPath $file -Algorithm SHA256
            [pscustomobject]@{ path = $hash.Path; relative_path = $hash.Path.Substring($repositoryRoot.Length + 1).Replace('\','/'); sha256 = $hash.Hash.ToLowerInvariant() }
        }
    )
}
$scratchRoot = [IO.Path]::GetFullPath((Join-Path $repositoryRoot '.scratch'))
if ([string]::IsNullOrWhiteSpace($ResultDirectory)) {
    $ResultDirectory = Join-Path $scratchRoot ('test-results/' + [guid]::NewGuid().ToString('N'))
}
$ResultDirectory = [IO.Path]::GetFullPath($ResultDirectory)
if (-not $ResultDirectory.StartsWith($scratchRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Test results must be a new child directory under the repository .scratch directory.'
}
$directory = [IO.DirectoryInfo]::new($ResultDirectory)
while ($null -ne $directory) {
    if ($directory.Exists -and (($directory.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0)) {
        throw 'Test results must not traverse a reparse point.'
    }
    $directory = $directory.Parent
}
if (Test-Path -LiteralPath $ResultDirectory) { throw 'Test result directory already exists; choose a new directory.' }
[IO.Directory]::CreateDirectory($ResultDirectory) | Out-Null
$summaryPath = Join-Path $ResultDirectory 'summary.json'
$previousMagick = [Environment]::GetEnvironmentVariable('WINIMG_TEST_MAGICK', 'Process')
$exitCode = 1
$summary = [ordered]@{
    schema_version = 1
    observed_at_utc = [DateTime]::UtcNow.ToString('o')
    result = 'infrastructure_failed'
    powershell_version = $PSVersionTable.PSVersion.ToString()
    powershell_edition = $PSVersionTable.PSEdition
    os_version = [Environment]::OSVersion.Version.ToString()
    architecture_bits = [IntPtr]::Size * 8
    culture = [Globalization.CultureInfo]::CurrentCulture.Name
    deliberate_failure = [bool]$DeliberateFailure
    test_paths = $Path
    total_count = 0
    passed_count = 0
    failed_count = 0
    skipped_count = 0
    inconclusive_count = 0
    not_run_count = 0
    failed_blocks_count = 0
    failed_containers_count = 0
    failed_tests = @()
    skipped_tests = @()
    gate_failures = @()
    suite_scope = $(if ($mandatory) { 'mandatory' } else { 'focused' })
    passed_case_ids = @()
    required_case_ids = $(if ($mandatory) { @($manifest.required_case_ids) } else { @() })
}

try {
    if ($env:OS -ne 'Windows_NT') { throw 'The mandatory suite requires actual Windows.' }
    $summary.tested_commit = (git -C $repositoryRoot rev-parse HEAD).Trim()
    if ($LASTEXITCODE -ne 0 -or $summary.tested_commit -notmatch '^[0-9a-f]{40}$') { throw 'Cannot bind tests to a Git commit.' }
    $summary.source_worktree_dirty = [bool](git -C $repositoryRoot status --porcelain=v1 --untracked-files=all)
    $summary.execution_environment = $(if ($env:GITHUB_ACTIONS -eq 'true') { 'github_windows_server' } else { 'windows_local_automated' })
    $summary.runner_image = $env:ImageOS
    $summary.runner_image_version = $env:ImageVersion
    $summary.runner_label = $env:WINIMG_CI_RUNNER_LABEL
    $summary.github_checkout_sha = $env:GITHUB_SHA
    if ($env:GITHUB_ACTIONS -eq 'true' -and ($summary.source_worktree_dirty -or $summary.github_checkout_sha -ne $summary.tested_commit)) {
        throw 'CI requires an exact clean GITHUB_SHA checkout.'
    }
    if ($mandatory) { Assert-WinImgMandatoryTestManifest -TestsRoot $PSScriptRoot -Manifest $manifest }
    if ($Path.Count -eq 0) { throw 'Explicit test selection is empty.' }
    $operatingSystem = Get-CimInstance Win32_OperatingSystem
    $summary.os_caption = $operatingSystem.Caption
    $summary.os_build = $operatingSystem.BuildNumber
    $summary.os_version = $operatingSystem.Version
    $summary.ui_culture = [Globalization.CultureInfo]::CurrentUICulture.Name
    $summary.host_executable = (Get-Process -Id $PID).Path
    $summary.persistent_policy_before = @(Get-ExecutionPolicy -List | Where-Object { $_.Scope -ne 'Process' } | ForEach-Object {
        [pscustomobject]@{ scope = $_.Scope.ToString(); execution_policy = $_.ExecutionPolicy.ToString() }
    })
    $dependencies = & (Join-Path $PSScriptRoot 'Initialize-TestDependencies.ps1') -Download:$DownloadDependencies -MagickPath $MagickPath
    $loaded = @(Get-Module Pester)
    if (@($loaded | Where-Object { $_.Version.ToString() -ne $dependencies.PesterVersion }).Count -gt 0) {
        throw 'A different Pester version is already loaded. Start a fresh -NoProfile host.'
    }
    $pester = Import-Module $dependencies.PesterManifest -Force -PassThru
    if ($pester.Version.ToString() -ne $dependencies.PesterVersion) { throw 'Loaded Pester version differs from the development pin.' }
    $summary.pester_version = $pester.Version.ToString()
    $summary.pester_package_sha256 = $dependencies.PesterPackageSha256
    $summary.imagemagick_executable_sha256 = $dependencies.ImageMagickExecutableSha256
    $imageVersion = @(& $dependencies.MagickPath -version)
    if ($LASTEXITCODE -ne 0) { throw 'ImageMagick version command failed.' }
    $summary.imagemagick_version_output = $imageVersion
    $summary.imagemagick_version = $dependencies.ImageMagickVersion
    $summary.imagemagick_delegates = @([regex]::Match(($imageVersion -join "`n"), '(?m)^Delegates[^:]*:\s*(.*)$').Groups[1].Value.Trim() -split '\s+')
    $formatOutput = @(& $dependencies.MagickPath -list format)
    if ($LASTEXITCODE -ne 0) { throw 'ImageMagick capability command failed.' }
    $summary.codec_capabilities = @(
        foreach ($line in $formatOutput) {
            if ($line -match '^\s*(JPEG|PNG|BMP|TIFF|GIF|HEIC|HEIF|WEBP)\*?\s+(?:\S+\s+)?([r-][w-][+-])\s') {
                [pscustomobject]@{ format = $Matches[1]; read = ($Matches[2][0] -eq 'r'); write = ($Matches[2][1] -eq 'w') }
            }
        }
    )
    if ($mandatory) {
        foreach ($format in @('JPEG','PNG','BMP','TIFF','GIF','WEBP')) {
            if (@($summary.codec_capabilities | Where-Object { $_.format -eq $format -and $_.read }).Count -ne 1) {
                throw ('Mandatory codec reader is unavailable: ' + $format)
            }
        }
        if (@($summary.codec_capabilities | Where-Object { $_.format -in @('HEIC','HEIF') -and $_.read }).Count -eq 0) { throw 'Mandatory HEIC/HEIF decoder is unavailable.' }
        if (@($summary.codec_capabilities | Where-Object { $_.format -eq 'JPEG' -and $_.write }).Count -ne 1) { throw 'Mandatory JPEG writer is unavailable.' }
    }
    $summary.script_analyzer_versions_available = @(Get-Module PSScriptAnalyzer -ListAvailable | ForEach-Object { $_.Version.ToString() })
    # Existing Windows Framework compiler used only by the private-console fixture.
    $fixtureCompiler = Join-Path $env:SystemRoot 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
    if (Test-Path -LiteralPath $fixtureCompiler -PathType Leaf) {
        $summary.test_fixture_compiler = [ordered]@{
            file_version = ([Diagnostics.FileVersionInfo]::GetVersionInfo($fixtureCompiler)).FileVersion
            sha256 = (Get-FileHash -LiteralPath $fixtureCompiler -Algorithm SHA256).Hash.ToLowerInvariant()
        }
    }
    [Environment]::SetEnvironmentVariable('WINIMG_TEST_MAGICK', $dependencies.MagickPath, 'Process')
    Write-Host ('Windows PowerShell host: {0} {1}; Pester: {2}; OS: {3}; {4}-bit; locale: {5}' -f $PSVersionTable.PSEdition, $summary.powershell_version, $summary.pester_version, $summary.os_version, $summary.architecture_bits, $summary.culture)
    $imageVersion | ForEach-Object { Write-Host $_ }

    $summary.source_sha256 = @(Get-TestSourceBindings -SelectedPaths $Path)
    $configuration = New-PesterConfiguration
    $configuration.Run.Path = $Path
    $configuration.Run.PassThru = $true
    $configuration.Run.Exit = $false
    $configuration.Run.Throw = $false
    $configuration.Output.Verbosity = 'Detailed'
    $configuration.TestResult.Enabled = $true
    $configuration.TestResult.OutputPath = Join-Path $ResultDirectory 'results.xml'
    $configuration.TestResult.OutputFormat = 'NUnitXml'
    if ($DeliberateFailure) {
        $configuration.Run.ScriptBlock = @({
            Describe 'T007 runner credibility control' {
                It 'deliberate assertion must fail' { 1 | Should -Be 2 }
            }
        })
    }
    $result = Invoke-Pester -Configuration $configuration
    $summary.result = $result.Result.ToString()
    $summary.total_count = $result.TotalCount
    $summary.passed_count = $result.PassedCount
    $summary.failed_count = $result.FailedCount
    $summary.skipped_count = $result.SkippedCount
    $summary.inconclusive_count = $result.InconclusiveCount
    $summary.not_run_count = $result.NotRunCount
    $summary.failed_blocks_count = $result.FailedBlocksCount
    $summary.failed_containers_count = $result.FailedContainersCount
    $summary.failed_tests = @($result.Tests | Where-Object { $_.Result -eq 'Failed' } | ForEach-Object { $_.ExpandedPath -join '.' })
    $summary.skipped_tests = @($result.Tests | Where-Object { $_.Result -eq 'Skipped' } | ForEach-Object { $_.ExpandedPath -join '.' })
    $summary.duration_seconds = $result.Duration.TotalSeconds
    $summary.passed_case_ids = @(Get-WinImgPassedTestCases $result)
    $summary.codec_coverage = @(Get-WinImgExecutedCodecCoverage $result)
    $summary.suite_results = @($result.Containers | ForEach-Object {
        [pscustomobject]@{ suite = $(if ($_.Item -is [string]) { [IO.Path]::GetFileName($_.Item) } else { 'scriptblock-control' }); result = [string]$_.Result; total = $_.TotalCount; passed = $_.PassedCount }
    })
    $afterSource = @(Get-TestSourceBindings -SelectedPaths $Path)
    $summary.source_unchanged = (($summary.source_sha256 | ConvertTo-Json -Depth 5 -Compress) -eq ($afterSource | ConvertTo-Json -Depth 5 -Compress))
    $summary.commit_unchanged = ((git -C $repositoryRoot rev-parse HEAD).Trim() -eq $summary.tested_commit)
    $summary.source_worktree_dirty_after = [bool](git -C $repositoryRoot status --porcelain=v1 --untracked-files=all)
    $requiredCases = @(); $requiredSuites = @()
    if ($mandatory) { $requiredCases = @($manifest.required_case_ids); $requiredSuites = @($manifest.suites) }
    $failures = @(Get-WinImgTestGateFailures -Result $result -RequiredCases $requiredCases -RequiredSuites $requiredSuites)
    if ($mandatory) {
        foreach ($codec in $summary.codec_coverage) {
            if (-not $codec.executed) { $failures += ('Required real codec test did not pass exactly once: .' + $codec.extension) }
        }
    }
    if (-not $summary.source_unchanged -or -not $summary.commit_unchanged) { $failures += 'Source or commit changed during tests.' }
    if ($env:GITHUB_ACTIONS -eq 'true' -and $summary.source_worktree_dirty_after) { $failures += 'CI worktree became dirty during tests.' }
    $summary.gate_failures = $failures
    if ($failures.Count -eq 0) { $exitCode = 0 }
    else { $failures | ForEach-Object { Write-Host ('TEST GATE FAILED: ' + $_) -ForegroundColor Red } }
}
catch {
    $summary.result = 'infrastructure_failed'
    $summary.infrastructure_error = $_.Exception.Message
    Write-Host ('TEST INFRASTRUCTURE FAILED: ' + $_.Exception.Message) -ForegroundColor Red
}
finally {
    [Environment]::SetEnvironmentVariable('WINIMG_TEST_MAGICK', $previousMagick, 'Process')
    if ($summary.Contains('persistent_policy_before')) {
        $summary.persistent_policy_after = @(Get-ExecutionPolicy -List | Where-Object { $_.Scope -ne 'Process' } | ForEach-Object {
            [pscustomobject]@{ scope = $_.Scope.ToString(); execution_policy = $_.ExecutionPolicy.ToString() }
        })
        $summary.persistent_policy_unchanged = ($summary.persistent_policy_before | ConvertTo-Json -Compress) -eq ($summary.persistent_policy_after | ConvertTo-Json -Compress)
        if (-not $summary.persistent_policy_unchanged) {
            $summary.gate_failures += 'Persistent execution policy changed during the test run.'
            $exitCode = 1
        }
    }
    $summary.exit_code = $exitCode
    $summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $summaryPath -Encoding UTF8
    Write-Host ('Test summary: ' + $summaryPath)
    Write-Host ('Total={0}; Passed={1}; Failed={2}; Skipped={3}; NotRun={4}; Exit={5}' -f $summary.total_count, $summary.passed_count, $summary.failed_count, $summary.skipped_count, $summary.not_run_count, $exitCode)
}
exit $exitCode
