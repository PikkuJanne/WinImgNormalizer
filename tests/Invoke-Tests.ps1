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
if (-not $PSBoundParameters.ContainsKey('Path')) {
    $Path = @((Join-Path $PSScriptRoot 'Normalizer.Tests.ps1'), (Join-Path $PSScriptRoot 'Preflight.Tests.ps1'), (Join-Path $PSScriptRoot 'Traversal.Tests.ps1'), (Join-Path $PSScriptRoot 'Naming.Tests.ps1'))
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
}

try {
    if ($env:OS -ne 'Windows_NT') { throw 'The mandatory suite requires actual Windows.' }
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
    $summary.script_analyzer_versions_available = @(Get-Module PSScriptAnalyzer -ListAvailable | ForEach-Object { $_.Version.ToString() })
    [Environment]::SetEnvironmentVariable('WINIMG_TEST_MAGICK', $dependencies.MagickPath, 'Process')
    Write-Host ('Windows PowerShell host: {0} {1}; Pester: {2}; OS: {3}; {4}-bit; locale: {5}' -f $PSVersionTable.PSEdition, $summary.powershell_version, $summary.pester_version, $summary.os_version, $summary.architecture_bits, $summary.culture)
    $imageVersion | ForEach-Object { Write-Host $_ }

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
    $summary.source_sha256 = @(
        $files = @((Join-Path $repositoryRoot 'WinImgNormalizer.ps1'), (Join-Path $repositoryRoot 'WinImgNormalizer.bat'), $PSCommandPath, (Join-Path $PSScriptRoot 'Initialize-TestDependencies.ps1'), (Join-Path $PSScriptRoot 'dependencies.json'))
        foreach ($testPath in $Path) {
            if (Test-Path -LiteralPath $testPath -PathType Container) {
                $files += @(Get-ChildItem -LiteralPath $testPath -Recurse -Filter '*.Tests.ps1' | ForEach-Object { $_.FullName })
            }
            elseif (Test-Path -LiteralPath $testPath -PathType Leaf) { $files += $testPath }
        }
        foreach ($file in ($files | Select-Object -Unique)) {
            $hash = Get-FileHash -LiteralPath $file -Algorithm SHA256
            [pscustomobject]@{ path = $hash.Path; sha256 = $hash.Hash.ToLowerInvariant() }
        }
    )
    $failures = @()
    if ($result.TotalCount -eq 0) { $failures += 'Zero tests discovered.' }
    if ($result.FailedCount -gt 0) { $failures += 'One or more assertions failed.' }
    if ($result.SkippedCount -gt 0) { $failures += 'Mandatory tests were skipped.' }
    if ($result.NotRunCount -gt 0 -or $result.InconclusiveCount -gt 0) { $failures += 'Mandatory tests were not completed.' }
    if ($result.FailedBlocksCount -gt 0 -or $result.FailedContainersCount -gt 0 -or $result.Result -ne 'Passed') { $failures += 'Pester reported a failing run, block or container.' }
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
