# Run in a fresh -NoProfile Windows PowerShell 5.1 or PowerShell 7 process.
[CmdletBinding()]
param(
    [switch]$DownloadDependencies,
    [switch]$DeliberateFailure,
    [string]$ResultDirectory
)

$ErrorActionPreference = 'Stop'
$repositoryRoot = [IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))
$scratchRoot = [IO.Path]::GetFullPath((Join-Path $repositoryRoot '.scratch'))
function Assert-AnalysisDirectory {
    param([string]$Directory)
    $current = [IO.DirectoryInfo]::new([IO.Path]::GetFullPath($Directory))
    while ($null -ne $current) {
        if ($current.Exists -and (($current.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0)) {
            throw 'Static-analysis directories must not traverse reparse points.'
        }
        $current = $current.Parent
    }
}
function Get-AnalysisBinding {
    param([string[]]$Files)
    foreach ($file in ($Files | Sort-Object -Unique)) {
        $item = Get-Item -LiteralPath $file
        Assert-AnalysisDirectory (Split-Path -Parent $item.FullName)
        if ($item.PSIsContainer -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 -or
            -not $item.FullName.StartsWith($repositoryRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Static-analysis sources must be ordinary files contained in this repository.'
        }
        [pscustomobject]@{
            path = $item.FullName.Substring($repositoryRoot.Length + 1).Replace('\', '/')
            sha256 = (Get-FileHash -LiteralPath $item.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
        }
    }
}
function ConvertTo-AnalysisFinding {
    param([object[]]$Diagnostics, [string]$Source)
    foreach ($diagnostic in $Diagnostics) {
        # Rule identifiers and numeric locations suffice for public evidence.
        # Analyzer messages/extents can include literal source secrets or paths.
        [pscustomobject]@{
            source = $Source
            rule = $diagnostic.RuleName
            severity = $diagnostic.Severity.ToString()
            line = [int]$diagnostic.Line
            column = [int]$diagnostic.Column
        }
    }
}

if ([string]::IsNullOrWhiteSpace($ResultDirectory)) {
    $ResultDirectory = Join-Path $scratchRoot ('static-analysis/' + [guid]::NewGuid().ToString('N'))
}
$ResultDirectory = [IO.Path]::GetFullPath($ResultDirectory)
if (-not $ResultDirectory.StartsWith($scratchRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Static-analysis results must be a new child directory under repository .scratch.'
}
Assert-AnalysisDirectory $ResultDirectory
if (Test-Path -LiteralPath $ResultDirectory) { throw 'Static-analysis result directory already exists.' }
[IO.Directory]::CreateDirectory($ResultDirectory) | Out-Null
$summaryPath = Join-Path $ResultDirectory 'static-analysis-summary.json'
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
    scanned_file_count = 0
    scoped_findings_count = 0
    control_findings_count = 0
    control_detected = $false
    findings_count = 0
    findings = @()
    source_sha256 = @()
    rules = @()
}
$phase = 'environment'
try {
    if ($env:OS -ne 'Windows_NT') { throw 'The mandatory analyzer gate requires actual Windows.' }
    $operatingSystem = Get-CimInstance Win32_OperatingSystem
    $summary.os_caption = $operatingSystem.Caption
    $summary.os_build = $operatingSystem.BuildNumber
    $summary.os_version = $operatingSystem.Version
    $summary.ui_culture = [Globalization.CultureInfo]::CurrentUICulture.Name
    $summary.host_file_version = ([Diagnostics.FileVersionInfo]::GetVersionInfo((Get-Process -Id $PID).Path)).FileVersion
    $commit = @(& git -C $repositoryRoot rev-parse --verify HEAD)
    if ($LASTEXITCODE -ne 0 -or $commit.Count -ne 1 -or $commit[0] -notmatch '^[0-9a-f]{40}$') { throw 'Could not bind exact tested Git commit.' }
    $summary.tested_commit = $commit[0]
    # Maintain the application and all top-level development entry points. The
    # Pester DSL bodies and legacy/vendor trees have separate execution gates.
    $files = @((Join-Path $repositoryRoot 'WinImgNormalizer.ps1'))
    $files += @(Get-ChildItem -LiteralPath $PSScriptRoot -File -Filter '*.ps1' | Where-Object { $_.Name -notlike '*.Tests.ps1' } | ForEach-Object { $_.FullName })
    if ($files.Count -lt 2) { throw 'Static-analysis scope unexpectedly empty or incomplete.' }
    $settingsPath = Join-Path $PSScriptRoot 'PSScriptAnalyzerSettings.psd1'
    $bindingFiles = @($files) + @($settingsPath, (Join-Path $PSScriptRoot 'dependencies.json'))
    $summary.source_sha256 = @(Get-AnalysisBinding $bindingFiles)
    $sourceState = @(& git -C $repositoryRoot status --porcelain=v1 --untracked-files=normal -- @($summary.source_sha256 | ForEach-Object { $_.path }))
    if ($LASTEXITCODE -ne 0) { throw 'Could not observe source worktree state.' }
    $summary.source_worktree_dirty = $sourceState.Count -gt 0
    $settings = Import-PowerShellDataFile -LiteralPath $settingsPath
    $summary.rules = @($settings.IncludeRules)
    if ($summary.rules.Count -eq 0) { throw 'Static-analysis rule set is empty.' }
    $summary.syntax_target_versions = @($settings.Rules.PSUseCompatibleSyntax.TargetVersions)
    $phase = 'dependency_bootstrap'
    $dependencies = & (Join-Path $PSScriptRoot 'Initialize-TestDependencies.ps1') -Download:$DownloadDependencies -StaticAnalysisOnly
    $phase = 'analyzer_load'
    if (@(Get-Module PSScriptAnalyzer).Count -gt 0) { throw 'An analyzer module is already loaded; use a fresh -NoProfile host.' }
    $analyzer = Import-Module $dependencies.PSScriptAnalyzerManifest -Force -PassThru
    if ($analyzer.Version.ToString() -ne $dependencies.PSScriptAnalyzerVersion) { throw 'Loaded analyzer version differs from the development pin.' }
    $summary.analyzer_version = $analyzer.Version.ToString()
    $summary.analyzer_package_sha256 = $dependencies.PSScriptAnalyzerPackageSha256
    $availableRules = @(PSScriptAnalyzer\Get-ScriptAnalyzerRule | ForEach-Object { $_.RuleName })
    foreach ($rule in $summary.rules) {
        if ($availableRules -notcontains $rule) { throw 'A configured static-analysis rule is unavailable.' }
    }
    $phase = 'source_analysis'
    foreach ($file in $files) {
        $relative = [IO.Path]::GetFullPath($file).Substring($repositoryRoot.Length + 1).Replace('\', '/')
        $diagnostics = @(PSScriptAnalyzer\Invoke-ScriptAnalyzer -Path $file -Settings $settingsPath)
        $summary.findings += @(ConvertTo-AnalysisFinding -Diagnostics $diagnostics -Source $relative)
        $summary.scanned_file_count++
    }
    $summary.scoped_findings_count = $summary.findings.Count
    if ($DeliberateFailure) {
        $phase = 'credibility_control'
        # Analyze this synthetic expression as text. It is never executed.
        $controlSource = "Invoke-Expression 'Write-Output owned-analyzer-control'"
        $hasher = [Security.Cryptography.SHA256]::Create()
        try { $summary.control_source_sha256 = [BitConverter]::ToString($hasher.ComputeHash([Text.Encoding]::UTF8.GetBytes($controlSource))).Replace('-', '').ToLowerInvariant() }
        finally { $hasher.Dispose() }
        $control = @(PSScriptAnalyzer\Invoke-ScriptAnalyzer -ScriptDefinition $controlSource -Settings $settingsPath)
        $summary.control_findings_count = $control.Count
        $summary.control_detected = $control.Count -eq 1 -and $control[0].RuleName -eq 'PSAvoidUsingInvokeExpression' -and $control[0].Severity.ToString() -eq 'Warning'
        $summary.findings += @(ConvertTo-AnalysisFinding -Diagnostics $control -Source 'control/unsafe-expression.ps1')
        if (-not $summary.control_detected) { throw 'The real analyzer credibility control was not detected exactly once.' }
    }
    $phase = 'source_binding'
    $summary.source_bindings_unchanged = ($summary.source_sha256 | ConvertTo-Json -Compress) -eq (@(Get-AnalysisBinding $bindingFiles) | ConvertTo-Json -Compress)
    if (-not $summary.source_bindings_unchanged) { throw 'Analyzed sources changed during static analysis.' }
    $summary.findings_count = $summary.findings.Count
    if ($summary.findings_count -eq 0) { $summary.result = 'passed'; $exitCode = 0 }
    else { $summary.result = 'analysis_failed' }
}
catch {
    # The private console keeps details. The uploadable JSON never retains raw
    # exception messages, usernames, paths, source text or credential values.
    $summary.result = 'infrastructure_failed'
    $summary.infrastructure_error = 'Static-analysis infrastructure failed; details are local only.'
    $summary.infrastructure_phase = $phase
    Write-Host ('STATIC ANALYSIS INFRASTRUCTURE FAILED: ' + $_.Exception.Message) -ForegroundColor Red
}
finally {
    $summary.findings_count = $summary.findings.Count
    $summary.exit_code = $exitCode
    $summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $summaryPath -Encoding UTF8
    Write-Host ('Static-analysis summary: ' + $summaryPath)
    Write-Host ('Files={0}; Findings={1}; ControlFindings={2}; Exit={3}' -f $summary.scanned_file_count, $summary.findings_count, $summary.control_findings_count, $exitCode)
}
exit $exitCode
