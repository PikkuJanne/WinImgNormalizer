# Export only allowlisted synthetic evidence, never raw logs, XML, argv or media.
[CmdletBinding()]
param([string[]]$ResultDirectories, [string[]]$StaticAnalysisDirectories, [Parameter(Mandatory)][string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
$scratch = [IO.Path]::GetFullPath((Join-Path $repository '.scratch'))
function Assert-ExportPath {
    param([string]$Path)
    $full = [IO.Path]::GetFullPath($Path)
    if (-not $full.StartsWith($scratch + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Evidence must remain under owned scratch.' }
    $entry = $full
    while ($entry) {
        if (Test-Path -LiteralPath $entry) {
            if (((Get-Item -LiteralPath $entry -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Evidence must not traverse a reparse point.' }
        }
        $parent = [IO.Directory]::GetParent($entry)
        if (-not $parent) { break }; $entry = $parent.FullName
    }
    return $full
}
function Get-ExportSourceBindings {
    param([object[]]$Sources)
    # These public release inputs/outputs are bound by the maintained gates.
    # Keep this exact list rather than allowing arbitrary docs or tool files.
    $releaseSources = @(
        'tools/release/Update-ReleaseMetadata.ps1', 'docs/release/NOTES.md',
        'tools/release/Build-Release.ps1', 'docs/release/GETTING_STARTED.md',
        'docs/release/THIRD_PARTY_NOTICES.md', 'docs/release/PACKAGING.md', 'docs/BEHAVIOR.md', 'SECURITY.md',
        'CHANGELOG.md', 'release-metadata.json', 'README.md', 'LICENSE', '.gitattributes', '.gitignore'
    )
    $releaseSources += @(
        'tools/website/Test-WebsiteHandoff.ps1', 'docs/website/metadata.json',
        'docs/website/PRODUCT_COPY.md', 'docs/website/INTEGRATION.md',
        'WinImgNormalizer_icon_variant.png', 'WinImgNormalizer_variant.ico', 'WinImgNormalizer_poster.png'
    )
    foreach ($source in $Sources) {
        $relative = if ($source.relative_path) { $source.relative_path } else { $source.path }
        if ($relative -isnot [string] -or $source.sha256 -isnot [string]) { throw 'Evidence source path and hash must be scalar strings.' }
        if (($relative -notmatch '^(?:tests/|\.github/workflows/|WinImgNormalizer\.(?:ps1|bat)$)' -and $relative -cnotin $releaseSources) -or
            $relative -match '(?:^|/)\.\.(?:/|$)' -or $relative -match '[:\\\x00-\x1f]' -or $source.sha256 -notmatch '^[a-f0-9]{64}$') {
            throw 'Unsafe evidence source binding.'
        }
        [pscustomobject]@{ path = $relative; sha256 = [string]$source.sha256 }
    }
}
function Get-ExportStrings {
    param([object[]]$Values, [string]$Pattern)
    foreach ($value in $Values) {
        if ($value -isnot [string]) { throw 'Evidence text list elements must be scalar strings.' }
        if ($value -match $Pattern) { [string]$value }
    }
}
function Read-ExportSummary {
    param([string]$Directory, [string]$Name)
    $directoryPath = Assert-ExportPath $Directory
    $summaryPath = Assert-ExportPath (Join-Path $directoryPath $Name)
    if (-not (Test-Path -LiteralPath $summaryPath -PathType Leaf)) { return $null }
    if ((Get-Item -LiteralPath $summaryPath).Length -gt 2MB) { throw 'Evidence summary exceeds the bounded size.' }
    return [pscustomobject]@{ data = (Get-Content -LiteralPath $summaryPath -Raw -Encoding UTF8 | ConvertFrom-Json); sha256 = (Get-FileHash -LiteralPath $summaryPath -Algorithm SHA256).Hash.ToLowerInvariant() }
}
function Get-ExportCommon {
    param([object]$Summary)
    $fields = [ordered]@{}
    foreach ($key in @('tested_commit','powershell_version','powershell_edition','os_caption','os_version','os_build','architecture_bits','culture','execution_environment','runner_image','runner_image_version','runner_label','github_checkout_sha')) {
        $value = $Summary.$key
        if ($null -ne $value -and $value -isnot [string] -and $value -isnot [ValueType]) { throw 'Evidence environment fields must be scalar.' }
        if ($null -ne $value -and ([string]$value).Length -gt 160) { throw 'Evidence field exceeds the bounded size.' }
        if ($null -ne $value -and [string]$value -match '[:\\/\x00-\x1f]') { throw 'Evidence environment field contains an unsafe path or control.' }
        $fields[$key] = $value
    }
    if ($Summary.tested_commit -notmatch '^[a-f0-9]{40}$') { throw 'Evidence requires an exact tested Git commit.' }
    $fields.source_worktree_dirty = [bool]$Summary.source_worktree_dirty
    if ($Summary.observed_at_utc) {
        # Core's JSON reader materializes ISO timestamps as DateTime; casting it
        # back to string loses its zone and fractional seconds before parsing.
        if ($Summary.observed_at_utc -is [DateTime]) { $observed = $Summary.observed_at_utc }
        elseif ($Summary.observed_at_utc -is [string] -and $Summary.observed_at_utc -match '^[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}(?:\.[0-9]{1,7})?(?:Z|[+-][0-9]{2}:[0-9]{2})$') {
            $observed = [DateTime]::Parse($Summary.observed_at_utc, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind)
        }
        else { throw 'Evidence observation time must be a scalar ISO timestamp.' }
        $fields.observed_at_utc = $observed.ToUniversalTime().ToString('o')
    }
    $fields.source_sha256 = @(Get-ExportSourceBindings $Summary.source_sha256)
    return $fields
}
function Get-ExportVersion {
    param([string]$Value)
    if ($Value -and $Value -notmatch '^[0-9]+(?:\.[0-9]+){1,3}(?:-[0-9]+)?$') { throw 'Unsafe exported tool version.' }
    return $Value
}
function Get-ExportHash {
    param([string]$Value)
    if ($Value -and $Value -notmatch '^[a-f0-9]{64}$') { throw 'Unsafe exported dependency hash.' }
    return $Value
}
$output = Assert-ExportPath $OutputDirectory
if (Test-Path -LiteralPath $output) { throw 'Evidence export directory must be new.' }
$testExports = @(
    foreach ($directory in $ResultDirectories) {
        $record = Read-ExportSummary $directory 'summary.json'
        if (-not $record) { [ordered]@{ status = 'missing_summary' }; continue }
        $summary = $record.data
        $export = Get-ExportCommon $summary
        $export.summary_sha256 = $record.sha256
        $export.status = if ($summary.infrastructure_error) { 'infrastructure_failed' } elseif ($summary.exit_code -eq 0) { 'passed' } else { 'failed' }
        foreach ($key in @('total_count','passed_count','failed_count','skipped_count','inconclusive_count','not_run_count','failed_blocks_count','failed_containers_count','exit_code')) { $export[$key] = [int]$summary.$key }
        $export.deliberate_failure = [bool]$summary.deliberate_failure
        $export.suite_scope = if ($summary.suite_scope -eq 'mandatory') { 'mandatory' } else { 'focused' }
        $export.passed_case_ids = @(Get-ExportStrings $summary.passed_case_ids '^T[0-9]{3}$')
        $export.required_case_ids = @(Get-ExportStrings $summary.required_case_ids '^T[0-9]{3}$')
        $export.gate_failure_count = @($summary.gate_failures).Count
        $export.source_unchanged = [bool]$summary.source_unchanged
        $export.commit_unchanged = [bool]$summary.commit_unchanged
        $export.source_worktree_dirty_after = [bool]$summary.source_worktree_dirty_after
        $export.persistent_policy_unchanged = [bool]$summary.persistent_policy_unchanged
        $export.pester_version = Get-ExportVersion $summary.pester_version
        $export.pester_package_sha256 = Get-ExportHash $summary.pester_package_sha256
        $export.imagemagick_executable_sha256 = Get-ExportHash $summary.imagemagick_executable_sha256
        $export.imagemagick_version = Get-ExportVersion $summary.imagemagick_version
        $export.imagemagick_delegates = @(Get-ExportStrings $summary.imagemagick_delegates '^[A-Za-z0-9_-]+$')
        $export.codec_capabilities = @($summary.codec_capabilities | Where-Object { $null -ne $_ } | ForEach-Object {
            if ($_.format -isnot [string]) { throw 'Evidence codec format must be a scalar string.' }
            if ($_.format -notmatch '^(JPEG|PNG|BMP|TIFF|GIF|WEBP|HEIC|HEIF)$') { throw 'Unrecognized exported codec.' }
            [pscustomobject]@{ format = $_.format; read = [bool]$_.read; write = [bool]$_.write }
        })
        $export.codec_coverage = @($summary.codec_coverage | Where-Object { $null -ne $_ } | ForEach-Object {
            if ($_.extension -isnot [string] -or $_.case_id -isnot [string]) { throw 'Evidence codec extension and case ID must be scalar strings.' }
            if ($_.extension -notin @('jpg','jpeg','png','bmp','tif','tiff','gif','webp','heic','heif') -or $_.case_id -notin @('T031','T065')) { throw 'Unrecognized codec coverage binding.' }
            [pscustomobject]@{ extension=$_.extension; case_id=$_.case_id; passed_test_count=[int]$_.passed_test_count; executed=[bool]$_.executed }
        })
        $export
    }
)
$analysisExports = @(
    foreach ($directory in $StaticAnalysisDirectories) {
        $record = Read-ExportSummary $directory 'static-analysis-summary.json'
        if (-not $record) { [ordered]@{ status = 'missing_summary' }; continue }
        $summary = $record.data
        $export = Get-ExportCommon $summary
        $export.summary_sha256 = $record.sha256
        $export.status = if ($summary.result -in @('passed','analysis_failed','infrastructure_failed')) { $summary.result } else { 'invalid_result' }
        foreach ($key in @('exit_code','scanned_file_count','scoped_findings_count','control_findings_count','findings_count')) { $export[$key] = [int]$summary.$key }
        $export.deliberate_failure = [bool]$summary.deliberate_failure
        $export.control_detected = [bool]$summary.control_detected
        $export.source_bindings_unchanged = [bool]$summary.source_bindings_unchanged
        $export.control_source_sha256 = Get-ExportHash $summary.control_source_sha256
        $export.analyzer_version = Get-ExportVersion $summary.analyzer_version
        $export.analyzer_package_sha256 = Get-ExportHash $summary.analyzer_package_sha256
        $export.rules = @(Get-ExportStrings $summary.rules '^PS[A-Za-z]+$')
        $export
    }
)
[IO.Directory]::CreateDirectory($output) | Out-Null
foreach ($pair in @(@{ Name='test-evidence.json'; Records=$testExports }, @{ Name='static-analysis-evidence.json'; Records=$analysisExports })) {
    $file = Join-Path $output $pair.Name
    [ordered]@{ schema_version=1; scope='Synthetic allowlisted evidence only. Raw diagnostics, test names, absolute paths, media and XML are excluded.'; runs=@($pair.Records) } | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $file -Encoding UTF8
}
Write-Host 'Exported two sanitized evidence JSON files.'
