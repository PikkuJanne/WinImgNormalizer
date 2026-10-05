BeforeAll {
    $repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    $scratch = Join-Path $repository '.scratch'
    $hostExecutable = (Get-Process -Id $PID).Path
    $runner = Join-Path $PSScriptRoot 'Invoke-StaticAnalysis.ps1'
    $initializer = Join-Path $PSScriptRoot 'Initialize-TestDependencies.ps1'
    $pin = (Get-Content -LiteralPath (Join-Path $PSScriptRoot 'dependencies.json') -Raw | ConvertFrom-Json).psscriptanalyzer
    $archive = Join-Path $scratch ('tools/PSScriptAnalyzer-' + $pin.version + '/PSScriptAnalyzer.' + $pin.version + '.nupkg')
    $current = [IO.DirectoryInfo]::new($scratch)
    while ($null -ne $current) {
        if ($current.Exists -and (($current.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0)) { throw 'Analyzer fixture directory crosses a reparse point.' }
        $current = $current.Parent
    }
    & git -C $repository check-ignore --quiet --no-index -- (Join-Path $scratch 'analyzer-ignore-probe')
    if ($LASTEXITCODE -ne 0) { throw 'Analyzer fixtures must already be ignored.' }
    $owned = Join-Path $scratch ('M3-T06-analyzer-' + [guid]::NewGuid().ToString('N'))
    if (Test-Path -LiteralPath $owned) { throw 'Analyzer fixture directory already exists.' }
    [IO.Directory]::CreateDirectory($owned) | Out-Null
    $verified = & $initializer -StaticAnalysisOnly
    Import-Module $verified.PSScriptAnalyzerManifest -Force
    $settings = Join-Path $PSScriptRoot 'PSScriptAnalyzerSettings.psd1'

    function Invoke-AnalyzerChild {
        param([switch]$Control)
        $directory = Join-Path $owned ([guid]::NewGuid().ToString('N'))
        $arguments = @('-NoProfile', '-File', $runner, '-ResultDirectory', $directory)
        if ($Control) { $arguments += '-DeliberateFailure' }
        & $hostExecutable @arguments *> ($directory + '.console.log')
        $nativeExit = $LASTEXITCODE
        $summaryFile = Join-Path $directory 'static-analysis-summary.json'
        [pscustomobject]@{
            ExitCode = $nativeExit
            Summary = Get-Content -LiteralPath $summaryFile -Raw | ConvertFrom-Json
            Json = Get-Content -LiteralPath $summaryFile -Raw
        }
    }
    function New-SyntheticDependencyRoot {
        $fixtureRoot = Join-Path $owned ('dependency-' + [guid]::NewGuid().ToString('N'))
        $fixtureTests = Join-Path $fixtureRoot 'tests'
        [IO.Directory]::CreateDirectory($fixtureTests) | Out-Null
        [IO.File]::Copy($initializer, (Join-Path $fixtureTests 'Initialize-TestDependencies.ps1'))
        [IO.File]::Copy((Join-Path $PSScriptRoot 'dependencies.json'), (Join-Path $fixtureTests 'dependencies.json'))
        $fixtureArchive = Join-Path $fixtureRoot ('.scratch/tools/PSScriptAnalyzer-' + $pin.version + '/PSScriptAnalyzer.' + $pin.version + '.nupkg')
        [IO.Directory]::CreateDirectory((Split-Path -Parent $fixtureArchive)) | Out-Null
        [pscustomobject]@{ Root = $fixtureRoot; Initializer = Join-Path $fixtureTests 'Initialize-TestDependencies.ps1'; Archive = $fixtureArchive; Pin = Join-Path $fixtureTests 'dependencies.json' }
    }
}

Describe 'T066 scoped analyzer security and honest native failure controls' {
    It 'loads reviewed package bytes and reextracts instead of trusting a modified old module' {
        (Get-FileHash -LiteralPath $archive).Hash.ToLowerInvariant() | Should -Be $pin.sha256
        $hasher = [Security.Cryptography.SHA512]::Create()
        try { [Convert]::ToBase64String($hasher.ComputeHash([IO.File]::ReadAllBytes($archive))) | Should -Be $pin.sha512_base64 }
        finally { $hasher.Dispose() }
        [IO.File]::AppendAllText($verified.PSScriptAnalyzerManifest, '# owned deliberate cache tamper')
        $fresh = & $initializer -StaticAnalysisOnly
        $fresh.PSScriptAnalyzerManifest | Should -Not -Be $verified.PSScriptAnalyzerManifest
        (Get-FileHash -LiteralPath $fresh.PSScriptAnalyzerManifest).Hash.ToLowerInvariant() | Should -Be $pin.manifest_sha256
        $fresh.PSScriptAnalyzerVersion | Should -Be $pin.version
    }

    It 'rejects a corrupted pinned archive before importing or extracting its module' {
        $fixture = New-SyntheticDependencyRoot
        [IO.File]::Copy($archive, $fixture.Archive)
        $stream = [IO.File]::Open($fixture.Archive, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        try {
            $null = $stream.Seek(-1, [IO.SeekOrigin]::End)
            $originalByte = $stream.ReadByte()
            $null = $stream.Seek(-1, [IO.SeekOrigin]::End)
            $stream.WriteByte([byte]($originalByte -bxor 1))
        }
        finally { $stream.Dispose() }
        { & $fixture.Initializer -StaticAnalysisOnly } | Should -Throw '*Dependency SHA-256 mismatch*'
        @(Get-ChildItem -LiteralPath (Join-Path $fixture.Root '.scratch/test-dependencies') -Recurse -Filter PSScriptAnalyzer.psd1 -ErrorAction SilentlyContinue).Count | Should -Be 0
    }

    It 'rejects archive traversal even when synthetic package bytes match a test pin' {
        $fixture = New-SyntheticDependencyRoot
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [IO.Compression.ZipFile]::Open($fixture.Archive, [IO.Compression.ZipArchiveMode]::Create)
        try {
            $entry = $zip.CreateEntry('../escape.ps1')
            $writer = [IO.StreamWriter]::new($entry.Open())
            try { $writer.Write('# owned synthetic traversal sentinel') } finally { $writer.Dispose() }
        }
        finally { $zip.Dispose() }
        $testPin = Get-Content -LiteralPath $fixture.Pin -Raw | ConvertFrom-Json
        $testPin.psscriptanalyzer.sha256 = (Get-FileHash -LiteralPath $fixture.Archive).Hash.ToLowerInvariant()
        $testPin.psscriptanalyzer.bytes = (Get-Item -LiteralPath $fixture.Archive).Length
        $testPin | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $fixture.Pin -Encoding UTF8
        { & $fixture.Initializer -StaticAnalysisOnly } | Should -Throw '*unsafe or duplicate archive entry*'
        @(Get-ChildItem -LiteralPath $fixture.Root -Recurse -Filter escape.ps1 -ErrorAction SilentlyContinue).Count | Should -Be 0
    }

    It 'detects unsafe evaluation and Windows PowerShell 5.1 incompatible syntax as real diagnostics' {
        $unsafe = @(PSScriptAnalyzer\Invoke-ScriptAnalyzer -ScriptDefinition "Invoke-Expression 'Write-Output owned-control'" -Settings $settings)
        $unsafe.Count | Should -Be 1
        $unsafe[0].RuleName | Should -Be 'PSAvoidUsingInvokeExpression'
        # Core parses this syntax; Desktop reports its actual parse error. Either
        # is a failing diagnostic, so 5.1-incompatible syntax cannot pass silently.
        $incompatible = @(PSScriptAnalyzer\Invoke-ScriptAnalyzer -ScriptDefinition '$value = $null ?? 1' -Settings $settings)
        $incompatible.Count | Should -BeGreaterThan 0
        @($incompatible | Where-Object { $_.RuleName -eq 'PSUseCompatibleSyntax' -or $_.Severity.ToString() -eq 'ParseError' }).Count | Should -BeGreaterThan 0
    }

    It 'passes the real scoped application gate with exact sanitized source and environment bindings' {
        $observed = Invoke-AnalyzerChild
        $observed.ExitCode | Should -Be 0
        $observed.Summary.result | Should -Be 'passed'
        $observed.Summary.scanned_file_count | Should -BeGreaterThan 0
        $observed.Summary.findings_count | Should -Be 0
        $observed.Summary.analyzer_version | Should -Be $pin.version
        $observed.Summary.analyzer_package_sha256 | Should -Be $pin.sha256
        $observed.Summary.tested_commit | Should -Be (& git -C $repository rev-parse HEAD)
        $observed.Summary.powershell_version | Should -Be $PSVersionTable.PSVersion.ToString()
        $observed.Summary.source_bindings_unchanged | Should -BeTrue
        $boundApplication = @($observed.Summary.source_sha256 | Where-Object path -eq WinImgNormalizer.ps1)
        $boundApplication.Count | Should -Be 1
        $boundApplication[0].sha256 | Should -Be (Get-FileHash -LiteralPath (Join-Path $repository 'WinImgNormalizer.ps1')).Hash.ToLowerInvariant()
        foreach ($binding in $observed.Summary.source_sha256) {
            [IO.Path]::IsPathRooted($binding.path) | Should -BeFalse
            $binding.path | Should -Not -Match '(^|/)\.\.(/|$)'
            $binding.sha256 | Should -Match '^[a-f0-9]{64}$'
        }
        $observed.Json | Should -Not -BeLike ('*' + $repository + '*')
        $observed.Json | Should -Not -BeLike ('*' + $env:USERPROFILE + '*')
    }

    It 'returns native 1 for exactly one real deliberate finding and preserves the clean application scope' {
        $observed = Invoke-AnalyzerChild -Control
        $observed.ExitCode | Should -Be 1
        $observed.Summary.exit_code | Should -Be 1
        $observed.Summary.result | Should -Be 'analysis_failed'
        $observed.Summary.deliberate_failure | Should -BeTrue
        $observed.Summary.control_detected | Should -BeTrue
        $observed.Summary.scoped_findings_count | Should -Be 0
        $observed.Summary.control_findings_count | Should -Be 1
        $observed.Summary.findings_count | Should -Be 1
        $observed.Summary.findings[0].source | Should -Be 'control/unsafe-expression.ps1'
        $observed.Summary.findings[0].rule | Should -Be 'PSAvoidUsingInvokeExpression'
        $observed.Summary.findings[0].severity | Should -Be 'Warning'
        $observed.Summary.findings[0].line | Should -Be 1
        $observed.Summary.findings[0].column | Should -Be 1
        $observed.Summary.control_source_sha256 | Should -Match '^[a-f0-9]{64}$'
        $observed.Summary.PSObject.Properties.Name | Should -Not -Contain 'infrastructure_error'
        $observed.Json | Should -Not -Match 'owned-analyzer-control|ScriptDefinition|Exception|Message'
    }
}
