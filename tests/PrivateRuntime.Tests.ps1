BeforeAll {
    $repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    . (Join-Path $repository 'WinImgNormalizer.ps1')
    $scratch = Join-Path $repository '.scratch'
    Assert-WinImgNoReparseAncestors $scratch
    $owned = Join-Path $scratch ('private-runtime-tests-' + [guid]::NewGuid().ToString('N'))
    if (-not (New-WinImgExclusiveDirectory $owned)) { throw 'Private runtime fixture collision.' }
    $payload = Join-Path $repository 'resources/private-runtime.zip'
    $originalLocalData = $env:LOCALAPPDATA
    function New-PrivateRuntimeCase([switch]$WithPayload) {
        $root = Join-Path $owned ([guid]::NewGuid().ToString('N'))
        [IO.Directory]::CreateDirectory((Join-Path $root 'app')) | Out-Null
        if ($WithPayload) {
            [IO.Directory]::CreateDirectory((Join-Path $root 'app/resources')) | Out-Null
            [IO.File]::Copy($payload, (Join-Path $root 'app/resources/private-runtime.zip'), $false)
        }
        $env:LOCALAPPDATA = Join-Path $root 'profile-data'
        return [pscustomobject]@{ Root=$root; App=(Join-Path $root 'app'); Profile=$env:LOCALAPPDATA }
    }
    function Get-PrivateRuntimeCaseCache($Case) {
        return Join-Path $Case.Profile ('WinImgNormalizer/runtimes/' + (Get-WinImgVersion) + '-' + (Get-WinImgPrivateRuntimePin).PayloadSha256)
    }
    function New-PrivateMalformedArchive($Case, [string]$Relative) {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $manifest = [ordered]@{schema_version=1;engine_relative_path='imagemagick/magick.exe';files=@(@{path=$Relative;bytes=2;sha256=('0' * 64)})}
        $manifestBytes = [Text.Encoding]::UTF8.GetBytes(($manifest | ConvertTo-Json -Depth 5 -Compress))
        $memory = [IO.MemoryStream]::new($manifestBytes, $false)
        try { $manifestHash = Get-WinImgRuntimeStreamHash $memory } finally { $memory.Dispose() }
        $path = Join-Path $Case.App 'resources/private-runtime.zip'
        [IO.Directory]::CreateDirectory((Join-Path $Case.App 'resources')) | Out-Null
        $zip = [IO.Compression.ZipFile]::Open($path, [IO.Compression.ZipArchiveMode]::Create)
        try {
            $entry = $zip.CreateEntry('runtime/runtime-manifest.json')
            $stream = $entry.Open()
            try { $stream.Write($manifestBytes, 0, $manifestBytes.Length) } finally { $stream.Dispose() }
        } finally { $zip.Dispose() }
        return [pscustomobject]@{PayloadBytes=(Get-Item -LiteralPath $path).Length;PayloadSha256=(Get-FileHash -LiteralPath $path).Hash.ToLowerInvariant();ManifestSha256=$manifestHash}
    }
}

Describe 'T075 private runtime distribution and cache boundary' {
    AfterEach { $env:LOCALAPPDATA = $originalLocalData }
    It 'authenticates the tracked payload and cold extraction without an external decoder' {
        $case = New-PrivateRuntimeCase -WithPayload
        Mock Resolve-WinImgMagickApplication { throw 'PATH must not be resolved.' }
        Mock Invoke-WinImgNativeProcess { throw 'Authentication must not execute an engine.' }
        $selection = Get-WinImgPrivateRuntime -ApplicationRoot $case.App
        try {
            $selection.Source | Should -Be 'cache'
            $selection.Path | Should -Be (Join-Path (Get-PrivateRuntimeCaseCache $case) 'runtime/imagemagick/magick.exe')
            $selection.ManifestSha256 | Should -Be (Get-WinImgPrivateRuntimePin).ManifestSha256
            (Get-FileHash -LiteralPath $selection.Path).Hash.ToLowerInvariant() | Should -Be 'a8f31207c24880b546cce7bb5bce7bac14c0c8dc15b76ce3a376e4f4f05bf440'
            Should -Invoke Resolve-WinImgMagickApplication -Times 0 -Exactly
            Should -Invoke Invoke-WinImgNativeProcess -Times 0 -Exactly
        } finally { Remove-WinImgEngineSelection $selection }
    }

    It 'keeps authenticated engine and configuration bytes locked until the run ends' {
        $case = New-PrivateRuntimeCase -WithPayload
        $selection = Get-WinImgPrivateRuntime -ApplicationRoot $case.App
        try {
            { [IO.File]::AppendAllText($selection.Path, 'tamper') } | Should -Throw
            { [IO.File]::AppendAllText((Join-Path (Split-Path $selection.Path -Parent) 'policy.xml'), 'tamper') } | Should -Throw
        } finally { Remove-WinImgEngineSelection $selection }
        { [IO.File]::AppendAllText($selection.Path, 'after-disposal') } | Should -Not -Throw
    }

    It 'reuses an authenticated version-matched cache from a bare app copy' {
        $case = New-PrivateRuntimeCase -WithPayload
        $selection = Get-WinImgPrivateRuntime -ApplicationRoot $case.App
        $expected = $selection.Path
        Remove-WinImgEngineSelection $selection
        $bare = Join-Path $case.Root 'bare app'
        [IO.Directory]::CreateDirectory($bare) | Out-Null
        $selection = Get-WinImgPrivateRuntime -ApplicationRoot $bare
        try { $selection.Source | Should -Be 'cache'; $selection.Path | Should -Be $expected }
        finally { Remove-WinImgEngineSelection $selection }
    }

    It 'truthfully rejects a cold bare app copy before resolving any external engine' {
        $case = New-PrivateRuntimeCase
        Mock Resolve-WinImgMagickApplication { throw 'An incomplete app must not try PATH.' }
        Mock Invoke-WinImgNativeProcess { throw 'An incomplete app must not execute.' }
        { Get-WinImgPrivateRuntime -ApplicationRoot $case.App } | Should -Throw ('*Incomplete WinImgNormalizer app copy*' + $case.App + '*WinImgNormalizer-' + (Get-WinImgVersion) + '-portable.zip*')
        Should -Invoke Resolve-WinImgMagickApplication -Times 0 -Exactly
        Should -Invoke Invoke-WinImgNativeProcess -Times 0 -Exactly
    }

    It 'rejects a corrupt source payload even when a healthy cache exists' {
        $case = New-PrivateRuntimeCase -WithPayload
        $selection = Get-WinImgPrivateRuntime -ApplicationRoot $case.App
        Remove-WinImgEngineSelection $selection
        [IO.File]::AppendAllText((Join-Path $case.App 'resources/private-runtime.zip'), 'tamper')
        Mock Invoke-WinImgNativeProcess { throw 'Corrupt source payload must not execute.' }
        { Get-WinImgPrivateRuntime -ApplicationRoot $case.App } | Should -Throw '*payload integrity failure*'
        Should -Invoke Invoke-WinImgNativeProcess -Times 0 -Exactly
    }

    It 'rejects a missing recognized source payload even when a healthy cache exists' {
        $case = New-PrivateRuntimeCase -WithPayload
        $selection = Get-WinImgPrivateRuntime -ApplicationRoot $case.App
        Remove-WinImgEngineSelection $selection
        [IO.File]::WriteAllText((Join-Path $case.App 'resources/private-runtime.json'), '{}')
        [IO.File]::Delete((Join-Path $case.App 'resources/private-runtime.zip'))
        { Get-WinImgPrivateRuntime -ApplicationRoot $case.App } | Should -Throw '*payload is missing*'
    }

    It 'fails closed on a corrupt private cache without replacing it from healthy source bytes' {
        $case = New-PrivateRuntimeCase -WithPayload
        $selection = Get-WinImgPrivateRuntime -ApplicationRoot $case.App
        Remove-WinImgEngineSelection $selection
        [IO.File]::AppendAllText($selection.Path, 'tamper')
        { Get-WinImgPrivateRuntime -ApplicationRoot $case.App } | Should -Throw '*Runtime file integrity failure*'
        (Get-Item -LiteralPath $selection.Path).Length | Should -Be (32395952 + 6)
    }

    It 'rejects missing source bytes identified by <Marker> with warm-cache=<Warm>' -ForEach @(
        @{Marker='resources';Warm=$false}, @{Marker='.git';Warm=$false}, @{Marker='release-metadata.json';Warm=$false},
        @{Marker='resources';Warm=$true}, @{Marker='.git';Warm=$true}, @{Marker='release-metadata.json';Warm=$true}
    ) {
        $case = New-PrivateRuntimeCase -WithPayload:$Warm
        if ($Warm) {
            $selection = Get-WinImgPrivateRuntime -ApplicationRoot $case.App
            Remove-WinImgEngineSelection $selection
            [IO.File]::Delete((Join-Path $case.App 'resources/private-runtime.zip'))
            [IO.Directory]::Delete((Join-Path $case.App 'resources'), $false)
        }
        if ($Marker -eq 'resources' -or $Marker -eq '.git') { [IO.Directory]::CreateDirectory((Join-Path $case.App $Marker)) | Out-Null }
        else { [IO.File]::WriteAllText((Join-Path $case.App $Marker), '{}') }
        { Get-WinImgPrivateRuntime -ApplicationRoot $case.App } | Should -Throw '*payload is missing from this source copy*'
    }

    It 'pins the manifest independently of a modified marker and matching modified manifest' {
        $case = New-PrivateRuntimeCase -WithPayload
        $selection = Get-WinImgPrivateRuntime -ApplicationRoot $case.App
        Remove-WinImgEngineSelection $selection
        $cache = Get-PrivateRuntimeCaseCache $case
        $manifest = Join-Path $cache 'runtime/runtime-manifest.json'
        [IO.File]::AppendAllText($manifest, ' ')
        $marker = @{runtime_required=$true;runtime_manifest_sha256=(Get-FileHash -LiteralPath $manifest).Hash.ToLowerInvariant()}
        [IO.File]::WriteAllText((Join-Path $cache 'portable-runtime.json'), ($marker | ConvertTo-Json))
        { Get-WinImgPrivateRuntime -ApplicationRoot $case.App } | Should -Throw '*pinned private engine*'
    }

    It 'refuses a linked cache ancestor before extraction or execution' {
        $case = New-PrivateRuntimeCase -WithPayload
        $elsewhere = Join-Path $case.Root 'untrusted sibling'
        [IO.Directory]::CreateDirectory($elsewhere) | Out-Null
        [IO.Directory]::CreateDirectory($case.Profile) | Out-Null
        New-Item -ItemType Junction -Path (Join-Path $case.Profile 'WinImgNormalizer') -Target $elsewhere -ErrorAction Stop | Out-Null
        { Get-WinImgPrivateRuntime -ApplicationRoot $case.App } | Should -Throw '*linked path*'
        @(Get-ChildItem -LiteralPath $elsewhere -Force).Count | Should -Be 0
    }

    It 'refuses an authenticated archive manifest with unsafe path <Relative>' -ForEach @(
        @{Relative='../escape.exe'}, @{Relative='a/../../escape.exe'}, @{Relative='C:/escape.exe'},
        @{Relative='a\\escape.exe'}, @{Relative='a/file:stream'}, @{Relative='a./escape.exe'}
    ) {
        $case = New-PrivateRuntimeCase
        $malformedPin = New-PrivateMalformedArchive $case $Relative
        Mock Get-WinImgPrivateRuntimePin { return $malformedPin }
        { Get-WinImgPrivateRuntime -ApplicationRoot $case.App } | Should -Throw '*Unsafe or duplicate private runtime archive manifest path*'
        Test-Path -LiteralPath $case.Profile | Should -BeFalse
    }

    It 'publishes one complete cache when two fresh Windows PowerShell 5.1 processes start concurrently' {
        $case = New-PrivateRuntimeCase -WithPayload
        $application = Join-Path $repository 'WinImgNormalizer.ps1'
        $encodedApplication = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($application))
        $encodedAppRoot = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($case.App))
        $command = @'
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
. ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('__APPLICATION__')))
$s=Get-WinImgPrivateRuntime -ApplicationRoot ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('__ROOT__')))
try { [Console]::WriteLine($s.Path); [Console]::WriteLine((Get-FileHash -LiteralPath $s.Path).Hash.ToLowerInvariant()) }
finally { Remove-WinImgEngineSelection $s }
'@
        $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command.Replace('__APPLICATION__',$encodedApplication).Replace('__ROOT__',$encodedAppRoot)))
        $children = @()
        try {
            foreach ($index in 1..2) {
                $start = New-Object Diagnostics.ProcessStartInfo
                # Always launch untouched production 5.1 children. Pester's
                # dependency bootstrap already loads compression assemblies in
                # this test host and could otherwise mask a cold-start failure.
                $start.FileName = Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe'
                $start.Arguments = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand ' + $encoded
                $start.UseShellExecute=$false; $start.CreateNoWindow=$true
                $start.RedirectStandardOutput=$true; $start.RedirectStandardError=$true
                foreach ($key in @($start.EnvironmentVariables.Keys)) {
                    if ([string]::Equals([string]$key,'PSModulePath',[StringComparison]::OrdinalIgnoreCase)) { $start.EnvironmentVariables.Remove([string]$key) }
                }
                $start.EnvironmentVariables['LOCALAPPDATA']=$case.Profile
                $child = New-Object Diagnostics.Process
                $child.StartInfo=$start
                $null=$child.Start()
                $children += [pscustomobject]@{Process=$child;Out=$child.StandardOutput.ReadToEndAsync();Err=$child.StandardError.ReadToEndAsync()}
            }
            foreach ($child in $children) {
                $child.Process.WaitForExit(30000) | Should -BeTrue
                $child.Out.Wait(5000) | Should -BeTrue; $child.Err.Wait(5000) | Should -BeTrue
                $child.Process.ExitCode | Should -Be 0 -Because ($child.Err.GetAwaiter().GetResult())
                $child.Err.GetAwaiter().GetResult() | Should -BeNullOrEmpty
                $child.Out.GetAwaiter().GetResult() | Should -Match 'a8f31207c24880b546cce7bb5bce7bac14c0c8dc15b76ce3a376e4f4f05bf440'
            }
            @(Get-ChildItem -LiteralPath (Join-Path $case.Profile 'WinImgNormalizer/runtimes') -Directory | Where-Object { $_.Name -eq ((Get-WinImgVersion) + '-' + (Get-WinImgPrivateRuntimePin).PayloadSha256) }).Count | Should -Be 1
        } finally {
            foreach ($child in $children) { if (-not $child.Process.HasExited) { $child.Process.Kill(); $null=$child.Process.WaitForExit(5000) }; $child.Process.Dispose() }
        }
    }
}
