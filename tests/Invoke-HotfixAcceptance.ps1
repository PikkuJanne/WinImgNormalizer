# Real portable-byte acceptance. Runner mutations are forbidden on a desktop.
[CmdletBinding()]
param(
    [string]$ArchivePath,
    [string]$ResultDirectory = '.scratch/hotfix-acceptance-result',
    [string]$LegacyMagickPath,
    [string]$BaselineZipPath,
    [switch]$DownloadDependencies,
    [switch]$Build,
    [switch]$DisposableRunner,
    [switch]$Worker,
    [switch]$InstrumentedLocal,
    [string]$CandidateTag,
    [string]$PublicTag
)
$ErrorActionPreference = 'Stop'
# A PS7 parent can pass its incompatible module directories to powershell.exe.
# Reset only this worker process before its first module-backed cmdlet executes.
function Initialize-HotfixWorkerModulePath {
    $profileRoot = [Environment]::GetFolderPath('UserProfile')
    $programFilesRoot = [Environment]::GetFolderPath('ProgramFiles')
    $windowsRoot = [Environment]::GetFolderPath('Windows')
    $trustedPaths = @(
        [IO.Path]::Combine($PSHOME, 'Modules'),
        [IO.Path]::Combine($programFilesRoot, 'WindowsPowerShell\Modules'),
        [IO.Path]::Combine($windowsRoot, 'System32\WindowsPowerShell\v1.0\Modules'),
        [IO.Path]::Combine($profileRoot, 'Documents\WindowsPowerShell\Modules')
    )
    [Environment]::SetEnvironmentVariable('PSModulePath', [string]::Join(';', $trustedPaths), 'Process')
}
if ($Worker) { Initialize-HotfixWorkerModulePath }
$repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$utf8 = [Text.UTF8Encoding]::new($false)
function Assert-Hotfix([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
function ConvertTo-HotfixArgument([string]$Value) {
    return '"' + [regex]::Replace([regex]::Replace($Value, '(\\*)"', '$1$1\"'), '(\\+)$', '$1$1') + '"'
}
function Get-HotfixBoundedConsoleDiagnostic([string]$Text) {
    $safe = [regex]::Replace($Text, '[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', ' ')
    if ($safe.Length -le 1400) { return $safe }
    return $safe.Substring(0, 300) + "`n[console capture bounded]`n" + $safe.Substring($safe.Length - 1000)
}
function Invoke-HotfixProcess([string]$Executable, [string[]]$Tokens, [string]$Work, [hashtable]$Environment, [int]$Timeout = 120000, [switch]$InteractiveConsole) {
    $start = New-Object Diagnostics.ProcessStartInfo
    $start.FileName = $Executable
    $start.Arguments = ($Tokens | ForEach-Object { ConvertTo-HotfixArgument $_ }) -join ' '
    if ([IO.Path]::GetFileName($Executable) -ieq 'cmd.exe') { $start.Arguments = '/d /v:off /s /c ' + $Tokens[-1] }
    $start.WorkingDirectory = $Work
    $start.UseShellExecute = $false; $start.CreateNoWindow = $true
    $start.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
    $start.RedirectStandardOutput = $true; $start.RedirectStandardError = $true; $start.RedirectStandardInput = $true
    $start.StandardOutputEncoding = [Text.UTF8Encoding]::new($false)
    $start.StandardErrorEncoding = [Text.UTF8Encoding]::new($false)
    foreach ($key in @($start.EnvironmentVariables.Keys)) {
        if ([string]$key -ieq 'PSModulePath' -or [string]$key -like 'MAGICK_*') { $start.EnvironmentVariables.Remove([string]$key) }
    }
    if ($Environment) {
        foreach ($key in $Environment.Keys) {
            if ($null -eq $Environment[$key]) { $start.EnvironmentVariables.Remove([string]$key) }
            else { $start.EnvironmentVariables[$key] = $Environment[$key] }
        }
    }
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $start
    try {
        Assert-Hotfix ($process.Start()) 'Acceptance child failed to start.'
        $out = $process.StandardOutput.ReadToEndAsync(); $err = $process.StandardError.ReadToEndAsync()
        if ($InteractiveConsole) {
            $record = [string]$start.EnvironmentVariables['WINIMG_TEST_PRIVATE_CONSOLE_RECORD']
            Assert-Hotfix ([IO.Path]::IsPathRooted($record) -and $record.StartsWith($Work + '\', [StringComparison]::OrdinalIgnoreCase)) 'Private-console record must belong to acceptance scratch.'
            $clock = [Diagnostics.Stopwatch]::StartNew()
            while (-not [IO.File]::Exists($record + '.ready') -and -not $process.HasExited -and $clock.ElapsedMilliseconds -lt 15000) { Start-Sleep -Milliseconds 20 }
            Assert-Hotfix ([IO.File]::Exists($record + '.ready')) 'Private-console fixture did not become ready.'
            # The fixture injects only after the real pause label is visible and
            # its private console contains exactly the controller and owned CMD.
            $process.StandardInput.WriteLine('x'); $process.StandardInput.Flush()
        }
        $process.StandardInput.Close()
        if (-not $process.WaitForExit($Timeout)) { $process.Kill(); throw 'Acceptance child exceeded its deadline.' }
        Assert-Hotfix ($out.Wait(5000) -and $err.Wait(5000)) 'Acceptance child streams incomplete.'
        $console = $null
        if ($InteractiveConsole) {
            $console = Get-Content -LiteralPath $record -Raw -Encoding UTF8 | ConvertFrom-Json
            Assert-Hotfix (-not $console.HarnessError -and $console.PrivateConsoleAllocated -and $console.PrivateConsoleOutput -and $console.PrivateJobAssignedBeforeCmd -and $console.KeyRecordsWritten -eq 2 -and @($console.MembersBeforeKey).Count -eq 2 -and $console.ExitCode -eq $process.ExitCode) 'Actual private-console processing/pause evidence is invalid.'
        }
        return [pscustomobject]@{ Exit = $process.ExitCode; Out = $out.GetAwaiter().GetResult(); Error = $err.GetAwaiter().GetResult(); Console = $console }
    } finally { $process.Dispose() }
}
function Get-HotfixSourceState([string]$Directory) {
    return (@(Get-ChildItem -LiteralPath $Directory -Recurse -File | Sort-Object FullName | ForEach-Object {
        [ordered]@{ Name = $_.FullName.Substring($Directory.Length); Bytes = $_.Length; Sha256 = (Get-FileHash -LiteralPath $_.FullName).Hash; Modified = $_.LastWriteTimeUtc.Ticks }
    }) | ConvertTo-Json -Compress)
}
function Test-HotfixOutbound {
    $socket = New-Object Net.Sockets.TcpClient
    try {
        $attempt = $socket.BeginConnect('github.com', 443, $null, $null)
        if (-not $attempt.AsyncWaitHandle.WaitOne(5000)) { return $false }
        try { $socket.EndConnect($attempt); return $socket.Connected } catch { return $false }
    } finally { $socket.Close() }
}
function Test-HotfixPrivateOpenMp([string]$EnginePath, [string]$PackageRoot, [string]$Work) {
    $manifest = [IO.File]::ReadAllText((Join-Path $PackageRoot 'runtime/runtime-manifest.json')) | ConvertFrom-Json
    $records = @($manifest.files | Where-Object { [IO.Path]::GetFileName([string]$_.path) -ieq 'VCOMP140.DLL' })
    Assert-Hotfix ($records.Count -eq 1) 'Packaged Microsoft OpenMP runtime is absent or ambiguous.'
    $expected = [IO.Path]::GetFullPath((Join-Path $PackageRoot ('runtime/' + $records[0].path)))
    Assert-Hotfix ([IO.File]::Exists($expected) -and (Get-FileHash -LiteralPath $expected).Hash -ieq $records[0].sha256) 'Packaged OpenMP runtime differs from its integrity manifest.'
    $start = New-Object Diagnostics.ProcessStartInfo
    $start.FileName = $EnginePath
    $start.Arguments = '-size 1x1 -depth 8 RGB:- -format "%m|%w|%h" info:'
    $start.WorkingDirectory = Split-Path -Parent $EnginePath
    $start.UseShellExecute = $false; $start.CreateNoWindow = $true
    $start.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
    $start.RedirectStandardInput = $true; $start.RedirectStandardOutput = $true; $start.RedirectStandardError = $true
    $start.StandardOutputEncoding = [Text.UTF8Encoding]::new($false)
    $start.StandardErrorEncoding = [Text.UTF8Encoding]::new($false)
    foreach ($key in @($start.EnvironmentVariables.Keys)) {
        if ([string]$key -like 'MAGICK_*' -or [string]$key -ieq 'LIBHEIF_PLUGIN_PATH') { $start.EnvironmentVariables.Remove([string]$key) }
    }
    $start.EnvironmentVariables['PATH'] = $start.WorkingDirectory + ';' + (Join-Path $env:SystemRoot 'System32')
    $start.EnvironmentVariables['MAGICK_HOME'] = $start.WorkingDirectory
    $start.EnvironmentVariables['MAGICK_CONFIGURE_PATH'] = $start.WorkingDirectory
    $start.EnvironmentVariables['MAGICK_CODER_MODULE_PATH'] = $start.WorkingDirectory
    $start.EnvironmentVariables['MAGICK_FILTER_MODULE_PATH'] = $start.WorkingDirectory
    $start.EnvironmentVariables['LIBHEIF_PLUGIN_PATH'] = $start.WorkingDirectory
    $start.EnvironmentVariables['TEMP'] = $Work; $start.EnvironmentVariables['TMP'] = $Work
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $start
    $started = $false
    try {
        Assert-Hotfix ($process.Start()) 'Owned OpenMP loading probe did not start.'
        $started = $true
        $out = $process.StandardOutput.ReadToEndAsync(); $err = $process.StandardError.ReadToEndAsync()
        $clock = [Diagnostics.Stopwatch]::StartNew()
        $loaded = @()
        while (-not $process.HasExited -and $clock.ElapsedMilliseconds -lt 10000) {
            $process.Refresh()
            $loaded = @($process.Modules | Where-Object ModuleName -IEQ 'VCOMP140.DLL')
            if ($loaded.Count -eq 1) { break }
            Start-Sleep -Milliseconds 20
        }
        Assert-Hotfix ($loaded.Count -eq 1 -and -not $process.HasExited) 'Owned engine did not remain available for same-user module inspection while awaiting synthetic RGB input.'
        $actual = [IO.Path]::GetFullPath($loaded[0].FileName)
        Assert-Hotfix ([string]::Equals($actual, $expected, [StringComparison]::OrdinalIgnoreCase)) 'Engine loaded OpenMP from outside the verified private payload.'
        $actualHash = (Get-FileHash -LiteralPath $actual -Algorithm SHA256).Hash.ToLowerInvariant()
        Assert-Hotfix ($actualHash -eq $records[0].sha256) 'Actual loaded OpenMP module hash differs from the payload manifest.'
        $pixels = [byte[]](64,160,96)
        $process.StandardInput.BaseStream.Write($pixels, 0, $pixels.Length)
        $process.StandardInput.Close()
        Assert-Hotfix ($process.WaitForExit(15000)) 'Owned engine did not complete its tiny RGB decode.'
        Assert-Hotfix ($out.Wait(5000) -and $err.Wait(5000)) 'OpenMP loading probe streams incomplete.'
        $decoded = $out.GetAwaiter().GetResult()
        Assert-Hotfix ($process.ExitCode -eq 0 -and $decoded -eq 'RGB|1|1' -and -not $err.GetAwaiter().GetResult()) 'Owned engine failed the actual one-pixel RGB decode after module inspection.'
        return [ordered]@{ Module = 'VCOMP140.DLL'; RelativePath = 'runtime/' + $records[0].path; Sha256 = $actualHash; LoadedPathMatchesPrivatePayload = $true; SameUserOwnedProcess = $true; ObservedWhileAwaitingSyntheticInput = $true; InputRgbBytes = 3; Decoded = $decoded; Exit = $process.ExitCode }
    } finally {
        if ($started -and -not $process.HasExited) { $process.Kill(); $null = $process.WaitForExit(5000) }
        $process.Dispose()
    }
}
function Test-HotfixLegacyModuleEnvironment([string]$LegacyShell, [string]$ModernShell, [string]$Work) {
    $driver = Join-Path $Work 'verify-legacy-module-environment.ps1'
    $record = Join-Path $Work 'legacy-module-environment.json'
    $initializer = (Get-Command Initialize-HotfixWorkerModulePath).ScriptBlock.ToString()
    $text = @'
param([string]$Record)
$ErrorActionPreference = 'Stop'
$initialAvailable = $null -ne (Get-Command Get-FileHash -ErrorAction SilentlyContinue)
if ($initialAvailable) { throw 'The PS7-only inherited module-path control did not reproduce missing PS5 Get-FileHash.' }
function Initialize-HotfixWorkerModulePath {
'@ + "`r`n" + $initializer + "`r`n}" + @'

Initialize-HotfixWorkerModulePath
$hashCommand = Get-Command Get-FileHash -ErrorAction Stop
$cimCommand = Get-Command Get-CimInstance -ErrorAction Stop
$hash = Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256
$os = Get-CimInstance Win32_OperatingSystem
if (-not $hash.Hash -or -not $os.BuildNumber) { throw 'Restored built-in cmdlets did not execute.' }
$result = [ordered]@{
    Shell = $PSVersionTable.PSVersion.ToString(); Edition = $PSVersionTable.PSEdition
    InitialGetFileHashAvailable = $initialAvailable; RestoredGetFileHashAvailable = $true
    RestoredGetCimInstanceAvailable = $true; HashModule = $hashCommand.Source
    CimModule = $cimCommand.Source; ProcessOnlyEnvironmentReset = $true; Exit = 0
}
[IO.File]::WriteAllText($Record, ($result | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
exit 0
'@
    [IO.File]::WriteAllText($driver, $text, [Text.UTF8Encoding]::new($false))
    $ps7Modules = [IO.Path]::Combine([IO.Path]::GetDirectoryName($ModernShell), 'Modules')
    Assert-Hotfix ([IO.Directory]::Exists($ps7Modules)) 'Actual PS7-only module directory is unavailable for the regression control.'
    $result = Invoke-HotfixProcess $LegacyShell @('-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',$driver,$record) $Work @{ PSModulePath = $ps7Modules }
    Assert-Hotfix ($result.Exit -eq 0 -and [IO.File]::Exists($record)) ('Real PS5 module-environment regression failed: ' + $result.Error)
    $observation = [IO.File]::ReadAllText($record) | ConvertFrom-Json
    Assert-Hotfix (-not $observation.InitialGetFileHashAvailable -and $observation.RestoredGetFileHashAvailable -and $observation.RestoredGetCimInstanceAvailable -and $observation.Edition -eq 'Desktop') 'Legacy module-environment regression observations are invalid.'
    return $observation
}
Assert-Hotfix ($env:OS -eq 'Windows_NT') 'Real Windows is mandatory.'
$root = [IO.Path]::GetFullPath((Join-Path $repository $ResultDirectory))
$scratch = [IO.Path]::GetFullPath((Join-Path $repository '.scratch'))
Assert-Hotfix ($root.StartsWith($scratch + '\', [StringComparison]::OrdinalIgnoreCase)) 'Acceptance output must be in owned ignored scratch.'
$ancestor = [IO.DirectoryInfo]::new($root)
while ($ancestor) {
    Assert-Hotfix (-not ($ancestor.Exists -and ($ancestor.Attributes -band [IO.FileAttributes]::ReparsePoint))) 'Acceptance output crosses a reparse point.'
    $ancestor = $ancestor.Parent
}
[IO.Directory]::CreateDirectory($root) | Out-Null
$legacyShell = Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe'
$modernShell = (Get-Command pwsh.exe -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
$cmd = Join-Path $env:SystemRoot 'System32/cmd.exe'

if (-not $Worker) {
    if ($DownloadDependencies -or -not $LegacyMagickPath -or -not $BaselineZipPath) {
        $dependencies = & (Join-Path $PSScriptRoot 'Initialize-HotfixAcceptance.ps1') -Download:$DownloadDependencies
        $LegacyMagickPath = $dependencies.LegacyMagickPath
        $BaselineZipPath = $dependencies.BaselineZipPath
    }
    $candidate = $null
    if ($CandidateTag) {
        Assert-Hotfix ($CandidateTag -match '^v\d+\.\d+\.\d+$') 'Candidate tag must be a stable version.'
        $candidate = Get-Content -LiteralPath (Join-Path $repository '.scratch/ci-candidate/candidate.json') -Raw -Encoding UTF8 | ConvertFrom-Json
        $acceptedCandidateState = $candidate.Draft -or ($candidate.Prerelease -and -not $candidate.Draft -and $candidate.AnonymousPublicDownload) -or ($candidate.Stable -and $candidate.Latest)
        Assert-Hotfix ($candidate.Tag -eq $CandidateTag -and $acceptedCandidateState -and $candidate.SourceRevision -eq (& git -C $repository rev-parse HEAD).Trim()) 'Release acquisition is absent or belongs to another revision.'
        $ArchivePath = Join-Path $repository ('.scratch/ci-candidate/' + $candidate.AssetFilename)
        Assert-Hotfix ((Get-FileHash -LiteralPath $ArchivePath).Hash -ieq $candidate.Sha256) 'Candidate cache was changed after authenticated acquisition.'
    } elseif ($Build) {
        # Authenticate and stage only the existing verified cache here. Network
        # acquisition belongs to the preceding build-only workflow step.
        $null = & (Join-Path $repository 'tools/release/Stage-Runtime.ps1')
        $revision = (& git -C $repository rev-parse HEAD).Trim()
        $packageDirectory = Join-Path $root 'build'
        $built = & (Join-Path $repository 'tools/release/Build-Release.ps1') -Revision $revision -OutputDirectory $packageDirectory
        $ArchivePath = Join-Path $packageDirectory $built.asset_filename
    }
    $public = $null
    if ($PublicTag) {
        Assert-Hotfix ($PublicTag -match '^v\d+\.\d+\.\d+$') 'Only a stable tag can be verified.'
        $release = Invoke-RestMethod -Uri ('https://api.github.com/repos/PikkuJanne/WinImgNormalizer/releases/tags/' + $PublicTag)
        $latest = Invoke-RestMethod -Uri 'https://api.github.com/repos/PikkuJanne/WinImgNormalizer/releases/latest'
        Assert-Hotfix (-not $release.draft -and -not $release.prerelease -and $latest.tag_name -eq $PublicTag) 'Public release is not stable/latest.'
        $asset = @($release.assets | Where-Object name -eq ('WinImgNormalizer-' + $PublicTag.Substring(1) + '-portable.zip'))
        Assert-Hotfix ($asset.Count -eq 1) 'Public portable asset is absent or ambiguous.'
        $download = Join-Path $root 'public-download'
        [IO.Directory]::CreateDirectory($download) | Out-Null
        foreach ($name in @($asset[0].name, 'SHA256SUMS.txt', 'build-provenance.json')) {
            $remote = @($release.assets | Where-Object name -eq $name)
            Assert-Hotfix ($remote.Count -eq 1) ('Missing public asset: ' + $name)
            Invoke-WebRequest -UseBasicParsing -Uri $remote[0].browser_download_url -OutFile (Join-Path $download $name)
        }
        $downloaded = Join-Path $download $asset[0].name
        $digest = (Get-FileHash -LiteralPath $downloaded).Hash.ToLowerInvariant()
        $provenance = Get-Content -LiteralPath (Join-Path $download 'build-provenance.json') -Raw -Encoding UTF8 | ConvertFrom-Json
        $sums = Get-Content -LiteralPath (Join-Path $download 'SHA256SUMS.txt') -Raw -Encoding UTF8
        Assert-Hotfix ($sums -match ('(?m)^' + $digest + '  ' + [regex]::Escape($asset[0].name) + '\r?$')) 'Public checksum sidecar differs.'
        Assert-Hotfix ($provenance.asset_sha256 -eq $digest -and $provenance.version -eq $PublicTag.Substring(1)) 'Public provenance differs.'
        $provenanceDigest = (Get-FileHash -LiteralPath (Join-Path $download 'build-provenance.json')).Hash.ToLowerInvariant()
        Assert-Hotfix ($sums -match ('(?m)^' + $provenanceDigest + '  build-provenance\.json\r?$')) 'Public provenance checksum differs.'
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [IO.Compression.ZipFile]::OpenRead($downloaded)
        try {
            $reader = [IO.StreamReader]::new($zip.GetEntry('package-manifest.json').Open())
            try { $packageManifest = $reader.ReadToEnd() | ConvertFrom-Json } finally { $reader.Dispose() }
        } finally { $zip.Dispose() }
        Assert-Hotfix ($packageManifest.source_revision -eq $provenance.source_revision -and $packageManifest.version -eq $provenance.version) 'Published ZIP source/version differs from provenance.'
        if ($ArchivePath) { Assert-Hotfix ((Get-FileHash -LiteralPath $ArchivePath).Hash -ieq $digest) 'Published ZIP differs from the tested build.' }
        $tag = Invoke-RestMethod -Uri ('https://api.github.com/repos/PikkuJanne/WinImgNormalizer/git/ref/tags/' + $PublicTag)
        if ($tag.object.type -eq 'tag') { $tag = [pscustomobject]@{ object = (Invoke-RestMethod -Uri $tag.object.url).object } }
        Assert-Hotfix ($tag.object.sha -eq $provenance.source_revision) 'Public tag differs from provenance source.'
        $ArchivePath = $downloaded
        $public = [ordered]@{ Tag = $PublicTag; ReleaseUrl = $release.html_url; ZipUrl = $asset[0].browser_download_url; Sha256 = $digest; SourceRevision = $provenance.source_revision; Stable = $true; Latest = $true }
    }
    Assert-Hotfix ([IO.File]::Exists($ArchivePath)) 'A real portable archive is required.'
    if ($InstrumentedLocal) {
        & $PSCommandPath -Worker -InstrumentedLocal -ArchivePath $ArchivePath -ResultDirectory $ResultDirectory -LegacyMagickPath $LegacyMagickPath -BaselineZipPath $BaselineZipPath
        if ($public) { [IO.File]::WriteAllText((Join-Path $root 'public-release.json'), ($public | ConvertTo-Json), $utf8) }
        exit $LASTEXITCODE
    }
    # Account/firewall changes are restricted to an ephemeral GitHub-hosted VM.
    Assert-Hotfix ($DisposableRunner -and $env:GITHUB_ACTIONS -eq 'true' -and $env:RUNNER_ENVIRONMENT -eq 'github-hosted') 'Exact BAT/-File offline acceptance requires the disposable hosted runner; desktop settings are never changed.'
    $legacyModuleEnvironment = Test-HotfixLegacyModuleEnvironment $legacyShell $modernShell $root
    $userName = 'winimg' + [guid]::NewGuid().ToString('N').Substring(0, 12)
    $password = New-Object Security.SecureString
    $random = New-Object byte[] 32
    $rng = [Security.Cryptography.RandomNumberGenerator]::Create()
    try { $rng.GetBytes($random) } finally { $rng.Dispose() }
    foreach ($character in ('W!9x' + [Convert]::ToBase64String($random)).ToCharArray()) { $password.AppendChar($character) }
    $password.MakeReadOnly()
    $createdUser = $false
    $checkoutAcl = $null
    $rules = New-Object 'Collections.Generic.List[string]'
    try {
        Assert-Hotfix (Test-HotfixOutbound) 'Pre-isolation outbound credibility probe could not reach GitHub.'
        $user = New-LocalUser -Name $userName -Password $password -AccountNeverExpires -UserMayNotChangePassword
        $createdUser = $true
        $qualifiedUser = [Environment]::MachineName + '\' + $userName
        $usersMembers = @(Get-LocalGroupMember -SID 'S-1-5-32-545')
        if (@($usersMembers | Where-Object { $_.SID.Value -eq $user.SID.Value }).Count -eq 0) {
            Add-LocalGroupMember -SID 'S-1-5-32-545' -Member $qualifiedUser
        }
        # The clean checkout and all synthetic dependency siblings are owned by
        # this disposable VM. Grant only read/execute/traverse there; restore the
        # original checkout ACL in finally, including inherited propagation.
        $checkoutAcl = Get-Acl -LiteralPath $repository
        $readAcl = Get-Acl -LiteralPath $repository
        $readAcl.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new($user.SID, 'ReadAndExecute', 'ContainerInherit,ObjectInherit', 'None', 'Allow'))
        Set-Acl -LiteralPath $repository -AclObject $readAcl
        $acl = Get-Acl -LiteralPath $root
        $acl.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new($user.SID, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow'))
        Set-Acl -LiteralPath $root -AclObject $acl
        $programs = @($legacyShell, $modernShell, $cmd) + @(Get-ChildItem -LiteralPath (Split-Path -Parent $LegacyMagickPath) -Filter '*.exe' -File | ForEach-Object FullName)
        # Extract before filtering so every shipped delegate executable is denied.
        $firewallExtract = Join-Path $root ('portable ' + [char]0x00e5 + [char]0x96ea)
        Expand-Archive -LiteralPath $ArchivePath -DestinationPath $firewallExtract
        $programs += @(Get-ChildItem -LiteralPath $firewallExtract -Filter '*.exe' -Recurse -File | ForEach-Object FullName)
        foreach ($program in @($programs | Select-Object -Unique)) {
            $rule = 'WinImgDisposable-' + [guid]::NewGuid().ToString('N')
            New-NetFirewallRule -Name $rule -DisplayName $rule -Direction Outbound -Action Block -Program $program -Profile Any | Out-Null
            $rules.Add($rule)
        }
        $configuration = [ordered]@{ ArchivePath = [IO.Path]::GetFullPath($ArchivePath); ResultDirectory = $ResultDirectory; LegacyMagickPath = $LegacyMagickPath; BaselineZipPath = $BaselineZipPath; RuntimeDirectory = $firewallExtract; FirewallRules = $rules.Count; Candidate = $candidate; LegacyModuleEnvironment = $legacyModuleEnvironment }
        [IO.File]::WriteAllText((Join-Path $root 'runner-control.json'), ($configuration | ConvertTo-Json), $utf8)
        $tokens = @('-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',$PSCommandPath,'-Worker','-ArchivePath',$ArchivePath,'-ResultDirectory',$ResultDirectory,'-LegacyMagickPath',$LegacyMagickPath,'-BaselineZipPath',$BaselineZipPath)
        $child = Start-Process -FilePath $legacyShell -ArgumentList (($tokens | ForEach-Object { ConvertTo-HotfixArgument $_ }) -join ' ') -Credential ([Management.Automation.PSCredential]::new($qualifiedUser, $password)) -LoadUserProfile -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $root 'worker.stdout.txt') -RedirectStandardError (Join-Path $root 'worker.stderr.txt') -WorkingDirectory $repository
        $null = $child.Handle
        if (-not $child.WaitForExit(600000)) { $child.Kill(); throw 'Ordinary-account acceptance exceeded deadline.' }
        if ($child.ExitCode -ne 0) {
            $detail = ''
            $summaryPath = Join-Path $root 'acceptance-summary.json'
            if ([IO.File]::Exists($summaryPath)) {
                try {
                    $failure = [IO.File]::ReadAllText($summaryPath) | ConvertFrom-Json
                    $detail = 'Command=' + $failure.ErrorCommand + '; Line=' + $failure.ErrorLine + '; ' + $failure.Error
                } catch { $detail = '' }
            }
            if (-not $detail) {
                $errorPath = Join-Path $root 'worker.stderr.txt'
                if ([IO.File]::Exists($errorPath)) {
                    $reader = [IO.StreamReader]::new($errorPath)
                    try {
                        $characters = New-Object char[] 4096
                        $count = $reader.Read($characters, 0, $characters.Length)
                        $detail = [string]::new($characters, 0, $count)
                    } finally { $reader.Dispose() }
                }
            }
            $detail = $detail.Replace($root, '<acceptance>').Replace($repository, '<checkout>').Replace($env:USERPROFILE, '<runner-profile>').Replace($userName, '<disposable-user>')
            $detail = [regex]::Replace($detail, '(?i)[A-Z]:[\\/]Users[\\/][^\\/\r\n" ]+', '<profile>')
            $detail = [regex]::Replace($detail, '[\x00-\x08\x0b\x0c\x0e-\x1f]', ' ')
            if ($detail.Length -gt 1800) { $detail = $detail.Substring(0, 1800) + ' [bounded]' }
            Write-Host ('Acceptance worker cause: ' + $detail)
        }
        Assert-Hotfix ($child.ExitCode -eq 0) ('Ordinary-account acceptance failed; inspect owned worker output. Exit=' + $child.ExitCode)
        if ($public) { [IO.File]::WriteAllText((Join-Path $root 'public-release.json'), ($public | ConvertTo-Json), $utf8) }
    } finally {
        foreach ($rule in $rules) { Remove-NetFirewallRule -Name $rule -ErrorAction Continue }
        if ($checkoutAcl) { Set-Acl -LiteralPath $repository -AclObject $checkoutAcl }
        if ($createdUser) { Remove-LocalUser -Name $userName }
        $password.Dispose()
    }
    exit 0
}

$summary = [ordered]@{ Schema = 1; Passed = $false; InstrumentedLocal = [bool]$InstrumentedLocal; Windows10Executed = $false; Cases = @(); Limitation = 'Windows 10 disposable runner unavailable; Windows 11 desktop and maintained Windows Server are recorded separately.' }
try {
    $os = Get-CimInstance Win32_OperatingSystem
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $administrator = [Security.Principal.WindowsPrincipal]::new($identity).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $InstrumentedLocal) {
        # CreateProcessWithLogon loads the new user's registry hive; its child
        # environment may still contain parent values. Rebuild only this worker's
        # identity variables from its actual token/profile, never from the owner.
        $ownProfile = [Environment]::GetFolderPath('UserProfile')
        $ownName = ($identity.Name -split '\\')[-1]
        Assert-Hotfix ($ownName -like 'winimg*' -and $ownProfile -and $ownProfile -notmatch '(?i)runneradmin') 'Disposable token/profile identity is invalid.'
        $env:USERPROFILE = $ownProfile
        $env:USERNAME = $ownName
        $env:HOME = $ownProfile
        $roamingData = [Environment]::GetFolderPath('ApplicationData')
        $localData = [Environment]::GetFolderPath('LocalApplicationData')
        if (-not $roamingData) { $roamingData = Join-Path $ownProfile 'AppData/Roaming' }
        if (-not $localData) { $localData = Join-Path $ownProfile 'AppData/Local' }
        Assert-Hotfix ($roamingData.StartsWith($ownProfile + '\', [StringComparison]::OrdinalIgnoreCase) -and $localData.StartsWith($ownProfile + '\', [StringComparison]::OrdinalIgnoreCase)) 'Disposable application-data paths belong to another profile.'
        [IO.Directory]::CreateDirectory($roamingData) | Out-Null
        [IO.Directory]::CreateDirectory($localData) | Out-Null
        $env:APPDATA = $roamingData
        $env:LOCALAPPDATA = $localData
        $env:TEMP = $root
        $env:TMP = $root
        $profileDrive = [IO.Path]::GetPathRoot($ownProfile).TrimEnd('\')
        $env:HOMEDRIVE = $profileDrive
        $env:HOMEPATH = $ownProfile.Substring($profileDrive.Length)
    }
    $summary.Environment = [ordered]@{ OS = $os.Caption; Build = $os.BuildNumber; Shell = $PSVersionTable.PSVersion.ToString(); Administrator = $administrator }
    if (-not $InstrumentedLocal) {
        Assert-Hotfix (-not $administrator) 'Packaged offline acceptance must run as an ordinary account.'
        Assert-Hotfix (-not (Test-HotfixOutbound)) 'Offline credibility probe unexpectedly reached GitHub.'
        $summary.OutboundProbeBlocked = $true
        $control = Get-Content -LiteralPath (Join-Path $root 'runner-control.json') -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($control.Candidate) { $summary.Candidate = $control.Candidate }
        $summary.LegacyModuleEnvironment = $control.LegacyModuleEnvironment
    }
    $package = Join-Path $root ('portable ' + [char]0x00e5 + [char]0x96ea)
    if (-not $InstrumentedLocal) { $package = (Get-Content -LiteralPath (Join-Path $root 'runner-control.json') -Raw -Encoding UTF8 | ConvertFrom-Json).RuntimeDirectory }
    else { Expand-Archive -LiteralPath $ArchivePath -DestinationPath $package }
    $baseline = Join-Path $root 'baseline-extracted'
    Expand-Archive -LiteralPath $BaselineZipPath -DestinationPath $baseline
    $application = Join-Path $package 'WinImgNormalizer.ps1'
    $version = [regex]::Match([IO.File]::ReadAllText($application), '(?m)^function Get-WinImgVersion\s*\{\s*return ''(?<version>\d+\.\d+\.\d+)''')
    Assert-Hotfix $version.Success 'Packaged version source missing.'
    $summary.ApplicationVersion = $version.Groups['version'].Value
    $consoleFixture = $null
    if (-not $InstrumentedLocal) {
        $consoleFixture = Join-Path $root 'AcceptancePrivateConsole.exe'
        $compiler = Join-Path $env:SystemRoot 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
        $fixtureSource = Join-Path $PSScriptRoot 'fixtures/LauncherConsoleFixture.cs'
        $compile = Invoke-HotfixProcess $compiler @('/nologo','/target:winexe','/reference:System.Web.Extensions.dll',('/out:' + $consoleFixture),$fixtureSource) $root @{}
        Assert-Hotfix ($compile.Exit -eq 0 -and [IO.File]::Exists($consoleFixture)) 'Packaged processing private-console fixture did not compile.'
        $summary.ConsoleFixture = [ordered]@{ SourceSha256 = (Get-FileHash -LiteralPath $fixtureSource).Hash.ToLowerInvariant(); ExecutableSha256 = (Get-FileHash -LiteralPath $consoleFixture).Hash.ToLowerInvariant(); CompilerSha256 = (Get-FileHash -LiteralPath $compiler).Hash.ToLowerInvariant() }
    }
    $engines = @(Get-ChildItem -LiteralPath $package -Recurse -File -Filter magick.exe)
    Assert-Hotfix ($engines.Count -eq 1) 'The portable archive must contain exactly one private engine.'
    $engine = $engines[0].FullName
    $systemPath = (Split-Path -Parent $legacyShell) + ';' + (Join-Path $env:SystemRoot 'System32')
    $legacyPath = (Split-Path -Parent $LegacyMagickPath) + ';' + $systemPath
    $legacyBefore = (Get-FileHash -LiteralPath $LegacyMagickPath).Hash
    $banner = Invoke-HotfixProcess $LegacyMagickPath @('-version') $root @{ PATH = $legacyPath }
    Assert-Hotfix ($banner.Exit -eq 0 -and $banner.Out -match '7\.1\.2-5 Q16-HDRI x64 4a620c2:20251006' -and $banner.Out -match 'Modules') 'The genuine reported legacy DLL build was not executed.'
    $summary.Legacy = [ordered]@{ Banner = $banner.Out.Trim(); Sha256 = $legacyBefore.ToLowerInvariant(); Variant = 'DLL installer unpacked without installation' }
    $current = Invoke-HotfixProcess $engine @('-version') $root @{ PATH = $systemPath }
    Assert-Hotfix ($current.Exit -eq 0) 'The shipped engine fails to start.'
    $summary.Engine = [ordered]@{ Version = ($current.Out -split '\r?\n')[0]; Source = 'bundled'; RelativePath = $engine.Substring($package.Length + 1); Sha256 = (Get-FileHash -LiteralPath $engine).Hash.ToLowerInvariant() }
    $summary.PrivateOpenMp = Test-HotfixPrivateOpenMp $engine $package $root
    $summary.ZipSha256 = (Get-FileHash -LiteralPath $ArchivePath).Hash.ToLowerInvariant()
    $source = Join-Path $root ('same synthetic input ' + [char]0x00e5 + [char]0x96ea)
    [IO.Directory]::CreateDirectory($source) | Out-Null
    Add-Type -AssemblyName System.Drawing
    $bitmap = [Drawing.Bitmap]::new(16,12)
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.Clear([Drawing.Color]::CornflowerBlue)
        $bitmap.Save((Join-Path $source 'tiny.jpg'), [Drawing.Imaging.ImageFormat]::Jpeg)
        $bitmap.Save((Join-Path $source 'tiny.png'), [Drawing.Imaging.ImageFormat]::Png)
    } finally { $graphics.Dispose(); $bitmap.Dispose() }
    $videoFixture = Join-Path $PSScriptRoot 'fixtures/video/synthetic-color.mp4'
    $videoProvenance = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'fixtures/video/PROVENANCE.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-Hotfix ((Get-Item -LiteralPath $videoFixture).Length -eq $videoProvenance.bytes -and (Get-FileHash -LiteralPath $videoFixture).Hash -ieq $videoProvenance.sha256 -and $videoProvenance.verification.full_decode_exit -eq 0 -and $videoProvenance.verification.decoded_frames -eq 2) 'The authenticated, fully decoded synthetic video fixture differs from its provenance.'
    Copy-Item -LiteralPath $videoFixture -Destination (Join-Path $source 'clip.mp4')
    $summary.VideoFixture = [ordered]@{ Sha256 = $videoProvenance.sha256; Bytes = $videoProvenance.bytes; Codec = $videoProvenance.verification.codec; Width = $videoProvenance.verification.width; Height = $videoProvenance.verification.height; DecodedFrames = $videoProvenance.verification.decoded_frames; FullDecodeExit = $videoProvenance.verification.full_decode_exit; ProvenanceSha256 = (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'fixtures/video/PROVENANCE.json')).Hash.ToLowerInvariant() }
    $sourceBefore = Get-HotfixSourceState $source
    $ownedOutput = Join-Path $root 'instrumented-output'
    [IO.Directory]::CreateDirectory($ownedOutput) | Out-Null
    $pictures = if ($InstrumentedLocal) { $ownedOutput } else { [Environment]::GetFolderPath('MyPictures') }
    if (-not $pictures) { $pictures = Join-Path $env:USERPROFILE 'Pictures' }
    if (-not $InstrumentedLocal) { Assert-Hotfix ($pictures.StartsWith($env:USERPROFILE + '\', [StringComparison]::OrdinalIgnoreCase) -and $env:USERNAME -like 'winimg*') 'Exact entry-point outputs must belong to the disposable account.' }
    [IO.Directory]::CreateDirectory($pictures) | Out-Null
    function Invoke-HotfixEntry([string]$Label, [string]$Script, [string]$Shell, [string]$InputDirectory, [string]$PathValue, [int]$Expected, [switch]$Bat, [switch]$WithCap, [hashtable]$AdditionalEnvironment, [switch]$InteractiveConsole) {
        $before = @(Get-ChildItem -LiteralPath $pictures -Directory | ForEach-Object FullName)
        $environment = @{ PATH = $PathValue; WINIMG_NO_PAUSE = '1'; NO_COLOR = '1'; TEMP = $root; TMP = $root }
        if ($AdditionalEnvironment) { foreach ($key in $AdditionalEnvironment.Keys) { $environment[$key] = $AdditionalEnvironment[$key] } }
        if ($InstrumentedLocal) {
            Assert-Hotfix (-not $Bat) 'Local instrumentation cannot claim the actual BAT.'
            $driver = Join-Path $root ($Label + '-driver.ps1')
            $driverText = 'param([string]$Script,[string]$Source,[string]$Output); . $Script; $result = Invoke-WinImgNormalizerCommand -Arguments @($Source' + $(if ($WithCap) { ',1048576' } else { '' }) + ') -OutputParent $Output; exit $result'
            [IO.File]::WriteAllText($driver, $driverText, $utf8)
            $tokens = @('-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',$driver,$Script,$InputDirectory,$pictures)
        } elseif ($Bat) {
            $environment['WINIMG_ACCEPTANCE_BAT'] = [IO.Path]::ChangeExtension($Script, '.bat')
            $environment['WINIMG_ACCEPTANCE_SOURCE'] = $InputDirectory
            $tokens = @('/d','/v:off','/s','/c','""%WINIMG_ACCEPTANCE_BAT%" "%WINIMG_ACCEPTANCE_SOURCE%""')
            $Shell = $cmd
            if ($InteractiveConsole) {
                $environment['WINIMG_NO_PAUSE'] = $null
                $environment['NO_COLOR'] = $null
                $environment['WINIMG_TEST_PRIVATE_CONSOLE_CMD'] = $cmd
                $environment['WINIMG_TEST_PRIVATE_CONSOLE_ARGUMENTS'] = '/d /v:off /s /c ' + $tokens[-1]
                $environment['WINIMG_TEST_PRIVATE_CONSOLE_RECORD'] = Join-Path $root ($Label + '.console.json')
                $Shell = $consoleFixture
                $tokens = @()
            }
        } else {
            $tokens = @('-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',$Script,$InputDirectory)
            if ($WithCap) { $tokens += '1048576' }
        }
        $result = Invoke-HotfixProcess $Shell $tokens $root $environment -InteractiveConsole:$InteractiveConsole
        [IO.File]::WriteAllText((Join-Path $root ($Label + '.stdout.txt')), $result.Out, $utf8)
        [IO.File]::WriteAllText((Join-Path $root ($Label + '.stderr.txt')), $result.Error, $utf8)
        Assert-Hotfix ($result.Exit -eq $Expected) ($Label + ' exit ' + $result.Exit + ', expected ' + $Expected + ': ' + $result.Out + $result.Error)
        Assert-Hotfix ($result.Out -notmatch '[\x1b\x08\x0d](?!\n)') ($Label + ' redirected output contains terminal controls.')
        $after = @(Get-ChildItem -LiteralPath $pictures -Directory | Where-Object FullName -notin $before)
        if (-not $InteractiveConsole) {
            foreach ($run in $after) {
                Assert-Hotfix ($result.Out -match ('(?m)^' + [regex]::Escape($run.FullName) + '\r?$')) ($Label + ' did not preserve the complete Unicode output path on its own redirected line.')
                $reports = @(Get-ChildItem -LiteralPath $run.FullName -Recurse -File -Filter '*.csv')
                Assert-Hotfix ($reports.Count -eq 1 -and $result.Out -match ('(?m)^' + [regex]::Escape($reports[0].FullName) + '\r?$')) ($Label + ' did not preserve the complete Unicode report path on its own redirected line.')
            }
        }
        $summary.Cases += [ordered]@{ Name = $Label; Exit = $result.Exit; Shell = $(if ($Bat) { 'CMD / Windows PowerShell 5.1' } else { [IO.Path]::GetFileName($Shell) }); Entry = $(if ($InstrumentedLocal) { 'instrumented command function' } elseif ($InteractiveConsole) { 'exact packaged BAT in private real console' } elseif ($Bat) { 'exact packaged BAT' } else { 'exact packaged PS1 -File' }); Outputs = $after.Count; PrivateConsoleOutput = [bool]$result.Console.PrivateConsoleOutput }
        return [pscustomobject]@{ Result = $result; Runs = $after }
    }
    $red = Invoke-HotfixEntry 'A1-baseline-refusal' (Join-Path $baseline 'WinImgNormalizer.ps1') $legacyShell $source $legacyPath 1
    Assert-Hotfix ($red.Result.Out -match '7\.1\.2-32 or newer' -and $red.Runs.Count -eq 0) 'Baseline did not reproduce the pre-decode version refusal.'
    function Assert-HotfixMixed($Observation) {
        Assert-Hotfix ($Observation.Runs.Count -eq 1) 'Mixed run did not produce one output root.'
        $run = $Observation.Runs[0].FullName
        $jpegs = @(Get-ChildItem -LiteralPath $run -File -Recurse -Filter '*.jpeg')
        Assert-Hotfix ($jpegs.Count -eq 2) 'Mixed run did not produce two JPEGs.'
        foreach ($jpeg in $jpegs) {
            $decoded = Invoke-HotfixProcess $engine @('identify','+ping','-regard-warnings','-format','%m|%w|%h|%n',$jpeg.FullName) $root @{ PATH = $systemPath }
            Assert-Hotfix ($jpeg.Length -gt 0 -and $decoded.Exit -eq 0 -and $decoded.Out -eq 'JPEG|16|12|1') 'A mixed output failed real full JPEG decode.'
        }
        Assert-Hotfix ((Get-FileHash -LiteralPath (Join-Path $run 'clip.mp4')).Hash -eq (Get-FileHash -LiteralPath (Join-Path $source 'clip.mp4')).Hash) 'Video bytes changed.'
        $logs = @(Get-ChildItem -LiteralPath $run -File -Recurse -Filter '*.log')
        $text = ($logs | ForEach-Object { [IO.File]::ReadAllText($_.FullName) }) -join "`n"
        Assert-Hotfix ($text -match [regex]::Escape($engine) -and $text -match [regex]::Escape($summary.Engine.Sha256)) 'Processing engine full path/hash is absent from the actual log.'
    }
    foreach ($shell in @($legacyShell, $modernShell)) {
        $label = if ($shell -eq $legacyShell) { 'ps51' } else { 'ps7' }
        Assert-HotfixMixed (Invoke-HotfixEntry ('A1-' + $label) $application $shell $source $legacyPath 0)
        Assert-HotfixMixed (Invoke-HotfixEntry ('A2-clean-' + $label) $application $shell $source $systemPath 0 -WithCap)
    }
    # An independently staged genuine compatible build newer than the old CI
    # pin must remain irrelevant to the standard package's selected backend.
    $newerExternal = Join-Path $root 'genuine newer compatible external engine'
    [IO.Directory]::CreateDirectory($newerExternal) | Out-Null
    foreach ($file in Get-ChildItem -LiteralPath (Split-Path -Parent $engine) -File) { [IO.File]::Copy($file.FullName, (Join-Path $newerExternal $file.Name), $false) }
    $newerExecutable = Join-Path $newerExternal 'magick.exe'
    $newerPath = $newerExternal + ';' + $systemPath
    $newerBanner = Invoke-HotfixProcess $newerExecutable @('-version') $root @{ PATH = $newerPath }
    Assert-Hotfix ($newerBanner.Exit -eq 0 -and $newerBanner.Out -match '7\.1\.2-35.*Q16-HDRI x64' -and (Get-FileHash -LiteralPath $newerExecutable).Hash -ieq $summary.Engine.Sha256) 'The genuine newer-compatible external fixture is invalid.'
    Assert-HotfixMixed (Invoke-HotfixEntry 'A3-real-newer-compatible-PATH' $application $legacyShell $source $newerPath 0)
    $summary.NewerCompatiblePath = [ordered]@{ Banner = ($newerBanner.Out -split '\r?\n')[0]; Sha256 = (Get-FileHash -LiteralPath $newerExecutable).Hash.ToLowerInvariant(); ComparedWithOldCIPin = '7.1.2-32'; HarnessOnlyVersionProbe = $true; ApplicationEngineSource = 'bundled'; ApplicationEnginePathVerifiedInLog = $true }
    if (-not $InstrumentedLocal) {
        Assert-HotfixMixed (Invoke-HotfixEntry 'A1-BAT-oldPATH' $application $legacyShell $source $legacyPath 0 -Bat)
        Assert-HotfixMixed (Invoke-HotfixEntry 'A2-BAT-clean-different-CWD' $application $legacyShell $source $systemPath 0 -Bat)
        $interactiveSuccess = Invoke-HotfixEntry 'A6-interactive-BAT-success' $application $legacyShell $source $legacyPath 0 -Bat -InteractiveConsole
        Assert-HotfixMixed $interactiveSuccess
        Assert-Hotfix ($interactiveSuccess.Result.Out -match '\[04 / RESULT\]\s+COMPLETE' -and $interactiveSuccess.Result.Out -match 'Press any key to close\.') ('Interactive BAT success/pause did not render the real result. Actual captured console: ' + (Get-HotfixBoundedConsoleDiagnostic $interactiveSuccess.Result.Out))
        $partialSource = Join-Path $root 'interactive partial inputs'; [IO.Directory]::CreateDirectory($partialSource) | Out-Null
        foreach ($name in @('tiny.jpg','tiny.png','clip.mp4')) { Copy-Item -LiteralPath (Join-Path $source $name) -Destination (Join-Path $partialSource $name) }
        [IO.File]::WriteAllBytes((Join-Path $partialSource 'corrupt.jpg'), [byte[]](0,1,2,3,4))
        $partialBefore = Get-HotfixSourceState $partialSource
        $interactivePartial = Invoke-HotfixEntry 'A6-interactive-BAT-partial' $application $legacyShell $partialSource $legacyPath 2 -Bat -InteractiveConsole
        Assert-HotfixMixed $interactivePartial
        Assert-Hotfix ($interactivePartial.Result.Out -match '\[04 / RESULT\]\s+PARTIAL' -and $interactivePartial.Result.Out -match 'corrupt.jpg' -and $interactivePartial.Result.Out -match 'Press any key to close\.') ('Interactive BAT partial result or cause/pause is missing. Actual captured console: ' + (Get-HotfixBoundedConsoleDiagnostic $interactivePartial.Result.Out))
        Assert-Hotfix ((Get-HotfixSourceState $partialSource) -ceq $partialBefore) 'Interactive partial originals changed.'
        $control = Get-Content -LiteralPath (Join-Path $root 'runner-control.json') -Raw -Encoding UTF8 | ConvertFrom-Json
        Assert-Hotfix ($control.FirewallRules -ge 4) 'Offline application/children firewall restrictions were not applied.'
        $summary.Offline = 'All tested PowerShell/CMD hosts and shipped runtime executables had outbound deny rules on the disposable hosted VM. Downloads and fixtures were staged before the rules; no desktop firewall changes.'
    } else {
        $summary.Offline = 'Not claimed for this supplemental desktop run; no network settings changed.'
        $summary.InteractivePackagedProcessing = 'Not executed in instrumented local mode; exact packaged console processing is exercised only in the disposable profile.'
    }
    # A PATH/current-directory executable must never become a release backend.
    $decoySource = Join-Path $root 'Decoy.cs'
    $decoyMarker = Join-Path $root 'DECOY-WAS-EXECUTED.txt'
    [IO.File]::WriteAllText($decoySource, 'using System; using System.IO; public static class Decoy { public static int Main() { File.WriteAllText(Environment.GetEnvironmentVariable("WINIMG_DECOY_MARKER"), "executed"); Console.WriteLine("Version: ImageMagick 8.0.0-99"); return 0; } }', $utf8)
    $compiler = Join-Path $env:SystemRoot 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
    $compiled = Invoke-HotfixProcess $compiler @('/nologo','/target:exe',('/out:' + (Join-Path $root 'magick.exe')), $decoySource) $root @{}
    Assert-Hotfix ($compiled.Exit -eq 0) 'Inert decoy fixture compilation failed.'
    Copy-Item -LiteralPath (Join-Path $root 'magick.exe') -Destination (Join-Path $source 'magick.exe')
    $decoyRun = Invoke-HotfixEntry 'A3-unexpected-major-decoy' $application $legacyShell $source ($root + ';' + $systemPath) 0 -AdditionalEnvironment @{ WINIMG_DECOY_MARKER = $decoyMarker }
    Assert-HotfixMixed $decoyRun
    Assert-Hotfix (-not [IO.File]::Exists($decoyMarker)) 'A system/media/current-directory decoy executed.'
    # Adding the inert source file is deliberate test setup, not processing mutation.
    $sourceBefore = Get-HotfixSourceState $source
    $empty = Join-Path $root 'empty-input'; [IO.Directory]::CreateDirectory($empty) | Out-Null
    $videoOnly = Join-Path $root 'video-only-input'; [IO.Directory]::CreateDirectory($videoOnly) | Out-Null
    Copy-Item -LiteralPath (Join-Path $source 'clip.mp4') -Destination (Join-Path $videoOnly 'clip.mp4')
    $heldEngine = $engine + '.test-unavailable'
    [IO.File]::Move($engine, $heldEngine)
    try {
        $missing = Invoke-HotfixEntry 'A3-missing-packaged-engine' $application $legacyShell $source $legacyPath 1
        Assert-Hotfix ($missing.Result.Out -match '(?i)re-extract|official.*ZIP' -and $missing.Runs.Count -eq 0) 'Missing runtime did not fail before output with a useful remedy.'
        $emptyResult = Invoke-HotfixEntry 'A4-empty-without-engine' $application $legacyShell $empty $legacyPath 0
        Assert-Hotfix ($emptyResult.Result.Out -match '(?i)no.*(eligible|media|files)|no-op|nothing|empty') 'Empty input is not an explicit no-op.'
        $videoResult = Invoke-HotfixEntry 'A4-video-without-engine' $application $legacyShell $videoOnly $legacyPath 0
        Assert-Hotfix ($videoResult.Runs.Count -eq 1) 'Video-only work failed without an engine.'
        Assert-Hotfix ((Get-FileHash -LiteralPath (Join-Path $videoResult.Runs[0].FullName 'clip.mp4')).Hash -eq (Get-FileHash -LiteralPath (Join-Path $videoOnly 'clip.mp4')).Hash) 'Video-only bytes differ.'
        Assert-Hotfix ($videoResult.Result.Out -match '(?i)not needed') 'Copy-only output falsely claims an image engine.'
    } finally { [IO.File]::Move($heldEngine, $engine) }
    $engineBytes = [IO.File]::ReadAllBytes($engine)
    try {
        [IO.File]::WriteAllBytes($engine, [byte[]](0,1,2,3))
        $corrupt = Invoke-HotfixEntry 'A3-corrupt-packaged-engine' $application $legacyShell $source $legacyPath 1
        Assert-Hotfix ($corrupt.Runs.Count -eq 0 -and $corrupt.Result.Out -match '(?i)integrity|damaged|re-extract') 'Corrupt packaged engine did not fail closed.'
    } finally { [IO.File]::WriteAllBytes($engine, $engineBytes) }
    $policyDirectory = Join-Path $root 'explicit administrator policy'; [IO.Directory]::CreateDirectory($policyDirectory) | Out-Null
    [IO.File]::WriteAllText((Join-Path $policyDirectory 'policy.xml'), '<policymap><policy domain="coder" rights="none" pattern="JPEG" /></policymap>', $utf8)
    $denied = Invoke-HotfixEntry 'A3-explicit-policy-denial' $application $legacyShell $source $legacyPath 2 -AdditionalEnvironment @{ MAGICK_CONFIGURE_PATH = $policyDirectory }
    $deniedLogs = @(Get-ChildItem -LiteralPath $denied.Runs[0].FullName -Recurse -File -Filter '*.log' | ForEach-Object { [IO.File]::ReadAllText($_.FullName) }) -join "`n"
    Assert-Hotfix (($denied.Result.Out + $denied.Result.Error + $deniedLogs) -match '(?i)policy|not authorized|not allowed') 'Explicit policy denial was not reported.'
    Assert-Hotfix (@(Get-ChildItem -LiteralPath $denied.Runs[0].FullName -Recurse -File -Filter '*.jpeg').Count -eq 0) 'A JPEG policy denial was bypassed.'
    # Real packaged codecs, including the existing independently generated HEIC collection.
    $codecSource = Join-Path $root 'packaged codec inputs'; [IO.Directory]::CreateDirectory($codecSource) | Out-Null
    foreach ($coder in @('JPEG','PNG','WEBP','GIF','TIFF')) {
        $name = $coder.ToLowerInvariant() + '.' + $coder.ToLowerInvariant()
        $created = Invoke-HotfixProcess $engine @('-size','32x24','xc:#40a060','-depth','8','-strip',($coder + ':' + (Join-Path $codecSource $name))) $root @{ PATH = $systemPath }
        Assert-Hotfix ($created.Exit -eq 0) ('Packaged codec fixture creation failed: ' + $coder)
    }
    $frames = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'Frames.Tests.ps1'))
    $heicLiteral = [regex]::Match($frames, '\$heicBytes = \[Convert\]::FromBase64String\(''(?<bytes>[A-Za-z0-9+/=]+)''\)')
    Assert-Hotfix $heicLiteral.Success 'Existing synthetic HEIC fixture literal missing.'
    foreach ($extension in @('heic','heif')) { [IO.File]::WriteAllBytes((Join-Path $codecSource ('collection.' + $extension)), [Convert]::FromBase64String($heicLiteral.Groups['bytes'].Value)) }
    $codecBefore = Get-HotfixSourceState $codecSource
    $codecRun = Invoke-HotfixEntry 'A5-real-packaged-codecs' $application $legacyShell $codecSource $systemPath 0
    Assert-Hotfix ($codecRun.Runs.Count -eq 1) 'Packaged codec run output missing.'
    $codecJpegs = @(Get-ChildItem -LiteralPath $codecRun.Runs[0].FullName -Recurse -File -Filter '*.jpeg')
    Assert-Hotfix ($codecJpegs.Count -eq 7) 'All JPEG/PNG/WebP/GIF/TIFF/HEIC/HEIF inputs must produce distinct derivatives.'
    foreach ($jpeg in $codecJpegs) {
        $decode = Invoke-HotfixProcess $engine @('identify','+ping','-regard-warnings','-format','%m|%w|%h|%n',$jpeg.FullName) $root @{ PATH = $systemPath }
        Assert-Hotfix ($decode.Exit -eq 0 -and $decode.Out -match '^JPEG\|(?:32\|24|64\|48)\|1$') 'Packaged codec derivative failed a full single-frame JPEG decode.'
    }
    Assert-Hotfix ((Get-HotfixSourceState $codecSource) -ceq $codecBefore) 'Packaged codec originals changed.'
    $summary.RealPackagedCodecs = @('JPEG','PNG','WebP','GIF','TIFF','HEIC','HEIF')
    Assert-Hotfix ((Get-HotfixSourceState $source) -ceq $sourceBefore) 'Synthetic originals changed.'
    Assert-Hotfix ((Get-FileHash -LiteralPath $LegacyMagickPath).Hash -eq $legacyBefore) 'The old PATH executable changed.'
    $summary.Passed = $true
} catch {
    $summary.Error = $_.Exception.Message
    $summary.ErrorCommand = [string]$_.InvocationInfo.MyCommand.Name
    $summary.ErrorLine = [int]$_.InvocationInfo.ScriptLineNumber
    if ($summary.ErrorCommand.Length -gt 120) { $summary.ErrorCommand = $summary.ErrorCommand.Substring(0, 120) }
    [IO.File]::WriteAllText((Join-Path $root 'acceptance-summary.json'), ($summary | ConvertTo-Json -Depth 10), $utf8)
    throw
}
[IO.File]::WriteAllText((Join-Path $root 'acceptance-summary.json'), ($summary | ConvertTo-Json -Depth 10), $utf8)
Write-Output ('Packaged acceptance passed: ' + $summary.Cases.Count + ' real runs; ZIP ' + $summary.ZipSha256)
exit 0
