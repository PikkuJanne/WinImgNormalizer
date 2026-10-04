BeforeAll {
    $repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    $application = Join-Path $repository 'WinImgNormalizer.ps1'

    function Assert-TraversalNoReparseAncestors {
        param([string]$Path)
        $current = [IO.DirectoryInfo]::new([IO.Path]::GetFullPath($Path))
        while ($null -ne $current) {
            if ($current.Exists -and (($current.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0)) {
                throw 'Traversal test ownership path includes a reparse point.'
            }
            $current = $current.Parent
        }
    }

    $scratchParent = [IO.Path]::GetFullPath((Join-Path $repository '.scratch'))
    Assert-TraversalNoReparseAncestors $scratchParent
    foreach ($pictures in @([Environment]::GetFolderPath('MyPictures'), (Join-Path $env:USERPROFILE 'Pictures'))) {
        if ([string]::IsNullOrWhiteSpace($pictures)) { continue }
        $picturesFull = [IO.Path]::GetFullPath($pictures).TrimEnd('\', '/')
        if ($scratchParent.Equals($picturesFull, [StringComparison]::OrdinalIgnoreCase) -or
            $scratchParent.StartsWith($picturesFull + '\', [StringComparison]::OrdinalIgnoreCase) -or
            $picturesFull.StartsWith($scratchParent + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Traversal test scratch overlaps a real Pictures location.'
        }
    }
    $git = Get-Command git -CommandType Application -ErrorAction Stop | Select-Object -First 1
    & $git.Source -C $repository check-ignore --quiet --no-index -- (Join-Path $scratchParent 'traversal-ignore-probe')
    if ($LASTEXITCODE -ne 0) { throw 'Traversal synthetic scratch must already be ignored.' }
    $ownedRoot = Join-Path $scratchParent ('T02-' + [Guid]::NewGuid().ToString('N'))
    if (Test-Path -LiteralPath $ownedRoot) { throw 'Traversal ownership directory already exists.' }
    $null = [IO.Directory]::CreateDirectory($ownedRoot)
    [IO.File]::WriteAllText((Join-Path $ownedRoot '.winimg-fixture-root'), 'M1-T02 synthetic test ownership')

    function New-TraversalDirectory {
        param([string]$Label)
        $path = [IO.Path]::GetFullPath((Join-Path $ownedRoot ($Label + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))))
        if (-not $path.StartsWith($ownedRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Traversal test path escapes its owned scratch directory.'
        }
        Assert-TraversalNoReparseAncestors $path
        if (Test-Path -LiteralPath $path) { throw 'Traversal test directory already exists.' }
        $null = [IO.Directory]::CreateDirectory($path)
        return $path
    }

    function Get-TraversalSourceState {
        param([string]$Root)
        # Deliberately prune links: a state check must not follow the loop fixture.
        $pending = New-Object 'Collections.Generic.Stack[string]'
        $pending.Push($Root)
        $state = New-Object 'Collections.Generic.List[object]'
        while ($pending.Count -gt 0) {
            foreach ($entry in Get-ChildItem -LiteralPath $pending.Pop() -Force) {
                if (($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
                if ($entry.PSIsContainer) { $pending.Push($entry.FullName); continue }
                $state.Add([pscustomobject]@{
                    Path = $entry.FullName.Substring($Root.Length + 1)
                    Hash = (Get-FileHash -LiteralPath $entry.FullName -Algorithm SHA256).Hash
                    Length = $entry.Length
                    CreationTicks = $entry.CreationTimeUtc.Ticks
                    ModifiedTicks = $entry.LastWriteTimeUtc.Ticks
                })
            }
        }
        return @($state | Sort-Object Path) | ConvertTo-Json -Depth 4 -Compress
    }

    if ([string]::IsNullOrWhiteSpace($env:WINIMG_TEST_MAGICK) -or -not [IO.Path]::IsPathRooted($env:WINIMG_TEST_MAGICK)) {
        throw 'WINIMG_TEST_MAGICK must explicitly select the verified test magick.exe.'
    }
    $magick = [IO.Path]::GetFullPath($env:WINIMG_TEST_MAGICK)
    Assert-TraversalNoReparseAncestors $magick
    if (-not (Test-Path -LiteralPath $magick -PathType Leaf) -or [IO.Path]::GetFileName($magick) -ne 'magick.exe') {
        throw 'The explicitly selected traversal test magick.exe is missing.'
    }

    function Invoke-TraversalRun {
        param([string]$Source, [string]$OutputParent)
        $observed = @(& Invoke-WinImgNormalizerCommand -Arguments @($Source) -OutputParent $OutputParent -MagickPath $magick 6>&1 3>&1 2>&1)
        $codes = @($observed | Where-Object { $_ -is [int] -or $_ -is [long] })
        if ($codes.Count -ne 1) { throw ('Expected one callable exit status: ' + ($observed -join "`n")) }
        return [pscustomobject]@{ Code = $codes[0]; Text = (@($observed | Where-Object { $_ -isnot [int] -and $_ -isnot [long] }) -join "`n") }
    }

    function ConvertTo-TraversalWindowsArgument {
        param([string]$Value)
        $escaped = [regex]::Replace($Value, '(\\*)"', '$1$1\"')
        $escaped = [regex]::Replace($escaped, '(\\+)$', '$1$1')
        return '"' + $escaped + '"'
    }

    function Start-TraversalOwnedPowerShell {
        param([string]$Script, [string[]]$Arguments, [string]$WorkDirectory)
        Assert-TraversalNoReparseAncestors $Script
        Assert-TraversalNoReparseAncestors $WorkDirectory
        $temporary = Join-Path $WorkDirectory 'temporary'
        $null = [IO.Directory]::CreateDirectory($temporary)
        $start = New-Object Diagnostics.ProcessStartInfo
        $start.FileName = (Get-Process -Id $PID).Path
        $tokens = @('-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $Script) + $Arguments
        $start.Arguments = ($tokens | ForEach-Object { ConvertTo-TraversalWindowsArgument ([string]$_) }) -join ' '
        $start.WorkingDirectory = $WorkDirectory
        $start.UseShellExecute = $false
        $start.CreateNoWindow = $true
        $start.RedirectStandardOutput = $true
        $start.RedirectStandardError = $true
        foreach ($key in @($start.EnvironmentVariables.Keys)) {
            if ([string]::Equals([string]$key, 'PSModulePath', [StringComparison]::OrdinalIgnoreCase)) { $start.EnvironmentVariables.Remove([string]$key) }
        }
        $start.EnvironmentVariables['TEMP'] = $temporary
        $start.EnvironmentVariables['TMP'] = $temporary
        $start.EnvironmentVariables['MAGICK_TEMPORARY_PATH'] = $temporary
        $process = New-Object Diagnostics.Process
        $process.StartInfo = $start
        if (-not $process.Start()) { $process.Dispose(); throw 'Owned traversal child did not start.' }
        return [pscustomobject]@{ Process = $process; StdOut = $process.StandardOutput.ReadToEndAsync(); StdErr = $process.StandardError.ReadToEndAsync() }
    }

    function Complete-TraversalOwnedPowerShell {
        param([object]$Child)
        try {
            if (-not $Child.Process.WaitForExit(30000)) {
                & (Join-Path $env:SystemRoot 'System32\taskkill.exe') /PID $Child.Process.Id /T /F 2>&1 | Out-Null
                if (-not $Child.Process.WaitForExit(5000)) { $Child.Process.Kill() }
                throw 'Owned traversal child exceeded the 30-second bound.'
            }
            if (-not $Child.StdOut.Wait(5000) -or -not $Child.StdErr.Wait(5000)) { throw 'Owned traversal child streams did not finish.' }
            return [pscustomobject]@{ ExitCode = $Child.Process.ExitCode; StdOut = $Child.StdOut.GetAwaiter().GetResult(); StdErr = $Child.StdErr.GetAwaiter().GetResult() }
        } finally { $Child.Process.Dispose() }
    }

    . $application
    $realExclusiveDirectory = (Get-Command New-WinImgExclusiveDirectory -CommandType Function).ScriptBlock
    $realDestinationInfo = (Get-Command Get-WinImgDestinationInfo -CommandType Function).ScriptBlock
}

Describe 'M1-T02 canonical destination containment (T013-T014)' {
    It 'T013 rejects <Label> before enumeration, dependency probes or any destination write' -ForEach @(
        @{ Label = 'equal source'; Kind = 'equal' },
        @{ Label = 'existing child'; Kind = 'existing' },
        @{ Label = 'missing child'; Kind = 'missing' },
        @{ Label = 'normalized dot segments'; Kind = 'dot' },
        @{ Label = 'synthetic profile parent'; Kind = 'profile' }
    ) {
        $work = New-TraversalDirectory 'nested-destination'
        $source = Join-Path $work 'synthetic-profile'
        $null = [IO.Directory]::CreateDirectory($source)
        [IO.File]::WriteAllText((Join-Path $source 'sentinel.txt'), 'Synthetic containment sentinel.')
        $parent = $source
        if ($Kind -eq 'existing') { $parent = Join-Path $source 'Pictures'; $null = [IO.Directory]::CreateDirectory($parent) }
        if ($Kind -eq 'missing') { $parent = Join-Path $source 'missing\output' }
        if ($Kind -eq 'dot') { $parent = Join-Path $source 'missing\..\output' }
        if ($Kind -eq 'profile') { $parent = Join-Path $source 'synthetic-child\Pictures' }
        $before = Get-TraversalSourceState $source
        $entriesBefore = @(Get-ChildItem -LiteralPath $source -Recurse -Force | ForEach-Object FullName) -join '|'
        Mock Get-WinImgSourceTree { throw 'Unsafe destination reached source enumeration.' }
        Mock Get-WinImgMagickInfo { throw 'Unsafe destination reached dependency probe.' }
        Mock Test-WinImgDestinationWritable { throw 'Unsafe destination reached write probe.' }
        Mock New-WinImgExclusiveDirectory { throw 'Unsafe destination reached run creation.' }
        $result = Invoke-TraversalRun -Source $source -OutputParent $parent
        $result.Code | Should -Be 1
        $result.Text | Should -Match '(?i)(destination|output).*(inside|within|equal|contain|source)'
        Should -Invoke Get-WinImgSourceTree -Times 0 -Exactly
        Should -Invoke Get-WinImgMagickInfo -Times 0 -Exactly
        Should -Invoke Test-WinImgDestinationWritable -Times 0 -Exactly
        Should -Invoke New-WinImgExclusiveDirectory -Times 0 -Exactly
        Get-TraversalSourceState $source | Should -Be $before
        (@(Get-ChildItem -LiteralPath $source -Recurse -Force | ForEach-Object FullName) -join '|') | Should -Be $entriesBefore
    }

    It 'T014 applies case-insensitive Windows segment boundaries to <Label> without touching that path' -ForEach @(
        @{ Label = 'equal'; Root = 'Q:\Photos'; Path = 'q:\PHOTOS'; Expected = $true },
        @{ Label = 'descendant'; Root = 'Q:\Photos'; Path = 'q:\photos\child'; Expected = $true },
        @{ Label = 'lookalike prefix'; Root = 'Q:\Photos'; Path = 'Q:\PhotosBackup'; Expected = $false },
        @{ Label = 'dot segments'; Root = 'Q:\Photos'; Path = 'Q:\Other\..\Photos\child'; Expected = $true },
        @{ Label = 'drive root'; Root = 'Q:\'; Path = 'q:\child'; Expected = $true },
        @{ Label = 'different drive'; Root = 'Q:\'; Path = 'R:\child'; Expected = $false },
        @{ Label = 'UNC share'; Root = '\\synthetic.invalid\share\'; Path = '\\SYNTHETIC.INVALID\SHARE\child'; Expected = $true },
        @{ Label = 'UNC share prefix'; Root = '\\synthetic.invalid\share\'; Path = '\\synthetic.invalid\shareBackup\child'; Expected = $false }
    ) {
        Test-WinImgPathContained -Root $Root -Path $Path | Should -Be $Expected
    }

    It 'T014 accepts an actual lookalike sibling and resolves a missing destination without creating it' {
        $work = New-TraversalDirectory 'boundary-siblings'
        $source = Join-Path $work 'Photos'
        $null = [IO.Directory]::CreateDirectory($source)
        $parent = Join-Path $work 'PhotosBackup\missing'
        $resolved = Assert-WinImgSafeDestination -SourceRoot $source -OutputParent $parent
        $resolved | Should -Be ([IO.Path]::GetFullPath($parent))
        Test-Path -LiteralPath $parent | Should -BeFalse
        $result = Invoke-TraversalRun -Source $source -OutputParent $parent
        $result.Code | Should -Be 0
        @(Get-ChildItem -LiteralPath $parent -Directory).Count | Should -Be 1
    }

    It 'T013 rejects a destination junction ancestor and a source junction ancestor before writes' {
        $work = New-TraversalDirectory 'canonical-junction'
        $source = Join-Path $work 'source'
        $outside = Join-Path $work 'ordinary-target'
        $null = [IO.Directory]::CreateDirectory((Join-Path $source 'child'))
        $null = [IO.Directory]::CreateDirectory($outside)
        $outputAlias = Join-Path $work 'output-alias'
        $sourceAlias = Join-Path $work 'source-alias'
        New-Item -ItemType Junction -Path $outputAlias -Target $source -ErrorAction Stop | Out-Null
        New-Item -ItemType Junction -Path $sourceAlias -Target $source -ErrorAction Stop | Out-Null
        Mock Test-WinImgDestinationWritable { throw 'Live destination junction reached write probe.' }
        $result = Invoke-TraversalRun -Source $source -OutputParent (Join-Path $outputAlias 'new-output')
        $result.Code | Should -Be 1
        $result.Text | Should -Match '(?i)(reparse|junction|link)'
        Test-Path -LiteralPath (Join-Path $source 'new-output') | Should -BeFalse
        $result = Invoke-TraversalRun -Source (Join-Path $sourceAlias 'child') -OutputParent (Join-Path $outside 'new-output')
        $result.Code | Should -Be 1
        $result.Text | Should -Match '(?i)(reparse|junction|link)'
        Test-Path -LiteralPath (Join-Path $outside 'new-output') | Should -BeFalse
        Should -Invoke Test-WinImgDestinationWritable -Times 0 -Exactly
    }

    It 'T013 rejects a dangling destination junction with a missing tail before enumeration or writes' {
        $work = New-TraversalDirectory 'dangling-junction'
        $source = Join-Path $work 'source'
        $target = Join-Path $work 'empty-target'
        $alias = Join-Path $work 'dangling-alias'
        $null = [IO.Directory]::CreateDirectory($source)
        $null = [IO.Directory]::CreateDirectory($target)
        New-Item -ItemType Junction -Path $alias -Target $target -ErrorAction Stop | Out-Null
        # Only the empty, owned target is deleted. The junction itself is never traversed.
        [IO.Directory]::Delete($target, $false)
        Mock Get-WinImgSourceTree { throw 'Dangling output ancestor reached source enumeration.' }
        Mock Test-WinImgDestinationWritable { throw 'Dangling output ancestor reached write probe.' }
        $result = Invoke-TraversalRun -Source $source -OutputParent (Join-Path $alias 'missing\output')
        $result.Code | Should -Be 1
        $result.Text | Should -Match '(?i)(reparse|junction|link)'
        Should -Invoke Get-WinImgSourceTree -Times 0 -Exactly
        Should -Invoke Test-WinImgDestinationWritable -Times 0 -Exactly
        Test-Path -LiteralPath $target | Should -BeFalse
        @(Get-ChildItem -LiteralPath $source -Force).Count | Should -Be 0
    }
}

Describe 'M1-T02 pruned traversal and incomplete scans (T013)' {
    It 'skips actual looping and outside directory junctions, reports both and preserves ordinary files' {
        $work = New-TraversalDirectory 'junction-traversal'
        $source = Join-Path $work 'source'
        $outside = Join-Path $work 'outside'
        $parent = Join-Path $work 'output'
        $null = [IO.Directory]::CreateDirectory($source)
        $null = [IO.Directory]::CreateDirectory($outside)
        [IO.File]::WriteAllBytes((Join-Path $source 'ordinary.mp4'), [byte[]](7, 8, 9, 10))
        [IO.File]::WriteAllBytes((Join-Path $outside 'must-not-copy.mp4'), [byte[]](101, 102, 103))
        New-Item -ItemType Junction -Path (Join-Path $source 'loop') -Target $source -ErrorAction Stop | Out-Null
        New-Item -ItemType Junction -Path (Join-Path $source 'outside-link') -Target $outside -ErrorAction Stop | Out-Null
        $before = Get-TraversalSourceState $source
        $outsideBefore = Get-TraversalSourceState $outside
        $tree = Get-WinImgSourceTree -SourceRoot $source
        @($tree.Files).Count | Should -Be 1
        @($tree.Directories).Count | Should -Be 0
        @($tree.Warnings).Count | Should -Be 2
        ($tree.Warnings.Path -join '|') | Should -Match 'loop'
        ($tree.Warnings.Path -join '|') | Should -Match 'outside-link'
        $result = Invoke-TraversalRun -Source $source -OutputParent $parent
        $result.Code | Should -Be 2
        $result.Text | Should -Match '(?i)(reparse|junction|link)'
        $result.Text | Should -Match 'outside-link'
        $result.Text | Should -Match 'loop'
        $run = @(Get-ChildItem -LiteralPath $parent -Directory)[0].FullName
        Test-Path -LiteralPath (Join-Path $run 'ordinary.mp4') -PathType Leaf | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $run 'loop') | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $run 'outside-link') | Should -BeFalse
        @(Get-ChildItem -LiteralPath $run -Recurse -File -Filter 'must-not-copy.mp4').Count | Should -Be 0
        Get-TraversalSourceState $source | Should -Be $before
        Get-TraversalSourceState $outside | Should -Be $outsideBefore
        foreach ($linkName in @('loop', 'outside-link')) {
            ((Get-Item -LiteralPath (Join-Path $source $linkName) -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) | Should -Not -Be 0
        }
    }

    It 'skips file reparse points using a real unprivileged symlink when available or a mandatory controlled entry' {
        $work = New-TraversalDirectory 'file-link'
        $source = Join-Path $work 'source'
        $null = [IO.Directory]::CreateDirectory($source)
        $target = Join-Path $work 'outside-video.mp4'
        $link = Join-Path $source 'linked-video.mp4'
        [IO.File]::WriteAllBytes($target, [byte[]](23, 24, 25))
        if (-not ('WinImgTraversalTestLinks' -as [type])) {
            Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class WinImgTraversalTestLinks {
    [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    [return: MarshalAs(UnmanagedType.I1)]
    public static extern bool CreateSymbolicLinkW(string link, string target, uint flags);
}
'@
        }
        $realLink = [WinImgTraversalTestLinks]::CreateSymbolicLinkW($link, $target, 2)
        $linkError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        if ($realLink) {
            ((Get-Item -LiteralPath $link -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) | Should -Not -Be 0
            [IO.File]::WriteAllText((Join-Path $work 'link-evidence.txt'), 'Actual CreateSymbolicLinkW file link was pruned.')
        } else {
            [IO.File]::WriteAllText((Join-Path $work 'link-evidence.txt'), ('CreateSymbolicLinkW unavailable; Win32=' + $linkError + '; controlled reparse-file entry tested.'))
            $fakeLink = [pscustomobject]@{ FullName = $link; Name = 'linked-video.mp4'; PSIsContainer = $false; Attributes = [IO.FileAttributes]::ReparsePoint; Extension = '.mp4'; Length = 3 }
            Mock Get-ChildItem { return $fakeLink } -ParameterFilter { $LiteralPath -eq $source }
            Mock Get-Item { return $fakeLink } -ParameterFilter { $LiteralPath -eq $link }
        }
        $hash = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash
        $tree = Get-WinImgSourceTree -SourceRoot $source
        @($tree.Files).Count | Should -Be 0
        @($tree.Warnings).Count | Should -Be 1
        $tree.Warnings[0].Path | Should -Be $link
        $tree.Warnings[0].Reason | Should -Match '^Reparse point skipped'
        $result = Invoke-TraversalRun -Source $source -OutputParent (Join-Path $work 'output')
        $result.Code | Should -Be 2
        $result.Text | Should -Match '(?i)(reparse|link)'
        @(Get-ChildItem -LiteralPath (Join-Path $work 'output') -Recurse -File -Filter '*.mp4').Count | Should -Be 0
        (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash | Should -Be $hash
        if ($realLink) {
            ((Get-Item -LiteralPath $link -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) | Should -Not -Be 0
        }
    }

    It 'reports an unreadable subtree as an incomplete scan while completing readable video work' {
        $work = New-TraversalDirectory 'incomplete-scan'
        $source = Join-Path $work 'source'
        $denied = Join-Path $source 'denied-subtree'
        $null = [IO.Directory]::CreateDirectory($denied)
        [IO.File]::WriteAllBytes((Join-Path $source 'readable.mp4'), [byte[]](44, 45, 46))
        [IO.File]::WriteAllBytes((Join-Path $denied 'unknown.mp4'), [byte[]](47, 48, 49))
        $before = Get-TraversalSourceState $source
        Mock Get-ChildItem { throw [UnauthorizedAccessException]::new('Synthetic subtree enumeration denied.') } -ParameterFilter { $LiteralPath -eq $denied }
        $tree = Get-WinImgSourceTree -SourceRoot $source
        @($tree.Files).Count | Should -Be 1
        @($tree.Warnings).Count | Should -Be 1
        $tree.Warnings[0].Path | Should -Be $denied
        $tree.Warnings[0].Reason | Should -Match '(?i)(denied|enumerat|access|scan)'
        $result = Invoke-TraversalRun -Source $source -OutputParent (Join-Path $work 'output')
        $result.Code | Should -Be 2
        $result.Text | Should -Match 'denied-subtree'
        $run = @(Get-ChildItem -LiteralPath (Join-Path $work 'output') -Directory)[0].FullName
        Test-Path -LiteralPath (Join-Path $run 'readable.mp4') -PathType Leaf | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $run 'denied-subtree\unknown.mp4') | Should -BeFalse
        # Bypass only the controlled enumeration mock for the source preservation check.
        $after = @(
            foreach ($entry in Microsoft.PowerShell.Management\Get-ChildItem -LiteralPath $source -Recurse -File | Sort-Object FullName) {
                [pscustomobject]@{ Path = $entry.FullName.Substring($source.Length + 1); Hash = (Get-FileHash -LiteralPath $entry.FullName -Algorithm SHA256).Hash; Length = $entry.Length; CreationTicks = $entry.CreationTimeUtc.Ticks; ModifiedTicks = $entry.LastWriteTimeUtc.Ticks }
            }
        ) | ConvertTo-Json -Depth 4 -Compress
        $after | Should -Be $before
    }

    It 'rechecks a source file that becomes a reparse point after inventory before copying bytes' {
        $work = New-TraversalDirectory 'changed-source-link'
        $source = Join-Path $work 'source'
        $parent = Join-Path $work 'output'
        $null = [IO.Directory]::CreateDirectory($source)
        $video = Join-Path $source 'changed.mp4'
        [IO.File]::WriteAllBytes($video, [byte[]](55, 56, 57))
        $before = Get-TraversalSourceState $source
        $state = @{ Changed = $false }
        $changedEntry = [pscustomobject]@{ FullName = $video; Attributes = [IO.FileAttributes]::ReparsePoint }
        Mock Get-WinImgDestinationInfo {
            param([string]$Path, [object[]]$Files, [long]$MaxBytes)
            $info = & $realDestinationInfo -Path $Path -Files $Files -MaxBytes $MaxBytes
            $state.Changed = $true
            return $info
        }
        Mock Get-Item { return $changedEntry } -ParameterFilter { $state.Changed -and $LiteralPath -eq $video }
        Mock Copy-Item { throw 'Changed source reparse point was copied.' }
        $result = Invoke-TraversalRun -Source $source -OutputParent $parent
        $state.Changed | Should -BeTrue
        $result.Code | Should -Be 2
        $result.Text | Should -Match 'Source entry became a reparse point after inventory'
        Should -Invoke Copy-Item -Times 0 -Exactly
        @(Get-ChildItem -LiteralPath $parent -Recurse -File -Filter '*.mp4').Count | Should -Be 0
        Get-TraversalSourceState $source | Should -Be $before
    }
}

Describe 'M1-T02 exclusive run ownership (T015)' {
    It 'does not adopt an existing <Kind> when exclusive directory creation collides' -ForEach @(
        @{ Kind = 'directory' }, @{ Kind = 'file' }
    ) {
        $parent = New-TraversalDirectory 'existing-candidate'
        $state = @{ FirstPath = ''; Count = 0; SentinelHash = '' }
        Mock New-WinImgExclusiveDirectory {
            param([string]$Path)
            $state.Count++
            if ($state.Count -eq 1) {
                $state.FirstPath = $Path
                if ($Kind -eq 'directory') {
                    $null = [IO.Directory]::CreateDirectory($Path)
                    [IO.File]::WriteAllText((Join-Path $Path 'sentinel.txt'), 'Existing unrelated directory, never adopt.')
                    $state.SentinelHash = (Get-FileHash -LiteralPath (Join-Path $Path 'sentinel.txt') -Algorithm SHA256).Hash
                } else {
                    [IO.File]::WriteAllText($Path, 'Existing unrelated regular file, never replace.')
                    $state.SentinelHash = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
                }
            }
            return (& $realExclusiveDirectory -Path $Path)
        }
        $run = New-WinImgRunDirectory -OutputParent $parent -BaseName 'source' -Stamp '20000101_010203'
        $run | Should -Not -Be $state.FirstPath
        $state.Count | Should -Be 2
        Test-Path -LiteralPath $run -PathType Container | Should -BeTrue
        @(Get-ChildItem -LiteralPath $run -Force).Count | Should -Be 0
        $sentinel = $state.FirstPath
        if ($Kind -eq 'directory') { $sentinel = Join-Path $sentinel 'sentinel.txt' }
        (Get-FileHash -LiteralPath $sentinel -Algorithm SHA256).Hash | Should -Be $state.SentinelHash
        $run | Should -Match 'source_WinImgNormalized_20000101_010203_[0-9a-f]{32}$'
    }

    It 'bounds collision retries without writing or adopting any candidate' {
        $parent = New-TraversalDirectory 'collision-exhaustion'
        Mock New-WinImgExclusiveDirectory { return $false }
        { New-WinImgRunDirectory -OutputParent $parent -BaseName 'source' -Stamp '20000101_010203' } | Should -Throw
        Should -Invoke New-WinImgExclusiveDirectory -Times 8 -Exactly
        @(Get-ChildItem -LiteralPath $parent -Force).Count | Should -Be 0
    }

    It 'propagates a native directory error instead of interpreting it as an existing candidate' {
        $parent = New-TraversalDirectory 'native-mkdir-failure'
        $missingParent = Join-Path $parent 'missing\candidate'
        { New-WinImgExclusiveDirectory -Path $missingParent } | Should -Throw
        Test-Path -LiteralPath $missingParent | Should -BeFalse
    }

    It 'runs two real child applications released together at the same timestamp with isolated output and logs' {
        $work = New-TraversalDirectory 'concurrent-applications'
        $source = Join-Path $work 'source'
        $parent = Join-Path $work 'output'
        $null = [IO.Directory]::CreateDirectory($source)
        $null = [IO.Directory]::CreateDirectory($parent)
        [IO.File]::WriteAllBytes((Join-Path $source 'shared-source.mp4'), [byte[]](61, 62, 63, 64))
        $before = Get-TraversalSourceState $source
        $release = Join-Path $work 'release'
        $script = Join-Path $work 'Concurrent-Application.ps1'
        [IO.File]::WriteAllText($script, @'
param([string]$Application, [string]$Source, [string]$Parent, [string]$Magick, [string]$Ready, [string]$Release)
$ErrorActionPreference = 'Stop'
. $Application
function Get-Date {
    param([string]$Format)
    if ($Format -eq 'yyyyMMdd_HHmmss') { return '20000101_010203' }
    return Microsoft.PowerShell.Utility\Get-Date -Format $Format
}
[IO.File]::WriteAllText($Ready, 'ready')
$timer = [Diagnostics.Stopwatch]::StartNew()
while (-not [IO.File]::Exists($Release)) {
    if ($timer.Elapsed.TotalSeconds -gt 15) { throw 'Owned concurrency barrier timed out.' }
    Start-Sleep -Milliseconds 20
}
exit (Invoke-WinImgNormalizerCommand -Arguments @($Source) -OutputParent $Parent -MagickPath $Magick)
'@, (New-Object Text.UTF8Encoding($false)))
        $children = @()
        $results = @()
        try {
            for ($index = 0; $index -lt 2; $index++) {
                $childWork = Join-Path $work ('child-' + $index)
                $null = [IO.Directory]::CreateDirectory($childWork)
                $children += Start-TraversalOwnedPowerShell -Script $script -WorkDirectory $childWork -Arguments @($application, $source, $parent, $magick, (Join-Path $work ('ready-' + $index)), $release)
            }
            $timer = [Diagnostics.Stopwatch]::StartNew()
            while (-not ([IO.File]::Exists((Join-Path $work 'ready-0')) -and [IO.File]::Exists((Join-Path $work 'ready-1')))) {
                if ($timer.Elapsed.TotalSeconds -gt 15) { throw 'Owned children did not reach the concurrency barrier.' }
                Start-Sleep -Milliseconds 20
            }
            [IO.File]::WriteAllText($release, 'release both owned children')
            foreach ($child in $children) { $results += Complete-TraversalOwnedPowerShell $child }
            $children = @()
        } finally {
            foreach ($child in $children) {
                try { if (-not $child.Process.HasExited) { $child.Process.Kill() } } catch {}
                $child.Process.Dispose()
            }
        }
        foreach ($result in $results) { $result.ExitCode | Should -Be 0; $result.StdErr | Should -BeNullOrEmpty }
        $runs = @(Get-ChildItem -LiteralPath $parent -Directory)
        $runs.Count | Should -Be 2
        $runs[0].FullName | Should -Not -Be $runs[1].FullName
        foreach ($run in $runs) {
            $run.Name | Should -Match '^source_WinImgNormalized_20000101_010203_[0-9a-f]{32}$'
            (Get-FileHash -LiteralPath (Join-Path $run.FullName 'shared-source.mp4') -Algorithm SHA256).Hash |
                Should -Be (Get-FileHash -LiteralPath (Join-Path $source 'shared-source.mp4') -Algorithm SHA256).Hash
            $logs = @(Get-ChildItem -LiteralPath $run.FullName -Recurse -File -Filter '*.log')
            $logs.Count | Should -Be 1
            $logs[0].Directory.Name | Should -Be 'reports'
            Get-Content -LiteralPath $logs[0].FullName -Raw | Should -Match ([regex]::Escape($run.FullName))
            Get-Content -LiteralPath $logs[0].FullName -Raw | Should -Match 'CopiedVideos=1'
        }
        Get-TraversalSourceState $source | Should -Be $before
    }
}

Describe 'M1-T02 generated namespace reservations (T016)' {
    It 'reserves generated names case-insensitively against files and directories' {
        Get-WinImgGeneratedName -TopLevelNames @('.winimgnormalizer', '.WINIMGNORMALIZER__2', '.WinImgNormalizer__4') | Should -Be '.WinImgNormalizer__3'
        Get-WinImgGeneratedName -TopLevelNames @('work', 'reports', '.WinImgNormalizerBackup') | Should -Be '.WinImgNormalizer'
    }

    It 'fails without overwriting a generated log arrival or falling back to TEMP' {
        $work = New-TraversalDirectory 'log-arrival'
        $source = Join-Path $work 'source'
        $parent = Join-Path $work 'output'
        $temporary = Join-Path $work 'temporary'
        $null = [IO.Directory]::CreateDirectory($source)
        $null = [IO.Directory]::CreateDirectory($temporary)
        [IO.File]::WriteAllBytes((Join-Path $source 'original.mp4'), [byte[]](91, 92, 93))
        $before = Get-TraversalSourceState $source
        $arrival = @{ Path = ''; Hash = '' }
        Mock Get-Date { return '20000101_010203' } -ParameterFilter { $Format -eq 'yyyyMMdd_HHmmss' }
        Mock New-WinImgExclusiveDirectory {
            param([string]$Path)
            $created = & $realExclusiveDirectory -Path $Path
            if ($created -and [IO.Path]::GetFileName($Path) -eq 'reports') {
                $arrival.Path = Join-Path $Path 'WinImgNormalizer_20000101_010203.log'
                [IO.File]::WriteAllText($arrival.Path, 'Unrelated arriving log must not be overwritten.')
                $arrival.Hash = (Get-FileHash -LiteralPath $arrival.Path -Algorithm SHA256).Hash
            }
            return $created
        }
        $previousTemp = $env:TEMP
        try {
            $env:TEMP = $temporary
            $result = Invoke-TraversalRun -Source $source -OutputParent $parent
        } finally { $env:TEMP = $previousTemp }
        $result.Code | Should -Be 1
        $result.Text | Should -Match '(?i)(create|creation).*log'
        $arrival.Path | Should -Not -BeNullOrEmpty
        (Get-FileHash -LiteralPath $arrival.Path -Algorithm SHA256).Hash | Should -Be $arrival.Hash
        @(Get-ChildItem -LiteralPath $temporary -Force).Count | Should -Be 0
        @(Get-ChildItem -LiteralPath $parent -Recurse -File -Filter '*.mp4').Count | Should -Be 0
        Get-TraversalSourceState $source | Should -Be $before
    }

    It 'preserves source work, reports, private-name and log-like directories beside the generated namespace' {
        $work = New-TraversalDirectory 'generated-conflicts'
        $source = Join-Path $work 'source'
        $parent = Join-Path $work 'output'
        $null = [IO.Directory]::CreateDirectory($source)
        [IO.File]::WriteAllText((Join-Path $source '.WinImgNormalizer__3'), 'Unsupported user file also reserves the generated namespace.')
        $names = @('.WinImgNormalizer', '.WinImgNormalizer__2', 'work', 'reports', 'WinImgNormalizer_20000101_010203.log')
        $index = 0
        foreach ($name in $names) {
            $directory = Join-Path $source $name
            $null = [IO.Directory]::CreateDirectory($directory)
            # Different names keep this namespace test independent of the legacy duplicate heuristic.
            [IO.File]::WriteAllBytes((Join-Path $directory ('user-video-' + $index + '.mp4')), [byte[]](81, 82, 83, 84))
            $index++
        }
        $fixedUtc = [DateTime]::Parse('2020-02-03T04:05:06Z').ToUniversalTime()
        foreach ($file in Get-ChildItem -LiteralPath $source -Recurse -File) {
            [IO.File]::SetCreationTimeUtc($file.FullName, $fixedUtc)
            [IO.File]::SetLastWriteTimeUtc($file.FullName, $fixedUtc)
        }
        $before = Get-TraversalSourceState $source
        Mock Get-Date { return '20000101_010203' } -ParameterFilter { $Format -eq 'yyyyMMdd_HHmmss' }
        $result = Invoke-TraversalRun -Source $source -OutputParent $parent
        $result.Code | Should -Be 0
        $run = @(Get-ChildItem -LiteralPath $parent -Directory)[0].FullName
        $index = 0
        foreach ($name in $names) {
            $videoName = 'user-video-' + $index + '.mp4'
            $original = Get-Item -LiteralPath (Join-Path (Join-Path $source $name) $videoName)
            $copy = Get-Item -LiteralPath (Join-Path (Join-Path $run $name) $videoName)
            (Get-FileHash -LiteralPath $copy.FullName -Algorithm SHA256).Hash | Should -Be (Get-FileHash -LiteralPath $original.FullName -Algorithm SHA256).Hash
            $copy.CreationTimeUtc.Ticks | Should -Be $original.CreationTimeUtc.Ticks
            $copy.LastWriteTimeUtc.Ticks | Should -Be $original.LastWriteTimeUtc.Ticks
            $index++
        }
        $generated = Join-Path $run '.WinImgNormalizer__4'
        Test-Path -LiteralPath (Join-Path $generated 'work') -PathType Container | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $generated 'reports') -PathType Container | Should -BeTrue
        $logs = @(Get-ChildItem -LiteralPath (Join-Path $generated 'reports') -File -Filter '*.log')
        $logs.Count | Should -Be 1
        Get-Content -LiteralPath $logs[0].FullName -Raw | Should -Match 'CopiedVideos=5'
        Get-TraversalSourceState $source | Should -Be $before
    }
}
