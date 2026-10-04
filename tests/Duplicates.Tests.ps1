BeforeAll {
    $repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    $application = Join-Path $repository 'WinImgNormalizer.ps1'

    function Assert-DuplicateNoReparseAncestors {
        param([string]$Path)
        $current = [IO.DirectoryInfo]::new([IO.Path]::GetFullPath($Path))
        while ($null -ne $current) {
            if ($current.Exists -and (($current.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0)) {
                throw 'Duplicate test ownership path includes a reparse point.'
            }
            $current = $current.Parent
        }
    }

    $scratchParent = [IO.Path]::GetFullPath((Join-Path $repository '.scratch'))
    Assert-DuplicateNoReparseAncestors $scratchParent
    foreach ($pictures in @([Environment]::GetFolderPath('MyPictures'), (Join-Path $env:USERPROFILE 'Pictures'))) {
        if ([string]::IsNullOrWhiteSpace($pictures)) { continue }
        $picturesFull = [IO.Path]::GetFullPath($pictures).TrimEnd('\', '/')
        if ($scratchParent.Equals($picturesFull, [StringComparison]::OrdinalIgnoreCase) -or
            $scratchParent.StartsWith($picturesFull + '\', [StringComparison]::OrdinalIgnoreCase) -or
            $picturesFull.StartsWith($scratchParent + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Duplicate test scratch overlaps a real Pictures location.'
        }
    }
    $git = Get-Command git -CommandType Application -ErrorAction Stop | Select-Object -First 1
    & $git.Source -C $repository check-ignore --quiet --no-index -- (Join-Path $scratchParent 'duplicate-ignore-probe')
    if ($LASTEXITCODE -ne 0) { throw 'Duplicate synthetic scratch must already be ignored.' }
    $ownedRoot = Join-Path $scratchParent ('M1-T05-' + [Guid]::NewGuid().ToString('N'))
    if (Test-Path -LiteralPath $ownedRoot) { throw 'Duplicate ownership directory already exists.' }
    $null = [IO.Directory]::CreateDirectory($ownedRoot)
    [IO.File]::WriteAllText((Join-Path $ownedRoot '.winimg-fixture-root'), 'M1-T05 synthetic test ownership')

    function New-DuplicateDirectory {
        param([string]$Label)
        $path = [IO.Path]::GetFullPath((Join-Path $ownedRoot ($Label + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))))
        if (-not $path.StartsWith($ownedRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Duplicate test path escapes its owned scratch directory.'
        }
        Assert-DuplicateNoReparseAncestors $path
        if (Test-Path -LiteralPath $path) { throw 'Duplicate test directory already exists.' }
        $null = [IO.Directory]::CreateDirectory($path)
        return $path
    }

    function Write-DuplicateInput {
        param([string]$Root, [string]$RelativePath, [byte[]]$Bytes, [DateTime]$Modified = $fixedUtc)
        $path = [IO.Path]::GetFullPath((Join-Path $Root $RelativePath))
        if (-not $path.StartsWith($Root + '\', [StringComparison]::OrdinalIgnoreCase) -or
            -not $path.StartsWith($ownedRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Duplicate input path escapes its owned synthetic source.'
        }
        Assert-DuplicateNoReparseAncestors ([IO.Path]::GetDirectoryName($path))
        $null = [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($path))
        if (Test-Path -LiteralPath $path) { throw 'Duplicate test input already exists.' }
        [IO.File]::WriteAllBytes($path, $Bytes)
        [IO.File]::SetCreationTimeUtc($path, $fixedUtc)
        [IO.File]::SetLastWriteTimeUtc($path, $Modified)
        return $path
    }

    function Get-DuplicateSourceState {
        param([string]$Root)
        return @((Get-ChildItem -LiteralPath $Root -Recurse -Force -File | Sort-Object FullName | ForEach-Object {
            [pscustomobject]@{
                Path = $_.FullName.Substring($Root.Length + 1)
                Hash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
                Length = $_.Length
                CreationTicks = $_.CreationTimeUtc.Ticks
                ModifiedTicks = $_.LastWriteTimeUtc.Ticks
            }
        })) | ConvertTo-Json -Depth 4 -Compress
    }

    function Get-DuplicateRun {
        param([string]$Parent)
        $runs = @(Get-ChildItem -LiteralPath $Parent -Directory)
        if ($runs.Count -ne 1) { throw 'Expected exactly one duplicate-test run directory.' }
        return $runs[0].FullName
    }

    function Get-DuplicateLog {
        param([string]$Run)
        $logs = @(Get-ChildItem -LiteralPath $Run -Recurse -File -Filter '*.log')
        if ($logs.Count -ne 1) { throw 'Expected exactly one duplicate-test log.' }
        return Get-Content -LiteralPath $logs[0].FullName -Raw
    }

    function Get-DuplicateSkipLines {
        param([string]$Log)
        return @([regex]::Matches($Log, '(?m)^.*?\[SKIP\] (Heuristic duplicate skipped: .*?)\r?$') | ForEach-Object { $_.Groups[1].Value })
    }

    function Assert-DuplicateRetainedLink {
        param([string]$Log, [string]$Skipped, [string]$RetainedSource, [string]$RetainedOutput)
        $lines = @(Get-DuplicateSkipLines $Log | Where-Object { $_.StartsWith('Heuristic duplicate skipped: ' + $Skipped + ' (', [StringComparison]::Ordinal) })
        $lines.Count | Should -Be 1
        $lines[0] | Should -Match ([regex]::Escape('retained source: ' + $RetainedSource))
        $lines[0] | Should -Match ([regex]::Escape('retained output: ' + $RetainedOutput))
    }

    function Assert-DuplicateOwnedCandidate {
        param([string]$Path, [string]$OutputParent)
        $ordinary = $Path
        if ($ordinary.StartsWith('\\?\UNC\', [StringComparison]::Ordinal)) { $ordinary = '\\' + $ordinary.Substring(8) }
        elseif ($ordinary.StartsWith('\\?\', [StringComparison]::Ordinal)) { $ordinary = $ordinary.Substring(4) }
        $work = Join-Path (Get-DuplicateRun $OutputParent) '.WinImgNormalizer\work'
        if (-not $ordinary.StartsWith($work + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Controlled duplicate runner was directed outside the generated work namespace.'
        }
        Assert-DuplicateNoReparseAncestors ([IO.Path]::GetDirectoryName($ordinary))
        return $ordinary
    }

    function Invoke-DuplicateRun {
        param([string]$Source, [string]$OutputParent, [scriptblock]$ProcessRunner, [long]$MaxBytes = 1MB)
        $invoke = @{ Source = $Source; OutputParent = $OutputParent; MagickPath = $magick; MaxBytes = $MaxBytes }
        if ($ProcessRunner) { $invoke.ProcessRunner = $ProcessRunner }
        $observed = @(& Invoke-WinImgNormalizer @invoke 6>&1 3>&1 2>&1)
        $codes = @($observed | Where-Object { $_ -is [int] -or $_ -is [long] })
        if ($codes.Count -ne 1) { throw ('Expected one callable duplicate exit status: ' + ($observed -join "`n")) }
        return [pscustomobject]@{ Code = $codes[0]; Text = (@($observed | Where-Object { $_ -isnot [int] -and $_ -isnot [long] }) -join "`n") }
    }

    if ([string]::IsNullOrWhiteSpace($env:WINIMG_TEST_MAGICK) -or -not [IO.Path]::IsPathRooted($env:WINIMG_TEST_MAGICK)) {
        throw 'WINIMG_TEST_MAGICK must explicitly select the verified test magick.exe.'
    }
    $magick = [IO.Path]::GetFullPath($env:WINIMG_TEST_MAGICK)
    Assert-DuplicateNoReparseAncestors $magick
    if (-not (Test-Path -LiteralPath $magick -PathType Leaf) -or [IO.Path]::GetFileName($magick) -ne 'magick.exe') {
        throw 'The explicitly selected duplicate test magick.exe is missing.'
    }

    function Invoke-DuplicateMagick {
        param([string[]]$Arguments)
        $result = & $magick @Arguments 2>&1
        $code = $LASTEXITCODE
        if ($code -ne 0) { throw ('Duplicate test ImageMagick exited {0}: {1}' -f $code, ($result -join "`n")) }
        return ($result -join "`n")
    }

    function Assert-DuplicateJpeg {
        param([string]$Path, [int]$Width, [int]$Height)
        $nativePath = Get-WinImgNativeOutputPath -Path $Path
        Invoke-DuplicateMagick @('identify', '+ping', '-regard-warnings', '-format', '%m|%w|%h|%n', $nativePath) | Should -Be ('JPEG|{0}|{1}|1' -f $Width, $Height)
        (Get-Item -LiteralPath $Path).Length | Should -BeGreaterThan 0
    }

    $fixedUtc = [DateTime]::Parse('2020-02-03T04:05:06Z').ToUniversalTime()
    $smallFixture = Join-Path $ownedRoot 'small.png'
    $largeFixture = Join-Path $ownedRoot 'large.png'
    $alternateFixture = Join-Path $ownedRoot 'alternate.png'
    $jpegFixture = Join-Path $ownedRoot 'candidate.jpeg'
    $null = Invoke-DuplicateMagick @('-quiet', '-size', '24x18', 'gradient:#205080-#E0A070', '-depth', '8', ('PNG:' + $smallFixture))
    $null = Invoke-DuplicateMagick @('-quiet', '-seed', '812', '-size', '48x32', 'xc:gray', '+noise', 'Random', '-depth', '8', ('PNG:' + $largeFixture))
    $null = Invoke-DuplicateMagick @('-quiet', '-size', '24x18', 'gradient:#E0A070-#205080', '-depth', '8', ('PNG:' + $alternateFixture))
    $null = Invoke-DuplicateMagick @('-quiet', $smallFixture, ('JPEG:' + $jpegFixture))
    $smallBytes = [IO.File]::ReadAllBytes($smallFixture)
    $largeBytes = [IO.File]::ReadAllBytes($largeFixture)
    $alternateBytes = [IO.File]::ReadAllBytes($alternateFixture)
    $jpegBytes = [IO.File]::ReadAllBytes($jpegFixture)
    # Trailing zero padding equalizes two valid PNG sizes for the explicitly
    # documented same-length/different-content heuristic limitation.
    $equalLength = [Math]::Max($smallBytes.Length, $alternateBytes.Length)
    $sameLengthFirst = New-Object byte[] $equalLength
    $sameLengthSecond = New-Object byte[] $equalLength
    [Array]::Copy($smallBytes, $sameLengthFirst, $smallBytes.Length)
    [Array]::Copy($alternateBytes, $sameLengthSecond, $alternateBytes.Length)
    $videoBytes = New-Object byte[] 4096
    $longVideoBytes = New-Object byte[] 8193
    $alternateVideoBytes = New-Object byte[] 4096
    for ($offset = 0; $offset -lt $longVideoBytes.Length; $offset++) { $longVideoBytes[$offset] = [byte](($offset * 31 + 13) % 256) }
    for ($offset = 0; $offset -lt $videoBytes.Length; $offset++) {
        $videoBytes[$offset] = [byte](($offset * 31 + 13) % 256)
        $alternateVideoBytes[$offset] = [byte](($offset * 17 + 92) % 256)
    }
    . $application
    $realMove = (Get-Command Move-WinImgPlannedImage -CommandType Function).ScriptBlock
    $realVideoCopy = (Get-Command Copy-WinImgVideoToCandidate -CommandType Function).ScriptBlock
    $realPlannedVideoCopy = (Get-Command Copy-WinImgPlannedVideo -CommandType Function).ScriptBlock
    $realSourceTree = (Get-Command Get-WinImgSourceTree -CommandType Function).ScriptBlock
    $realSourceSnapshot = (Get-Command New-WinImgSourceSnapshot -CommandType Function).ScriptBlock
}

Describe 'M1-T05 register only finalized retained media (T025)' {
    It 'T025 attempts and retains the later same-key image after first <Failure> failure' -ForEach @(
        @{ Failure = 'native' }, @{ Failure = 'validation' }, @{ Failure = 'finalization' }
    ) {
        $source = New-DuplicateDirectory 'failed-image-source'
        $parent = New-DuplicateDirectory 'failed-image-output'
        $first = Write-DuplicateInput -Root $source -RelativePath 'a\photo.png' -Bytes $smallBytes
        $second = Write-DuplicateInput -Root $source -RelativePath 'b\photo.png' -Bytes $smallBytes
        $before = Get-DuplicateSourceState $source
        $trace = [pscustomobject]@{ Sources = New-Object 'Collections.Generic.List[string]'; ArrivalPath = $null; SnapshotSources = @{} }
        Mock New-WinImgSourceSnapshot {
            param([string]$SourcePath, [string]$WorkRoot, [long]$ExpectedLength, [DateTime]$ExpectedModified, [Collections.Generic.List[object]]$OwnedCandidates)
            $snapshot = & $realSourceSnapshot -SourcePath $SourcePath -WorkRoot $WorkRoot -ExpectedLength $ExpectedLength -ExpectedModified $ExpectedModified -OwnedCandidates $OwnedCandidates
            $trace.SnapshotSources[(Get-WinImgNativeOutputPath $snapshot)] = $SourcePath
            return $snapshot
        }
        $runner = {
            param([string]$Executable, [string[]]$Arguments)
            $originalSource = $trace.SnapshotSources[$Arguments[2]]
            if (-not $originalSource) { throw 'Duplicate conversion input did not match an actual owned source snapshot.' }
            $trace.Sources.Add($originalSource)
            $candidate = Assert-DuplicateOwnedCandidate -Path $Arguments[-1].Substring('JPEG:'.Length) -OutputParent $parent
            if ($originalSource -eq $first -and $Failure -eq 'native') { [IO.File]::WriteAllBytes($candidate, $jpegBytes); return 7 }
            if ($originalSource -eq $first -and $Failure -eq 'validation') { [IO.File]::WriteAllText($candidate, 'Synthetic invalid native candidate.'); return 0 }
            & $Executable @Arguments 1>$null 2>$null
            return $LASTEXITCODE
        }
        Mock Move-WinImgPlannedImage {
            param([string]$CandidatePath, [string]$DestinationPath)
            if ($Failure -eq 'finalization' -and $DestinationPath.EndsWith('\a\photo.jpeg', [StringComparison]::OrdinalIgnoreCase)) {
                $trace.ArrivalPath = $DestinationPath
                [IO.File]::WriteAllText($DestinationPath, 'External first final-name arrival remains untouched.')
            }
            & $realMove -CandidatePath $CandidatePath -DestinationPath $DestinationPath
        }
        $result = Invoke-DuplicateRun -Source $source -OutputParent $parent -ProcessRunner $runner
        $result.Code | Should -Be 2 -Because $result.Text
        @($trace.Sources | Where-Object { $_ -eq $first }).Count | Should -BeGreaterThan 0
        @($trace.Sources | Where-Object { $_ -eq $second }).Count | Should -Be 1
        $run = Get-DuplicateRun $parent
        Assert-DuplicateJpeg -Path (Join-Path $run 'b\photo.jpeg') -Width 24 -Height 18
        if ($Failure -eq 'finalization') { [IO.File]::ReadAllText($trace.ArrivalPath) | Should -Be 'External first final-name arrival remains untouched.' }
        else { Test-Path -LiteralPath (Join-Path $run 'a\photo.jpeg') | Should -BeFalse }
        @(Get-DuplicateSkipLines (Get-DuplicateLog $run)).Count | Should -Be 0
        Get-DuplicateLog $run | Should -Match 'SUMMARY ConvertedImages=1 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=1'
        Get-DuplicateSourceState $source | Should -Be $before
    }

    It 'T025 attempts and retains the later same-key video after first staged-copy failure' {
        $source = New-DuplicateDirectory 'failed-video-source'
        $parent = New-DuplicateDirectory 'failed-video-output'
        $first = Write-DuplicateInput -Root $source -RelativePath 'a\clip.mp4' -Bytes $videoBytes
        $second = Write-DuplicateInput -Root $source -RelativePath 'b\clip.mp4' -Bytes $videoBytes
        $before = Get-DuplicateSourceState $source
        $trace = [pscustomobject]@{ Sources = New-Object 'Collections.Generic.List[string]'; PartialPath = $null }
        Mock Copy-WinImgVideoToCandidate {
            param([string]$SourcePath, [string]$CandidatePath)
            $trace.Sources.Add($SourcePath)
            $null = Assert-DuplicateOwnedCandidate -Path $CandidatePath -OutputParent $parent
            if ($SourcePath -eq $first) {
                $trace.PartialPath = $CandidatePath
                [IO.File]::WriteAllBytes($CandidatePath, [byte[]]$videoBytes[0..127])
                throw 'Controlled first video copy failure.'
            }
            & $realVideoCopy -SourcePath $SourcePath -CandidatePath $CandidatePath
        }
        $result = Invoke-DuplicateRun -Source $source -OutputParent $parent -ProcessRunner { throw 'Video-only duplicate input attempted image conversion.' }
        $result.Code | Should -Be 2 -Because $result.Text
        ($trace.Sources -join '|') | Should -Be ($first + '|' + $second)
        $run = Get-DuplicateRun $parent
        Test-Path -LiteralPath (Join-Path $run 'a\clip.mp4') | Should -BeFalse
        Test-Path -LiteralPath $trace.PartialPath | Should -BeFalse
        (Get-FileHash -LiteralPath (Join-Path $run 'b\clip.mp4') -Algorithm SHA256).Hash | Should -Be (Get-FileHash -LiteralPath $second -Algorithm SHA256).Hash
        Get-DuplicateLog $run | Should -Match 'SUMMARY ConvertedImages=0 CopiedVideos=1 Duplicates=0 Unsupported=0 Errors=1'
        Get-DuplicateSourceState $source | Should -Be $before
    }
}

Describe 'M1-T05 source length guard and retained entries (T026)' {
    It 'T026 refreshes source metadata after inventory before matching the retained heuristic key' {
        $source = New-DuplicateDirectory 'changed-inventory-source'
        $parent = New-DuplicateDirectory 'changed-inventory-output'
        $first = Write-DuplicateInput -Root $source -RelativePath 'a\photo.png' -Bytes $smallBytes
        $null = Write-DuplicateInput -Root $source -RelativePath 'b\photo.png' -Bytes $smallBytes
        $trace = [pscustomobject]@{ Calls = 0; InjectedSourceState = $null }
        Mock Get-WinImgSourceTree {
            param([string]$SourceRoot)
            $tree = & $realSourceTree -SourceRoot $SourceRoot
            $trace.Calls++
            [IO.File]::WriteAllBytes($first, $largeBytes)
            [IO.File]::SetLastWriteTimeUtc($first, $fixedUtc)
            $trace.InjectedSourceState = Get-DuplicateSourceState $source
            return $tree
        }
        $result = Invoke-DuplicateRun -Source $source -OutputParent $parent
        $result.Code | Should -Be 0 -Because $result.Text
        $trace.Calls | Should -Be 1
        $run = Get-DuplicateRun $parent
        Assert-DuplicateJpeg -Path (Join-Path $run 'a\photo.jpeg') -Width 48 -Height 32
        Assert-DuplicateJpeg -Path (Join-Path $run 'b\photo.jpeg') -Width 24 -Height 18
        Get-DuplicateLog $run | Should -Match 'SUMMARY ConvertedImages=2 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=0'
        Get-DuplicateSourceState $source | Should -Be $trace.InjectedSourceState
    }

    It 'T026 processes same-name same-timestamp <Media> files with different lengths' -ForEach @(
        @{ Media = 'image' }, @{ Media = 'video' }
    ) {
        $source = New-DuplicateDirectory 'length-source'
        $parent = New-DuplicateDirectory 'length-output'
        $name = if ($Media -eq 'image') { 'photo.png' } else { 'clip.mp4' }
        $firstBytes = if ($Media -eq 'image') { $smallBytes } else { $videoBytes }
        $secondBytes = if ($Media -eq 'image') { $largeBytes } else { $longVideoBytes }
        $firstBytes.Length | Should -Not -Be $secondBytes.Length
        $first = Write-DuplicateInput -Root $source -RelativePath ('a\' + $name) -Bytes $firstBytes
        $second = Write-DuplicateInput -Root $source -RelativePath ('b\' + $name) -Bytes $secondBytes
        $before = Get-DuplicateSourceState $source
        $result = Invoke-DuplicateRun -Source $source -OutputParent $parent
        $result.Code | Should -Be 0 -Because $result.Text
        $run = Get-DuplicateRun $parent
        if ($Media -eq 'image') {
            Assert-DuplicateJpeg -Path (Join-Path $run 'a\photo.jpeg') -Width 24 -Height 18
            Assert-DuplicateJpeg -Path (Join-Path $run 'b\photo.jpeg') -Width 48 -Height 32
            Get-DuplicateLog $run | Should -Match 'SUMMARY ConvertedImages=2 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=0'
        } else {
            foreach ($relative in @('a\clip.mp4', 'b\clip.mp4')) {
                (Get-FileHash -LiteralPath (Join-Path $run $relative) -Algorithm SHA256).Hash | Should -Be (Get-FileHash -LiteralPath (Join-Path $source $relative) -Algorithm SHA256).Hash
            }
            Get-DuplicateLog $run | Should -Match 'SUMMARY ConvertedImages=0 CopiedVideos=2 Duplicates=0 Unsupported=0 Errors=0'
        }
        Get-DuplicateSourceState $source | Should -Be $before
    }

    It 'T026 retains independent lengths so later <Media> length-one candidate links to the first retained source' -ForEach @(
        @{ Media = 'image' }, @{ Media = 'video' }
    ) {
        $source = New-DuplicateDirectory 'per-length-source'
        $parent = New-DuplicateDirectory 'per-length-output'
        $name = if ($Media -eq 'image') { 'photo.png' } else { 'clip.mp4' }
        $firstBytes = if ($Media -eq 'image') { $smallBytes } else { $videoBytes }
        $secondBytes = if ($Media -eq 'image') { $largeBytes } else { $longVideoBytes }
        $null = Write-DuplicateInput -Root $source -RelativePath ('a\' + $name) -Bytes $firstBytes
        $null = Write-DuplicateInput -Root $source -RelativePath ('b\' + $name) -Bytes $secondBytes
        $null = Write-DuplicateInput -Root $source -RelativePath ('c\' + $name) -Bytes $firstBytes
        $null = Write-DuplicateInput -Root $source -RelativePath ('d\' + $name) -Bytes $secondBytes
        $before = Get-DuplicateSourceState $source
        $result = Invoke-DuplicateRun -Source $source -OutputParent $parent
        $result.Code | Should -Be 0 -Because $result.Text
        $run = Get-DuplicateRun $parent
        $outputName = if ($Media -eq 'image') { 'photo.jpeg' } else { 'clip.mp4' }
        Test-Path -LiteralPath (Join-Path $run ('c\' + $outputName)) | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $run ('d\' + $outputName)) | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $run ('a\' + $outputName)) -PathType Leaf | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $run ('b\' + $outputName)) -PathType Leaf | Should -BeTrue
        $log = Get-DuplicateLog $run
        Assert-DuplicateRetainedLink -Log $log -Skipped ('c\' + $name) -RetainedSource ('a\' + $name) -RetainedOutput ('a\' + $outputName)
        Assert-DuplicateRetainedLink -Log $log -Skipped ('d\' + $name) -RetainedSource ('b\' + $name) -RetainedOutput ('b\' + $outputName)
        if ($Media -eq 'image') { $log | Should -Match 'SUMMARY ConvertedImages=2 CopiedVideos=0 Duplicates=2 Unsupported=0 Errors=0' }
        else { $log | Should -Match 'SUMMARY ConvertedImages=0 CopiedVideos=2 Duplicates=2 Unsupported=0 Errors=0' }
        Get-DuplicateSourceState $source | Should -Be $before
    }
}

Describe 'M1-T05 deterministic retained links and explicit heuristic limitation (T027)' {
    It 'T027 keeps finalized images without stale registration when the source changes during conversion' {
        $source = New-DuplicateDirectory 'changed-conversion-source'
        $parent = New-DuplicateDirectory 'changed-conversion-output'
        $first = Write-DuplicateInput -Root $source -RelativePath 'a\photo.png' -Bytes $smallBytes
        $second = Write-DuplicateInput -Root $source -RelativePath 'b\photo.png' -Bytes $smallBytes
        $trace = [pscustomobject]@{ Sources = New-Object 'Collections.Generic.List[string]'; InjectedSourceState = $null; SnapshotSources = @{} }
        Mock New-WinImgSourceSnapshot {
            param([string]$SourcePath, [string]$WorkRoot, [long]$ExpectedLength, [DateTime]$ExpectedModified, [Collections.Generic.List[object]]$OwnedCandidates)
            $snapshot = & $realSourceSnapshot -SourcePath $SourcePath -WorkRoot $WorkRoot -ExpectedLength $ExpectedLength -ExpectedModified $ExpectedModified -OwnedCandidates $OwnedCandidates
            $trace.SnapshotSources[(Get-WinImgNativeOutputPath $snapshot)] = $SourcePath
            return $snapshot
        }
        $runner = {
            param([string]$Executable, [string[]]$Arguments)
            $originalSource = $trace.SnapshotSources[$Arguments[2]]
            if (-not $originalSource) { throw 'Duplicate conversion input did not match an actual owned source snapshot.' }
            $trace.Sources.Add($originalSource)
            $null = Assert-DuplicateOwnedCandidate -Path $Arguments[-1].Substring('JPEG:'.Length) -OutputParent $parent
            & $Executable @Arguments 1>$null 2>$null
            $code = $LASTEXITCODE
            if ($originalSource -eq $first) {
                [IO.File]::SetLastWriteTimeUtc($first, $fixedUtc.AddSeconds(10))
                $trace.InjectedSourceState = Get-DuplicateSourceState $source
            }
            return $code
        }
        $result = Invoke-DuplicateRun -Source $source -OutputParent $parent -ProcessRunner $runner
        $result.Code | Should -Be 2 -Because $result.Text
        ($trace.Sources -join '|') | Should -Be ($first + '|' + $second)
        $run = Get-DuplicateRun $parent
        foreach ($relative in @('a\photo.jpeg', 'b\photo.jpeg')) { Assert-DuplicateJpeg -Path (Join-Path $run $relative) -Width 24 -Height 18 }
        $log = Get-DuplicateLog $run
        $log | Should -Match 'Finalized image kept without heuristic registration:'
        $log | Should -Match 'SUMMARY ConvertedImages=2 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=0'
        Get-DuplicateSourceState $source | Should -Be $trace.InjectedSourceState
    }

    It 'T027 registers video metadata from the stable copy snapshot when it changes after heuristic lookup' {
        $source = New-DuplicateDirectory 'video-copy-snapshot-source'
        $parent = New-DuplicateDirectory 'video-copy-snapshot-output'
        $first = Write-DuplicateInput -Root $source -RelativePath 'a\clip.mp4' -Bytes $videoBytes
        $second = Write-DuplicateInput -Root $source -RelativePath 'b\clip.mp4' -Bytes $longVideoBytes -Modified $fixedUtc.AddSeconds(10)
        $trace = [pscustomobject]@{ Sources = New-Object 'Collections.Generic.List[string]'; InjectedSourceState = $null }
        Mock Copy-WinImgPlannedVideo {
            param([string]$SourcePath, [string]$DestinationPath, [string]$WorkRoot)
            $trace.Sources.Add($SourcePath)
            if ($SourcePath -eq $first) {
                [IO.File]::WriteAllBytes($first, $longVideoBytes)
                [IO.File]::SetLastWriteTimeUtc($first, $fixedUtc.AddSeconds(10))
                $trace.InjectedSourceState = Get-DuplicateSourceState $source
            }
            & $realPlannedVideoCopy -SourcePath $SourcePath -DestinationPath $DestinationPath -WorkRoot $WorkRoot
        }
        $result = Invoke-DuplicateRun -Source $source -OutputParent $parent -ProcessRunner { throw 'Video-only duplicate input attempted image conversion.' }
        $result.Code | Should -Be 0 -Because $result.Text
        ($trace.Sources -join '|') | Should -Be $first
        $run = Get-DuplicateRun $parent
        $target = Join-Path $run 'a\clip.mp4'
        (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash | Should -Be (Get-FileHash -LiteralPath $first -Algorithm SHA256).Hash
        (Get-Item -LiteralPath $target).Length | Should -Be $longVideoBytes.Length
        Test-Path -LiteralPath (Join-Path $run 'b\clip.mp4') | Should -BeFalse
        $log = Get-DuplicateLog $run
        Assert-DuplicateRetainedLink -Log $log -Skipped 'b\clip.mp4' -RetainedSource 'a\clip.mp4' -RetainedOutput 'a\clip.mp4'
        $log | Should -Match 'retained status: CopiedVideo'
        $log | Should -Match 'SUMMARY ConvertedImages=0 CopiedVideos=1 Duplicates=1 Unsupported=0 Errors=0'
        Get-DuplicateSourceState $source | Should -Be $trace.InjectedSourceState
    }

    It 'T027 documents and demonstrates same-key same-length <Media> files with different bytes can still match' -ForEach @(
        @{ Media = 'image' }, @{ Media = 'video' }
    ) {
        $source = New-DuplicateDirectory 'heuristic-source'
        $parent = New-DuplicateDirectory 'heuristic-output'
        $name = if ($Media -eq 'image') { 'photo.png' } else { 'clip.mp4' }
        $firstBytes = if ($Media -eq 'image') { $sameLengthFirst } else { $videoBytes }
        $secondBytes = if ($Media -eq 'image') { $sameLengthSecond } else { $alternateVideoBytes }
        $firstBytes.Length | Should -Be $secondBytes.Length
        $first = Write-DuplicateInput -Root $source -RelativePath ('a\' + $name) -Bytes $firstBytes
        $second = Write-DuplicateInput -Root $source -RelativePath ('b\' + $name) -Bytes $secondBytes
        (Get-FileHash -LiteralPath $first -Algorithm SHA256).Hash | Should -Not -Be (Get-FileHash -LiteralPath $second -Algorithm SHA256).Hash
        if ($Media -eq 'image') {
            Invoke-DuplicateMagick @('identify', '+ping', '-regard-warnings', '-format', '%m|%w|%h|%n', $first) | Should -Be 'PNG|24|18|1'
            Invoke-DuplicateMagick @('identify', '+ping', '-regard-warnings', '-format', '%m|%w|%h|%n', $second) | Should -Be 'PNG|24|18|1'
        }
        $before = Get-DuplicateSourceState $source
        $result = Invoke-DuplicateRun -Source $source -OutputParent $parent
        $result.Code | Should -Be 0 -Because $result.Text
        $run = Get-DuplicateRun $parent
        $outputName = if ($Media -eq 'image') { 'photo.jpeg' } else { 'clip.mp4' }
        Test-Path -LiteralPath (Join-Path $run ('b\' + $outputName)) | Should -BeFalse
        $log = Get-DuplicateLog $run
        Assert-DuplicateRetainedLink -Log $log -Skipped ('b\' + $name) -RetainedSource ('a\' + $name) -RetainedOutput ('a\' + $outputName)
        if ($Media -eq 'image') {
            Assert-DuplicateJpeg -Path (Join-Path $run ('a\' + $outputName)) -Width 24 -Height 18
            $log | Should -Match 'SUMMARY ConvertedImages=1 CopiedVideos=0 Duplicates=1 Unsupported=0 Errors=0'
        } else {
            (Get-FileHash -LiteralPath (Join-Path $run ('a\' + $outputName)) -Algorithm SHA256).Hash | Should -Be (Get-FileHash -LiteralPath $first -Algorithm SHA256).Hash
            $log | Should -Match 'SUMMARY ConvertedImages=0 CopiedVideos=1 Duplicates=1 Unsupported=0 Errors=0'
        }
        Get-DuplicateSourceState $source | Should -Be $before
    }

    It 'T027 keeps filename matching case-normalized while reporting the exact retained source and output' {
        $source = New-DuplicateDirectory 'case-source'
        $parent = New-DuplicateDirectory 'case-output'
        $null = Write-DuplicateInput -Root $source -RelativePath 'a\PHOTO.PNG' -Bytes $smallBytes
        $null = Write-DuplicateInput -Root $source -RelativePath 'b\photo.png' -Bytes $smallBytes
        $before = Get-DuplicateSourceState $source
        $result = Invoke-DuplicateRun -Source $source -OutputParent $parent
        $result.Code | Should -Be 0 -Because $result.Text
        $run = Get-DuplicateRun $parent
        Assert-DuplicateJpeg -Path (Join-Path $run 'a\PHOTO.jpeg') -Width 24 -Height 18
        Test-Path -LiteralPath (Join-Path $run 'b\photo.jpeg') | Should -BeFalse
        $log = Get-DuplicateLog $run
        Assert-DuplicateRetainedLink -Log $log -Skipped 'b\photo.png' -RetainedSource 'a\PHOTO.PNG' -RetainedOutput 'a\PHOTO.jpeg'
        $log | Should -Match 'SUMMARY ConvertedImages=1 CopiedVideos=0 Duplicates=1 Unsupported=0 Errors=0'
        Get-DuplicateSourceState $source | Should -Be $before
    }

    It 'T027 retains same-name same-length inputs with distinct modification timestamps' {
        $source = New-DuplicateDirectory 'timestamp-source'
        $parent = New-DuplicateDirectory 'timestamp-output'
        $null = Write-DuplicateInput -Root $source -RelativePath 'a\photo.png' -Bytes $smallBytes
        $null = Write-DuplicateInput -Root $source -RelativePath 'b\photo.png' -Bytes $smallBytes -Modified $fixedUtc.AddSeconds(1)
        $before = Get-DuplicateSourceState $source
        $result = Invoke-DuplicateRun -Source $source -OutputParent $parent
        $result.Code | Should -Be 0 -Because $result.Text
        $run = Get-DuplicateRun $parent
        foreach ($relative in @('a\photo.jpeg', 'b\photo.jpeg')) { Assert-DuplicateJpeg -Path (Join-Path $run $relative) -Width 24 -Height 18 }
        Get-DuplicateLog $run | Should -Match 'SUMMARY ConvertedImages=2 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=0'
        Get-DuplicateSourceState $source | Should -Be $before
    }

    It 'T027 retains the same ordinal source and output across shuffled inventories under <Culture>' -ForEach @(
        @{ Culture = 'en-US' }, @{ Culture = 'de-DE' }, @{ Culture = 'tr-TR' }, @{ Culture = 'fi-FI' }, @{ Culture = '' }
    ) {
        $source = New-DuplicateDirectory 'ordered-source'
        foreach ($relative in @('a\IMAGE.PNG', 'b\image.png', 'c\Image.Png')) {
            $null = Write-DuplicateInput -Root $source -RelativePath $relative -Bytes $smallBytes
        }
        $before = Get-DuplicateSourceState $source
        $trace = [pscustomobject]@{ Seed = 0 }
        Mock Get-WinImgSourceTree {
            param([string]$SourceRoot)
            $tree = & $realSourceTree -SourceRoot $SourceRoot
            $random = New-Object Random($trace.Seed)
            foreach ($property in @('Files', 'Directories', 'TopLevelNames')) {
                $values = @($tree.$property)
                for ($index = $values.Count - 1; $index -gt 0; $index--) {
                    $other = $random.Next($index + 1)
                    $saved = $values[$index]; $values[$index] = $values[$other]; $values[$other] = $saved
                }
                $tree.$property = $values
            }
            return $tree
        }
        $priorCulture = [Threading.Thread]::CurrentThread.CurrentCulture
        $priorUiCulture = [Threading.Thread]::CurrentThread.CurrentUICulture
        try {
            [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo($Culture)
            [Threading.Thread]::CurrentThread.CurrentUICulture = [Globalization.CultureInfo]::GetCultureInfo($Culture)
            $baseline = $null
            foreach ($seed in 1..3) {
                $trace.Seed = $seed
                $parent = New-DuplicateDirectory 'ordered-output'
                $result = Invoke-DuplicateRun -Source $source -OutputParent $parent
                $result.Code | Should -Be 0 -Because $result.Text
                $run = Get-DuplicateRun $parent
                $log = Get-DuplicateLog $run
                Assert-DuplicateJpeg -Path (Join-Path $run 'a\IMAGE.jpeg') -Width 24 -Height 18
                Test-Path -LiteralPath (Join-Path $run 'b\image.jpeg') | Should -BeFalse
                Test-Path -LiteralPath (Join-Path $run 'c\Image.jpeg') | Should -BeFalse
                Assert-DuplicateRetainedLink -Log $log -Skipped 'b\image.png' -RetainedSource 'a\IMAGE.PNG' -RetainedOutput 'a\IMAGE.jpeg'
                Assert-DuplicateRetainedLink -Log $log -Skipped 'c\Image.Png' -RetainedSource 'a\IMAGE.PNG' -RetainedOutput 'a\IMAGE.jpeg'
                $signature = (Get-DuplicateSkipLines $log) -join "`n"
                if ($null -eq $baseline) { $baseline = $signature }
                else { $signature | Should -Be $baseline }
                $log | Should -Match 'SUMMARY ConvertedImages=1 CopiedVideos=0 Duplicates=2 Unsupported=0 Errors=0'
            }
        } finally {
            [Threading.Thread]::CurrentThread.CurrentCulture = $priorCulture
            [Threading.Thread]::CurrentThread.CurrentUICulture = $priorUiCulture
        }
        Get-DuplicateSourceState $source | Should -Be $before
    }

    It 'T027 registers a verified finalized above-target JPEG as the retained successful source with its warning' {
        $source = New-DuplicateDirectory 'above-cap-source'
        $parent = New-DuplicateDirectory 'above-cap-output'
        $first = Write-DuplicateInput -Root $source -RelativePath 'a\photo.png' -Bytes $smallBytes
        $null = Write-DuplicateInput -Root $source -RelativePath 'b\photo.png' -Bytes $smallBytes
        $before = Get-DuplicateSourceState $source
        $trace = [pscustomobject]@{ Calls = 0; Sources = New-Object 'Collections.Generic.List[string]'; SnapshotSources = @{} }
        Mock New-WinImgSourceSnapshot {
            param([string]$SourcePath, [string]$WorkRoot, [long]$ExpectedLength, [DateTime]$ExpectedModified, [Collections.Generic.List[object]]$OwnedCandidates)
            $snapshot = & $realSourceSnapshot -SourcePath $SourcePath -WorkRoot $WorkRoot -ExpectedLength $ExpectedLength -ExpectedModified $ExpectedModified -OwnedCandidates $OwnedCandidates
            $trace.SnapshotSources[(Get-WinImgNativeOutputPath $snapshot)] = $SourcePath
            return $snapshot
        }
        $runner = {
            param([string]$Executable, [string[]]$Arguments)
            $trace.Calls++
            $originalSource = $trace.SnapshotSources[$Arguments[2]]
            if (-not $originalSource) { throw 'Duplicate conversion input did not match an actual owned source snapshot.' }
            $trace.Sources.Add($originalSource)
            $candidate = Assert-DuplicateOwnedCandidate -Path $Arguments[-1].Substring('JPEG:'.Length) -OutputParent $parent
            [IO.File]::WriteAllBytes($candidate, $jpegBytes)
            return 0
        }
        $result = Invoke-DuplicateRun -Source $source -OutputParent $parent -ProcessRunner $runner -MaxBytes 1
        $result.Code | Should -Be 0 -Because $result.Text
        $trace.Calls | Should -Be 6
        @($trace.Sources | Where-Object { $_ -ne $first }).Count | Should -Be 0
        $run = Get-DuplicateRun $parent
        $target = Join-Path $run 'a\photo.jpeg'
        Assert-DuplicateJpeg -Path $target -Width 24 -Height 18
        (Get-Item -LiteralPath $target).Length | Should -Be $jpegBytes.Length
        (Get-Item -LiteralPath $target).Length | Should -BeGreaterThan 1
        Test-Path -LiteralPath (Join-Path $run 'b\photo.jpeg') | Should -BeFalse
        $log = Get-DuplicateLog $run
        $log | Should -Match 'WARN: Could not reach target; best-effort saved'
        $log | Should -Match ('\[' + [regex]::Escape(('{0:n0}' -f $jpegBytes.Length)) + ' bytes, Scale=50%\]')
        Assert-DuplicateRetainedLink -Log $log -Skipped 'b\photo.png' -RetainedSource 'a\photo.png' -RetainedOutput 'a\photo.jpeg'
        $log | Should -Match 'retained status: ConvertedWithWarning'
        $log | Should -Match 'SUMMARY ConvertedImages=1 CopiedVideos=0 Duplicates=1 Unsupported=0 Errors=0'
        Get-DuplicateSourceState $source | Should -Be $before
    }
}
