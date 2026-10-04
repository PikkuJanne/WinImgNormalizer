BeforeAll {
    $repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    $application = Join-Path $repository 'WinImgNormalizer.ps1'

    function Assert-TransactionNoReparseAncestors {
        param([string]$Path)
        $current = [IO.DirectoryInfo]::new([IO.Path]::GetFullPath($Path))
        while ($null -ne $current) {
            if ($current.Exists -and (($current.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0)) {
                throw 'Transaction test ownership path includes a reparse point.'
            }
            $current = $current.Parent
        }
    }

    $scratchParent = [IO.Path]::GetFullPath((Join-Path $repository '.scratch'))
    Assert-TransactionNoReparseAncestors $scratchParent
    foreach ($pictures in @([Environment]::GetFolderPath('MyPictures'), (Join-Path $env:USERPROFILE 'Pictures'))) {
        if ([string]::IsNullOrWhiteSpace($pictures)) { continue }
        $picturesFull = [IO.Path]::GetFullPath($pictures).TrimEnd('\', '/')
        if ($scratchParent.Equals($picturesFull, [StringComparison]::OrdinalIgnoreCase) -or
            $scratchParent.StartsWith($picturesFull + '\', [StringComparison]::OrdinalIgnoreCase) -or
            $picturesFull.StartsWith($scratchParent + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Transaction test scratch overlaps a real Pictures location.'
        }
    }
    $git = Get-Command git -CommandType Application -ErrorAction Stop | Select-Object -First 1
    & $git.Source -C $repository check-ignore --quiet --no-index -- (Join-Path $scratchParent 'transaction-ignore-probe')
    if ($LASTEXITCODE -ne 0) { throw 'Transaction synthetic scratch must already be ignored.' }
    $ownedRoot = Join-Path $scratchParent ('M1-T04-' + [Guid]::NewGuid().ToString('N'))
    if (Test-Path -LiteralPath $ownedRoot) { throw 'Transaction ownership directory already exists.' }
    $null = [IO.Directory]::CreateDirectory($ownedRoot)
    [IO.File]::WriteAllText((Join-Path $ownedRoot '.winimg-fixture-root'), 'M1-T04 synthetic test ownership')

    function New-TransactionDirectory {
        param([string]$Label)
        $path = [IO.Path]::GetFullPath((Join-Path $ownedRoot ($Label + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))))
        if (-not $path.StartsWith($ownedRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Transaction test path escapes its owned scratch directory.'
        }
        Assert-TransactionNoReparseAncestors $path
        if (Test-Path -LiteralPath $path) { throw 'Transaction test directory already exists.' }
        $null = [IO.Directory]::CreateDirectory($path)
        return $path
    }

    function ConvertFrom-TransactionNativePath {
        param([string]$Path)
        if ($Path.StartsWith('\\?\UNC\', [StringComparison]::Ordinal)) { return '\\' + $Path.Substring(8) }
        if ($Path.StartsWith('\\?\', [StringComparison]::Ordinal)) { return $Path.Substring(4) }
        return $Path
    }

    function Get-TransactionSourceState {
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

    function Get-TransactionRun {
        param([string]$Parent)
        $runs = @(Get-ChildItem -LiteralPath $Parent -Directory)
        if ($runs.Count -ne 1) { throw 'Expected exactly one transaction-test run directory.' }
        return $runs[0].FullName
    }

    function Get-TransactionLog {
        param([string]$Run)
        $logs = @(Get-ChildItem -LiteralPath $Run -Recurse -File -Filter '*.log')
        if ($logs.Count -ne 1) { throw 'Expected exactly one transaction-test log.' }
        return Get-Content -LiteralPath $logs[0].FullName -Raw
    }

    function Assert-TransactionOwnedCandidate {
        param([string]$CandidatePath, [string]$OutputParent)
        $run = Get-TransactionRun $OutputParent
        $work = Join-Path $run '.WinImgNormalizer\work'
        $ordinary = ConvertFrom-TransactionNativePath $CandidatePath
        if (-not $ordinary.StartsWith($work + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Controlled transaction runner was directed outside the generated work namespace.'
        }
        Assert-TransactionNoReparseAncestors ([IO.Path]::GetDirectoryName($ordinary))
        return $ordinary
    }

    function Invoke-TransactionRun {
        param([string]$Source, [string]$OutputParent, [scriptblock]$ProcessRunner, [long]$MaxBytes = 1MB)
        $invoke = @{ Source = $Source; OutputParent = $OutputParent; MagickPath = $magick; MaxBytes = $MaxBytes }
        if ($ProcessRunner) { $invoke.ProcessRunner = $ProcessRunner }
        $observed = @(& Invoke-WinImgNormalizer @invoke 6>&1 3>&1 2>&1)
        $codes = @($observed | Where-Object { $_ -is [int] -or $_ -is [long] })
        if ($codes.Count -ne 1) { throw ('Expected one callable transaction exit status: ' + ($observed -join "`n")) }
        return [pscustomobject]@{ Code = $codes[0]; Text = (@($observed | Where-Object { $_ -isnot [int] -and $_ -isnot [long] }) -join "`n") }
    }

    if ([string]::IsNullOrWhiteSpace($env:WINIMG_TEST_MAGICK) -or -not [IO.Path]::IsPathRooted($env:WINIMG_TEST_MAGICK)) {
        throw 'WINIMG_TEST_MAGICK must explicitly select the verified test magick.exe.'
    }
    $magick = [IO.Path]::GetFullPath($env:WINIMG_TEST_MAGICK)
    Assert-TransactionNoReparseAncestors $magick
    if (-not (Test-Path -LiteralPath $magick -PathType Leaf) -or [IO.Path]::GetFileName($magick) -ne 'magick.exe') {
        throw 'The explicitly selected transaction test magick.exe is missing.'
    }

    function Invoke-TransactionMagick {
        param([string[]]$Arguments)
        $result = & $magick @Arguments 2>&1
        $code = $LASTEXITCODE
        if ($code -ne 0) { throw ('Transaction test ImageMagick exited {0}: {1}' -f $code, ($result -join "`n")) }
        return ($result -join "`n")
    }

    $jpegFixture = Join-Path $ownedRoot 'synthetic-noise.jpeg'
    $pngFixture = Join-Path $ownedRoot 'synthetic-gradient.png'
    $truncatedFixture = Join-Path $ownedRoot 'truncated.jpeg'
    $null = Invoke-TransactionMagick @('-quiet', '-seed', '4319', '-size', '96x64', 'xc:gray', '+noise', 'Random', '-depth', '8', '-quality', '90', ('JPEG:' + $jpegFixture))
    $null = Invoke-TransactionMagick @('-quiet', '-size', '24x18', 'gradient:#205080-#E0A070', '-depth', '8', ('PNG:' + $pngFixture))
    $jpegBytes = [IO.File]::ReadAllBytes($jpegFixture)
    $pngBytes = [IO.File]::ReadAllBytes($pngFixture)
    $truncatedBytes = [byte[]]$jpegBytes[0..([int]($jpegBytes.Length * 0.7))]
    [IO.File]::WriteAllBytes($truncatedFixture, $truncatedBytes)
    $videoBytes = New-Object byte[] 65537
    for ($offset = 0; $offset -lt $videoBytes.Length; $offset++) { $videoBytes[$offset] = [byte](($offset * 31 + 13) % 256) }
    . $application
    $realImageCandidate = (Get-Command New-WinImgImageCandidate -CommandType Function).ScriptBlock
    $realVideoCopy = (Get-Command Copy-WinImgVideoToCandidate -CommandType Function).ScriptBlock
    $realOutputAvailable = (Get-Command Assert-WinImgOutputAvailable -CommandType Function).ScriptBlock
}

Describe 'M1-T04 native result and full candidate validation (T020)' {
    It 'T020 fixture has readable JPEG header dimensions but fails full pixel decoding' {
        $priorPreference = $ErrorActionPreference
        try {
            $ErrorActionPreference = 'Continue'
            $header = @(& $magick identify -ping -format '%m|%w|%h|%n' $truncatedFixture 2>$null)
            ($header -join '') | Should -Be 'JPEG|96|64|1'
            & $magick identify +ping -regard-warnings $truncatedFixture 1>$null 2>$null
            $LASTEXITCODE | Should -Not -Be 0
        } finally { $ErrorActionPreference = $priorPreference }
        { Test-WinImgImageCandidate -CandidatePath $truncatedFixture -MagickPath $magick } | Should -Throw
    }

    It 'T020 accepts a real JPEG only after format, positive dimensions and full decode verification' {
        $result = Test-WinImgImageCandidate -CandidatePath $jpegFixture -MagickPath $magick
        $result.BytesOut | Should -Be $jpegBytes.Length
        $result.Width | Should -Be 96
        $result.Height | Should -Be 64
        $null = Invoke-TransactionMagick @('-regard-warnings', $jpegFixture, 'null:')
    }

    It 'T020 fully validates and finalizes a real neutral candidate longer than MAX_PATH' {
        $work = New-TransactionDirectory 'long-validation'
        $workRoot = Join-Path $work ('w' * 128)
        $null = [IO.Directory]::CreateDirectory($workRoot)
        $candidate = New-WinImgImageCandidate -WorkRoot $workRoot
        $candidate.Length | Should -BeGreaterOrEqual 260
        $nativeCandidate = Get-WinImgNativeOutputPath $candidate
        $null = Invoke-TransactionMagick @('-quiet', $jpegFixture, ('JPEG:' + $nativeCandidate))
        $result = Test-WinImgImageCandidate -CandidatePath $candidate -MagickPath $magick
        $result.BytesOut | Should -BeGreaterThan 0
        $result.Width | Should -Be 96
        $result.Height | Should -Be 64
        $target = Join-Path $work 'validated.jpeg'
        Move-WinImgPlannedImage -CandidatePath $candidate -DestinationPath $target
        Test-Path -LiteralPath $candidate | Should -BeFalse
        Invoke-TransactionMagick @('identify', '-format', '%m|%w|%h|%n', $target) | Should -Be 'JPEG|96|64|1'
        $null = Invoke-TransactionMagick @('-regard-warnings', $target, 'null:')
    }

    It 'T020 rejects numbered native JPEG sequence outputs rather than treating an empty neutral candidate as one frame' {
        $source = New-TransactionDirectory 'sequence-source'
        $parent = New-TransactionDirectory 'sequence-output'
        [IO.File]::WriteAllBytes((Join-Path $source 'single.png'), $pngBytes)
        $before = Get-TransactionSourceState $source
        $trace = [pscustomobject]@{ Calls = 0; Numbered = @(); Candidates = New-Object 'Collections.Generic.List[string]' }
        $runner = {
            param([string]$Executable, [string[]]$Arguments)
            $trace.Calls++
            $nativeCandidate = $Arguments[-1].Substring('JPEG:'.Length)
            $candidate = Assert-TransactionOwnedCandidate -CandidatePath $nativeCandidate -OutputParent $parent
            $trace.Candidates.Add($candidate)
            if ($trace.Calls -ne 1) { return 1 }
            & $Executable -quiet $jpegFixture $jpegFixture ('JPEG:' + $nativeCandidate) 1>$null 2>$null
            $code = $LASTEXITCODE
            $trace.Numbered = @(Get-ChildItem -LiteralPath ([IO.Path]::GetDirectoryName($candidate)) -File -Filter 'image-*.jpeg' | ForEach-Object FullName)
            return $code
        }
        $result = Invoke-TransactionRun -Source $source -OutputParent $parent -ProcessRunner $runner
        $result.Code | Should -Be 2 -Because $result.Text
        $trace.Numbered.Count | Should -Be 2
        foreach ($path in $trace.Numbered) {
            Invoke-TransactionMagick @('identify', '-format', '%m|%w|%h|%n', $path) | Should -Be 'JPEG|96|64|1'
            $null = Invoke-TransactionMagick @('-regard-warnings', $path, 'null:')
        }
        foreach ($candidate in $trace.Candidates) { Test-Path -LiteralPath $candidate | Should -BeFalse }
        $run = Get-TransactionRun $parent
        Test-Path -LiteralPath (Join-Path $run 'single.jpeg') | Should -BeFalse
        Get-TransactionLog $run | Should -Match 'SUMMARY ConvertedImages=0 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=1'
        Get-TransactionSourceState $source | Should -Be $before
    }

    It 'T020 refuses <Outcome> without finalizing or counting the image as converted' -ForEach @(
        @{ Outcome = 'no-file' }, @{ Outcome = 'empty' }, @{ Outcome = 'malformed' },
        @{ Outcome = 'wrong-format' }, @{ Outcome = 'truncated' }, @{ Outcome = 'nonzero-with-valid-file' },
        @{ Outcome = 'missing-exit' }, @{ Outcome = 'string-zero' }, @{ Outcome = 'extra-output' },
        @{ Outcome = 'timed-out' }, @{ Outcome = 'cancelled' }
    ) {
        $source = New-TransactionDirectory 'invalid-source'
        $parent = New-TransactionDirectory 'invalid-output'
        [IO.File]::WriteAllBytes((Join-Path $source 'single.png'), $pngBytes)
        $before = Get-TransactionSourceState $source
        $trace = [pscustomobject]@{ Calls = 0; Paths = New-Object 'Collections.Generic.List[string]'; InitialLengths = New-Object 'Collections.Generic.List[long]' }
        $runner = {
            param([string]$Executable, [string[]]$Arguments)
            $trace.Calls++
            $candidate = Assert-TransactionOwnedCandidate -CandidatePath $Arguments[-1].Substring('JPEG:'.Length) -OutputParent $parent
            $trace.Paths.Add($candidate)
            $trace.InitialLengths.Add(([IO.FileInfo]::new($candidate)).Length)
            switch ($Outcome) {
                'no-file' { [IO.File]::Delete($candidate); return 0 }
                'empty' { return 0 }
                'malformed' { [IO.File]::WriteAllText($candidate, 'This synthetic text is not a JPEG.'); return 0 }
                'wrong-format' { [IO.File]::WriteAllBytes($candidate, $pngBytes); return 0 }
                'truncated' { [IO.File]::WriteAllBytes($candidate, $truncatedBytes); return 0 }
                default { [IO.File]::WriteAllBytes($candidate, $jpegBytes) }
            }
            switch ($Outcome) {
                'nonzero-with-valid-file' { return 9 }
                'missing-exit' { return $null }
                'string-zero' { return '0' }
                'extra-output' { 'Unexpected runner output'; return 0 }
                'timed-out' { return [pscustomobject]@{ ExitCode = 0; TimedOut = $true; Cancelled = $false } }
                'cancelled' { return [pscustomobject]@{ ExitCode = 0; TimedOut = $false; Cancelled = $true } }
            }
        }
        $result = Invoke-TransactionRun -Source $source -OutputParent $parent -ProcessRunner $runner
        $result.Code | Should -Be 2 -Because $result.Text
        $trace.Calls | Should -BeGreaterThan 0
        @($trace.Paths | Select-Object -Unique).Count | Should -Be $trace.Calls
        @($trace.InitialLengths | Where-Object { $_ -ne 0 }).Count | Should -Be 0
        foreach ($candidate in $trace.Paths) { Test-Path -LiteralPath $candidate | Should -BeFalse }
        $run = Get-TransactionRun $parent
        Test-Path -LiteralPath (Join-Path $run 'single.jpeg') | Should -BeFalse
        @(Get-ChildItem -LiteralPath (Join-Path $run '.WinImgNormalizer\work') -Force).Count | Should -Be 0
        Get-TransactionLog $run | Should -Match 'SUMMARY ConvertedImages=0 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=1'
        Get-TransactionSourceState $source | Should -Be $before
    }

    It 'T020 validates a fresh fallback after a native failure rather than rejecting a later valid attempt' {
        $source = New-TransactionDirectory 'fallback-source'
        $parent = New-TransactionDirectory 'fallback-output'
        [IO.File]::WriteAllBytes((Join-Path $source 'single.png'), $pngBytes)
        $before = Get-TransactionSourceState $source
        $trace = [pscustomobject]@{ Calls = 0; Paths = New-Object 'Collections.Generic.List[string]' }
        $runner = {
            param([string]$Executable, [string[]]$Arguments)
            $trace.Calls++
            $candidate = Assert-TransactionOwnedCandidate -CandidatePath $Arguments[-1].Substring('JPEG:'.Length) -OutputParent $parent
            $trace.Paths.Add($candidate)
            if ($trace.Calls -eq 1) { [IO.File]::WriteAllText($candidate, 'Failed native partial.'); return 1 }
            [IO.File]::WriteAllBytes($candidate, $jpegBytes)
            return [pscustomobject]@{ ExitCode = 0; TimedOut = $false; Cancelled = $false }
        }
        $result = Invoke-TransactionRun -Source $source -OutputParent $parent -ProcessRunner $runner
        $result.Code | Should -Be 0 -Because $result.Text
        $trace.Calls | Should -Be 2
        $trace.Paths[0] | Should -Not -Be $trace.Paths[1]
        $run = Get-TransactionRun $parent
        $target = Join-Path $run 'single.jpeg'
        Invoke-TransactionMagick @('identify', '-format', '%m|%w|%h|%n', $target) | Should -Be 'JPEG|96|64|1'
        $null = Invoke-TransactionMagick @('-regard-warnings', $target, 'null:')
        Get-TransactionSourceState $source | Should -Be $before
        Get-TransactionLog $run | Should -Match 'SUMMARY ConvertedImages=1 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=0'
    }
}

Describe 'M1-T04 attempt freshness and final conflicts (T021-T022)' {
    It 'T021 does not reuse an earlier above-cap JPEG after later <Failure> attempts' -ForEach @(
        @{ Failure = 'nonzero' }, @{ Failure = 'zero-with-no-file' }
    ) {
        $source = New-TransactionDirectory 'stale-source'
        $parent = New-TransactionDirectory 'stale-output'
        [IO.File]::WriteAllBytes((Join-Path $source 'single.png'), $pngBytes)
        $before = Get-TransactionSourceState $source
        $trace = [pscustomobject]@{ Calls = 0; Paths = New-Object 'Collections.Generic.List[string]'; InitialLengths = New-Object 'Collections.Generic.List[long]' }
        $runner = {
            param([string]$Executable, [string[]]$Arguments)
            $trace.Calls++
            $candidate = Assert-TransactionOwnedCandidate -CandidatePath $Arguments[-1].Substring('JPEG:'.Length) -OutputParent $parent
            $trace.Paths.Add($candidate)
            $trace.InitialLengths.Add(([IO.FileInfo]::new($candidate)).Length)
            if ($trace.Calls -eq 1) { [IO.File]::WriteAllBytes($candidate, $jpegBytes); return 0 }
            if ($Failure -eq 'nonzero') { return 1 }
            [IO.File]::Delete($candidate)
            return 0
        }
        $result = Invoke-TransactionRun -Source $source -OutputParent $parent -MaxBytes 1 -ProcessRunner $runner
        $result.Code | Should -Be 2 -Because $result.Text
        $trace.Calls | Should -BeGreaterThan 1
        @($trace.Paths | Select-Object -Unique).Count | Should -Be $trace.Calls
        @($trace.InitialLengths | Where-Object { $_ -ne 0 }).Count | Should -Be 0
        foreach ($candidate in $trace.Paths) { Test-Path -LiteralPath $candidate | Should -BeFalse }
        $run = Get-TransactionRun $parent
        Test-Path -LiteralPath (Join-Path $run 'single.jpeg') | Should -BeFalse
        Get-TransactionLog $run | Should -Match 'SUMMARY ConvertedImages=0 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=1'
        Get-TransactionSourceState $source | Should -Be $before
    }

    It 'T022 preserves an external <Arrival> arriving after candidate validation and before <Media> finalization' -ForEach @(
        @{ Arrival = 'file'; Media = 'image' }, @{ Arrival = 'directory'; Media = 'image' },
        @{ Arrival = 'file'; Media = 'video' }, @{ Arrival = 'directory'; Media = 'video' }
    ) {
        $source = New-TransactionDirectory 'arrival-source'
        $parent = New-TransactionDirectory 'arrival-output'
        $sourceName = if ($Media -eq 'image') { 'single.png' } else { 'clip.mp4' }
        $bytes = if ($Media -eq 'image') { $pngBytes } else { $videoBytes }
        [IO.File]::WriteAllBytes((Join-Path $source $sourceName), $bytes)
        $before = Get-TransactionSourceState $source
        $trace = [pscustomobject]@{ Checks = 0; ArrivalPath = $null }
        Mock Assert-WinImgOutputAvailable {
            param([string]$Path)
            & $realOutputAvailable $Path
            $trace.Checks++
            if ($trace.Checks -eq 2) {
                $trace.ArrivalPath = $Path
                if ($Arrival -eq 'directory') {
                    $null = [IO.Directory]::CreateDirectory($Path)
                    [IO.File]::WriteAllText((Join-Path $Path 'external.txt'), 'External arrival after validation.')
                } else { [IO.File]::WriteAllText($Path, 'External arrival after validation.') }
            }
        }
        $result = Invoke-TransactionRun -Source $source -OutputParent $parent
        $result.Code | Should -Be 2 -Because $result.Text
        $trace.Checks | Should -Be 2
        if ($Arrival -eq 'directory') {
            [IO.File]::ReadAllText((Join-Path $trace.ArrivalPath 'external.txt')) | Should -Be 'External arrival after validation.'
            @(Get-ChildItem -LiteralPath $trace.ArrivalPath -Force).Count | Should -Be 1
        } else { [IO.File]::ReadAllText($trace.ArrivalPath) | Should -Be 'External arrival after validation.' }
        $run = Get-TransactionRun $parent
        @(Get-ChildItem -LiteralPath (Join-Path $run '.WinImgNormalizer\work') -Force).Count | Should -Be 0
        Get-TransactionLog $run | Should -Match 'SUMMARY ConvertedImages=0 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=1'
        Get-TransactionSourceState $source | Should -Be $before
    }
}

Describe 'M1-T04 staged video identity and source stability (T023)' {
    It 'T023 copies stable opaque video bytes through an owned partial before the final name exists' {
        $source = New-TransactionDirectory 'video-source'
        $parent = New-TransactionDirectory 'video-output'
        $inputPath = Join-Path $source 'clip.mp4'
        [IO.File]::WriteAllBytes($inputPath, $videoBytes)
        $before = Get-TransactionSourceState $source
        $trace = [pscustomobject]@{ Calls = 0; CandidatePath = $null; FinalExistedDuringCopy = $null }
        Mock Copy-WinImgVideoToCandidate {
            param([string]$SourcePath, [string]$CandidatePath)
            $trace.Calls++
            $trace.CandidatePath = Assert-TransactionOwnedCandidate -CandidatePath $CandidatePath -OutputParent $parent
            $trace.FinalExistedDuringCopy = Test-Path -LiteralPath (Join-Path (Get-TransactionRun $parent) 'clip.mp4')
            & $realVideoCopy -SourcePath $SourcePath -CandidatePath $CandidatePath
        }
        $result = Invoke-TransactionRun -Source $source -OutputParent $parent -ProcessRunner { throw 'Video-only input attempted image conversion.' }
        $result.Code | Should -Be 0 -Because $result.Text
        $trace.Calls | Should -Be 1
        $trace.FinalExistedDuringCopy | Should -BeFalse
        $run = Get-TransactionRun $parent
        $target = Join-Path $run 'clip.mp4'
        (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash | Should -Be (Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash
        (Get-Item -LiteralPath $target).Length | Should -Be $videoBytes.Length
        (Get-Item -LiteralPath $target).LastWriteTimeUtc.Ticks | Should -Be (Get-Item -LiteralPath $inputPath).LastWriteTimeUtc.Ticks
        (Get-Item -LiteralPath $target).CreationTimeUtc.Ticks | Should -Be (Get-Item -LiteralPath $inputPath).CreationTimeUtc.Ticks
        Test-Path -LiteralPath $trace.CandidatePath | Should -BeFalse
        @(Get-ChildItem -LiteralPath (Join-Path $run '.WinImgNormalizer\work') -Force).Count | Should -Be 0
        Get-TransactionSourceState $source | Should -Be $before
        Get-TransactionLog $run | Should -Match 'SUMMARY ConvertedImages=0 CopiedVideos=1 Duplicates=0 Unsupported=0 Errors=0'
    }

    It 'T023 refuses <Change> during staged video copying' -ForEach @(
        @{ Change = 'length-change' }, @{ Change = 'timestamp-change' },
        @{ Change = 'partial-with-success' }, @{ Change = 'partial-with-exception' }
    ) {
        $source = New-TransactionDirectory 'changed-video-source'
        $parent = New-TransactionDirectory 'changed-video-output'
        $inputPath = Join-Path $source 'clip.mp4'
        [IO.File]::WriteAllBytes($inputPath, $videoBytes)
        $before = Get-TransactionSourceState $source
        $trace = [pscustomobject]@{ Calls = 0; CandidatePath = $null; InjectedSourceState = $null }
        Mock Copy-WinImgVideoToCandidate {
            param([string]$SourcePath, [string]$CandidatePath)
            $trace.Calls++
            $trace.CandidatePath = Assert-TransactionOwnedCandidate -CandidatePath $CandidatePath -OutputParent $parent
            if ($Change.StartsWith('partial-', [StringComparison]::Ordinal)) {
                [IO.File]::WriteAllBytes($CandidatePath, [byte[]]$videoBytes[0..127])
                if ($Change -eq 'partial-with-exception') { throw 'Controlled video copy interruption.' }
            } else {
                & $realVideoCopy -SourcePath $SourcePath -CandidatePath $CandidatePath
                if ($Change -eq 'length-change') { [IO.File]::WriteAllBytes($SourcePath, [byte[]](11, 42, 255)) }
                else { [IO.File]::SetLastWriteTimeUtc($SourcePath, ([IO.File]::GetLastWriteTimeUtc($SourcePath)).AddSeconds(10)) }
                $trace.InjectedSourceState = Get-TransactionSourceState $source
            }
        }
        $result = Invoke-TransactionRun -Source $source -OutputParent $parent -ProcessRunner { throw 'Video-only input attempted image conversion.' }
        $result.Code | Should -Be 2 -Because $result.Text
        $trace.Calls | Should -Be 1
        $run = Get-TransactionRun $parent
        Test-Path -LiteralPath (Join-Path $run 'clip.mp4') | Should -BeFalse
        Test-Path -LiteralPath $trace.CandidatePath | Should -BeFalse
        @(Get-ChildItem -LiteralPath (Join-Path $run '.WinImgNormalizer\work') -Force).Count | Should -Be 0
        Get-TransactionLog $run | Should -Match 'SUMMARY ConvertedImages=0 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=1'
        if ($trace.InjectedSourceState) { Get-TransactionSourceState $source | Should -Be $trace.InjectedSourceState }
        else { Get-TransactionSourceState $source | Should -Be $before }
    }

    It 'T023 rejects a reserved partial that became nonempty before opening the copy stream' {
        $work = New-TransactionDirectory 'exclusive-video-copy'
        $inputPath = Join-Path $work 'clip.mp4'
        $candidate = Join-Path $work 'video.partial'
        [IO.File]::WriteAllBytes($inputPath, $videoBytes)
        [IO.File]::WriteAllText($candidate, 'Existing partial path remains untouched.')
        { Copy-WinImgVideoToCandidate -SourcePath $inputPath -CandidatePath $candidate } | Should -Throw
        [IO.File]::ReadAllText($candidate) | Should -Be 'Existing partial path remains untouched.'
        (Get-Item -LiteralPath $inputPath).Length | Should -Be $videoBytes.Length
    }
}

Describe 'M1-T04 exact temporary ownership (T024)' {
    It 'T024 preserves an unrelated <Media> candidate arrival before exclusive file reservation' -ForEach @(
        @{ Media = 'image' }, @{ Media = 'video' }
    ) {
        $source = New-TransactionDirectory 'reservation-arrival-source'
        $parent = New-TransactionDirectory 'reservation-arrival-output'
        $name = if ($Media -eq 'image') { 'single.png' } else { 'clip.mp4' }
        $bytes = if ($Media -eq 'image') { $pngBytes } else { $videoBytes }
        [IO.File]::WriteAllBytes((Join-Path $source $name), $bytes)
        $before = Get-TransactionSourceState $source
        $trace = [pscustomobject]@{ Arrivals = New-Object 'Collections.Generic.List[string]'; ProcessCalls = 0; CopyCalls = 0 }
        Mock New-WinImgImageCandidate {
            param([string]$WorkRoot)
            $candidate = & $realImageCandidate -WorkRoot $WorkRoot
            $arrival = if ($Media -eq 'image') { $candidate } else { Join-Path ([IO.Path]::GetDirectoryName($candidate)) 'video.partial' }
            [IO.File]::WriteAllText($arrival, 'External candidate arrival remains untouched.')
            $trace.Arrivals.Add($arrival)
            return $candidate
        }
        Mock Copy-WinImgVideoToCandidate { $trace.CopyCalls++; throw 'Unexpected copy of an unowned candidate.' }
        $result = Invoke-TransactionRun -Source $source -OutputParent $parent -ProcessRunner { $trace.ProcessCalls++; return 0 }
        $result.Code | Should -Be 2 -Because $result.Text
        $trace.Arrivals.Count | Should -BeGreaterThan 0
        $trace.ProcessCalls | Should -Be 0
        $trace.CopyCalls | Should -Be 0
        foreach ($arrival in $trace.Arrivals) { [IO.File]::ReadAllText($arrival) | Should -Be 'External candidate arrival remains untouched.' }
        $run = Get-TransactionRun $parent
        $finalName = if ($Media -eq 'image') { 'single.jpeg' } else { 'clip.mp4' }
        Test-Path -LiteralPath (Join-Path $run $finalName) | Should -BeFalse
        Get-TransactionSourceState $source | Should -Be $before
        Get-TransactionLog $run | Should -Match 'SUMMARY ConvertedImages=0 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=1'
    }

    It 'T024 removes only known image candidates after failure and preserves numbered and unrelated neighbors' {
        $source = New-TransactionDirectory 'ownership-image-source'
        $parent = New-TransactionDirectory 'ownership-image-output'
        [IO.File]::WriteAllBytes((Join-Path $source 'single.png'), $pngBytes)
        $before = Get-TransactionSourceState $source
        $trace = [pscustomobject]@{ Candidates = New-Object 'Collections.Generic.List[string]'; Neighbors = New-Object 'Collections.Generic.List[string]' }
        $runner = {
            param([string]$Executable, [string[]]$Arguments)
            $candidate = Assert-TransactionOwnedCandidate -CandidatePath $Arguments[-1].Substring('JPEG:'.Length) -OutputParent $parent
            $trace.Candidates.Add($candidate)
            [IO.File]::WriteAllBytes($candidate, $jpegBytes)
            foreach ($name in @('image-0.jpeg', 'unrelated.txt')) {
                $neighbor = Join-Path ([IO.Path]::GetDirectoryName($candidate)) $name
                [IO.File]::WriteAllText($neighbor, 'Unrelated synthetic neighbor remains untouched.')
                $trace.Neighbors.Add($neighbor)
            }
            return 1
        }
        $result = Invoke-TransactionRun -Source $source -OutputParent $parent -ProcessRunner $runner
        $result.Code | Should -Be 2 -Because $result.Text
        $trace.Candidates.Count | Should -BeGreaterThan 0
        foreach ($candidate in $trace.Candidates) { Test-Path -LiteralPath $candidate | Should -BeFalse }
        foreach ($neighbor in $trace.Neighbors) { [IO.File]::ReadAllText($neighbor) | Should -Be 'Unrelated synthetic neighbor remains untouched.' }
        $run = Get-TransactionRun $parent
        Test-Path -LiteralPath (Join-Path $run 'single.jpeg') | Should -BeFalse
        Get-TransactionSourceState $source | Should -Be $before
        Get-TransactionLog $run | Should -Match 'SUMMARY ConvertedImages=0 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=1'
    }

    It 'T024 cleans an interrupted owned video partial while preserving unrelated files beside it' {
        $source = New-TransactionDirectory 'ownership-video-source'
        $parent = New-TransactionDirectory 'ownership-video-output'
        [IO.File]::WriteAllBytes((Join-Path $source 'clip.mp4'), $videoBytes)
        $before = Get-TransactionSourceState $source
        $trace = [pscustomobject]@{ CandidatePath = $null; NeighborPath = $null; WorkNeighborPath = $null }
        Mock Copy-WinImgVideoToCandidate {
            param([string]$SourcePath, [string]$CandidatePath)
            $trace.CandidatePath = Assert-TransactionOwnedCandidate -CandidatePath $CandidatePath -OutputParent $parent
            [IO.File]::WriteAllBytes($CandidatePath, [byte[]]$videoBytes[0..127])
            $trace.NeighborPath = Join-Path ([IO.Path]::GetDirectoryName($CandidatePath)) 'video.partial.keep'
            $trace.WorkNeighborPath = Join-Path (Join-Path (Get-TransactionRun $parent) '.WinImgNormalizer\work') 'user.keep'
            foreach ($path in @($trace.NeighborPath, $trace.WorkNeighborPath)) { [IO.File]::WriteAllText($path, 'Unrelated synthetic video neighbor remains untouched.') }
            throw 'Controlled interruption after partial video write.'
        }
        $result = Invoke-TransactionRun -Source $source -OutputParent $parent -ProcessRunner { throw 'Video-only input attempted image conversion.' }
        $result.Code | Should -Be 2 -Because $result.Text
        Test-Path -LiteralPath $trace.CandidatePath | Should -BeFalse
        foreach ($path in @($trace.NeighborPath, $trace.WorkNeighborPath)) { [IO.File]::ReadAllText($path) | Should -Be 'Unrelated synthetic video neighbor remains untouched.' }
        $run = Get-TransactionRun $parent
        Test-Path -LiteralPath (Join-Path $run 'clip.mp4') | Should -BeFalse
        Get-TransactionSourceState $source | Should -Be $before
        Get-TransactionLog $run | Should -Match 'SUMMARY ConvertedImages=0 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=1'
    }
}
