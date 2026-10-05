BeforeAll {
    $repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    $scratch = Join-Path $repository '.scratch'

    function Assert-CodecAncestors {
        param([string]$Path)
        $entry = [IO.DirectoryInfo]::new([IO.Path]::GetFullPath($Path))
        while ($null -ne $entry) {
            if ($entry.Exists -and (($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0)) {
                throw 'Codec fixture ownership crosses a reparse point.'
            }
            $entry = $entry.Parent
        }
    }

    Assert-CodecAncestors $scratch
    foreach ($pictures in @([Environment]::GetFolderPath('MyPictures'), (Join-Path $env:USERPROFILE 'Pictures'))) {
        if ([string]::IsNullOrWhiteSpace($pictures)) { continue }
        $protected = [IO.Path]::GetFullPath($pictures).TrimEnd('\', '/')
        if ($scratch.Equals($protected, [StringComparison]::OrdinalIgnoreCase) -or
            $scratch.StartsWith($protected + '\', [StringComparison]::OrdinalIgnoreCase) -or
            $protected.StartsWith($scratch + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Codec fixture scratch overlaps a real Pictures location.'
        }
    }
    $git = Get-Command git -CommandType Application -ErrorAction Stop | Select-Object -First 1
    & $git.Source -C $repository check-ignore --quiet --no-index -- (Join-Path $scratch 'codec-ignore-probe')
    if ($LASTEXITCODE -ne 0) { throw 'Codec fixture scratch must already be ignored.' }
    $ownedRoot = Join-Path $scratch ('M3-T06-codecs-' + [Guid]::NewGuid().ToString('N'))
    if (Test-Path -LiteralPath $ownedRoot) { throw 'Codec fixture ownership directory already exists.' }
    $null = [IO.Directory]::CreateDirectory($ownedRoot)
    [IO.File]::WriteAllText((Join-Path $ownedRoot '.winimg-fixture-root'), 'M3-T06 synthetic codec fixture ownership')
    $fixed = [DateTime]::Parse('2020-02-03T04:05:06Z').ToUniversalTime()
    $codecObservations = New-Object 'Collections.Generic.List[object]'

    function New-CodecDirectory {
        param([string]$Label)
        $path = [IO.Path]::GetFullPath((Join-Path $ownedRoot ($Label + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))))
        if (-not $path.StartsWith($ownedRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Codec fixture escaped its owned scratch directory.'
        }
        Assert-CodecAncestors $path
        if (Test-Path -LiteralPath $path) { throw 'Codec fixture directory already exists.' }
        $null = [IO.Directory]::CreateDirectory($path)
        return $path
    }

    if ([string]::IsNullOrWhiteSpace($env:WINIMG_TEST_MAGICK) -or -not [IO.Path]::IsPathRooted($env:WINIMG_TEST_MAGICK)) {
        throw 'WINIMG_TEST_MAGICK must explicitly select the verified test magick.exe.'
    }
    $magick = [IO.Path]::GetFullPath($env:WINIMG_TEST_MAGICK)
    Assert-CodecAncestors $magick
    if (-not (Test-Path -LiteralPath $magick -PathType Leaf) -or [IO.Path]::GetFileName($magick) -ne 'magick.exe') {
        throw 'The explicitly selected codec fixture magick.exe is missing.'
    }
    . (Join-Path $repository 'WinImgNormalizer.ps1')

    function Invoke-CodecMagick {
        param([string[]]$Arguments)
        $native = Invoke-WinImgNativeProcess -Executable $magick -Arguments $Arguments -TimeoutMilliseconds 15000
        if ($native.ExitCode -ne 0 -or $native.StdErr -or -not $native.StreamsComplete -or $native.StdOutTruncated -or $native.StdErrTruncated) {
            throw ('Codec fixture native command failed: ' + $native.StdErr)
        }
        return ($native.StdOut -replace "`r`n", "`n").TrimEnd("`r", "`n")
    }
    $versionText = Invoke-CodecMagick @('-version')
    $formatText = Invoke-CodecMagick @('-list', 'format')

    function Set-CodecSourceTimes {
        param([string]$Path)
        [IO.File]::SetCreationTimeUtc($Path, $fixed)
        [IO.File]::SetLastWriteTimeUtc($Path, $fixed)
    }

    function New-CodecImage {
        param([string]$Path, [string]$Coder)
        if (-not ([IO.Path]::GetFullPath($Path)).StartsWith($ownedRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Codec fixture image escaped ownership.'
        }
        Assert-CodecAncestors ([IO.Path]::GetDirectoryName($Path))
        if (Test-Path -LiteralPath $Path) { throw 'Codec source fixture already exists.' }
        # These are actual format encodings of an owned 32x24 solid RGB patch;
        # the HEIC/HEIF collection bytes and animated codecs are tested in T031.
        $null = Invoke-CodecMagick @('-define', 'registry:filename:literal=true', '-size', '32x24', 'xc:#40A060',
            '-depth', '8', '-strip', ($Coder + ':' + (Get-WinImgNativeOutputPath $Path)))
        Invoke-CodecMagick @('identify', '+ping', '-regard-warnings', '-define', 'registry:filename:literal=true',
            '-format', '%m|%w|%h|%n', (Get-WinImgNativeOutputPath $Path)) | Should -Be ($Coder + '|32|24|1')
        Set-CodecSourceTimes $Path
        $codecObservations.Add([pscustomobject]@{
            Kind = 'actual synthetic codec source'; Extension = [IO.Path]::GetExtension($Path); Format = $Coder
            Width = 32; Height = 24; Recipe = 'ImageMagick -size 32x24 xc:#40A060 -depth 8 -strip explicit coder'
            Sha256 = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
            RelativePath = $Path.Substring($ownedRoot.Length + 1)
        })
    }

    function Get-CodecSourceState {
        param([string]$Source)
        return @(
            Get-ChildItem -LiteralPath $Source -Recurse -Force -File | Sort-Object FullName | ForEach-Object {
                [pscustomobject]@{
                    Path = $_.FullName.Substring($Source.Length + 1)
                    Hash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
                    Length = $_.Length; CreationTicks = $_.CreationTimeUtc.Ticks; ModifiedTicks = $_.LastWriteTimeUtc.Ticks
                    Attributes = [int]$_.Attributes
                }
            }
        ) | ConvertTo-Json -Depth 4 -Compress
    }

    function New-CodecMaskedPreflight {
        param([string]$MaskedFormats)
        $capturedVersion = $versionText
        $capturedFormats = $MaskedFormats
        $capturedExecutable = $magick
        $runner = {
            param([string]$Executable, [string[]]$Arguments)
            if ($Executable -ne $capturedExecutable) { throw 'Controlled codec preflight selected a different executable.' }
            if (($Arguments -join '|') -eq '-version') {
                return [pscustomobject]@{ ExitCode = 0; StdOut = $capturedVersion; StdErr = ''; TimedOut = $false }
            }
            if (($Arguments -join '|') -eq '-list|format') {
                return [pscustomobject]@{ ExitCode = 0; StdOut = $capturedFormats; StdErr = ''; TimedOut = $false }
            }
            throw 'Unexpected controlled codec preflight command.'
        }
        return $runner.GetNewClosure()
    }

    function Set-CodecFormatMode {
        param([string]$Decoder, [string]$Mode)
        $pattern = '(?m)^(\s*' + [regex]::Escape($Decoder) + '\*?\s+(?:\S+\s+)?)([r-][w-][+-])(\s.*)$'
        [regex]::Matches($formatText, $pattern).Count | Should -Be 1
        if ($Mode -eq 'NoRead') {
            $observedMode = [regex]::Match($formatText, $pattern).Groups[2].Value
            $Mode = '-' + $observedMode.Substring(1)
        }
        return [regex]::Replace($formatText, $pattern, ('${1}' + $Mode + '${3}'))
    }

    function Invoke-CodecRun {
        param([string]$Source, [string]$Parent, [scriptblock]$Preflight, [scriptblock]$Runner)
        $script:codecReport = $null
        $parameters = @{
            Source = $Source; OutputParent = $Parent; MagickPath = $magick; NativeTemporaryRoot = $ownedRoot
            ReportObserver = { param($State) $script:codecReport = $State }
        }
        if ($Preflight) { $parameters.PreflightRunner = $Preflight }
        if ($Runner) { $parameters.ProcessRunner = $Runner }
        $observed = @(& Invoke-WinImgNormalizer @parameters 6>&1 3>&1 2>&1)
        $codes = @($observed | Where-Object { $_ -is [int] -or $_ -is [long] })
        $codes.Count | Should -Be 1
        $run = @(Get-ChildItem -LiteralPath $Parent -Force -Directory)
        $run.Count | Should -Be 1
        $logs = @(Get-ChildItem -LiteralPath (Join-Path $run[0].FullName '.WinImgNormalizer\reports') -File -Filter '*.log')
        $logs.Count | Should -Be 1
        return [pscustomobject]@{
            Code = $codes[0]; Text = $observed -join "`n"; Run = $run[0].FullName
            Log = Get-Content -LiteralPath $logs[0].FullName -Raw; State = $script:codecReport
        }
    }

    function Assert-CodecJpeg {
        param([string]$Path)
        Invoke-CodecMagick @('identify', '+ping', '-regard-warnings', '-define', 'registry:filename:literal=true',
            '-format', '%m|%w|%h|%n', (Get-WinImgNativeOutputPath $Path)) | Should -Be 'JPEG|32|24|1'
        $pixel = Invoke-CodecMagick @('-define', 'registry:filename:literal=true', (Get-WinImgNativeOutputPath $Path),
            '-format', '%[fx:round(255*p{16,12}.r)]|%[fx:round(255*p{16,12}.g)]|%[fx:round(255*p{16,12}.b)]', 'info:')
        $rgb = @($pixel.Split('|') | ForEach-Object { [int]$_ })
        $rgb.Count | Should -Be 3
        foreach ($channel in 0..2) { [Math]::Abs($rgb[$channel] - @(64,160,96)[$channel]) | Should -BeLessOrEqual 12 }
        Invoke-CodecMagick @('identify', '+ping', '-verbose', (Get-WinImgNativeOutputPath $Path)) |
            Should -Not -Match '(?m)^\s+(Profiles:|Profile-)'
    }
}

AfterAll {
    if ($codecObservations) {
        [IO.File]::WriteAllText((Join-Path $ownedRoot 'codec-observations.json'),
            (ConvertTo-Json -InputObject $codecObservations.ToArray() -Depth 8), [Text.UTF8Encoding]::new($false))
    }
}

Describe 'M3-T06 mandatory real codec and missing-capability corpus (T065)' {
    It 'T065 fully decodes an actual .<Extension> source and its JPEG derivative while preserving source bytes and timestamps' -ForEach @(
        @{ Extension = 'jpg'; Coder = 'JPEG' }, @{ Extension = 'jpeg'; Coder = 'JPEG' },
        @{ Extension = 'png'; Coder = 'PNG' }, @{ Extension = 'bmp'; Coder = 'BMP' },
        @{ Extension = 'tif'; Coder = 'TIFF' }, @{ Extension = 'tiff'; Coder = 'TIFF' },
        @{ Extension = 'gif'; Coder = 'GIF' }, @{ Extension = 'webp'; Coder = 'WEBP' }
    ) {
        $source = New-CodecDirectory ('actual-' + $Extension)
        $parent = New-CodecDirectory ('output-' + $Extension)
        $path = Join-Path $source ('actual.' + $Extension)
        New-CodecImage -Path $path -Coder $Coder
        $before = Get-CodecSourceState $source
        $result = Invoke-CodecRun -Source $source -Parent $parent
        $result.Code | Should -Be 0 -Because $result.Text
        $result.Log | Should -Match 'SUMMARY ConvertedImages=1 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=0'
        Assert-CodecJpeg (Join-Path $result.Run 'actual.jpeg')
        $result.State.Rows.Count | Should -Be 1
        $result.State.Rows[0].Status | Should -Be 'Converted'
        $result.State.Rows[0].Attempts | Should -Be 1
        @(Get-ChildItem -LiteralPath $result.Run -Recurse -Force -File | Where-Object { $_.Extension -notin @('.log', '.csv') }).Count | Should -Be 1
        @(Get-ChildItem -LiteralPath (Join-Path $result.Run '.WinImgNormalizer\work') -Force).Count | Should -Be 0
        Get-CodecSourceState $source | Should -BeExactly $before
    }

    It 'T065 reports an unavailable <Decoder> reader for .<Extension> without conversion retries while supported media completes' -ForEach @(
        @{ Extension = 'jpg'; Decoder = 'JPEG' }, @{ Extension = 'jpeg'; Decoder = 'JPEG' },
        @{ Extension = 'png'; Decoder = 'PNG' }, @{ Extension = 'bmp'; Decoder = 'BMP' },
        @{ Extension = 'tif'; Decoder = 'TIFF' }, @{ Extension = 'tiff'; Decoder = 'TIFF' },
        @{ Extension = 'gif'; Decoder = 'GIF' }, @{ Extension = 'webp'; Decoder = 'WEBP' },
        @{ Extension = 'heic'; Decoder = 'HEIC' }, @{ Extension = 'heif'; Decoder = 'HEIF' }
    ) {
        $source = New-CodecDirectory ('missing-' + $Extension)
        $parent = New-CodecDirectory ('missing-output-' + $Extension)
        $supportedExtension = if ($Decoder -eq 'PNG') { 'jpg' } else { 'png' }
        $supportedCoder = if ($Decoder -eq 'PNG') { 'JPEG' } else { 'PNG' }
        $supported = Join-Path $source ('supported.' + $supportedExtension)
        New-CodecImage -Path $supported -Coder $supportedCoder
        $missing = Join-Path $source ('missing.' + $Extension)
        # A valid generated raster with a controlled unavailable extension tests
        # pre-decode routing only. Real codec support comes from the positive
        # corpus and unconditional animated WebP/HEIC/HEIF T031 integration tests.
        [IO.File]::WriteAllBytes($missing, [IO.File]::ReadAllBytes($supported))
        Set-CodecSourceTimes $missing
        $video = Join-Path $source 'opaque.mp4'
        [IO.File]::WriteAllBytes($video, [byte[]]@(17,18,19,20,255,0,99))
        Set-CodecSourceTimes $video
        $before = Get-CodecSourceState $source
        $videoHash = (Get-FileHash -LiteralPath $video -Algorithm SHA256).Hash
        # Retain JPEG.Write while masking JPEG.Read so the sibling PNG still
        # reaches the per-file path rather than failing setup for no encoder.
        $masked = Set-CodecFormatMode -Decoder $Decoder -Mode NoRead
        $trace = New-Object 'Collections.Generic.List[object]'
        $runner = {
            param([string]$Executable, [string[]]$Arguments)
            $trace.Add(@($Arguments))
            return Invoke-WinImgNativeProcess -Executable $Executable -Arguments $Arguments -TimeoutMilliseconds 15000
        }
        $result = Invoke-CodecRun -Source $source -Parent $parent -Preflight (New-CodecMaskedPreflight $masked) -Runner $runner
        $result.Code | Should -Be 2 -Because $result.Text
        $trace.Count | Should -Be 1
        $result.Log | Should -Match ([regex]::Escape(('ERR IMG: missing.' + $Extension + ' (Missing ' + $Decoder + ' decoder')))
        $result.Log | Should -Not -Match ('(SOURCE|COLOUR) IMG: missing\.' + [regex]::Escape($Extension))
        $result.Log | Should -Match 'SUMMARY ConvertedImages=1 CopiedVideos=1 Duplicates=0 Unsupported=0 Errors=1'
        $missingRow = @($result.State.Rows | Where-Object SourceRelativePath -eq ('missing.' + $Extension))
        $missingRow.Count | Should -Be 1
        $missingRow[0].Status | Should -Be 'Error'
        $missingRow[0].Reason | Should -Be ('Missing ' + $Decoder + ' decoder; no conversion attempted')
        $missingRow[0].Attempts | Should -Be 0
        Test-Path -LiteralPath (Join-Path $result.Run 'missing.jpeg') | Should -BeFalse
        Assert-CodecJpeg (Join-Path $result.Run 'supported.jpeg')
        (Get-FileHash -LiteralPath (Join-Path $result.Run 'opaque.mp4') -Algorithm SHA256).Hash | Should -BeExactly $videoHash
        @(Get-ChildItem -LiteralPath (Join-Path $result.Run '.WinImgNormalizer\work') -Force).Count | Should -Be 0
        Get-CodecSourceState $source | Should -BeExactly $before
    }

    It 'T065 rejects an unavailable JPEG writer before output creation or native conversion' {
        $source = New-CodecDirectory 'missing-writer'
        $parent = Join-Path (New-CodecDirectory 'missing-writer-parent') 'never-created'
        New-CodecImage -Path (Join-Path $source 'supported.png') -Coder PNG
        $before = Get-CodecSourceState $source
        $masked = Set-CodecFormatMode -Decoder JPEG -Mode 'r--'
        $parameters = @{
            Source = $source; OutputParent = $parent; MagickPath = $magick
            PreflightRunner = New-CodecMaskedPreflight $masked
            ProcessRunner = { throw 'A missing writer attempted native conversion.' }
        }
        $observed = @(& Invoke-WinImgNormalizer @parameters 6>&1 3>&1 2>&1)
        $codes = @($observed | Where-Object { $_ -is [int] -or $_ -is [long] })
        $codes.Count | Should -Be 1
        $codes[0] | Should -Be 1
        ($observed -join "`n") | Should -Match '(?i)JPEG.*(write|encoder|encoding)'
        Test-Path -LiteralPath $parent | Should -BeFalse
        Get-CodecSourceState $source | Should -BeExactly $before
    }
}
