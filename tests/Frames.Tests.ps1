BeforeAll {
    $repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))

    function Assert-FramesNoReparseAncestors {
        param([string]$Path)
        $current = [IO.DirectoryInfo]::new([IO.Path]::GetFullPath($Path))
        while ($null -ne $current) {
            if ($current.Exists -and (($current.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0)) {
                throw 'Frame fixture ownership path includes a reparse point.'
            }
            $current = $current.Parent
        }
    }

    $scratch = [IO.Path]::GetFullPath((Join-Path $repository '.scratch'))
    Assert-FramesNoReparseAncestors $scratch
    foreach ($pictures in @([Environment]::GetFolderPath('MyPictures'), (Join-Path $env:USERPROFILE 'Pictures'))) {
        if ([string]::IsNullOrWhiteSpace($pictures)) { continue }
        $full = [IO.Path]::GetFullPath($pictures).TrimEnd('\', '/')
        if ($scratch.Equals($full, [StringComparison]::OrdinalIgnoreCase) -or
            $scratch.StartsWith($full + '\', [StringComparison]::OrdinalIgnoreCase) -or
            $full.StartsWith($scratch + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Frame fixture scratch overlaps a real Pictures location.'
        }
    }
    $git = Get-Command git -CommandType Application -ErrorAction Stop | Select-Object -First 1
    & $git.Source -C $repository check-ignore --quiet --no-index -- (Join-Path $scratch 'frames-ignore-probe')
    if ($LASTEXITCODE -ne 0) { throw 'Frame fixture scratch must already be ignored.' }
    $ownedRoot = Join-Path $scratch ('M2-T01-frames-' + [Guid]::NewGuid().ToString('N'))
    if (Test-Path -LiteralPath $ownedRoot) { throw 'Frame fixture ownership directory already exists.' }
    $null = [IO.Directory]::CreateDirectory($ownedRoot)
    [IO.File]::WriteAllText((Join-Path $ownedRoot '.winimg-fixture-root'), 'M2-T01 synthetic frame fixture ownership')

    function New-FramesDirectory {
        param([string]$Label)
        $path = [IO.Path]::GetFullPath((Join-Path $ownedRoot ($Label + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))))
        if (-not $path.StartsWith($ownedRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Frame fixture path escapes its owned scratch directory.'
        }
        Assert-FramesNoReparseAncestors $path
        if (Test-Path -LiteralPath $path) { throw 'Frame fixture directory already exists.' }
        $null = [IO.Directory]::CreateDirectory($path)
        return $path
    }

    if ([string]::IsNullOrWhiteSpace($env:WINIMG_TEST_MAGICK) -or -not [IO.Path]::IsPathRooted($env:WINIMG_TEST_MAGICK)) {
        throw 'WINIMG_TEST_MAGICK must explicitly select the verified test magick.exe.'
    }
    $magick = [IO.Path]::GetFullPath($env:WINIMG_TEST_MAGICK)
    Assert-FramesNoReparseAncestors $magick
    if (-not (Test-Path -LiteralPath $magick -PathType Leaf) -or [IO.Path]::GetFileName($magick) -ne 'magick.exe') {
        throw 'The explicitly selected frame fixture test magick.exe is missing.'
    }
    . (Join-Path $repository 'WinImgNormalizer.ps1')
    $realFramesAllocation = (Get-Command New-WinImgImageCandidate).ScriptBlock

    function Invoke-FramesMagick {
        param([string[]]$Arguments)
        $observed = @(& $magick @Arguments 2>&1)
        $code = $LASTEXITCODE
        if ($code -ne 0) { throw ('Frame fixture ImageMagick exited {0}: {1}' -f $code, ($observed -join "`n")) }
        return ($observed -join "`n")
    }

    function Get-FramesSourceState {
        param([string]$Source)
        return @(
            Get-ChildItem -LiteralPath $Source -Recurse -Force -File | Sort-Object FullName | ForEach-Object {
                [pscustomobject]@{
                    Path = $_.FullName.Substring($Source.Length + 1)
                    Hash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
                    Length = $_.Length
                    CreationTicks = $_.CreationTimeUtc.Ticks
                    ModifiedTicks = $_.LastWriteTimeUtc.Ticks
                    Attributes = [int]$_.Attributes
                }
            }
        ) | ConvertTo-Json -Depth 3 -Compress
    }

    function Set-FramesSourceTimes {
        param([string]$Path)
        $fixed = [DateTime]::Parse('2020-02-03T04:05:06Z').ToUniversalTime()
        [IO.File]::SetCreationTimeUtc($Path, $fixed)
        [IO.File]::SetLastWriteTimeUtc($Path, $fixed)
    }

    function Invoke-FramesRun {
        param([string]$Source, [string]$Parent, [scriptblock]$ProcessRunner)
        $parameters = @{ Source = $Source; OutputParent = $Parent; MagickPath = $magick }
        if ($ProcessRunner) { $parameters.ProcessRunner = $ProcessRunner }
        $observed = @(& Invoke-WinImgNormalizer @parameters 6>&1 3>&1 2>&1)
        $codes = @($observed | Where-Object { $_ -is [int] -or $_ -is [long] })
        if ($codes.Count -ne 1) { throw ('Expected one frame regression exit status: ' + ($observed -join "`n")) }
        $runs = @(Get-ChildItem -LiteralPath $Parent -Force -Directory)
        $runs.Count | Should -Be 1
        $logs = @(Get-ChildItem -LiteralPath (Join-Path $runs[0].FullName '.WinImgNormalizer\reports') -File -Filter '*.log')
        $logs.Count | Should -Be 1
        return [pscustomobject]@{
            Code = $codes[0]; Run = $runs[0].FullName
            Text = ($observed -join "`n"); Log = Get-Content -LiteralPath $logs[0].FullName -Raw
        }
    }

    function Assert-FramesJpeg {
        param([string]$Path, [int]$Width, [int]$Height, [object[]]$Samples)
        $native = Get-WinImgNativeOutputPath $Path
        Invoke-FramesMagick @('identify', '+ping', '-regard-warnings', '-format', '%m|%w|%h|%n', $native) |
            Should -Be ('JPEG|{0}|{1}|1' -f $Width, $Height)
        foreach ($sample in $Samples) {
            $format = '%[fx:round(255*p{{{0},{1}}}.r)]|%[fx:round(255*p{{{0},{1}}}.g)]|%[fx:round(255*p{{{0},{1}}}.b)]' -f $sample.X, $sample.Y
            $actual = (Invoke-FramesMagick @($native, '-format', $format, 'info:')) -split '\|'
            $actual.Count | Should -Be 3
            for ($channel = 0; $channel -lt 3; $channel++) {
                [Math]::Abs([int]$actual[$channel] - $sample.Rgb[$channel]) | Should -BeLessOrEqual 12 -Because ('pixel ({0},{1}) channel {2}' -f $sample.X, $sample.Y, $channel)
            }
        }
    }

    function Assert-FramesCompleteRun {
        param([object]$Result, [string]$Output, [string]$Source, [int]$Count, [string]$Unit, [string]$Policy, [string]$Decoder)
        $Result.Code | Should -Be 0 -Because $Result.Text
        $Result.Log | Should -Match ([regex]::Escape(('SOURCE IMG: {0} (SourceCount={1}; Selected=1; Omitted={2}; Unit={3}; Policy={4}; Decoder={5})' -f $Source, $Count, ($Count - 1), $Unit, $Policy, $Decoder)))
        $Result.Log | Should -Match 'SUMMARY ConvertedImages=1 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=0'
        $files = @(Get-ChildItem -LiteralPath $Result.Run -Recurse -Force -File)
        @($files | Where-Object { $_.Extension -eq '.jpeg' }).Count | Should -Be 1
        @($files | Where-Object { $_.Extension -notin @('.log','.csv') }).Count | Should -Be 1
        @($files | Where-Object { $_.Extension -eq '.csv' }).Count | Should -Be 1
        $files.Count | Should -Be 3
        Test-Path -LiteralPath (Join-Path $Result.Run $Output) -PathType Leaf | Should -BeTrue
        @(Get-ChildItem -LiteralPath (Join-Path $Result.Run '.WinImgNormalizer\work') -Force).Count | Should -Be 0
    }

    function New-FramesGif {
        param([string]$Path, [int]$Count = 2)
        $arguments = @('-background', 'none', '(', '-size', '24x20', 'xc:none', '-fill', '#E02020',
            '-draw', 'rectangle 4,4 19,15', '-set', 'page', '64x48+8+10', '-dispose', 'Background', ')')
        if ($Count -eq 2) {
            $arguments += @('(', '-size', '64x48', 'xc:#2020E0', '-set', 'page', '64x48+0+0', ')')
        }
        $arguments += @('-delay', '12', '-loop', '0', ('GIF:' + (Get-WinImgNativeOutputPath $Path)))
        $null = Invoke-FramesMagick $arguments
    }

    function Set-FramesFirstTiffOrientation {
        param([string]$Path)
        # ImageMagick's TIFF writer canonicalizes the tag to TopLeft. Patch the
        # existing first-IFD SHORT tag only, so the fixture actually stores6.
        $bytes = [IO.File]::ReadAllBytes($Path)
        if ($bytes[0] -ne 0x49 -or $bytes[1] -ne 0x49 -or [BitConverter]::ToUInt16($bytes, 2) -ne 42) {
            throw 'Expected little-endian classic TIFF synthetic fixture.'
        }
        $ifd = [BitConverter]::ToUInt32($bytes, 4)
        $entries = [BitConverter]::ToUInt16($bytes, $ifd)
        $found = 0
        for ($index = 0; $index -lt $entries; $index++) {
            $offset = [int]$ifd + 2 + 12 * $index
            if ([BitConverter]::ToUInt16($bytes, $offset) -eq 0x112) {
                [BitConverter]::ToUInt16($bytes, $offset + 2) | Should -Be 3
                [BitConverter]::ToUInt32($bytes, $offset + 4) | Should -Be 1
                $bytes[$offset + 8] = 6
                $bytes[$offset + 9] = 0
                $found++
            }
        }
        $found | Should -Be 1
        [IO.File]::WriteAllBytes($Path, $bytes)
    }

    # Self-generated HEVC collection, not a timed sequence or a thumbnail pair.
    # pillow-heif1.8.0 (sourceBSD3, generator wheelGPLv2) with Pillow12.3.0:
    # first64x48 RGB redE02020, green20D040 rect0,0..23,15, yellowE0D020
    # rect40,32..63,47; second64x48 blue2020E0; from_pillow(first),
    # add_from_pillow(second), save(quality=-1,chroma='444',primary_index=0,
    # save_all=True). No thumbnails, auxiliary images or external source media.
    # Official wheel: pillow_heif-1.8.0-cp312-cp312-win_amd64.whl,6604433bytes,
    # SHA256921784d39b5e3b9bcbfcf96ff05e78bc5a15ced5d171cf4e3987d06c04da71bf.
    # Provenance: https://pypi.org/pypi/pillow-heif/1.8.0/json
    # Generator: libheif1.23.4/x2654.3+1-e9b8812; actual pinned IM7.1.2-32
    # libheif1.23.5 independently exposes two64x48HEIC images in this order.
    $heicBytes = [Convert]::FromBase64String('AAAAHGZ0eXBoZWl4AAAAAG1pZjFoZWl4bWlhZgAAAa1tZXRhAAAAAAAAACFoZGxyAAAAAAAAAABwaWN0AAAAAAAAAAAAAAAAAAAAADRpbG9jAAAAAERAAAIAAQAAAAAB0QABAAAAAAAAAvsAAgAAAAAEzAABAAAAAAAAAGQAAAA4aWluZgAAAAAAAgAAABVpbmZlAgAAAAABAABodmMxAAAAABVpbmZlAgAAAAACAABodmMxAAAAAA5waXRtAAAAAAABAAABBmlwcnAAAADeaXBjbwAAAHdodmNDAQQIAAAAAAAAAAAA//AA/P/4+AAADwNgAAEAF0ABDAH//wQIAAADAJ44AAADAAD/ugJAYQABACpCAQEECAAAAwCeOAAAAwAA/5AEECCy3VyU1zcBDQYEAAADAGQAAAMABCBiAAEACEQBwXAwYJEgAAAAE2NvbHJuY2x4AAEADQAGgAAAABRpc3BlAAAAAAAAAEAAAABAAAAAKGNsYXAAAABAAAAAAQAAADAAAAABAAAAAAAAAAL////wAAAAAgAAABBwaXhpAAAAAAMICAgAAAAgaXBtYQAAAAAAAAACAAEFgQIDBYQAAgWBAgMFhAAAA2dtZGF0AAAC9ygBrwW4H2xh+gG9DAhCEISWMYxjDyf//yJ2v//1PKhmJqy5cuXNFixYsWLFEz/9kHP//4f/DfwX2vK8ryvK8vy/L8vy/L8vy7wJmK2JYhZMCnSWPsP//XXrRppMtra2treHh4eHh4eG2m/tTEAV97rWta1xKUpSlB9/mgoA88yH5iVlq0luFnhZ4WeFnhdsXbF2xdsXbF2xdsWt8ZP//VXBbL3CBJO6vcYH5y2nrrciv//FVGPK5hbAiuinOitpOhP//i4Djne2zn3dFq9FuJAwEob//zKBcu7tycHBwcIDw8PDw8PDkitNBAIPOYQhCEKb3ve94eVX2AMLW3etWivT0Y4VeFXhV4VeGIRiEYhGIRiEYhGIRWMF//yVhr8glqY7IqcpPgWJKmcsf/9j2A/YGubTs+b+OMohcPf/8Nv0Y+pkZERiPuBRLVVzdkX67f/8vgrhwyfiut3Vx3R56eWCJf2Esm8JPmB9f/aggv6/loBYGKQOgAf9H/ov578t8N8N8N8N8P8P8P8P8P8P8P8OAOuJ//8QYFP5hD7u7u7vTMzMzMzMzHdfnf//k9El2nr1Xsuy7LsvBIEgSBIEgSBHaIAEFO3tkbHwP87ELa5wS2/a9SmKQACv+5qn6K5R4mcQgihAhqVkcT5a6Cn//sB1zc/6JkRLMlkbPiW96r+AC7cvpOxROaqhVNxtSTDZsjMNNzYAQs29+KsALZZ8st4COdqi3OG9tnvgBKx76uVcW2ZgDZQQC1RsiV8SVgADUYDxBCBLIklRUhWhpOapNrNqfJADLZvxpjSQEmpGrFzUgRaa2j+pAAMlofHKORErVSfc3m4HWiXiJfAwhVgf//MnlxJZ59khXss+d05Jft6AAjRXbtzOmlcV2KYShW538QUuO06AErHvbqofkn036XWIQ8XIUtQwUI43//498aOt8IxunGPCAlBSABHIAF+3/Rfh6VRso2Ohcf7V7S00Kp2ABAD84N+9zT1gqvWLegjddm65SKHigAAAAGAoAa8FuB9sY7///yeiS7T16r2XZdl2XgkCQJAkCQJAjr7//+RNXAF7a47UIyD2De973ve+B0HQdB0HQdBz05/+yDn//n8urWta173ve973gTMVsSxCyhy/TimpgRyrOKI=')
    $heicSha = [Security.Cryptography.SHA256]::Create()
    try {
        ([BitConverter]::ToString($heicSha.ComputeHash($heicBytes))).Replace('-', '') |
            Should -Be '618ED4D0C36F19BFAC97401F28D45D42DF9E9125D093875EF794C8C4E675A69A'
    } finally { $heicSha.Dispose() }
}

Describe 'M2-T01 deliberate first logical canvas and first page (T029-T031)' {
    It 'T029 selects the offset transparent first GIF canvas from <Count> actual frame(s) at a literal path' -ForEach @(
        @{ Count = 1 }, @{ Count = 2 }
    ) {
        $source = New-FramesDirectory 'gif-source'
        $parent = New-FramesDirectory 'gif-output'
        $nested = Join-Path $source 'nested [safe]'
        $null = [IO.Directory]::CreateDirectory($nested)
        $relative = 'nested [safe]\offset [0].gif'
        $path = Join-Path $source $relative
        New-FramesGif -Path $path -Count $Count
        $lines = (Invoke-FramesMagick @('identify', '+ping', '-regard-warnings', '-format', '%m|%n|%w|%h|%[page]\n', (Get-WinImgNativeOutputPath $path))) -split '\r?\n'
        $lines.Count | Should -Be $Count
        $lines[0] | Should -Be ('GIF|{0}|24|20|64x48' -f $Count)
        # Actual encoded first tile location, independent of the output policy.
        Invoke-FramesMagick @('identify', '-format', '%X|%Y\n', (Get-WinImgNativeOutputPath $path)) |
            Should -Match '^\+8\|\+10'
        Set-FramesSourceTimes $path
        $before = Get-FramesSourceState $source
        $result = Invoke-FramesRun -Source $source -Parent $parent
        $output = 'nested [safe]\offset [0].jpeg'
        Assert-FramesCompleteRun -Result $result -Output $output -Source $relative -Count $Count -Unit Frames -Policy FirstDisplayedFrame -Decoder GIF
        Assert-FramesJpeg -Path (Join-Path $result.Run $output) -Width 64 -Height 48 -Samples @(
            @{ X = 0; Y = 0; Rgb = @(255,255,255) },
            @{ X = 20; Y = 20; Rgb = @(224,32,32) },
            @{ X = 40; Y = 32; Rgb = @(255,255,255) }
        )
        Get-FramesSourceState $source | Should -Be $before
    }

    It 'T030 keeps the actual oriented first TIFF page, excludes the distinct blue second page and reports one omission' {
        $source = New-FramesDirectory 'tiff-source'
        $parent = New-FramesDirectory 'tiff-output'
        $path = Join-Path $source 'pages.tiff'
        $null = Invoke-FramesMagick @('(', '-size', '80x48', 'xc:#E02020', '-fill', '#20D040', '-draw', 'rectangle 0,0 23,15',
            '-fill', '#E0D020', '-draw', 'rectangle 56,32 79,47', ')', '(', '-size', '56x32', 'xc:#2020E0', ')',
            '-orient', 'TopLeft', '-compress', 'None', ('TIFF:' + (Get-WinImgNativeOutputPath $path)))
        Set-FramesFirstTiffOrientation $path
        Invoke-FramesMagick @('identify', '+ping', '-regard-warnings', '-format', '%m|%n|%w|%h|%[orientation]\n', (Get-WinImgNativeOutputPath $path)) |
            Should -Be "TIFF|2|80|48|RightTop`nTIFF|2|56|32|TopLeft"
        Set-FramesSourceTimes $path
        $before = Get-FramesSourceState $source
        $result = Invoke-FramesRun -Source $source -Parent $parent
        Assert-FramesCompleteRun -Result $result -Output 'pages.jpeg' -Source 'pages.tiff' -Count 2 -Unit Pages -Policy FirstPage -Decoder TIFF
        Assert-FramesJpeg -Path (Join-Path $result.Run 'pages.jpeg') -Width 48 -Height 80 -Samples @(
            @{ X = 40; Y = 10; Rgb = @(32,208,64) },
            @{ X = 8; Y = 68; Rgb = @(224,208,32) },
            @{ X = 24; Y = 40; Rgb = @(224,32,32) }
        )
        Get-FramesSourceState $source | Should -Be $before
    }

    It 'T031 decodes an actual two-frame animated WebP, keeps its first alpha canvas and reports the dropped frame' {
        $source = New-FramesDirectory 'webp-source'
        $parent = New-FramesDirectory 'webp-output'
        $path = Join-Path $source 'animated.webp'
        # A transparent gap inside the encoded patch keeps real alpha in its
        # optimized WebP fragment, instead of relying on opaque cropped padding.
        $null = Invoke-FramesMagick @('-background', 'none', '(', '-size', '64x48', 'xc:none', '-fill', '#E02020', '-draw', 'rectangle 12,14 15,25 rectangle 18,14 27,25', ')',
            '(', '-size', '64x48', 'xc:#2020E0', ')', '-define', 'webp:lossless=true', '-delay', '12', '-loop', '0',
            ('WEBP:' + (Get-WinImgNativeOutputPath $path)))
        Invoke-FramesMagick @('identify', '+ping', '-regard-warnings', '-format', '%m|%n|%w|%h\n', (Get-WinImgNativeOutputPath $path)) |
            Should -Be "WEBP|2|64|48`nWEBP|2|64|48"
        # Require an actual alpha channel plus a fully transparent decoded pixel;
        # fx alpha access on an RGB-only later image is not an opacity oracle.
        Invoke-FramesMagick @((Get-WinImgNativeOutputPath $path), '-format', '%s|%[channels]|%[pixel:p{0,0}]\n', 'info:') |
            Should -Match '^0\|srgba[^|]*\|srgba\(0,0,0,0\)'
        Set-FramesSourceTimes $path
        $before = Get-FramesSourceState $source
        $result = Invoke-FramesRun -Source $source -Parent $parent
        Assert-FramesCompleteRun -Result $result -Output 'animated.jpeg' -Source 'animated.webp' -Count 2 -Unit Frames -Policy FirstDisplayedFrame -Decoder WEBP
        Assert-FramesJpeg -Path (Join-Path $result.Run 'animated.jpeg') -Width 64 -Height 48 -Samples @(
            @{ X = 0; Y = 0; Rgb = @(255,255,255) },
            @{ X = 24; Y = 20; Rgb = @(224,32,32) },
            @{ X = 40; Y = 32; Rgb = @(255,255,255) }
        )
        Get-FramesSourceState $source | Should -Be $before
    }

    It 'T031 decodes both actual HEVC collection images via .<Extension>, retains the decoder primary first image and reports one omission' -ForEach @(
        @{ Extension = 'heic' }, @{ Extension = 'heif' }
    ) {
        $source = New-FramesDirectory ('hevc-source-' + $Extension)
        $parent = New-FramesDirectory ('hevc-output-' + $Extension)
        $relative = 'collection.' + $Extension
        $decoder = $Extension.ToUpperInvariant()
        $path = Join-Path $source $relative
        [IO.File]::WriteAllBytes($path, $heicBytes)
        $heicBytes.Length | Should -Be 1328
        Invoke-FramesMagick @('identify', '+ping', '-regard-warnings', '-format', '%m|%n|%w|%h\n', (Get-WinImgNativeOutputPath $path)) |
            Should -Be ("{0}|2|64|48`n{0}|2|64|48" -f $decoder)
        # Confirm actual decoder-visible second content; this is not a static
        # decode or a controlled naming mock counted as collection coverage.
        Invoke-FramesMagick @((Get-WinImgNativeOutputPath $path), '-format', '%s|%[fx:round(255*p{32,24}.r)]|%[fx:round(255*p{32,24}.b)]\n', 'info:') |
            Should -Be "0|224|32`n1|32|224"
        Set-FramesSourceTimes $path
        $before = Get-FramesSourceState $source
        $result = Invoke-FramesRun -Source $source -Parent $parent
        Assert-FramesCompleteRun -Result $result -Output 'collection.jpeg' -Source $relative -Count 2 -Unit Images -Policy DecoderPrimaryOrFirstImage -Decoder $decoder
        Assert-FramesJpeg -Path (Join-Path $result.Run 'collection.jpeg') -Width 64 -Height 48 -Samples @(
            @{ X = 10; Y = 8; Rgb = @(32,208,64) },
            @{ X = 52; Y = 40; Rgb = @(224,208,32) },
            @{ X = 32; Y = 24; Rgb = @(224,32,32) }
        )
        Get-FramesSourceState $source | Should -Be $before
    }
}

Describe 'M2-T01 source inspection and snapshot safety' {
    It 'rejects an invalid sequence source before any conversion process and preserves its bytes and timestamps' {
        $source = New-FramesDirectory 'invalid-source'
        $parent = New-FramesDirectory 'invalid-output'
        $path = Join-Path $source 'invalid.gif'
        [IO.File]::WriteAllText($path, 'Owned malformed GIF source.')
        Set-FramesSourceTimes $path
        $before = Get-FramesSourceState $source
        $trace = [pscustomobject]@{ Calls = 0 }
        $runner = { param($Executable, $Arguments); $trace.Calls++; return 0 }
        $result = Invoke-FramesRun -Source $source -Parent $parent -ProcessRunner $runner
        $result.Code | Should -Be 2
        $trace.Calls | Should -Be 0
        $result.Log | Should -Match 'Source frame/page inspection failed'
        $result.Log | Should -Not -Match 'SOURCE IMG:'
        Test-Path -LiteralPath (Join-Path $result.Run 'invalid.jpeg') | Should -BeFalse
        @(Get-ChildItem -LiteralPath (Join-Path $result.Run '.WinImgNormalizer\work') -Force).Count | Should -Be 0
        Get-FramesSourceState $source | Should -Be $before
    }

    It 'rejects an actual source <Change> change during snapshot allocation before source inspection or conversion' -ForEach @(
        @{ Change = 'length' }, @{ Change = 'timestamp' }
    ) {
        $source = New-FramesDirectory ('changed-' + $Change)
        $parent = New-FramesDirectory ('changed-output-' + $Change)
        $path = Join-Path $source 'changed.gif'
        New-FramesGif $path
        Set-FramesSourceTimes $path
        $trace = [pscustomobject]@{ Calls = 0; Allocations = 0; State = $null; SourcePath = $path; SourceRoot = $source }
        Mock New-WinImgImageCandidate {
            param([string]$WorkRoot)
            $candidate = & $realFramesAllocation -WorkRoot $WorkRoot
            $trace.Allocations++
            if ($trace.Allocations -eq 1) {
                if ($Change -eq 'length') {
                    $bytes = [IO.File]::ReadAllBytes($trace.SourcePath)
                    [IO.File]::WriteAllBytes($trace.SourcePath, [byte[]]@($bytes + [byte]0))
                } else { [IO.File]::SetLastWriteTimeUtc($trace.SourcePath, [IO.File]::GetLastWriteTimeUtc($trace.SourcePath).AddSeconds(2)) }
                $trace.State = Get-FramesSourceState $trace.SourceRoot
            }
            return $candidate
        }
        $runner = { param($Executable, $Arguments); $trace.Calls++; return 0 }
        $result = Invoke-FramesRun -Source $source -Parent $parent -ProcessRunner $runner
        $result.Code | Should -Be 2
        $trace.Calls | Should -Be 0
        $trace.Allocations | Should -Be 1
        $result.Log | Should -Match 'Image snapshot is incomplete or the source changed during copying'
        $result.Log | Should -Not -Match 'SOURCE IMG:'
        Test-Path -LiteralPath (Join-Path $result.Run 'changed.jpeg') | Should -BeFalse
        @(Get-ChildItem -LiteralPath (Join-Path $result.Run '.WinImgNormalizer\work') -Force).Count | Should -Be 0
        Get-FramesSourceState $source | Should -Be $trace.State
    }

    It 'preserves an external reserved snapshot arrival and never inspects or converts its bytes' {
        $source = New-FramesDirectory 'snapshot-arrival-source'
        $parent = New-FramesDirectory 'snapshot-arrival-output'
        $path = Join-Path $source 'arrival.gif'
        New-FramesGif $path
        Set-FramesSourceTimes $path
        $before = Get-FramesSourceState $source
        $trace = [pscustomobject]@{ Calls = 0; Arrival = $null; Hash = $null }
        Mock New-WinImgImageCandidate {
            param([string]$WorkRoot)
            $candidate = & $realFramesAllocation -WorkRoot $WorkRoot
            $trace.Arrival = Join-Path ([IO.Path]::GetDirectoryName($candidate)) 'source.gif'
            [IO.File]::WriteAllText($trace.Arrival, 'Unknown snapshot arrival stays preserved.')
            $trace.Hash = (Get-FileHash -LiteralPath $trace.Arrival -Algorithm SHA256).Hash
            return $candidate
        }
        Mock Get-WinImgSourceImageInfo { throw 'Reserved arrival must never reach inspection.' }
        $runner = { param($Executable, $Arguments); $trace.Calls++; return 0 }
        $result = Invoke-FramesRun -Source $source -Parent $parent -ProcessRunner $runner
        $result.Code | Should -Be 2
        $trace.Calls | Should -Be 0
        Should -Invoke Get-WinImgSourceImageInfo -Times 0 -Exactly
        (Get-FileHash -LiteralPath $trace.Arrival -Algorithm SHA256).Hash | Should -Be $trace.Hash
        [IO.File]::ReadAllText($trace.Arrival) | Should -Be 'Unknown snapshot arrival stays preserved.'
        Test-Path -LiteralPath (Join-Path $result.Run 'arrival.jpeg') | Should -BeFalse
        Get-FramesSourceState $source | Should -Be $before
    }

    It 'removes its successful owned source snapshot and keeps an unknown neighboring file' {
        $source = New-FramesDirectory 'neighbor-source'
        $parent = New-FramesDirectory 'neighbor-output'
        $path = Join-Path $source 'neighbor.gif'
        New-FramesGif $path
        Set-FramesSourceTimes $path
        $before = Get-FramesSourceState $source
        $trace = [pscustomobject]@{ Allocations = 0; Neighbor = $null; Snapshot = $null }
        Mock New-WinImgImageCandidate {
            param([string]$WorkRoot)
            $candidate = & $realFramesAllocation -WorkRoot $WorkRoot
            $trace.Allocations++
            if ($trace.Allocations -eq 1) {
                $trace.Neighbor = Join-Path ([IO.Path]::GetDirectoryName($candidate)) 'unknown [keep].txt'
                $trace.Snapshot = Join-Path ([IO.Path]::GetDirectoryName($candidate)) 'source.gif'
                [IO.File]::WriteAllText($trace.Neighbor, 'Unknown neighbor stays preserved.')
            }
            return $candidate
        }
        $result = Invoke-FramesRun -Source $source -Parent $parent
        $result.Code | Should -Be 2 -Because 'owned snapshot was removed but its unknown neighbor prevents empty-directory cleanup'
        $result.Log | Should -Match 'Could not remove owned image scratch:'
        $result.Log | Should -Match 'SUMMARY ConvertedImages=1 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=0'
        Test-Path -LiteralPath $trace.Snapshot | Should -BeFalse
        [IO.File]::ReadAllText($trace.Neighbor) | Should -Be 'Unknown neighbor stays preserved.'
        $workFiles = @(Get-ChildItem -LiteralPath (Join-Path $result.Run '.WinImgNormalizer\work') -Recurse -Force -File)
        $workFiles.Count | Should -Be 1
        $workFiles[0].FullName | Should -Be $trace.Neighbor
        Assert-FramesJpeg -Path (Join-Path $result.Run 'neighbor.jpeg') -Width 64 -Height 48 -Samples @(
            @{ X = 20; Y = 20; Rgb = @(224,32,32) }
        )
        Get-FramesSourceState $source | Should -Be $before
    }

    It 'rejects a controlled short copied-snapshot length before inspection and conversion' {
        $source = New-FramesDirectory 'short-copy-source'
        $parent = New-FramesDirectory 'short-copy-output'
        $path = Join-Path $source 'short.gif'
        New-FramesGif $path
        Set-FramesSourceTimes $path
        $before = Get-FramesSourceState $source
        Mock Get-Item {
            $real = [IO.FileInfo]::new($LiteralPath)
            return [pscustomobject]@{ PSIsContainer = $false; Attributes = $real.Attributes; Length = $real.Length - 1 }
        } -ParameterFilter { [IO.Path]::GetFileName($LiteralPath) -eq 'source.gif' }
        $trace = [pscustomobject]@{ Calls = 0 }
        $runner = { param($Executable, $Arguments); $trace.Calls++; return 0 }
        $result = Invoke-FramesRun -Source $source -Parent $parent -ProcessRunner $runner
        $result.Code | Should -Be 2
        $trace.Calls | Should -Be 0
        $result.Log | Should -Match 'Image snapshot is incomplete or the source changed during copying'
        $result.Log | Should -Not -Match 'SOURCE IMG:'
        Test-Path -LiteralPath (Join-Path $result.Run 'short.jpeg') | Should -BeFalse
        @(Get-ChildItem -LiteralPath (Join-Path $result.Run '.WinImgNormalizer\work') -Force).Count | Should -Be 0
        Get-FramesSourceState $source | Should -Be $before
    }
}
