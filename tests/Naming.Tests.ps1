BeforeAll {
    $repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    $application = Join-Path $repository 'WinImgNormalizer.ps1'

    function Assert-NamingNoReparseAncestors {
        param([string]$Path)
        $current = [IO.DirectoryInfo]::new([IO.Path]::GetFullPath($Path))
        while ($null -ne $current) {
            if ($current.Exists -and (($current.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0)) {
                throw 'Naming test ownership path includes a reparse point.'
            }
            $current = $current.Parent
        }
    }

    $scratchParent = [IO.Path]::GetFullPath((Join-Path $repository '.scratch'))
    Assert-NamingNoReparseAncestors $scratchParent
    foreach ($pictures in @([Environment]::GetFolderPath('MyPictures'), (Join-Path $env:USERPROFILE 'Pictures'))) {
        if ([string]::IsNullOrWhiteSpace($pictures)) { continue }
        $picturesFull = [IO.Path]::GetFullPath($pictures).TrimEnd('\', '/')
        if ($scratchParent.Equals($picturesFull, [StringComparison]::OrdinalIgnoreCase) -or
            $scratchParent.StartsWith($picturesFull + '\', [StringComparison]::OrdinalIgnoreCase) -or
            $picturesFull.StartsWith($scratchParent + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Naming test scratch overlaps a real Pictures location.'
        }
    }
    $git = Get-Command git -CommandType Application -ErrorAction Stop | Select-Object -First 1
    & $git.Source -C $repository check-ignore --quiet --no-index -- (Join-Path $scratchParent 'naming-ignore-probe')
    if ($LASTEXITCODE -ne 0) { throw 'Naming synthetic scratch must already be ignored.' }
    $ownedRoot = Join-Path $scratchParent ('M1-T03-' + [Guid]::NewGuid().ToString('N'))
    if (Test-Path -LiteralPath $ownedRoot) { throw 'Naming ownership directory already exists.' }
    $null = [IO.Directory]::CreateDirectory($ownedRoot)
    [IO.File]::WriteAllText((Join-Path $ownedRoot '.winimg-fixture-root'), 'M1-T03 synthetic test ownership')

    function New-NamingDirectory {
        param([string]$Label)
        $path = [IO.Path]::GetFullPath((Join-Path $ownedRoot ($Label + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))))
        if (-not $path.StartsWith($ownedRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Naming test path escapes its owned scratch directory.'
        }
        Assert-NamingNoReparseAncestors $path
        if (Test-Path -LiteralPath $path) { throw 'Naming test directory already exists.' }
        $null = [IO.Directory]::CreateDirectory($path)
        return $path
    }

    function New-NamingTree {
        param([string]$Root, [string[]]$Files, [string[]]$Directories = @())
        # Controlled FileInfo paths need no codec and can express case-only
        # variants without changing NTFS case-sensitivity or creating real media.
        return [pscustomobject]@{
            Files = @($Files | ForEach-Object { [IO.FileInfo]::new([IO.Path]::Combine($Root, $_)) })
            Directories = @($Directories | ForEach-Object { [IO.DirectoryInfo]::new([IO.Path]::Combine($Root, $_)) })
            TopLevelNames = @(@($Files) + @($Directories) | ForEach-Object { ($_ -split '\\')[0] })
            Warnings = @()
        }
    }

    function Get-NamingPlanSignature {
        param([object[]]$Plan)
        return (@($Plan | ForEach-Object {
            '{0}|{1}|{2}|{3}' -f $_.SourceRelativePath, $_.OutputRelativePath, $_.Kind, $_.NamingReason
        }) -join "`n")
    }

    function ConvertFrom-NamingNativePath {
        param([string]$Path)
        if ($Path.StartsWith('\\?\UNC\', [StringComparison]::Ordinal)) { return '\\' + $Path.Substring(8) }
        if ($Path.StartsWith('\\?\', [StringComparison]::Ordinal)) { return $Path.Substring(4) }
        return $Path
    }

    function Get-NamingSourceState {
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

    function Get-NamingRun {
        param([string]$Parent)
        $runs = @(Get-ChildItem -LiteralPath $Parent -Directory)
        if ($runs.Count -ne 1) { throw 'Expected exactly one naming-test run directory.' }
        return $runs[0].FullName
    }

    function Get-NamingLog {
        param([string]$Run)
        $logs = @(Get-ChildItem -LiteralPath $Run -Recurse -File -Filter '*.log')
        if ($logs.Count -ne 1) { throw 'Expected exactly one naming-test log.' }
        return Get-Content -LiteralPath $logs[0].FullName -Raw
    }

    function Invoke-NamingRun {
        param([string]$Source, [string]$OutputParent, [scriptblock]$ProcessRunner)
        $invoke = @{ Source = $Source; OutputParent = $OutputParent; MagickPath = $magick }
        if ($ProcessRunner) { $invoke.ProcessRunner = $ProcessRunner }
        $observed = @(& Invoke-WinImgNormalizer @invoke 6>&1 3>&1 2>&1)
        $codes = @($observed | Where-Object { $_ -is [int] -or $_ -is [long] })
        if ($codes.Count -ne 1) { throw ('Expected one callable naming exit status: ' + ($observed -join "`n")) }
        return [pscustomobject]@{ Code = $codes[0]; Text = (@($observed | Where-Object { $_ -isnot [int] -and $_ -isnot [long] }) -join "`n") }
    }

    if ([string]::IsNullOrWhiteSpace($env:WINIMG_TEST_MAGICK) -or -not [IO.Path]::IsPathRooted($env:WINIMG_TEST_MAGICK)) {
        throw 'WINIMG_TEST_MAGICK must explicitly select the verified test magick.exe.'
    }
    $magick = [IO.Path]::GetFullPath($env:WINIMG_TEST_MAGICK)
    Assert-NamingNoReparseAncestors $magick
    if (-not (Test-Path -LiteralPath $magick -PathType Leaf) -or [IO.Path]::GetFileName($magick) -ne 'magick.exe') {
        throw 'The explicitly selected naming test magick.exe is missing.'
    }

    function Invoke-NamingMagick {
        param([string[]]$Arguments)
        $result = & $magick @Arguments 2>&1
        $code = $LASTEXITCODE
        if ($code -ne 0) { throw ('Naming test ImageMagick exited {0}: {1}' -f $code, ($result -join "`n")) }
        return ($result -join "`n")
    }

    $jpegFixture = Join-Path $ownedRoot 'controlled-candidate.jpeg'
    $null = Invoke-NamingMagick @('-quiet', '-size', '20x16', 'gradient:#205080-#E0A070', '-depth', '8', ('JPEG:' + $jpegFixture))
    $jpegBytes = [IO.File]::ReadAllBytes($jpegFixture)
    . $application
    $realOutputAvailable = (Get-Command Assert-WinImgOutputAvailable -CommandType Function).ScriptBlock
    $realImageCandidate = (Get-Command New-WinImgImageCandidate -CommandType Function).ScriptBlock
}

Describe 'M1-T03 complete deterministic namespace planning (T017-T019)' {
    It 'T017 assigns every same-stem JPG, PNG and HEIC input a distinct deterministic name without requiring HEIC decoding' {
        $source = New-NamingDirectory 'pure-same-stem'
        $tree = New-NamingTree -Root $source -Files @('photo.png', 'photo.heic', 'photo.jpg')
        $plan = @(Get-WinImgOutputPlan -SourceRoot $source -SourceTree $tree -GeneratedName '.WinImgNormalizer')
        $plan.Count | Should -Be 3
        Get-NamingPlanSignature $plan | Should -Be (@(
            'photo.heic|photo__heic.jpeg|Image|ExtensionSuffix',
            'photo.jpg|photo__jpg.jpeg|Image|ExtensionSuffix',
            'photo.png|photo__png.jpeg|Image|ExtensionSuffix'
        ) -join "`n")
        foreach ($row in $plan) {
            $row.Source | Should -BeOfType ([IO.FileInfo])
            $row.Source.FullName | Should -Be ([IO.Path]::Combine($source, $row.SourceRelativePath))
        }
        @(Get-ChildItem -LiteralPath $source -Force).Count | Should -Be 0
    }

    It 'T017 retains legacy names for unique images and videos in separate directories and excludes unsupported inputs' {
        $source = New-NamingDirectory 'pure-legacy'
        $tree = New-NamingTree -Root $source -Files @('sub\photo.png', 'photo.jpg', 'clip.MP4', 'readme.txt') -Directories @('sub')
        $plan = @(Get-WinImgOutputPlan -SourceRoot $source -SourceTree $tree -GeneratedName '.WinImgNormalizer')
        Get-NamingPlanSignature $plan | Should -Be (@(
            'clip.MP4|clip.MP4|Video|Legacy', 'photo.jpg|photo.jpeg|Image|Legacy', 'sub\photo.png|sub\photo.jpeg|Image|Legacy'
        ) -join "`n")
    }

    It 'T018 reserves all unique legacy outputs and checked directory names before assigning secondary suffixes' {
        $source = New-NamingDirectory 'pure-secondary'
        $tree = New-NamingTree -Root $source -Files @('photo.png', 'photo__png.jpeg', 'photo.jpg') -Directories @('photo.jpeg', 'PHOTO__PNG__2.JPEG', 'photo__png__3.jpeg')
        $plan = @(Get-WinImgOutputPlan -SourceRoot $source -SourceTree $tree -GeneratedName '.WinImgNormalizer')
        Get-NamingPlanSignature $plan | Should -Be (@(
            'photo.jpg|photo__jpg.jpeg|Image|ExtensionSuffix',
            'photo.png|photo__png__4.jpeg|Image|NumericSuffix',
            'photo__png.jpeg|photo__png.jpeg|Image|Legacy'
        ) -join "`n")
    }

    It 'T018 avoids a directory named as a single image legacy target without another colliding source' {
        $source = New-NamingDirectory 'pure-directory'
        $tree = New-NamingTree -Root $source -Files @('photo.png') -Directories @('PHOTO.JPEG')
        $plan = @(Get-WinImgOutputPlan -SourceRoot $source -SourceTree $tree -GeneratedName '.WinImgNormalizer')
        Get-NamingPlanSignature $plan | Should -Be 'photo.png|photo__png.jpeg|Image|ExtensionSuffix'
    }

    It 'T018 reserves file and directory paths case-insensitively for controlled case-only image variants' {
        $source = New-NamingDirectory 'pure-case'
        $tree = New-NamingTree -Root $source -Files @('photo.png', 'PHOTO.PNG', 'Photo__png.jpeg') -Directories @('photo__png__2.jpeg')
        $plan = @(Get-WinImgOutputPlan -SourceRoot $source -SourceTree $tree -GeneratedName '.WinImgNormalizer')
        Get-NamingPlanSignature $plan | Should -Be (@(
            'PHOTO.PNG|PHOTO__png__3.jpeg|Image|NumericSuffix',
            'Photo__png.jpeg|Photo__png.jpeg|Image|Legacy',
            'photo.png|photo__png__4.jpeg|Image|NumericSuffix'
        ) -join "`n")
        $targets = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($row in $plan) { $targets.Add($row.OutputRelativePath) | Should -BeTrue }
    }

    It 'T018 plans case-conflicting videos against mirrored directories without changing their output extension' {
        $source = New-NamingDirectory 'pure-video'
        $tree = New-NamingTree -Root $source -Files @('clip.mp4', 'CLIP.MP4') -Directories @('CLIP__MP4.MP4')
        $plan = @(Get-WinImgOutputPlan -SourceRoot $source -SourceTree $tree -GeneratedName '.WinImgNormalizer')
        Get-NamingPlanSignature $plan | Should -Be (@(
            'CLIP.MP4|CLIP__mp4__2.mp4|Video|NumericSuffix',
            'clip.mp4|clip__mp4__3.mp4|Video|NumericSuffix'
        ) -join "`n")
    }

    It 'T018 preserves a source generated-name lookalike directory while allocating a separate generated namespace' {
        $source = New-NamingDirectory 'pure-generated'
        $tree = New-NamingTree -Root $source -Files @('.WinImgNormalizer\work\photo.png', 'reports\photo.jpg') -Directories @('.WinImgNormalizer', '.WinImgNormalizer\work', 'reports')
        $generated = Get-WinImgGeneratedName -TopLevelNames $tree.TopLevelNames
        $generated | Should -Be '.WinImgNormalizer__2'
        $plan = @(Get-WinImgOutputPlan -SourceRoot $source -SourceTree $tree -GeneratedName $generated)
        Get-NamingPlanSignature $plan | Should -Be (@(
            '.WinImgNormalizer\work\photo.png|.WinImgNormalizer\work\photo.jpeg|Image|Legacy',
            'reports\photo.jpg|reports\photo.jpeg|Image|Legacy'
        ) -join "`n")
        @($plan | Where-Object { $_.OutputRelativePath.StartsWith($generated + '\', [StringComparison]::OrdinalIgnoreCase) }).Count | Should -Be 0
    }

    It 'T019 returns no plan rows for an empty or unsupported-only inventory' {
        $source = New-NamingDirectory 'pure-empty'
        foreach ($files in @(@(), @('readme.txt', 'image.jpeg.bak'))) {
            $tree = New-NamingTree -Root $source -Files $files
            @(Get-WinImgOutputPlan -SourceRoot $source -SourceTree $tree -GeneratedName '.WinImgNormalizer').Count | Should -Be 0
        }
    }

    It 'T019 produces the same mapping and ordinal row ordering across shuffled enumerations in <Culture>' -ForEach @(
        @{ Culture = 'en-US' }, @{ Culture = 'de-DE' }, @{ Culture = 'tr-TR' }, @{ Culture = 'fi-FI' }, @{ Culture = '' }
    ) {
        $source = New-NamingDirectory 'pure-order'
        $files = @('photo.png', 'photo.heic', 'photo.jpg', 'photo__png.jpeg', 'PHOTO.PNG', 'clip.MP4', 'sub\photo.bmp', 'sub\photo.png', 'I.JPG', 'i.jpg', ([char]0x0130 + 'mage.png'), 'image.PNG', 'readme.txt')
        $directories = @('sub', 'photo.jpeg', 'photo__png__2.jpeg', 'empty', 'photo__png__3.jpeg')
        $baselineTree = New-NamingTree -Root $source -Files $files -Directories $directories
        $baseline = @(Get-WinImgOutputPlan -SourceRoot $source -SourceTree $baselineTree -GeneratedName '.WinImgNormalizer')
        $signature = Get-NamingPlanSignature $baseline
        $expectedOrder = New-Object 'Collections.Generic.List[string]'
        foreach ($file in $files) { if ($file -ne 'readme.txt') { $expectedOrder.Add($file) } }
        $expectedOrder.Sort([StringComparer]::Ordinal)
        (@($baseline | ForEach-Object SourceRelativePath) -join '|') | Should -Be ($expectedOrder.ToArray() -join '|')
        $priorCulture = [Threading.Thread]::CurrentThread.CurrentCulture
        $priorUiCulture = [Threading.Thread]::CurrentThread.CurrentUICulture
        try {
            [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo($Culture)
            [Threading.Thread]::CurrentThread.CurrentUICulture = [Globalization.CultureInfo]::GetCultureInfo($Culture)
            for ($seed = 1; $seed -le 16; $seed++) {
                $random = New-Object Random($seed)
                $shuffledFiles = @($files)
                $shuffledDirectories = @($directories)
                foreach ($values in @($shuffledFiles, $shuffledDirectories)) {
                    for ($index = $values.Count - 1; $index -gt 0; $index--) {
                        $other = $random.Next($index + 1)
                        $saved = $values[$index]; $values[$index] = $values[$other]; $values[$other] = $saved
                    }
                }
                $tree = New-NamingTree -Root $source -Files $shuffledFiles -Directories $shuffledDirectories
                $plan = @(Get-WinImgOutputPlan -SourceRoot $source -SourceTree $tree -GeneratedName '.WinImgNormalizer')
                Get-NamingPlanSignature $plan | Should -Be $signature
            }
        } finally {
            [Threading.Thread]::CurrentThread.CurrentCulture = $priorCulture
            [Threading.Thread]::CurrentThread.CurrentUICulture = $priorUiCulture
        }
    }
}

Describe 'M1-T03 collision integration and final-arrival preservation (T017-T019)' {
    It 'T018 writes and finalizes a real neutral candidate longer than MAX_PATH without shortening its mapped name' {
        $work = New-NamingDirectory 'long-neutral-candidate'
        $workRoot = Join-Path $work ('w' * 128)
        $null = [IO.Directory]::CreateDirectory($workRoot)
        $candidate = New-WinImgImageCandidate -WorkRoot $workRoot
        $candidate.Length | Should -BeGreaterOrEqual 260
        Test-Path -LiteralPath $candidate | Should -BeFalse
        [IO.Path]::GetFileName($candidate) | Should -Be 'image.jpeg'
        $target = Join-Path $work 'photo__png__2.jpeg'
        $nativeCandidate = Get-WinImgNativeOutputPath $candidate
        $nativeCandidate | Should -Be ('\\?\' + $candidate)
        $null = Invoke-NamingMagick @('-quiet', $jpegFixture, ('JPEG:' + $nativeCandidate))
        Move-WinImgPlannedImage -CandidatePath $candidate -DestinationPath $target
        Test-Path -LiteralPath $candidate | Should -BeFalse
        Invoke-NamingMagick @('identify', '-format', '%m|%w|%h|%n', $target) | Should -Be 'JPEG|20|16|1'
        $null = Invoke-NamingMagick @('-quiet', $target, 'null:')
    }

    It 'T018 reports failed scratch allocation as partial status and still retains a later valid planned output' {
        $source = New-NamingDirectory 'allocation-failure-source'
        $parent = New-NamingDirectory 'allocation-failure-output'
        foreach ($name in @('first.jpg', 'second.jpg')) { [IO.File]::WriteAllBytes((Join-Path $source $name), $jpegBytes) }
        $before = Get-NamingSourceState $source
        $trace = [pscustomobject]@{ Allocations = 0; Conversions = 0 }
        Mock New-WinImgImageCandidate {
            param([string]$WorkRoot)
            $trace.Allocations++
            if ($trace.Allocations -eq 1) { throw 'Controlled scratch allocation failure.' }
            return & $realImageCandidate $WorkRoot
        }
        $runner = {
            param([string]$Executable, [string[]]$Arguments)
            $trace.Conversions++
            [IO.File]::WriteAllBytes($Arguments[-1].Substring('JPEG:'.Length), $jpegBytes)
            return 0
        }
        $result = Invoke-NamingRun -Source $source -OutputParent $parent -ProcessRunner $runner
        $result.Code | Should -Be 2
        $trace.Allocations | Should -Be 2
        $trace.Conversions | Should -Be 1
        $run = Get-NamingRun $parent
        Test-Path -LiteralPath (Join-Path $run 'first.jpeg') | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $run 'second.jpeg') -PathType Leaf | Should -BeTrue
        $null = Invoke-NamingMagick @('-quiet', (Join-Path $run 'second.jpeg'), 'null:')
        Get-NamingSourceState $source | Should -Be $before
        Get-NamingLog $run | Should -Match 'SUMMARY ConvertedImages=1 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=1'
    }

    It 'T017-T019 converts all real JPG, PNG and BMP collisions, preserves sources and video bytes, and logs the complete plan before conversion' {
        $source = New-NamingDirectory 'real-collisions'
        $parent = New-NamingDirectory 'real-output'
        foreach ($relative in @('photo.jpeg', 'photo__bmp.jpeg', 'sub')) {
            $null = [IO.Directory]::CreateDirectory((Join-Path $source $relative))
        }
        $cases = @(
            @{ Source = 'photo.jpg'; Output = 'photo__jpg.jpeg'; Coder = 'JPEG'; Width = 32; Height = 24 },
            @{ Source = 'photo.png'; Output = 'photo__png__2.jpeg'; Coder = 'PNG'; Width = 30; Height = 22 },
            @{ Source = 'photo.bmp'; Output = 'photo__bmp__2.jpeg'; Coder = 'BMP'; Width = 28; Height = 20 },
            @{ Source = 'photo__png.jpeg'; Output = 'photo__png.jpeg'; Coder = 'JPEG'; Width = 26; Height = 18 },
            @{ Source = 'sub\photo.png'; Output = 'sub\photo.jpeg'; Coder = 'PNG'; Width = 24; Height = 16 }
        )
        foreach ($image in $cases) {
            $null = Invoke-NamingMagick @('-quiet', '-size', ($image.Width.ToString() + 'x' + $image.Height),
                'gradient:#205080-#E0A070', '-depth', '8', ($image.Coder + ':' + (Join-Path $source $image.Source)))
        }
        $video = Join-Path $source 'clip.MP4'
        [IO.File]::WriteAllBytes($video, [byte[]](0, 255, 11, 92, 64, 0, 17, 128))
        $before = Get-NamingSourceState $source
        $tree = Get-WinImgSourceTree -SourceRoot $source
        $generated = Get-WinImgGeneratedName -TopLevelNames $tree.TopLevelNames
        $plan = @(Get-WinImgOutputPlan -SourceRoot $source -SourceTree $tree -GeneratedName $generated)
        $trace = [pscustomobject]@{ Calls = 0; PlanBeforeFirstCall = $null; ScratchPaths = New-Object 'Collections.Generic.List[string]' }
        $runner = {
            param([string]$Executable, [string[]]$Arguments)
            $trace.Calls++
            $run = Get-NamingRun $parent
            if ($trace.Calls -eq 1) { $trace.PlanBeforeFirstCall = Get-NamingLog $run }
            $scratch = $Arguments[-1].Substring('JPEG:'.Length)
            $ordinaryScratch = ConvertFrom-NamingNativePath $scratch
            if (-not $ordinaryScratch.StartsWith((Join-Path $run ($generated + '\work\')), [StringComparison]::OrdinalIgnoreCase)) {
                throw 'ImageMagick was directed outside the generated scratch namespace.'
            }
            $trace.ScratchPaths.Add($scratch)
            & $Executable @Arguments 1>$null 2>$null
            return $LASTEXITCODE
        }
        $result = Invoke-NamingRun -Source $source -OutputParent $parent -ProcessRunner $runner
        $result.Code | Should -Be 0 -Because $result.Text
        $trace.Calls | Should -Be 5
        $run = Get-NamingRun $parent
        $log = Get-NamingLog $run
        $logged = @([regex]::Matches($log, '(?m)^.*\[INFO\] (PLAN .*?)\r?$') | ForEach-Object { $_.Groups[1].Value })
        $expected = @($plan | ForEach-Object { 'PLAN {0}: {1} -> {2} ({3})' -f $_.Kind, $_.SourceRelativePath, $_.OutputRelativePath, $_.NamingReason })
        ($logged -join "`n") | Should -Be ($expected -join "`n")
        foreach ($line in $expected) { $trace.PlanBeforeFirstCall | Should -Match ([regex]::Escape($line)) }
        foreach ($image in $cases) {
            $target = Join-Path $run $image.Output
            (Get-Item -LiteralPath $target).Length | Should -BeGreaterThan 0
            Invoke-NamingMagick @('identify', '-format', '%m|%w|%h|%n', $target) |
                Should -Be ('JPEG|{0}|{1}|1' -f $image.Width, $image.Height)
            $null = Invoke-NamingMagick @('-quiet', $target, 'null:')
        }
        (Get-FileHash -LiteralPath (Join-Path $run 'clip.MP4') -Algorithm SHA256).Hash |
            Should -Be (Get-FileHash -LiteralPath $video -Algorithm SHA256).Hash
        foreach ($directory in @('photo.jpeg', 'photo__bmp.jpeg', 'sub')) {
            Test-Path -LiteralPath (Join-Path $run $directory) -PathType Container | Should -BeTrue
        }
        @(Get-ChildItem -LiteralPath (Join-Path $run ($generated + '\work')) -Force).Count | Should -Be 0
        Get-NamingSourceState $source | Should -Be $before
        $log | Should -Match 'SUMMARY ConvertedImages=5 CopiedVideos=1 Duplicates=0 Unsupported=0 Errors=0'
    }

    It 'T017 exercises a controlled HEIC conversion alongside JPG and PNG with no optional-codec skip' {
        $source = New-NamingDirectory 'mocked-heic'
        $parent = New-NamingDirectory 'mocked-heic-output'
        foreach ($name in @('photo.heic', 'photo.jpg', 'photo.png')) {
            [IO.File]::WriteAllText((Join-Path $source $name), ('Controlled source bytes: ' + $name))
        }
        $before = Get-NamingSourceState $source
        Mock Get-WinImgMagickInfo {
            return [pscustomobject]@{
                Path = $magick; VersionText = 'Controlled naming-only codec capabilities'; Formats = @{
                    JPEG = @{ Read = $true; Write = $true }; PNG = @{ Read = $true; Write = $true }; HEIC = @{ Read = $true; Write = $false }
                }
            }
        }
        # ASCII input here proves naming/routing only; actual HEVC decoding and
        # decoder-visible image counts belong to Frames.Tests.ps1.
        Mock Get-WinImgSourceImageInfo {
            return [pscustomobject]@{ SourceCount = 1; Omitted = 0; Unit = 'Images'; Policy = 'DecoderPrimaryOrFirstImage'; Decoder = 'HEIC' }
        }
        $trace = New-Object 'Collections.Generic.List[string]'
        $runner = {
            param([string]$Executable, [string[]]$Arguments)
            $trace.Add($Arguments[1])
            [IO.File]::WriteAllBytes($Arguments[-1].Substring('JPEG:'.Length), $jpegBytes)
            return 0
        }
        $result = Invoke-NamingRun -Source $source -OutputParent $parent -ProcessRunner $runner
        $result.Code | Should -Be 0 -Because $result.Text
        $trace.Count | Should -Be 3
        $run = Get-NamingRun $parent
        foreach ($extension in @('heic', 'jpg', 'png')) {
            $target = Join-Path $run ('photo__' + $extension + '.jpeg')
            Test-Path -LiteralPath $target -PathType Leaf | Should -BeTrue
            $null = Invoke-NamingMagick @('-quiet', $target, 'null:')
        }
        Get-NamingSourceState $source | Should -Be $before
        Get-NamingLog $run | Should -Match 'SUMMARY ConvertedImages=3 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=0'
    }

    It 'T018 preserves an external <Arrival> arriving at an image final name while conversion writes scratch and returns partial status' -ForEach @(
        @{ Arrival = 'file' }, @{ Arrival = 'directory' }
    ) {
        $source = New-NamingDirectory 'image-arrival-source'
        $parent = New-NamingDirectory 'image-arrival-output'
        [IO.File]::WriteAllBytes((Join-Path $source 'single.jpg'), $jpegBytes)
        $before = Get-NamingSourceState $source
        $trace = [pscustomobject]@{ Calls = 0; ArrivalPath = $null; ScratchPath = $null }
        $runner = {
            param([string]$Executable, [string[]]$Arguments)
            $trace.Calls++
            $run = Get-NamingRun $parent
            $trace.ArrivalPath = Join-Path $run 'single.jpeg'
            $trace.ScratchPath = $Arguments[-1].Substring('JPEG:'.Length)
            if ((ConvertFrom-NamingNativePath $trace.ScratchPath).Equals($trace.ArrivalPath, [StringComparison]::OrdinalIgnoreCase)) {
                throw 'Conversion was directed at its final target.'
            }
            if ($Arrival -eq 'directory') {
                $null = [IO.Directory]::CreateDirectory($trace.ArrivalPath)
                [IO.File]::WriteAllText((Join-Path $trace.ArrivalPath 'external.txt'), 'External directory arrival remains untouched.')
            } else { [IO.File]::WriteAllText($trace.ArrivalPath, 'External file arrival remains untouched.') }
            [IO.File]::WriteAllBytes($trace.ScratchPath, $jpegBytes)
            return 0
        }
        $result = Invoke-NamingRun -Source $source -OutputParent $parent -ProcessRunner $runner
        $result.Code | Should -Be 2
        $trace.Calls | Should -Be 1
        if ($Arrival -eq 'directory') {
            [IO.File]::ReadAllText((Join-Path $trace.ArrivalPath 'external.txt')) | Should -Be 'External directory arrival remains untouched.'
            @(Get-ChildItem -LiteralPath $trace.ArrivalPath -Force).Count | Should -Be 1
        } else { [IO.File]::ReadAllText($trace.ArrivalPath) | Should -Be 'External file arrival remains untouched.' }
        Test-Path -LiteralPath $trace.ScratchPath | Should -BeFalse
        Get-NamingSourceState $source | Should -Be $before
        Get-NamingLog (Get-NamingRun $parent) | Should -Match 'SUMMARY ConvertedImages=0 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=1'
    }

    It 'T018 refuses an external <Arrival> that appears after the final availability check for <Media>' -ForEach @(
        @{ Arrival = 'file'; Media = 'image' }, @{ Arrival = 'directory'; Media = 'image' },
        @{ Arrival = 'file'; Media = 'video' }, @{ Arrival = 'directory'; Media = 'video' }
    ) {
        $work = New-NamingDirectory 'post-check-arrival'
        $candidate = Join-Path $work 'owned-candidate.bin'
        $target = Join-Path $work 'final.bin'
        [IO.File]::WriteAllBytes($candidate, [byte[]](9, 8, 7, 6, 0, 255))
        $before = (Get-FileHash -LiteralPath $candidate -Algorithm SHA256).Hash
        Mock Assert-WinImgOutputAvailable {
            param([string]$Path)
            & $realOutputAvailable $Path
            if ($Arrival -eq 'directory') {
                $null = [IO.Directory]::CreateDirectory($Path)
                [IO.File]::WriteAllText((Join-Path $Path 'external.txt'), 'Post-check directory arrival.')
            } else { [IO.File]::WriteAllText($Path, 'Post-check file arrival.') }
        }
        if ($Media -eq 'image') {
            { Move-WinImgPlannedImage -CandidatePath $candidate -DestinationPath $target } | Should -Throw
        } else {
            { Copy-WinImgPlannedVideo -SourcePath $candidate -DestinationPath $target } | Should -Throw
        }
        Should -Invoke Assert-WinImgOutputAvailable -Times 1 -Exactly
        (Get-FileHash -LiteralPath $candidate -Algorithm SHA256).Hash | Should -Be $before
        if ($Arrival -eq 'directory') {
            [IO.File]::ReadAllText((Join-Path $target 'external.txt')) | Should -Be 'Post-check directory arrival.'
            @(Get-ChildItem -LiteralPath $target -Force).Count | Should -Be 1
        } else { [IO.File]::ReadAllText($target) | Should -Be 'Post-check file arrival.' }
    }

    It 'T018 returns partial status and preserves an external video arrival between checking and copying' {
        $source = New-NamingDirectory 'video-arrival-source'
        $parent = New-NamingDirectory 'video-arrival-output'
        [IO.File]::WriteAllBytes((Join-Path $source 'clip.mp4'), [byte[]](1, 2, 3, 0, 255))
        $before = Get-NamingSourceState $source
        $trace = [pscustomobject]@{ Checks = 0; ArrivalPath = $null }
        Mock Assert-WinImgOutputAvailable {
            param([string]$Path)
            & $realOutputAvailable $Path
            $trace.Checks++
            if ($trace.Checks -eq 2) {
                $trace.ArrivalPath = $Path
                [IO.File]::WriteAllText($Path, 'External video final arrival.')
            }
        }
        $result = Invoke-NamingRun -Source $source -OutputParent $parent -ProcessRunner { throw 'Video-only input attempted an image conversion.' }
        $result.Code | Should -Be 2
        $trace.Checks | Should -Be 2
        [IO.File]::ReadAllText($trace.ArrivalPath) | Should -Be 'External video final arrival.'
        Get-NamingSourceState $source | Should -Be $before
        Get-NamingLog (Get-NamingRun $parent) | Should -Match 'SUMMARY ConvertedImages=0 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=1'
    }
}
