BeforeAll {
    $repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    $application = Join-Path $repository 'WinImgNormalizer.ps1'

    function Assert-SafetyNoReparseAncestors {
        param([string]$Path)
        $current = [IO.DirectoryInfo]::new([IO.Path]::GetFullPath($Path))
        while ($null -ne $current) {
            if ($current.Exists -and (($current.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0)) {
                throw 'Safety regression ownership path includes a reparse point.'
            }
            $current = $current.Parent
        }
    }

    $scratchParent = [IO.Path]::GetFullPath((Join-Path $repository '.scratch'))
    Assert-SafetyNoReparseAncestors $scratchParent
    foreach ($pictures in @([Environment]::GetFolderPath('MyPictures'), (Join-Path $env:USERPROFILE 'Pictures'))) {
        if ([string]::IsNullOrWhiteSpace($pictures)) { continue }
        $picturesFull = [IO.Path]::GetFullPath($pictures).TrimEnd('\', '/')
        if ($scratchParent.Equals($picturesFull, [StringComparison]::OrdinalIgnoreCase) -or
            $scratchParent.StartsWith($picturesFull + '\', [StringComparison]::OrdinalIgnoreCase) -or
            $picturesFull.StartsWith($scratchParent + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Safety regression scratch overlaps a real Pictures location.'
        }
    }
    $git = Get-Command git -CommandType Application -ErrorAction Stop | Select-Object -First 1
    & $git.Source -C $repository check-ignore --quiet --no-index -- (Join-Path $scratchParent 'safety-ignore-probe')
    if ($LASTEXITCODE -ne 0) { throw 'Safety regression synthetic scratch must already be ignored.' }
    $ownedRoot = Join-Path $scratchParent ('M1-T06-' + [Guid]::NewGuid().ToString('N'))
    if (Test-Path -LiteralPath $ownedRoot) { throw 'Safety regression ownership directory already exists.' }
    $null = [IO.Directory]::CreateDirectory($ownedRoot)
    [IO.File]::WriteAllText((Join-Path $ownedRoot '.winimg-fixture-root'), 'M1-T06 synthetic test ownership')

    function New-SafetyDirectory {
        param([string]$Label)
        $path = [IO.Path]::GetFullPath((Join-Path $ownedRoot ($Label + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))))
        if (-not $path.StartsWith($ownedRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Safety regression path escapes its owned scratch directory.'
        }
        Assert-SafetyNoReparseAncestors $path
        if (Test-Path -LiteralPath $path) { throw 'Safety regression directory already exists.' }
        $null = [IO.Directory]::CreateDirectory($path)
        return $path
    }

    function New-SafetyChildDirectory {
        param([string]$Root, [string]$RelativePath)
        $path = [IO.Path]::GetFullPath((Join-Path $Root $RelativePath))
        if (-not $path.StartsWith($Root + '\', [StringComparison]::OrdinalIgnoreCase) -or
            -not $path.StartsWith($ownedRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Safety regression child path escapes its owned tree.'
        }
        Assert-SafetyNoReparseAncestors $path
        $null = [IO.Directory]::CreateDirectory($path)
        return $path
    }

    function Get-SafetyFileState {
        param([string]$Root, [IO.FileInfo[]]$Files)
        return @($Files | Sort-Object FullName | ForEach-Object {
            [pscustomobject]@{
                Path = $_.FullName.Substring($Root.Length + 1)
                Hash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
                Length = $_.Length
                CreationTicks = $_.CreationTimeUtc.Ticks
                ModifiedTicks = $_.LastWriteTimeUtc.Ticks
                Attributes = [int]$_.Attributes
            }
        })
    }

    function Get-SafetyTreeState {
        param([string]$Root)
        $directories = @([IO.DirectoryInfo]::new($Root)) + @(Get-ChildItem -LiteralPath $Root -Recurse -Force -Directory)
        return [pscustomobject]@{
            Files = @(Get-SafetyFileState -Root $Root -Files @(Get-ChildItem -LiteralPath $Root -Recurse -Force -File))
            Directories = @($directories | Sort-Object FullName | ForEach-Object {
                [pscustomobject]@{
                    Path = if ($_.FullName -eq $Root) { '.' } else { $_.FullName.Substring($Root.Length + 1) }
                    CreationTicks = $_.CreationTimeUtc.Ticks
                    ModifiedTicks = $_.LastWriteTimeUtc.Ticks
                    Attributes = [int]$_.Attributes
                }
            })
        } | ConvertTo-Json -Depth 5 -Compress
    }

    function Get-SafetyOrdinalSignature {
        param([string[]]$Values)
        $ordered = New-Object 'Collections.Generic.List[string]'
        foreach ($value in $Values) { $ordered.Add($value) }
        $ordered.Sort([StringComparer]::Ordinal)
        return $ordered.ToArray() -join "`n"
    }

    function Get-SafetyLog {
        param([string]$Run, [string]$GeneratedName)
        $logs = @(Get-ChildItem -LiteralPath (Join-Path $Run ($GeneratedName + '\reports')) -File -Filter '*.log')
        $logs.Count | Should -Be 1
        return Get-Content -LiteralPath $logs[0].FullName -Raw
    }

    function Get-SafetyLogLines {
        param([string]$Log, [string]$Level, [string]$Prefix)
        return @([regex]::Matches($Log, ('(?m)^.*?\[' + [regex]::Escape($Level) + '\] (' + [regex]::Escape($Prefix) + '.*?)\r?$')) | ForEach-Object { $_.Groups[1].Value })
    }

    function Invoke-SafetyRun {
        param([string]$Source, [string]$OutputParent)
        $observed = @(& Invoke-WinImgNormalizer -Source $Source -OutputParent $OutputParent -MagickPath $magick 6>&1 3>&1 2>&1)
        $codes = @($observed | Where-Object { $_ -is [int] -or $_ -is [long] })
        if ($codes.Count -ne 1) { throw ('Expected one callable safety regression exit status: ' + ($observed -join "`n")) }
        return [pscustomobject]@{ Code = $codes[0]; Text = (@($observed | Where-Object { $_ -isnot [int] -and $_ -isnot [long] }) -join "`n") }
    }

    if ([string]::IsNullOrWhiteSpace($env:WINIMG_TEST_MAGICK) -or -not [IO.Path]::IsPathRooted($env:WINIMG_TEST_MAGICK)) {
        throw 'WINIMG_TEST_MAGICK must explicitly select the verified test magick.exe.'
    }
    $magick = [IO.Path]::GetFullPath($env:WINIMG_TEST_MAGICK)
    Assert-SafetyNoReparseAncestors $magick
    if (-not (Test-Path -LiteralPath $magick -PathType Leaf) -or [IO.Path]::GetFileName($magick) -ne 'magick.exe') {
        throw 'The explicitly selected safety regression test magick.exe is missing.'
    }

    function Invoke-SafetyMagick {
        param([string[]]$Arguments)
        $result = & $magick @Arguments 2>&1
        $code = $LASTEXITCODE
        if ($code -ne 0) { throw ('Safety regression ImageMagick exited {0}: {1}' -f $code, ($result -join "`n")) }
        return ($result -join "`n")
    }

    function Assert-SafetyJpeg {
        param([string]$SourceRoot, [string]$Run, [object]$Case)
        $target = Join-Path $Run $Case.Output
        $nativeTarget = Get-WinImgNativeOutputPath -Path $target
        Invoke-SafetyMagick @('identify', '+ping', '-regard-warnings', '-format', '%m|%w|%h|%n', $nativeTarget) |
            Should -Be ('JPEG|{0}|{1}|1' -f $Case.Width, $Case.Height)
        (Get-Item -LiteralPath $target).Length | Should -BeGreaterThan 0
        # Channel means provide a small tolerance-based check that the expected
        # synthetic source reached this output, without cross-build byte matching.
        $nativeSource = Get-WinImgNativeOutputPath -Path (Join-Path $SourceRoot $Case.Source)
        $sourceMeans = (Invoke-SafetyMagick @('-quiet', $nativeSource, '-format', '%[fx:mean.r]|%[fx:mean.g]|%[fx:mean.b]', 'info:')) -split '\|'
        $outputMeans = (Invoke-SafetyMagick @('-quiet', $nativeTarget, '-format', '%[fx:mean.r]|%[fx:mean.g]|%[fx:mean.b]', 'info:')) -split '\|'
        $sourceMeans.Count | Should -Be 3
        $outputMeans.Count | Should -Be 3
        for ($channel = 0; $channel -lt 3; $channel++) {
            $sourceMean = [double]::Parse($sourceMeans[$channel], [Globalization.CultureInfo]::InvariantCulture)
            $outputMean = [double]::Parse($outputMeans[$channel], [Globalization.CultureInfo]::InvariantCulture)
            [Math]::Abs($sourceMean - $outputMean) | Should -BeLessOrEqual 0.02
        }
    }

    . $application
}

Describe 'M1-T06 complete mixed-tree source preservation and run isolation (T028)' {
    It 'T028 preserves all originals and prior outputs while two complete runs keep deterministic separate media and retained links' {
        $source = New-SafetyDirectory 'source'
        $parent = New-SafetyDirectory 'outputs'
        $literalDirectory = 'nested [safe] ' + [char]0x00FC
        $generatedName = '.WinImgNormalizer__2'
        $directories = @('a', 'b', 'clips', 'photo.jpeg', 'photo__bmp.jpeg', 'empty [keep]', '.WinImgNormalizer', '.WinImgNormalizer\work', $literalDirectory, ($literalDirectory + '\empty'))
        foreach ($relative in $directories) { $null = New-SafetyChildDirectory -Root $source -RelativePath $relative }
        $images = @(
            @{ Source = '.hidden.png'; Output = '.hidden.jpeg'; Coder = 'PNG'; Width = 20; Height = 18; Reason = 'Legacy'; Skipped = $false },
            @{ Source = '.WinImgNormalizer\work\preserved.png'; Output = '.WinImgNormalizer\work\preserved.jpeg'; Coder = 'PNG'; Width = 20; Height = 14; Reason = 'Legacy'; Skipped = $false },
            @{ Source = 'a\shared.png'; Output = 'a\shared.jpeg'; Coder = 'PNG'; Width = 24; Height = 18; Reason = 'Legacy'; Skipped = $false },
            @{ Source = 'b\shared.png'; Output = 'b\shared.jpeg'; Coder = 'PNG'; Width = 24; Height = 18; Reason = 'Legacy'; Skipped = $true },
            @{ Source = ($literalDirectory + '\still.png'); Output = ($literalDirectory + '\still.jpeg'); Coder = 'PNG'; Width = 22; Height = 16; Reason = 'Legacy'; Skipped = $false },
            @{ Source = 'photo.bmp'; Output = 'photo__bmp__2.jpeg'; Coder = 'BMP'; Width = 28; Height = 20; Reason = 'NumericSuffix'; Skipped = $false },
            @{ Source = 'photo.jpg'; Output = 'photo__jpg.jpeg'; Coder = 'JPEG'; Width = 32; Height = 24; Reason = 'ExtensionSuffix'; Skipped = $false },
            @{ Source = 'photo.png'; Output = 'photo__png__2.jpeg'; Coder = 'PNG'; Width = 30; Height = 22; Reason = 'NumericSuffix'; Skipped = $false },
            @{ Source = 'photo__png.jpeg'; Output = 'photo__png.jpeg'; Coder = 'JPEG'; Width = 26; Height = 18; Reason = 'Legacy'; Skipped = $false }
        )
        foreach ($image in @($images | Where-Object { -not $_.Skipped })) {
            $null = Invoke-SafetyMagick @('-quiet', '-size', ('{0}x{1}' -f $image.Width, $image.Height), 'gradient:#205080-#E0A070', '-depth', '8', ($image.Coder + ':' + (Join-Path $source $image.Source)))
        }
        [IO.File]::Copy((Join-Path $source 'a\shared.png'), (Join-Path $source 'b\shared.png'), $false)
        $videos = @(
            @{ Source = 'a\shared clip.MOV'; Output = 'a\shared clip.MOV'; Length = 8193; Skipped = $false },
            @{ Source = 'b\shared clip.MOV'; Output = 'b\shared clip.MOV'; Length = 8193; Skipped = $true },
            @{ Source = 'clips\opaque clip.mp4'; Output = 'clips\opaque clip.mp4'; Length = 4096; Skipped = $false }
        )
        foreach ($video in $videos) {
            $bytes = New-Object byte[] $video.Length
            for ($offset = 0; $offset -lt $bytes.Length; $offset++) { $bytes[$offset] = [byte](($offset * 31 + 13) % 256) }
            [IO.File]::WriteAllBytes((Join-Path $source $video.Source), $bytes)
        }
        foreach ($relative in @('notes [keep].txt', '.hidden.keep', '.WinImgNormalizer\work\user.txt')) {
            [IO.File]::WriteAllText((Join-Path $source $relative), 'Unsupported synthetic source stays preserved.')
        }
        foreach ($relative in @('.hidden.png', '.hidden.keep')) {
            $path = Join-Path $source $relative
            [IO.File]::SetAttributes($path, ([IO.File]::GetAttributes($path) -bor [IO.FileAttributes]::Hidden))
        }
        $fixedUtc = [DateTime]::Parse('2020-02-03T04:05:06Z').ToUniversalTime()
        foreach ($file in @(Get-ChildItem -LiteralPath $source -Recurse -Force -File)) {
            [IO.File]::SetCreationTimeUtc($file.FullName, $fixedUtc)
            [IO.File]::SetLastWriteTimeUtc($file.FullName, $fixedUtc)
        }
        $priorDirectory = New-SafetyChildDirectory -Root $parent -RelativePath 'prior [user]'
        [IO.File]::WriteAllText((Join-Path $priorDirectory 'previous.jpeg'), 'Unrelated prior output remains untouched.')
        $parentKeep = Join-Path $parent 'untouched [user].txt'
        [IO.File]::WriteAllText($parentKeep, 'Unrelated output-parent file remains untouched.')
        $sourceBefore = Get-SafetyTreeState $source
        $priorDirectoryBefore = Get-SafetyTreeState $priorDirectory
        $parentFileBefore = @(Get-SafetyFileState -Root $parent -Files @([IO.FileInfo]::new($parentKeep))) | ConvertTo-Json -Depth 3 -Compress
        $initialDirectories = @(Get-ChildItem -LiteralPath $parent -Force -Directory | ForEach-Object FullName)
        $expectedPlans = @($images | ForEach-Object { 'PLAN Image: {0} -> {1} ({2})' -f $_.Source, $_.Output, $_.Reason }) +
            @($videos | ForEach-Object { 'PLAN Video: {0} -> {1} (Legacy)' -f $_.Source, $_.Output })
        $expectedPlanSignature = Get-SafetyOrdinalSignature $expectedPlans
        $expectedSkipSignature = Get-SafetyOrdinalSignature @(
            'Heuristic duplicate skipped: b\shared clip.MOV (retained source: a\shared clip.MOV; retained output: a\shared clip.MOV; retained status: CopiedVideo)',
            'Heuristic duplicate skipped: b\shared.png (retained source: a\shared.png; retained output: a\shared.jpeg; retained status: Converted)'
        )
        $expectedMedia = @($images | Where-Object { -not $_.Skipped } | ForEach-Object Output) + @($videos | Where-Object { -not $_.Skipped } | ForEach-Object Output)
        $firstRun = $null
        $firstRunBefore = $null
        $firstPlanOrder = $null
        $firstSkipOrder = $null
        foreach ($iteration in 1..2) {
            $existingDirectories = @(Get-ChildItem -LiteralPath $parent -Force -Directory | ForEach-Object FullName)
            $result = Invoke-SafetyRun -Source $source -OutputParent $parent
            $result.Code | Should -Be 0 -Because $result.Text
            $newDirectories = @(Get-ChildItem -LiteralPath $parent -Force -Directory | Where-Object { $existingDirectories -notcontains $_.FullName })
            $newDirectories.Count | Should -Be 1
            $run = $newDirectories[0].FullName
            [IO.Directory]::GetParent($run).FullName | Should -Be $parent
            Get-SafetyTreeState $source | Should -Be $sourceBefore
            Get-SafetyTreeState $priorDirectory | Should -Be $priorDirectoryBefore
            (@(Get-SafetyFileState -Root $parent -Files @([IO.FileInfo]::new($parentKeep))) | ConvertTo-Json -Depth 3 -Compress) | Should -Be $parentFileBefore
            foreach ($relative in $directories) { Test-Path -LiteralPath (Join-Path $run $relative) -PathType Container | Should -BeTrue }
            Test-Path -LiteralPath (Join-Path $run $generatedName) -PathType Container | Should -BeTrue
            @(Get-ChildItem -LiteralPath (Join-Path $run ($generatedName + '\work')) -Force).Count | Should -Be 0
            foreach ($image in $images) {
                if ($image.Skipped) { Test-Path -LiteralPath (Join-Path $run $image.Output) | Should -BeFalse }
                else { Assert-SafetyJpeg -SourceRoot $source -Run $run -Case $image }
            }
            foreach ($video in $videos) {
                if ($video.Skipped) { Test-Path -LiteralPath (Join-Path $run $video.Output) | Should -BeFalse }
                else {
                    $target = Join-Path $run $video.Output
                    $inputPath = Join-Path $source $video.Source
                    (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash | Should -Be (Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash
                    (Get-Item -LiteralPath $target).Length | Should -Be $video.Length
                }
            }
            foreach ($case in @($images + $videos | Where-Object { -not $_.Skipped })) {
                $outputFile = Get-Item -LiteralPath (Join-Path $run $case.Output) -Force
                $inputFile = Get-Item -LiteralPath (Join-Path $source $case.Source) -Force
                $outputFile.CreationTimeUtc.Ticks | Should -Be $inputFile.CreationTimeUtc.Ticks
                $outputFile.LastWriteTimeUtc.Ticks | Should -Be $inputFile.LastWriteTimeUtc.Ticks
            }
            $actualMedia = @(Get-ChildItem -LiteralPath $run -Recurse -Force -File | ForEach-Object { $_.FullName.Substring($run.Length + 1) } | Where-Object { -not $_.StartsWith($generatedName + '\', [StringComparison]::OrdinalIgnoreCase) })
            Get-SafetyOrdinalSignature $actualMedia | Should -Be (Get-SafetyOrdinalSignature $expectedMedia)
            @(Get-ChildItem -LiteralPath $run -Recurse -Force -File).Count | Should -Be ($expectedMedia.Count + 2)
            @(Get-ChildItem -LiteralPath (Join-Path $run ($generatedName + '\reports')) -File -Filter '*.csv').Count | Should -Be 1
            $log = Get-SafetyLog -Run $run -GeneratedName $generatedName
            $planLines = @(Get-SafetyLogLines -Log $log -Level 'INFO' -Prefix 'PLAN ')
            Get-SafetyOrdinalSignature $planLines | Should -Be $expectedPlanSignature
            $planLines.Count | Should -Be 12
            $skipLines = @(Get-SafetyLogLines -Log $log -Level 'SKIP' -Prefix 'Heuristic duplicate skipped: ')
            Get-SafetyOrdinalSignature $skipLines | Should -Be $expectedSkipSignature
            $skipLines.Count | Should -Be 2
            $log | Should -Match 'SUMMARY ConvertedImages=8 CopiedVideos=2 Duplicates=2 Unsupported=3 Errors=0'
            if ($iteration -eq 1) {
                $firstRun = $run
                $firstRunBefore = Get-SafetyTreeState $run
                $firstPlanOrder = $planLines -join "`n"
                $firstSkipOrder = $skipLines -join "`n"
            } else {
                $run | Should -Not -Be $firstRun
                Get-SafetyTreeState $firstRun | Should -Be $firstRunBefore
                ($planLines -join "`n") | Should -Be $firstPlanOrder
                ($skipLines -join "`n") | Should -Be $firstSkipOrder
            }
        }
        $finalRunDirectories = @(Get-ChildItem -LiteralPath $parent -Force -Directory | Where-Object { $initialDirectories -notcontains $_.FullName })
        $finalRunDirectories.Count | Should -Be 2
        Get-SafetyTreeState $source | Should -Be $sourceBefore
    }
}
