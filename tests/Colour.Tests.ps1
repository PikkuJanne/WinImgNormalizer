BeforeAll {
    $repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    $assets = Join-Path $PSScriptRoot 'fixtures/colour'

    function Assert-ColourNoReparseAncestors {
        param([string]$Path)
        $current = [IO.DirectoryInfo]::new([IO.Path]::GetFullPath($Path))
        while ($null -ne $current) {
            if ($current.Exists -and (($current.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0)) {
                throw 'Colour fixture path includes a reparse point.'
            }
            $current = $current.Parent
        }
    }
    $scratch = [IO.Path]::GetFullPath((Join-Path $repository '.scratch'))
    Assert-ColourNoReparseAncestors $scratch
    Assert-ColourNoReparseAncestors $assets
    foreach ($pictures in @([Environment]::GetFolderPath('MyPictures'), (Join-Path $env:USERPROFILE 'Pictures'))) {
        if ([string]::IsNullOrWhiteSpace($pictures)) { continue }
        $full = [IO.Path]::GetFullPath($pictures).TrimEnd('\', '/')
        if ($scratch.Equals($full, [StringComparison]::OrdinalIgnoreCase) -or
            $scratch.StartsWith($full + '\', [StringComparison]::OrdinalIgnoreCase) -or
            $full.StartsWith($scratch + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Colour fixture scratch overlaps real Pictures.'
        }
    }
    $git = Get-Command git -CommandType Application -ErrorAction Stop | Select-Object -First 1
    & $git.Source -C $repository check-ignore --quiet --no-index -- (Join-Path $scratch 'colour-ignore-probe')
    if ($LASTEXITCODE -ne 0) { throw 'Colour fixture scratch must already be ignored.' }
    $ownedRoot = Join-Path $scratch ('M2-T02-colour-' + [Guid]::NewGuid().ToString('N'))
    if (Test-Path -LiteralPath $ownedRoot) { throw 'Colour fixture ownership directory already exists.' }
    $null = [IO.Directory]::CreateDirectory($ownedRoot)
    [IO.File]::WriteAllText((Join-Path $ownedRoot '.winimg-fixture-root'), 'M2-T02 synthetic colour test ownership')

    function New-ColourDirectory {
        param([string]$Label)
        $path = [IO.Path]::GetFullPath((Join-Path $ownedRoot ($Label + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))))
        if (-not $path.StartsWith($ownedRoot + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Colour fixture escaped ownership.' }
        Assert-ColourNoReparseAncestors $path
        if (Test-Path -LiteralPath $path) { throw 'Colour fixture directory already exists.' }
        $null = [IO.Directory]::CreateDirectory($path)
        return $path
    }
    function New-ColourCase {
        param([string]$Label, [string]$Name)
        return [pscustomobject]@{
            Source = New-ColourDirectory ($Label + '-source')
            Parent = New-ColourDirectory ($Label + '-output')
            Probe = New-ColourDirectory ($Label + '-probe')
            Name = $Name
        }
    }
    function Get-ColourSourceState {
        param([string]$Source)
        return @(Get-ChildItem -LiteralPath $Source -Recurse -Force -File | Sort-Object FullName | ForEach-Object {
            [pscustomobject]@{
                Path = $_.FullName.Substring($Source.Length + 1)
                Hash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
                Length = $_.Length; CreationTicks = $_.CreationTimeUtc.Ticks
                ModifiedTicks = $_.LastWriteTimeUtc.Ticks; Attributes = [int]$_.Attributes
            }
        }) | ConvertTo-Json -Depth 3 -Compress
    }
    function Set-ColourSourceTimes {
        param([string]$Path)
        $fixed = [DateTime]::Parse('2020-02-03T04:05:06Z').ToUniversalTime()
        [IO.File]::SetCreationTimeUtc($Path, $fixed)
        [IO.File]::SetLastWriteTimeUtc($Path, $fixed)
    }
    function Get-ColourBigEndian {
        param([byte[]]$Bytes, [int]$Offset, [int]$Count)
        $value = [long]0
        for ($index = 0; $index -lt $Count; $index++) { $value = ($value -shl 8) + $Bytes[$Offset + $index] }
        return $value
    }
    $manifest = Get-Content -LiteralPath (Join-Path $assets 'MANIFEST.json') -Raw | ConvertFrom-Json
    $reference = Get-Content -LiteralPath (Join-Path $assets 'REFERENCE.json') -Raw | ConvertFrom-Json
    foreach ($profile in $manifest.profiles) {
        $file = Join-Path $assets ([IO.Path]::GetFileName($profile.source_path))
        (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash.ToLowerInvariant() | Should -Be $profile.sha256
        $bytes = [IO.File]::ReadAllBytes($file)
        $bytes.Length | Should -Be $profile.bytes
        Get-ColourBigEndian $bytes 0 4 | Should -Be $bytes.Length
        [Text.Encoding]::ASCII.GetString($bytes, 36, 4) | Should -Be 'acsp'
        [Text.Encoding]::ASCII.GetString($bytes, 16, 4).Trim() | Should -Be $profile.header.colour_space
    }
    $srgbProfile = Join-Path $assets 'sRGB-v4.icc'
    $wideProfile = Join-Path $assets 'AdobeCompat-v2.icc'
    $cmykProfile = Join-Path $assets 'CGATS001Compat-v2-micro.icc'
    if ([string]::IsNullOrWhiteSpace($env:WINIMG_TEST_MAGICK) -or -not [IO.Path]::IsPathRooted($env:WINIMG_TEST_MAGICK)) {
        throw 'WINIMG_TEST_MAGICK must select verified magick.exe explicitly.'
    }
    $magick = [IO.Path]::GetFullPath($env:WINIMG_TEST_MAGICK)
    Assert-ColourNoReparseAncestors $magick
    if (-not (Test-Path -LiteralPath $magick -PathType Leaf) -or [IO.Path]::GetFileName($magick) -ne 'magick.exe') {
        throw 'Selected colour test magick.exe is missing.'
    }
    . (Join-Path $repository 'WinImgNormalizer.ps1')

    function Invoke-ColourMagick {
        param([string[]]$Arguments)
        $observed = @(& $magick @Arguments 2>&1)
        $code = $LASTEXITCODE
        if ($code -ne 0 -or $observed.Count -gt 0 -and $Arguments[-1] -notin @('info:', 'info:-') -and $Arguments[0] -ne 'identify') {
            throw ('Colour fixture ImageMagick failed ({0}): {1}' -f $code, ($observed -join "`n"))
        }
        return ($observed -join "`n")
    }
    function New-ColourRgb {
        param([string]$Path, [object[]]$Values, [string]$Profile, [switch]$Alpha, [string]$Coder = 'PNG', [int]$Width = 128, [int]$Height = 96)
        $arguments = @('-size', ('{0}x{1}' -f $Width, $Height), 'xc:none')
        for ($index = 0; $index -lt 4; $index++) {
            $value = $Values[$index]
            $color = 'rgb({0},{1},{2})' -f $value[0], $value[1], $value[2]
            if ($Alpha) {
                $fraction = ([double]$value[3] / 255).ToString('R', [Globalization.CultureInfo]::InvariantCulture)
                $color = 'rgba({0},{1},{2},{3})' -f $value[0], $value[1], $value[2], $fraction
            }
            $x = ($index % 2) * ($Width / 2); $y = [Math]::Floor($index / 2) * ($Height / 2)
            $arguments += @('-fill', $color, '-draw', ('rectangle {0},{1} {2},{3}' -f $x, $y, ($x + $Width / 2 - 1), ($y + $Height / 2 - 1)))
        }
        if (-not $Alpha) { $arguments += @('-alpha', 'off') }
        if ($Profile) { $arguments += @('-profile', (Get-WinImgNativeOutputPath $Profile)) }
        $arguments += @('-depth', '8', '-quality', '100', ($Coder + ':' + (Get-WinImgNativeOutputPath $Path)))
        $null = Invoke-ColourMagick $arguments
    }
    function New-ColourCmyk {
        param([object]$Case, [string]$Profile)
        $rawPath = Join-Path $Case.Probe 'source-cmyk.raw'
        $bytes = New-Object byte[] (128 * 96 * 4)
        for ($y = 0; $y -lt 96; $y++) {
            for ($x = 0; $x -lt 128; $x++) {
                $index = [int]($x -ge 64) + 2 * [int]($y -ge 48)
                for ($channel = 0; $channel -lt 4; $channel++) { $bytes[4 * ($y * 128 + $x) + $channel] = $reference.cmyk_source[$index][$channel] }
            }
        }
        [IO.File]::WriteAllBytes($rawPath, $bytes)
        $arguments = @('-size', '128x96', '-depth', '8', ('CMYK:' + (Get-WinImgNativeOutputPath $rawPath)))
        if ($Profile) { $arguments += @('-profile', (Get-WinImgNativeOutputPath $Profile)) }
        $arguments += @('-compress', 'None', ('TIFF:' + (Get-WinImgNativeOutputPath (Join-Path $Case.Source $Case.Name))))
        $null = Invoke-ColourMagick $arguments
    }
    function Add-ColourJpegSegment {
        param([string]$Path, [byte]$Marker, [byte[]]$Payload)
        $bytes = [IO.File]::ReadAllBytes($Path)
        if ($bytes[0] -ne 255 -or $bytes[1] -ne 216 -or $Payload.Length -gt 65533) { throw 'Invalid synthetic JPEG segment request.' }
        $length = $Payload.Length + 2
        $segment = [byte[]]@([byte]255, $Marker, [byte]($length -shr 8), [byte]($length -band 255)) + $Payload
        [IO.File]::WriteAllBytes($Path, [byte[]](@($bytes[0..1]) + @($segment) + @($bytes[2..($bytes.Length - 1)])))
    }
    function Get-ColourJpegSegments {
        param([string]$Path)
        $bytes = [IO.File]::ReadAllBytes($Path)
        if ($bytes[0] -ne 255 -or $bytes[1] -ne 216) { throw 'Expected JPEG for independent segment inspection.' }
        $offset = 2
        while ($offset + 3 -lt $bytes.Length) {
            if ($bytes[$offset] -ne 255) { throw 'Invalid JPEG marker boundary.' }
            $marker = $bytes[$offset + 1]
            if ($marker -in @(218,217)) { break }
            $length = Get-ColourBigEndian $bytes ($offset + 2) 2
            if ($length -lt 2 -or $offset + 2 + $length -gt $bytes.Length) { throw 'Invalid JPEG segment length.' }
            [pscustomobject]@{ Marker = $marker; Payload = [byte[]]$bytes[($offset + 4)..($offset + 1 + $length)] }
            $offset += 2 + $length
        }
    }
    function Assert-ColourNoMetadata {
        param([string]$Path)
        $segments = @(Get-ColourJpegSegments $Path)
        @($segments | Where-Object { $_.Marker -in @(225,226,237,254) }).Count | Should -Be 0 -Because 'EXIF/GPS/XMP, ICC, Photoshop/IPTC and comments must be stripped'
    }
    function Set-ColourTiffWrongProfile {
        param([string]$Path, [byte[]]$ProfileBytes)
        $bytes = [IO.File]::ReadAllBytes($Path)
        if ($bytes[0] -ne 73 -or $bytes[1] -ne 73 -or [BitConverter]::ToUInt16($bytes, 2) -ne 42) { throw 'Expected classic little-endian TIFF.' }
        $ifd = [BitConverter]::ToUInt32($bytes, 4)
        $count = [BitConverter]::ToUInt16($bytes, $ifd)
        $found = 0
        for ($index = 0; $index -lt $count; $index++) {
            $offset = [int]$ifd + 2 + 12 * $index
            if ([BitConverter]::ToUInt16($bytes, $offset) -eq 34675) {
                $oldLength = [BitConverter]::ToUInt32($bytes, $offset + 4)
                $profileOffset = [BitConverter]::ToUInt32($bytes, $offset + 8)
                if ($ProfileBytes.Length -gt $oldLength) { throw 'Wrong-model fixture replacement exceeds its existing profile extent.' }
                [BitConverter]::GetBytes([uint32]$ProfileBytes.Length).CopyTo($bytes, $offset + 4)
                $ProfileBytes.CopyTo($bytes, $profileOffset)
                $found++
            }
        }
        $found | Should -Be 1
        [IO.File]::WriteAllBytes($Path, $bytes)
    }
    function Add-ColourPrivacyMetadata {
        param([string]$Path, [int]$Orientation)
        $exif = [Convert]::FromBase64String($reference.orientation.exif_template_base64)
        [Text.Encoding]::ASCII.GetString($exif, 6, 2) | Should -Be 'MM'
        $ifd = 6 + (Get-ColourBigEndian $exif 10 4)
        $count = Get-ColourBigEndian $exif $ifd 2
        $found = 0
        for ($index = 0; $index -lt $count; $index++) {
            $offset = $ifd + 2 + 12 * $index
            if ((Get-ColourBigEndian $exif $offset 2) -eq 274) {
                $exif[$offset + 8] = 0; $exif[$offset + 9] = [byte]$Orientation; $found++
            }
        }
        $found | Should -Be 1
        Add-ColourJpegSegment -Path $Path -Marker 225 -Payload $exif
        $xmp = [Text.Encoding]::UTF8.GetBytes("http://ns.adobe.com/xap/1.0/`0<x:xmpmeta xmlns:x='adobe:ns:meta/'><rdf:RDF xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#'><rdf:Description xmlns:exif='http://ns.adobe.com/exif/1.0/' exif:GPSLatitude='12,34.56N' exif:GPSLongitude='34,12.06E'/></rdf:RDF></x:xmpmeta>")
        Add-ColourJpegSegment -Path $Path -Marker 225 -Payload $xmp
        $title = [Text.Encoding]::ASCII.GetBytes('Owned synthetic IPTC title')
        $iptc = [byte[]]@(28,2,5,0,[byte]$title.Length) + $title
        $size = [byte[]]@(0,0,[byte]($iptc.Length -shr 8),[byte]($iptc.Length -band 255))
        $photoshop = [Text.Encoding]::ASCII.GetBytes("Photoshop 3.0`0" + '8BIM') + [byte[]]@(4,4,0,0) + $size + $iptc
        if ($iptc.Length % 2 -ne 0) { $photoshop += [byte]0 }
        Add-ColourJpegSegment -Path $Path -Marker 237 -Payload $photoshop
        Add-ColourJpegSegment -Path $Path -Marker 254 -Payload ([Text.Encoding]::ASCII.GetBytes('Owned synthetic private comment'))
    }
    function Assert-ColourPixels {
        param([string]$Path, [int]$Width, [int]$Height, [object[]]$Expected, [int]$Tolerance, [string]$Format)
        $native = Get-WinImgNativeOutputPath $Path
        Invoke-ColourMagick @('identify', '+ping', '-regard-warnings', '-format', '%m|%w|%h|%n', $native) |
            Should -Be ('{0}|{1}|{2}|1' -f $Format, $Width, $Height)
        $points = @(@(($Width / 4),($Height / 4)),@((3 * $Width / 4),($Height / 4)),@(($Width / 4),(3 * $Height / 4)),@((3 * $Width / 4),(3 * $Height / 4)))
        for ($index = 0; $index -lt 4; $index++) {
            $formatString = '%[fx:round(255*p{{{0},{1}}}.r)]|%[fx:round(255*p{{{0},{1}}}.g)]|%[fx:round(255*p{{{0},{1}}}.b)]' -f $points[$index][0], $points[$index][1]
            $actual = (Invoke-ColourMagick @($native, '-format', $formatString, 'info:')) -split '\|'
            for ($channel = 0; $channel -lt 3; $channel++) {
                [Math]::Abs([int]$actual[$channel] - $Expected[$index][$channel]) | Should -BeLessOrEqual $Tolerance -Because ('patch {0} channel {1}, {2} reference' -f $index, $channel, $Format)
            }
        }
    }
    function Invoke-ColourRun {
        param([object]$Case, [scriptblock]$Runner, [long]$MaxBytes = 1048576)
        $parameters = @{ Source = $Case.Source; OutputParent = $Case.Parent; MagickPath = $magick; MaxBytes = $MaxBytes }
        if ($Runner) { $parameters.ProcessRunner = $Runner }
        $observed = @(& Invoke-WinImgNormalizer @parameters 6>&1 3>&1 2>&1)
        $codes = @($observed | Where-Object { $_ -is [int] -or $_ -is [long] })
        $codes.Count | Should -Be 1
        $runs = @(Get-ChildItem -LiteralPath $Case.Parent -Force -Directory)
        $runs.Count | Should -Be 1
        $logs = @(Get-ChildItem -LiteralPath (Join-Path $runs[0].FullName '.WinImgNormalizer\reports') -File -Filter '*.log')
        $logs.Count | Should -Be 1
        return [pscustomobject]@{ Code = $codes[0]; Run = $runs[0].FullName; Log = Get-Content -LiteralPath $logs[0].FullName -Raw; Text = $observed -join "`n" }
    }
    function Assert-ColourSuccess {
        param([object]$Result, [object]$Case, [object[]]$Expected, [string]$Policy, [bool]$HasIcc, [string]$Space = 'sRGB')
        $Result.Code | Should -Be 0 -Because $Result.Text
        $Result.Log | Should -Match ([regex]::Escape(('SourceSpace={0}; SourceICC={1}; Policy={2};' -f $Space, $HasIcc, $Policy)))
        $Result.Log | Should -Match 'Alpha=WhiteAfterSrgb; OutputICC=None'
        $Result.Log | Should -Match 'SUMMARY ConvertedImages=1 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=0'
        $output = Join-Path $Result.Run ([IO.Path]::GetFileNameWithoutExtension($Case.Name) + '.jpeg')
        Assert-ColourPixels -Path $output -Width 128 -Height 96 -Expected $Expected -Tolerance $reference.tolerance.jpeg_channel_units -Format JPEG
        Assert-ColourNoMetadata $output
        @(Get-ChildItem -LiteralPath $Result.Run -Recurse -Force -File | Where-Object Extension -ne '.log').Count | Should -Be 1
        @(Get-ChildItem -LiteralPath (Join-Path $Result.Run '.WinImgNormalizer\work') -Force).Count | Should -Be 0
    }
    function Assert-ColourFailed {
        param([object]$Result)
        $Result.Code | Should -Be 2
        $Result.Log | Should -Not -Match 'OK IMG:'
        $Result.Log | Should -Match 'SUMMARY ConvertedImages=0 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=1'
        @(Get-ChildItem -LiteralPath $Result.Run -Recurse -Force -File | Where-Object Extension -ne '.log').Count | Should -Be 0
        @(Get-ChildItem -LiteralPath (Join-Path $Result.Run '.WinImgNormalizer\work') -Force).Count | Should -Be 0
    }
}

Describe 'M2-T02 independent profiled and untagged colour references (T032-T033)' {
    It 'T032 transforms <Kind> tagged RGB before stripping and verifies actual runtime arguments before JPEG loss' -ForEach @(
        @{ Kind = 'wide'; ProfileName = 'AdobeCompat-v2.icc' },
        @{ Kind = 'sRGB'; ProfileName = 'sRGB-v4.icc' },
        @{ Kind = 'wide with spoofed free profile property'; ProfileName = 'AdobeCompat-v2.icc' }
    ) {
        $case = New-ColourCase 'rgb' 'patches.png'
        $sourcePath = Join-Path $case.Source $case.Name
        $profilePath = Join-Path $assets $ProfileName
        if ($Kind -like '*spoofed*') {
            $initial = Join-Path $case.Probe 'tagged.png'
            New-ColourRgb -Path $initial -Values $reference.rgb_source -Profile $profilePath
            $null = Invoke-ColourMagick @((Get-WinImgNativeOutputPath $initial), '-set', 'profiles', 'none', ('MIFF:' + (Get-WinImgNativeOutputPath $sourcePath)))
        } else { New-ColourRgb -Path $sourcePath -Values $reference.rgb_source -Profile $profilePath }
        Set-ColourSourceTimes $sourcePath
        $before = Get-ColourSourceState $case.Source
        $expected = if ($Kind -eq 'sRGB') { $reference.rgb_source } else { $reference.wide_rgb }
        $colourTrace = [pscustomobject]@{ Arguments = New-Object 'Collections.Generic.List[object]'; Probe = Join-Path $case.Probe 'runtime-before-jpeg.png' }
        $runner = {
            param([string]$Executable, [string[]]$Arguments)
            $colourTrace.Arguments.Add(@($Arguments))
            $pngArgs = [string[]]$Arguments.Clone()
            $pngArgs[-1] = 'PNG:' + (Get-WinImgNativeOutputPath $colourTrace.Probe)
            $probeText = @(& $Executable @pngArgs 2>&1); $probeCode = $LASTEXITCODE
            if ($probeCode -ne 0 -or $probeText.Count -ne 0) { throw ('Pre-JPEG runtime argument probe failed: ' + ($probeText -join "`n")) }
            $nativeText = @(& $Executable @Arguments 2>&1)
            return [pscustomobject]@{ ExitCode = $LASTEXITCODE; DiagnosticOutput = $nativeText -join "`n" }
        }
        $result = Invoke-ColourRun -Case $case -Runner $runner
        $colourTrace.Arguments.Count | Should -Be 1
        Assert-ColourPixels -Path $colourTrace.Probe -Width 128 -Height 96 -Expected $expected -Tolerance $reference.tolerance.pre_jpeg_channel_units -Format PNG
        Assert-ColourSuccess -Result $result -Case $case -Expected $expected -Policy ProfileToSrgb -HasIcc $true
        Get-ColourSourceState $case.Source | Should -Be $before
    }

    It 'T033 transforms genuine four-channel tagged CMYK through the available A2B0 mapping and verifies its independent forward reference' {
        $case = New-ColourCase 'cmyk' 'ink.tiff'
        New-ColourCmyk -Case $case -Profile $cmykProfile
        $sourcePath = Join-Path $case.Source $case.Name
        Invoke-ColourMagick @('identify', '+ping', '-regard-warnings', '-format', '%[colorspace]|%[profiles]', (Get-WinImgNativeOutputPath $sourcePath)) | Should -Be 'CMYK|icc'
        Set-ColourSourceTimes $sourcePath
        $before = Get-ColourSourceState $case.Source
        $colourTrace = [pscustomobject]@{ Probe = Join-Path $case.Probe 'runtime-before-jpeg.png'; Calls = 0 }
        $runner = {
            param([string]$Executable, [string[]]$Arguments)
            $colourTrace.Calls++
            $pngArgs = [string[]]$Arguments.Clone(); $pngArgs[-1] = 'PNG:' + (Get-WinImgNativeOutputPath $colourTrace.Probe)
            $probeText = @(& $Executable @pngArgs 2>&1); $probeCode = $LASTEXITCODE
            if ($probeCode -ne 0 -or $probeText.Count -ne 0) { throw ('Pre-JPEG CMYK runtime argument probe failed: ' + ($probeText -join "`n")) }
            $nativeText = @(& $Executable @Arguments 2>&1)
            return [pscustomobject]@{ ExitCode = $LASTEXITCODE; DiagnosticOutput = $nativeText -join "`n" }
        }
        $result = Invoke-ColourRun -Case $case -Runner $runner
        $colourTrace.Calls | Should -Be 1
        Assert-ColourPixels -Path $colourTrace.Probe -Width 128 -Height 96 -Expected $reference.tagged_cmyk -Tolerance $reference.tolerance.pre_jpeg_channel_units -Format PNG
        Assert-ColourSuccess -Result $result -Case $case -Expected $reference.tagged_cmyk -Policy ProfileToSrgb -HasIcc $true -Space CMYK
        Get-ColourSourceState $case.Source | Should -Be $before
    }

    It 'T033 explicitly assumes untagged RGB is sRGB and preserves its controlled encoded patch values' {
        $case = New-ColourCase 'untagged-rgb' 'plain.png'
        $sourcePath = Join-Path $case.Source $case.Name
        New-ColourRgb -Path $sourcePath -Values $reference.rgb_source
        Set-ColourSourceTimes $sourcePath
        $before = Get-ColourSourceState $case.Source
        $result = Invoke-ColourRun $case
        Assert-ColourSuccess -Result $result -Case $case -Expected $reference.rgb_source -Policy AssumeSrgb -HasIcc $false
        Get-ColourSourceState $case.Source | Should -Be $before
    }

    It 'T033 converts a genuinely declared linear RGB raster to encoded sRGB instead of treating its samples as encoded bytes' {
        $case = New-ColourCase 'linear-rgb' 'linear.png'
        $sourcePath = Join-Path $case.Source $case.Name
        $null = Invoke-ColourMagick @('-size','128x96','xc:rgb(128,128,128)','-set','colorspace','RGB',('MIFF:' + (Get-WinImgNativeOutputPath $sourcePath)))
        Invoke-ColourMagick @('identify','+ping','-format','%m|%[colorspace]',(Get-WinImgNativeOutputPath $sourcePath)) | Should -Be 'MIFF|RGB'
        Set-ColourSourceTimes $sourcePath
        $before = Get-ColourSourceState $case.Source
        $result = Invoke-ColourRun $case
        # IEC sRGB transfer at128/255 linear =0.055 adjustment plus1/2.4
        # exponent, rounding to188; independent of the native colour pipeline.
        $expected = @(@(188,188,188),@(188,188,188),@(188,188,188),@(188,188,188))
        Assert-ColourSuccess -Result $result -Case $case -Expected $expected -Policy ConvertLinearRgb -HasIcc $false -Space RGB
        Get-ColourSourceState $case.Source | Should -Be $before
    }

    It 'T033 rejects untagged CMYK before conversion and makes no characterized-colour claim' {
        $case = New-ColourCase 'untagged-cmyk' 'ambiguous.tiff'
        New-ColourCmyk $case
        $sourcePath = Join-Path $case.Source $case.Name
        Set-ColourSourceTimes $sourcePath
        $before = Get-ColourSourceState $case.Source
        $colourTrace = [pscustomobject]@{ Calls = 0 }
        $runner = { param($Executable,$Arguments); $colourTrace.Calls++; return 0 }
        $result = Invoke-ColourRun -Case $case -Runner $runner
        $colourTrace.Calls | Should -Be 0
        Assert-ColourFailed $result
        $result.Log | Should -Match '(?i)CMYK.*(ICC|profile|characteriz)'
        Get-ColourSourceState $case.Source | Should -Be $before
    }
}

Describe 'M2-T02 malformed profiles and native diagnostic handling (T034)' {
    It 'T034 rejects actual <Kind> ICC problems without finalizing an inaccurate JPEG' -ForEach @(
        @{ Kind = 'malformed PNG'; Name = 'bad.png' },
        @{ Kind = 'malformed JPEG'; Name = 'bad.jpg' },
        @{ Kind = 'RGB pixels with CMYK profile'; Name = 'bad.jpg' },
        @{ Kind = 'CMYK pixels with RGB profile'; Name = 'bad.tiff' }
    ) {
        $case = New-ColourCase 'bad-icc' $Name
        $sourcePath = Join-Path $case.Source $case.Name
        if ($Kind -eq 'malformed PNG') {
            [IO.File]::WriteAllBytes($sourcePath, [Convert]::FromBase64String($reference.malformed_png.base64))
        } elseif ($Kind -eq 'CMYK pixels with RGB profile') {
            New-ColourCmyk -Case $case -Profile $cmykProfile
            Set-ColourTiffWrongProfile -Path $sourcePath -ProfileBytes ([IO.File]::ReadAllBytes($wideProfile))
        } else {
            New-ColourRgb -Path $sourcePath -Values $reference.rgb_source -Coder JPEG
            $bytes = if ($Kind -eq 'malformed JPEG') { [Text.Encoding]::ASCII.GetBytes('Owned benign malformed ICC data') } else { [IO.File]::ReadAllBytes($cmykProfile) }
            $payload = [Text.Encoding]::ASCII.GetBytes("ICC_PROFILE`0") + [byte[]]@(1,1) + $bytes
            Add-ColourJpegSegment -Path $sourcePath -Marker 226 -Payload $payload
            @(@(Get-ColourJpegSegments $sourcePath) | Where-Object Marker -eq 226).Count | Should -Be 1
        }
        Set-ColourSourceTimes $sourcePath
        $before = Get-ColourSourceState $case.Source
        $result = Invoke-ColourRun $case
        Assert-ColourFailed $result
        $result.Log | Should -Match '(?i)(colour|color|ICC|profile)'
        Get-ColourSourceState $case.Source | Should -Be $before
    }

    It 'T034 refuses a structured native zero exit with profile diagnostics even when the emitted JPEG fully decodes' {
        $case = New-ColourCase 'diagnostic' 'source.png'
        $sourcePath = Join-Path $case.Source $case.Name
        New-ColourRgb -Path $sourcePath -Values $reference.rgb_source
        $validJpeg = Join-Path $case.Probe 'valid.jpg'
        $null = Invoke-ColourMagick @((Get-WinImgNativeOutputPath $sourcePath),('JPEG:' + (Get-WinImgNativeOutputPath $validJpeg)))
        $colourTrace = [pscustomobject]@{ Bytes = [IO.File]::ReadAllBytes($validJpeg); Calls = 0 }
        Set-ColourSourceTimes $sourcePath
        $before = Get-ColourSourceState $case.Source
        $runner = {
            param($Executable,$Arguments)
            $colourTrace.Calls++
            [IO.File]::WriteAllBytes($Arguments[-1].Substring(5), $colourTrace.Bytes)
            return [pscustomobject]@{ ExitCode = 0; DiagnosticOutput = 'ColorspaceColorProfileMismatch: controlled diagnostic despite zero native exit.' }
        }
        $result = Invoke-ColourRun -Case $case -Runner $runner
        $colourTrace.Calls | Should -Be 1
        Assert-ColourFailed $result
        Get-ColourSourceState $case.Source | Should -Be $before
    }

    It 'T034 rejects an actual retained RGB ICC with valid structure but an invalid curve through the default native runner' {
        $case = New-ColourCase 'invalid-curve' 'curve.jpg'
        $sourcePath = Join-Path $case.Source $case.Name
        New-ColourRgb -Path $sourcePath -Values $reference.rgb_source -Coder JPEG
        # Keep the genuine sRGB header/model/tag bounds, but replace the shared
        # rTRC/gTRC/bTRC payload type. Header checks cannot prove LCMS semantics.
        $profileBytes = [IO.File]::ReadAllBytes($srgbProfile)
        $tagCount = Get-ColourBigEndian $profileBytes 128 4
        $curveOffset = -1
        for ($index = 0; $index -lt $tagCount; $index++) {
            $entry = 132 + 12 * $index
            if ([Text.Encoding]::ASCII.GetString($profileBytes, $entry, 4) -eq 'rTRC') {
                $curveOffset = Get-ColourBigEndian $profileBytes ($entry + 4) 4
            }
        }
        $curveOffset | Should -BeGreaterThan 0
        [Text.Encoding]::ASCII.GetString($profileBytes, $curveOffset, 4) | Should -Be 'para'
        [Array]::Copy([Text.Encoding]::ASCII.GetBytes('xxxx'), 0, $profileBytes, $curveOffset, 4)
        $payload = [Text.Encoding]::ASCII.GetBytes("ICC_PROFILE`0") + [byte[]]@(1,1) + $profileBytes
        Add-ColourJpegSegment -Path $sourcePath -Marker 226 -Payload $payload
        Invoke-ColourMagick @('identify','-ping','-regard-warnings','-format','%[colorspace]|%[profiles]',(Get-WinImgNativeOutputPath $sourcePath)) | Should -Be 'sRGB|icc'
        Set-ColourSourceTimes $sourcePath
        $before = Get-ColourSourceState $case.Source
        $result = Invoke-ColourRun $case
        # This log follows source ICC extraction and the structural/model guard;
        # the actual LCMS conversion must still reject the semantic error.
        $result.Log | Should -Match 'COLOUR IMG: curve.jpg \(SourceSpace=sRGB; SourceICC=True; Policy=ProfileToSrgb;'
        $result.Log | Should -Match 'Native conversion stopped: Category=DamagedInput'
        $result.Log | Should -Match 'NATIVE DETAILS:'
        Assert-ColourFailed $result
        Get-ColourSourceState $case.Source | Should -Be $before
    }

    It 'T034 captures real native diagnostics when a structurally accepted ICC returns zero and writes a decodable JPEG' {
        $case = New-ColourCase 'unsupported-icc-version' 'version.jpg'
        $sourcePath = Join-Path $case.Source $case.Name
        New-ColourRgb -Path $sourcePath -Values $reference.rgb_source -Coder JPEG
        # Version 99 is semantically unsupported while size, acsp, RGB model,
        # the genuine tag table and every payload bound remain unchanged.
        $profileBytes = [IO.File]::ReadAllBytes($srgbProfile)
        [Array]::Copy([byte[]]@(153,0,0,0), 0, $profileBytes, 8, 4)
        $payload = [Text.Encoding]::ASCII.GetBytes("ICC_PROFILE`0") + [byte[]]@(1,1) + $profileBytes
        Add-ColourJpegSegment -Path $sourcePath -Marker 226 -Payload $payload
        Invoke-ColourMagick @('identify','-ping','-regard-warnings','-format','%[colorspace]|%[profiles]',(Get-WinImgNativeOutputPath $sourcePath)) | Should -Be 'sRGB|icc'
        Set-ColourSourceTimes $sourcePath
        $before = Get-ColourSourceState $case.Source
        $probePath = Join-Path $case.Probe 'native-zero.jpg'
        $nativeArguments = @('-quiet','-regard-warnings',(Get-WinImgNativeOutputPath $sourcePath),
            '-auto-orient','+black-point-compensation','-intent','Relative','-profile',(Get-WinImgNativeOutputPath $srgbProfile),
            '-background','white','-alpha','remove','-alpha','off','-strip','-sampling-factor','4:2:0','-interlace','Line',
            '-resize','100%','-define','jpeg:extent=1048576B',('JPEG:' + (Get-WinImgNativeOutputPath $probePath)))
        $previousPreference = $ErrorActionPreference
        try {
            $ErrorActionPreference = 'Continue'
            $nativeDiagnostics = @(& $magick @nativeArguments 2>&1)
            $nativeExit = $LASTEXITCODE
        } finally { $ErrorActionPreference = $previousPreference }
        $nativeExit | Should -Be 0
        ($nativeDiagnostics -join "`n") | Should -Match 'ColorspaceColorProfileMismatch'
        Invoke-ColourMagick @('identify','+ping','-regard-warnings','-format','%m|%w|%h|%n',(Get-WinImgNativeOutputPath $probePath)) | Should -Be 'JPEG|128|96|1'
        [IO.File]::WriteAllText((Join-Path $case.Probe 'native-zero-observation.json'), ([pscustomobject]@{
            ExitCode = $nativeExit; DiagnosticOutput = $nativeDiagnostics -join "`n"
            FullyDecoded = $true; JpegBytes = [IO.File]::ReadAllBytes($probePath).Length
        } | ConvertTo-Json))
        # No injected ProcessRunner: the default implementation must capture the
        # same native diagnostics and stop immediately despite valid JPEG bytes.
        $result = Invoke-ColourRun $case
        $result.Log | Should -Match 'COLOUR IMG: version.jpg \(SourceSpace=sRGB; SourceICC=True; Policy=ProfileToSrgb;'
        $result.Log | Should -Match 'Native conversion stopped: Category=DamagedInput'
        $result.Log | Should -Match 'NATIVE DETAILS:'
        Assert-ColourFailed $result
        Get-ColourSourceState $case.Source | Should -Be $before
    }
}

Describe 'M2-T02 privacy stripping after actual displayed orientation (T035)' {
    It 'T035 honors genuine EXIF orientation <Orientation> and removes synthetic GPS/EXIF/XMP/IPTC/comments/ICC after processing' -ForEach @(
        @{ Orientation = 1 }, @{ Orientation = 2 }, @{ Orientation = 3 }, @{ Orientation = 4 },
        @{ Orientation = 5 }, @{ Orientation = 6 }, @{ Orientation = 7 }, @{ Orientation = 8 }
    ) {
        $case = New-ColourCase 'metadata' 'oriented.jpg'
        $sourcePath = Join-Path $case.Source $case.Name
        New-ColourRgb -Path $sourcePath -Values $reference.orientation.nominal_corners -Profile $srgbProfile -Coder JPEG -Width 96 -Height 64
        Add-ColourPrivacyMetadata -Path $sourcePath -Orientation $Orientation
        Invoke-ColourMagick @('identify','+ping','-regard-warnings','-format','%[EXIF:Orientation]|%[EXIF:GPSLatitude]|%[profiles]',(Get-WinImgNativeOutputPath $sourcePath)) |
            Should -Match ('^' + $Orientation + '\|12/1,34/1,56/1\|.*exif.*icc.*iptc.*xmp')
        @(@(Get-ColourJpegSegments $sourcePath) | Where-Object Marker -eq 225).Count | Should -Be 2
        @(@(Get-ColourJpegSegments $sourcePath) | Where-Object Marker -eq 237).Count | Should -Be 1
        Set-ColourSourceTimes $sourcePath
        $before = Get-ColourSourceState $case.Source
        $result = Invoke-ColourRun $case
        $result.Code | Should -Be 0 -Because $result.Text
        $width = if ($Orientation -ge 5) { 64 } else { 96 }
        $height = if ($Orientation -ge 5) { 96 } else { 64 }
        $expected = @($reference.orientation.expected_order_by_orientation[$Orientation - 1] | ForEach-Object { ,$reference.orientation.nominal_corners[$_] })
        $output = Join-Path $result.Run 'oriented.jpeg'
        Assert-ColourPixels -Path $output -Width $width -Height $height -Expected $expected -Tolerance $reference.tolerance.jpeg_channel_units -Format JPEG
        Assert-ColourNoMetadata $output
        @(Get-ChildItem -LiteralPath (Join-Path $result.Run '.WinImgNormalizer\work') -Force).Count | Should -Be 0
        @(Get-ChildItem -LiteralPath $result.Run -Recurse -Force -File | Where-Object Extension -ne '.log').Count | Should -Be 1
        Get-ColourSourceState $case.Source | Should -Be $before
    }
}

Describe 'M2-T02 alpha composition and retry invariants (T036)' {
    It 'T036 composites <Kind> transparent patches on white after conversion and checks separate pre-JPEG and JPEG tolerances' -ForEach @(
        @{ Kind = 'tagged wide RGB'; Tagged = $true },
        @{ Kind = 'untagged sRGB'; Tagged = $false }
    ) {
        $case = New-ColourCase 'alpha' 'transparent.png'
        $sourcePath = Join-Path $case.Source $case.Name
        $profile = if ($Tagged) { $wideProfile } else { $null }
        New-ColourRgb -Path $sourcePath -Values $reference.alpha_source -Profile $profile -Alpha
        Set-ColourSourceTimes $sourcePath
        $before = Get-ColourSourceState $case.Source
        $expected = if ($Tagged) { $reference.tagged_alpha_white } else { $reference.untagged_alpha_white }
        $colourTrace = [pscustomobject]@{ Probe = Join-Path $case.Probe 'runtime-before-jpeg.png' }
        $runner = {
            param([string]$Executable,[string[]]$Arguments)
            $pngArgs = [string[]]$Arguments.Clone(); $pngArgs[-1] = 'PNG:' + (Get-WinImgNativeOutputPath $colourTrace.Probe)
            $probeText = @(& $Executable @pngArgs 2>&1); $probeCode = $LASTEXITCODE
            if ($probeCode -ne 0 -or $probeText.Count -ne 0) { throw ('Pre-JPEG alpha runtime argument probe failed: ' + ($probeText -join "`n")) }
            $nativeText = @(& $Executable @Arguments 2>&1)
            return [pscustomobject]@{ ExitCode = $LASTEXITCODE; DiagnosticOutput = $nativeText -join "`n" }
        }
        $result = Invoke-ColourRun -Case $case -Runner $runner
        Assert-ColourPixels -Path $colourTrace.Probe -Width 128 -Height 96 -Expected $expected -Tolerance $reference.tolerance.pre_jpeg_channel_units -Format PNG
        $policy = if ($Tagged) { 'ProfileToSrgb' } else { 'AssumeSrgb' }
        Assert-ColourSuccess -Result $result -Case $case -Expected $expected -Policy $policy -HasIcc $Tagged
        Get-ColourSourceState $case.Source | Should -Be $before
    }

    It 'T036 keeps profile conversion, white alpha composition and stripping on every real native scale attempt without an alpha-dropping fallback' {
        $case = New-ColourCase 'alpha-retry' 'transparent.png'
        $sourcePath = Join-Path $case.Source $case.Name
        New-ColourRgb -Path $sourcePath -Values $reference.alpha_source -Profile $wideProfile -Alpha
        Set-ColourSourceTimes $sourcePath
        $before = Get-ColourSourceState $case.Source
        $colourTrace = [pscustomobject]@{
            Arguments = New-Object 'Collections.Generic.List[object]'
            TargetPaths = New-Object 'Collections.Generic.List[string]'
            TargetHashes = New-Object 'Collections.Generic.List[string]'
        }
        $runner = {
            param([string]$Executable,[string[]]$Arguments)
            $colourTrace.Arguments.Add(@($Arguments))
            $profileIndex = [Array]::IndexOf($Arguments, '-profile')
            $targetPath = $Arguments[$profileIndex + 1]
            $colourTrace.TargetPaths.Add($targetPath)
            # The argument is already a native extended path; direct .NET reads
            # keep this byte binding usable on long PS5.1 hosted roots.
            $sha256 = [Security.Cryptography.SHA256]::Create()
            try {
                $hash = $sha256.ComputeHash([IO.File]::ReadAllBytes($targetPath))
                $colourTrace.TargetHashes.Add(([BitConverter]::ToString($hash).Replace('-', '').ToLowerInvariant()))
            } finally { $sha256.Dispose() }
            # Only relax this test's native extent while the caller's one-byte
            # limit forces all six valid outputs through the normal retry loop.
            $nativeArgs = @($Arguments | ForEach-Object { if ($_ -like 'jpeg:extent=*') { 'jpeg:extent=1048576B' } else { $_ } })
            $nativeText = @(& $Executable @nativeArgs 2>&1)
            return [pscustomobject]@{ ExitCode = $LASTEXITCODE; DiagnosticOutput = $nativeText -join "`n" }
        }
        $result = Invoke-ColourRun -Case $case -Runner $runner -MaxBytes 1
        $result.Code | Should -Be 2 -Because $result.Text
        $colourTrace.Arguments.Count | Should -Be 6
        @($colourTrace.TargetPaths | Select-Object -Unique).Count | Should -Be 1
        $colourTrace.TargetHashes.Count | Should -Be 6
        $expectedProfileHash = @($manifest.profiles | Where-Object source_path -eq 'profiles/sRGB-v4.icc')[0].sha256
        foreach ($targetHash in $colourTrace.TargetHashes) { $targetHash | Should -Be $expectedProfileHash }
        $scales = @()
        foreach ($arguments in $colourTrace.Arguments) {
            $orient = [Array]::IndexOf($arguments,'-auto-orient')
            $intent = [Array]::IndexOf($arguments,'-intent')
            $profile = [Array]::IndexOf($arguments,'-profile')
            $background = [Array]::IndexOf($arguments,'-background')
            $strip = [Array]::IndexOf($arguments,'-strip')
            $arguments | Should -Contain '-regard-warnings'
            $arguments | Should -Contain '+black-point-compensation'
            $intent | Should -BeGreaterThan $orient
            $arguments[$intent + 1] | Should -Be 'Relative'
            $profile | Should -BeGreaterThan $intent
            $background | Should -BeGreaterThan $profile
            $arguments[$background + 1] | Should -Be 'white'
            (@($arguments[($background + 2)..($strip - 1)]) -join '|') | Should -Be '-alpha|remove|-alpha|off'
            $strip | Should -BeGreaterThan $background
            $resize = [Array]::IndexOf($arguments,'-resize')
            $scales += $arguments[$resize + 1]
        }
        ($scales -join ',') | Should -Be '100%,90%,80%,70%,60%,50%'
        $output = Join-Path $result.Run 'transparent.jpeg'
        Assert-ColourPixels -Path $output -Width 64 -Height 48 -Expected $reference.tagged_alpha_white -Tolerance $reference.tolerance.jpeg_channel_units -Format JPEG
        Assert-ColourNoMetadata $output
        $result.Log | Should -Match '(?i)(above|exceed|best|over)'
        @(Get-ChildItem -LiteralPath (Join-Path $result.Run '.WinImgNormalizer\work') -Force).Count | Should -Be 0
        Get-ColourSourceState $case.Source | Should -Be $before
    }
}
