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
    . (Join-Path $PSScriptRoot 'IccAcceptanceHelpers.ps1')
    $iccRepairAssets = Join-Path $PSScriptRoot 'fixtures/icc-repair'

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
    function Assert-ColourCompositionOrder {
        param([string[]]$Arguments, [bool]$Tagged)
        $orient = [Array]::IndexOf($Arguments, '-auto-orient')
        $conversion = [Array]::IndexOf($Arguments, $(if ($Tagged) { '-profile' } else { '-colorspace' }))
        $clamp = [Array]::IndexOf($Arguments, '-clamp')
        $background = [Array]::IndexOf($Arguments, '-background')
        $strip = [Array]::IndexOf($Arguments, '-strip')
        $Arguments | Should -Contain '-regard-warnings'
        $orient | Should -BeGreaterThan -1
        $conversion | Should -BeGreaterThan $orient
        if ($Tagged) {
            $intent = [Array]::IndexOf($Arguments, '-intent')
            $Arguments | Should -Contain '+black-point-compensation'
            $intent | Should -BeGreaterThan $orient
            $Arguments[$intent + 1] | Should -Be 'Relative'
            $conversion | Should -BeGreaterThan $intent
        } else {
            $Arguments[$conversion + 1] | Should -Be 'sRGB'
            $Arguments | Should -Not -Contain '-profile'
        }
        @($Arguments | Where-Object { $_ -eq '-clamp' }).Count | Should -Be 1
        $clamp | Should -Be ($conversion + 2) -Because 'HDRI channels must be bounded after sRGB conversion, before blending'
        $background | Should -Be ($clamp + 1)
        $Arguments[$background + 1] | Should -Be 'white'
        (@($Arguments[($background + 2)..($strip - 1)]) -join '|') | Should -Be '-alpha|remove|-alpha|off'
        $strip | Should -BeGreaterThan $background
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
        @(Get-ChildItem -LiteralPath $Result.Run -Recurse -Force -File | Where-Object { $_.Extension -notin @('.log','.csv') }).Count | Should -Be 1
        @(Get-ChildItem -LiteralPath (Join-Path $Result.Run '.WinImgNormalizer\work') -Force).Count | Should -Be 0
    }
    function Assert-ColourFailed {
        param([object]$Result)
        $Result.Code | Should -Be 2
        $Result.Log | Should -Not -Match 'OK IMG:'
        $Result.Log | Should -Match 'SUMMARY ConvertedImages=0 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=1'
        @(Get-ChildItem -LiteralPath $Result.Run -Recurse -Force -File | Where-Object { $_.Extension -notin @('.log','.csv') }).Count | Should -Be 0
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
    It 'T034 handles actual <Kind> ICC problems under the restricted JPEG preview policy' -ForEach @(
        @{ Kind = 'malformed PNG'; Name = 'bad.png'; Preview = $false },
        @{ Kind = 'malformed JPEG'; Name = 'bad.jpg'; Preview = $true },
        @{ Kind = 'RGB pixels with CMYK profile'; Name = 'bad.jpg'; Preview = $true },
        @{ Kind = 'CMYK pixels with RGB profile'; Name = 'bad.tiff'; Preview = $false }
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
        if ($Preview) {
            $result.Code | Should -Be 2
            $result.Log | Should -Match 'Policy=JpegIccAssumeSrgb'
            $result.Log | Should -Match ([regex]::Escape('JPEG saved; embedded colour profile unusable; sRGB assumed, colours may differ.'))
            $result.Log | Should -Match 'SUMMARY ConvertedImages=1 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=0'
            $output = Join-Path $result.Run 'bad.jpeg'
            Assert-ColourPixels -Path $output -Width 128 -Height 96 -Expected $reference.rgb_source -Tolerance $reference.tolerance.jpeg_channel_units -Format JPEG
            Assert-ColourNoMetadata $output
        } else { Assert-ColourFailed $result }
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

    It 'T034 saves an explicitly approximate JPEG when a structurally valid RGB profile has an invalid curve' {
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
        $result.Code | Should -Be 2
        $result.Log | Should -Match 'Policy=JpegIccAssumeSrgb'
        $result.Log | Should -Match ([regex]::Escape('JPEG saved; embedded colour profile unusable; sRGB assumed, colours may differ.'))
        $result.Log | Should -Match 'SUMMARY ConvertedImages=1 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=0'
        Assert-ColourPixels -Path (Join-Path $result.Run 'curve.jpeg') -Width 128 -Height 96 -Expected $reference.rgb_source -Tolerance $reference.tolerance.jpeg_channel_units -Format JPEG
        Assert-ColourNoMetadata (Join-Path $result.Run 'curve.jpeg')
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
        # No injected ProcessRunner: retain the real profile diagnostics locally,
        # then use one clean profile-excluded read for the identified ICC problem.
        $result = Invoke-ColourRun $case
        $result.Code | Should -Be 2
        $result.Log | Should -Match 'Policy=JpegIccAssumeSrgb'
        $result.Log | Should -Match ([regex]::Escape('JPEG saved; embedded colour profile unusable; sRGB assumed, colours may differ.'))
        $result.Log | Should -Match 'SUMMARY ConvertedImages=1 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=0'
        Assert-ColourPixels -Path (Join-Path $result.Run 'version.jpeg') -Width 128 -Height 96 -Expected $reference.rgb_source -Tolerance $reference.tolerance.jpeg_channel_units -Format JPEG
        Assert-ColourNoMetadata (Join-Path $result.Run 'version.jpeg')
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
        @(Get-ChildItem -LiteralPath $result.Run -Recurse -Force -File | Where-Object { $_.Extension -notin @('.log','.csv') }).Count | Should -Be 1
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
        $colourTrace = [pscustomobject]@{
            Probe = Join-Path $case.Probe 'runtime-before-jpeg.png'
            Arguments = New-Object 'Collections.Generic.List[object]'
        }
        $runner = {
            param([string]$Executable,[string[]]$Arguments)
            $colourTrace.Arguments.Add(@($Arguments))
            $pngArgs = [string[]]$Arguments.Clone(); $pngArgs[-1] = 'PNG:' + (Get-WinImgNativeOutputPath $colourTrace.Probe)
            $probeText = @(& $Executable @pngArgs 2>&1); $probeCode = $LASTEXITCODE
            if ($probeCode -ne 0 -or $probeText.Count -ne 0) { throw ('Pre-JPEG alpha runtime argument probe failed: ' + ($probeText -join "`n")) }
            $nativeText = @(& $Executable @Arguments 2>&1)
            return [pscustomobject]@{ ExitCode = $LASTEXITCODE; DiagnosticOutput = $nativeText -join "`n" }
        }
        $result = Invoke-ColourRun -Case $case -Runner $runner
        $colourTrace.Arguments.Count | Should -Be 1
        Assert-ColourCompositionOrder -Arguments $colourTrace.Arguments[0] -Tagged $Tagged
        Assert-ColourPixels -Path $colourTrace.Probe -Width 128 -Height 96 -Expected $expected -Tolerance $reference.tolerance.pre_jpeg_channel_units -Format PNG
        $policy = if ($Tagged) { 'ProfileToSrgb' } else { 'AssumeSrgb' }
        Assert-ColourSuccess -Result $result -Case $case -Expected $expected -Policy $policy -HasIcc $Tagged
        Get-ColourSourceState $case.Source | Should -Be $before
    }

    It 'T036 keeps <Kind> conversion, gamut clamp and white alpha policy through two sharing controls and all six real native scales' -ForEach @(
        @{ Kind = 'tagged wide RGB'; Tagged = $true },
        @{ Kind = 'untagged sRGB'; Tagged = $false }
    ) {
        $case = New-ColourCase 'alpha-retry' 'transparent.png'
        $sourcePath = Join-Path $case.Source $case.Name
        $profile = if ($Tagged) { $wideProfile } else { $null }
        $expected = if ($Tagged) { $reference.tagged_alpha_white } else { $reference.untagged_alpha_white }
        New-ColourRgb -Path $sourcePath -Values $reference.alpha_source -Profile $profile -Alpha
        Set-ColourSourceTimes $sourcePath
        $before = Get-ColourSourceState $case.Source
        $colourTrace = [pscustomobject]@{
            Arguments = New-Object 'Collections.Generic.List[object]'
            TargetPaths = New-Object 'Collections.Generic.List[string]'
            TargetHashes = New-Object 'Collections.Generic.List[string]'
            NativeCalls = 0
        }
        $runner = {
            param([string]$Executable,[string[]]$Arguments)
            $colourTrace.Arguments.Add(@($Arguments))
            $profileIndex = [Array]::IndexOf($Arguments, '-profile')
            if ($profileIndex -ge 0) {
                $targetPath = $Arguments[$profileIndex + 1]
                $colourTrace.TargetPaths.Add($targetPath)
                # Direct .NET reads preserve extended-path byte binding on PS5.1.
                $sha256 = [Security.Cryptography.SHA256]::Create()
                try {
                    $hash = $sha256.ComputeHash([IO.File]::ReadAllBytes($targetPath))
                    $colourTrace.TargetHashes.Add(([BitConverter]::ToString($hash).Replace('-', '').ToLowerInvariant()))
                } finally { $sha256.Dispose() }
            }
            if ($colourTrace.Arguments.Count -le 2) {
                # Controlled start failures exercise same-scale reassembly; the
                # six following positive conversions use the selected real build.
                return [pscustomobject]@{ ExitCode = 1; StartError = 'Controlled sharing start failure'; Win32ErrorCode = 32; StdOut = ''; StdErr = '' }
            }
            # Only relax this test's native extent while the caller's one-byte
            # limit forces all six valid outputs through the normal retry loop.
            $nativeArgs = @($Arguments | ForEach-Object { if ($_ -like 'jpeg:extent=*') { 'jpeg:extent=1048576B' } else { $_ } })
            $colourTrace.NativeCalls++
            $nativeText = @(& $Executable @nativeArgs 2>&1)
            return [pscustomobject]@{ ExitCode = $LASTEXITCODE; DiagnosticOutput = $nativeText -join "`n" }
        }
        $result = Invoke-ColourRun -Case $case -Runner $runner -MaxBytes 1
        $result.Code | Should -Be 2 -Because $result.Text
        $colourTrace.Arguments.Count | Should -Be 8
        $colourTrace.NativeCalls | Should -Be 6
        if ($Tagged) {
            @($colourTrace.TargetPaths | Select-Object -Unique).Count | Should -Be 1
            $colourTrace.TargetHashes.Count | Should -Be 8
            $expectedProfileHash = @($manifest.profiles | Where-Object source_path -eq 'profiles/sRGB-v4.icc')[0].sha256
            foreach ($targetHash in $colourTrace.TargetHashes) { $targetHash | Should -Be $expectedProfileHash }
        } else {
            $colourTrace.TargetPaths.Count | Should -Be 0
            $colourTrace.TargetHashes.Count | Should -Be 0
        }
        $scales = @()
        foreach ($arguments in $colourTrace.Arguments) {
            Assert-ColourCompositionOrder -Arguments $arguments -Tagged $Tagged
            $resize = [Array]::IndexOf($arguments,'-resize')
            $scales += $arguments[$resize + 1]
        }
        ($scales -join ',') | Should -Be '100%,100%,100%,90%,80%,70%,60%,50%'
        $output = Join-Path $result.Run 'transparent.jpeg'
        Assert-ColourPixels -Path $output -Width 64 -Height 48 -Expected $expected -Tolerance $reference.tolerance.jpeg_channel_units -Format JPEG
        Assert-ColourNoMetadata $output
        $result.Log | Should -Match 'Category=TransientIO'
        $result.Log | Should -Match 'DelayMs=100'
        $result.Log | Should -Match 'DelayMs=200'
        $result.Log | Should -Match '(?i)(above|exceed|best|over)'
        @(Get-ChildItem -LiteralPath (Join-Path $result.Run '.WinImgNormalizer\work') -Force).Count | Should -Be 0
        Get-ColourSourceState $case.Source | Should -Be $before
    }
}

Describe '1.0.3 actual JPEG ICC repair corpus and independent colour references (T076)' {
    It 'T076 reconciles the actual <Group> corpus, fully decoded outputs, independent colours and warning subsets' -ForEach @(
        @{Group='views'}, @{Group='negative'}, @{Group='mixed'}
    ) {
        $container=New-ColourDirectory ('icc-'+$Group)
        $case=[pscustomobject]@{Source=(Join-Path $container 'Kesä ICC input');Parent=(New-ColourDirectory 'icc-output');Probe=(New-ColourDirectory 'icc-probe');Name=''}
        $null=Copy-IccAcceptanceInputs -Assets $iccRepairAssets -Group $Group -Source $case.Source
        $before=Get-ColourSourceState $case.Source
        $result=Invoke-ColourRun $case
        $accepted=Assert-IccAcceptanceOutput -Assets $iccRepairAssets -Group $Group -Output $result.Run -MagickPath $magick -ExitCode $result.Code -ConsoleText $result.Text
        $accepted.IndependentReferencesPassed | Should -BeTrue
        Get-ColourSourceState $case.Source | Should -Be $before
    }

    It 'T076 applies exactly the reconstructed wide source bytes at every real scale with the embedded ICC excluded on read' {
        $case=New-ColourCase 'icc-split-scales' 'split.jpg'
        $sourcePath=Join-Path $case.Source $case.Name
        [IO.File]::Copy((Join-Path $iccRepairAssets 'fixtures/views/04-wide-split-reversed.jpg'),$sourcePath,$false)
        $normal=Join-Path $iccRepairAssets 'fixtures/views/03-good-wide.jpg'
        $segment=@(Get-ColourJpegSegments $normal|Where-Object {$_.Marker -eq 226 -and [Text.Encoding]::ASCII.GetString($_.Payload,0,12) -eq "ICC_PROFILE`0"})[0]
        $expected=[Convert]::ToBase64String([byte[]]$segment.Payload[14..($segment.Payload.Length-1)])
        $trace=[pscustomobject]@{Calls=0;Profiles=New-Object 'Collections.Generic.List[string]'}
        $runner={
            param([string]$Executable,[string[]]$Arguments)
            $trace.Calls++
            $skip=[Array]::IndexOf($Arguments,'profile:skip=ICC')
            $skip | Should -BeGreaterThan 0
            $Arguments[$skip-1] | Should -Be '-define'
            $indices=@(for($i=0;$i -lt $Arguments.Length;$i++){if($Arguments[$i] -eq '-profile'){$i}})
            $indices.Count | Should -Be 2 -Because 'validated source is assigned once before the sRGB transform'
            $sourceProfile=$Arguments[$indices[0]+1]
            $trace.Profiles.Add([Convert]::ToBase64String([IO.File]::ReadAllBytes($sourceProfile)))
            $native=@($Arguments|ForEach-Object {if($_ -like 'jpeg:extent=*'){'jpeg:extent=1048576B'}else{$_}})
            $nativeText=@(& $Executable @native 2>&1)
            return [pscustomobject]@{ExitCode=$LASTEXITCODE;DiagnosticOutput=$nativeText -join "`n"}
        }
        $before=Get-ColourSourceState $case.Source
        $result=Invoke-ColourRun -Case $case -Runner $runner -MaxBytes 1
        $result.Code | Should -Be 2
        $trace.Calls | Should -Be 6
        foreach($profile in $trace.Profiles){$profile | Should -Be $expected}
        @(Get-ChildItem -LiteralPath $result.Run -File -Filter '*.jpeg').Count | Should -Be 1
        Get-ColourSourceState $case.Source | Should -Be $before
    }

    It 'T076 treats <Kind> APP2 chunks as unusable metadata without losing clean RGB pixels' -ForEach @(
        @{Kind='missing';ExpectedReason='IccSegmentsMissing'},@{Kind='duplicate';ExpectedReason='IccSegmentsDuplicate'},@{Kind='contradictory';ExpectedReason='IccSegmentsContradictory'}
    ) {
        $case=New-ColourCase ('icc-'+$Kind) 'broken-chunks.jpg'
        $sourcePath=Join-Path $case.Source $case.Name
        [IO.File]::Copy((Join-Path $iccRepairAssets 'fixtures/views/01-untagged.jpg'),$sourcePath,$false)
        $normal=Join-Path $iccRepairAssets 'fixtures/views/03-good-wide.jpg'
        $segment=@(Get-ColourJpegSegments $normal|Where-Object {$_.Marker -eq 226 -and [Text.Encoding]::ASCII.GetString($_.Payload,0,12) -eq "ICC_PROFILE`0"})[0]
        $profile=[byte[]]$segment.Payload[14..($segment.Payload.Length-1)]
        $first=[byte[]]$profile[0..283];$second=[byte[]]$profile[284..($profile.Length-1)]
        Add-ColourJpegSegment -Path $sourcePath -Marker 226 -Payload ([Text.Encoding]::ASCII.GetBytes("ICC_PROFILE`0")+[byte[]]@(1,2)+$first)
        if($Kind -ne 'missing'){
            $sequence=if($Kind -eq 'duplicate'){1}else{2};$count=if($Kind -eq 'contradictory'){3}else{2}
            Add-ColourJpegSegment -Path $sourcePath -Marker 226 -Payload ([Text.Encoding]::ASCII.GetBytes("ICC_PROFILE`0")+[byte[]]@($sequence,$count)+$second)
        }
        $before=Get-ColourSourceState $case.Source
        $result=Invoke-ColourRun $case
        $result.Code | Should -Be 2
        $result.Log | Should -Match $ExpectedReason
        $result.Log | Should -Match 'SUMMARY ConvertedImages=1 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=0'
        Assert-IccAcceptanceJpeg -Path (Join-Path $result.Run 'broken-chunks.jpeg') -Reference (Join-Path $iccRepairAssets 'references/srgb.png') -MagickPath $magick
        Get-ColourSourceState $case.Source | Should -Be $before
    }

    It 'T076 caps a recurring ICC notice while preserving every warning in the local outcome CSV' {
        $case=New-ColourCase 'icc-recurring' ''
        for($i=1;$i -le 3;$i++){[IO.File]::Copy((Join-Path $iccRepairAssets 'fixtures/views/06-bad-signature.jpg'),(Join-Path $case.Source ('bad-'+$i+'.jpg')),$false)}
        $result=Invoke-ColourRun $case
        $notice=[regex]::Escape('JPEG saved; embedded colour profile unusable; sRGB assumed, colours may differ.')
        [regex]::Matches($result.Text,$notice).Count | Should -Be 1
        [regex]::Matches($result.Log,$notice).Count | Should -BeGreaterOrEqual 3
        $csv=@(Get-ChildItem -LiteralPath (Join-Path $result.Run '.WinImgNormalizer/reports') -Filter '*.csv' -File)[0]
        $rows=@(Import-Csv -LiteralPath $csv.FullName|ForEach-Object {foreach($property in $_.PSObject.Properties){if([string]$property.Value -like 'text:*'){$property.Value=([string]$property.Value).Substring(5)}};$_}|Where-Object RecordType -eq 'File')
        $rows.Count | Should -Be 3
        foreach($row in $rows){$row.Status|Should -Be 'Converted';$row.ColourPath|Should -Be 'JpegIccAssumeSrgb';[int]$row.ColourWarning|Should -Be 1}
        $result.Log | Should -Match 'SUMMARY ConvertedImages=3 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=0'
    }

    It 'T076 identifies the actual independent header assertion for <Name>' -ForEach @(
        @{Name='06-bad-signature.jpg';Reason='IccSignature'},
        @{Name='07-zero-tags.jpg';Reason='IccEmptyTagTable'},
        @{Name='08-huge-table.jpg';Reason='IccTagTableBounds'},
        @{Name='09-declared-too-large.jpg';Reason='IccDeclaredLength'}
    ) {
        $path=Join-Path $iccRepairAssets ('fixtures/views/'+$Name)
        $segment=@(Get-ColourJpegSegments $path|Where-Object {$_.Marker -eq 226 -and [Text.Encoding]::ASCII.GetString($_.Payload,0,12) -eq "ICC_PROFILE`0"})[0]
        $bytes=[byte[]]$segment.Payload[14..($segment.Payload.Length-1)]
        $caught=$null
        try {$null=Assert-WinImgIccBytes -Bytes $bytes -ColourSpace sRGB}catch{$caught=$_}
        $caught | Should -Not -BeNullOrEmpty
        $caught.Exception.Data['WinImgIccReason'] | Should -Be $Reason
    }

    It 'T076 retains an actual conversion failure for <Category> even after identifying unusable JPEG ICC' -ForEach @(
        @{Category='policy denial';Diagnostic='not authorized by the security policy JPEG';NativeExit=1},
        @{Category='resource limit';Diagnostic='cache resources exhausted';NativeExit=1},
        @{Category='unidentified native failure';Diagnostic='Unidentified native decode failure';NativeExit=1}
    ) {
        $case=New-ColourCase 'icc-real-failure' 'profile.jpg'
        $sourcePath=Join-Path $case.Source $case.Name
        [IO.File]::Copy((Join-Path $iccRepairAssets 'fixtures/views/06-bad-signature.jpg'),$sourcePath,$false)
        $trace=[pscustomobject]@{Calls=0;Diagnostic=$Diagnostic;Exit=$NativeExit}
        $runner={param($Executable,$Arguments);$trace.Calls++;return [pscustomobject]@{ExitCode=$trace.Exit;StdErr=$trace.Diagnostic}}
        $before=Get-ColourSourceState $case.Source
        $result=Invoke-ColourRun -Case $case -Runner $runner
        Assert-ColourFailed $result
        $trace.Calls | Should -Be 1
        $result.Log | Should -Not -Match 'JPEG saved; embedded colour profile unusable'
        Get-ColourSourceState $case.Source | Should -Be $before
    }

    It 'T076 rejects ICC-only fallback classification for incomplete or failed native lifecycle <Mode>' -ForEach @(
        @{Mode='incomplete capture';Field='StreamsComplete';Value=$false},
        @{Mode='drain timeout';Field='DrainTimedOut';Value=$true},
        @{Mode='unterminated child tree';Field='TreeTerminated';Value=$false},
        @{Mode='unexpected native exit';Field='ExitCode';Value=7},
        @{Mode='native timeout';Field='TimedOut';Value=$true},
        @{Mode='native cancellation';Field='Cancelled';Value=$true},
        @{Mode='capture overflow';Field='StdErrTruncated';Value=$true}
    ) {
        $native=[pscustomobject]@{ExitCode=1;StdOut='';StdErr='magick.exe: ColorspaceColorProfileMismatch `icc` @ warning/profile.c/LCMSExceptionHandler/123.';StartError='';StreamsComplete=$true;DrainTimedOut=$false;ProcessId=123;TreeTerminated=$true;TimedOut=$false;Cancelled=$false;StdErrTruncated=$false}
        Test-WinImgIccOnlyOutcome -Result $native | Should -BeTrue -Because 'the complete pure profile diagnostic is the positive control'
        $native.$Field=$Value
        Test-WinImgIccOnlyOutcome -Result $native | Should -BeFalse
    }

    It 'T076 refuses <Category> diagnostics from an attempted valid-profile JPEG transform without an ICC fallback' -ForEach @(
        @{Category='policy';Diagnostic='not authorized by the security policy JPEG'},
        @{Category='resource';Diagnostic='cache resources exhausted'},
        @{Category='pixel damage mixed with profile warning';Diagnostic="magick.exe: ColorspaceColorProfileMismatch ``icc`` @ warning/profile.c/LCMSExceptionHandler/123.`nmagick.exe: Premature end of JPEG file @ warning/jpeg.c/JPEGWarningHandler/123."}
    ) {
        $case=New-ColourCase 'icc-transform-failure' 'valid.jpg'
        [IO.File]::Copy((Join-Path $iccRepairAssets 'fixtures/views/03-good-wide.jpg'),(Join-Path $case.Source $case.Name),$false)
        $trace=[pscustomobject]@{Calls=0;Diagnostic=$Diagnostic}
        $runner={param($Executable,$Arguments);$trace.Calls++;return [pscustomobject]@{ExitCode=1;StdErr=$trace.Diagnostic}}
        $result=Invoke-ColourRun -Case $case -Runner $runner
        Assert-ColourFailed $result
        $trace.Calls | Should -Be 1
        $result.Log | Should -Not -Match 'JPEG saved; embedded colour profile unusable'
    }
}

Describe '1.0.3 explicit alternate EXIF characterization after unusable JPEG ICC (T076)' {
    BeforeAll {
        function New-ColourExifCharacterization {
            param([ValidateSet('Srgb','AdobeRgb','Duplicate')][string]$Characterization)
            $stream=[IO.MemoryStream]::new();$writer=[IO.BinaryWriter]::new($stream)
            try {
                $writer.Write([Text.Encoding]::ASCII.GetBytes("Exif`0`0II"));$writer.Write([uint16]42);$writer.Write([uint32]8)
                $writer.Write([uint16]1)
                $writer.Write([uint16]34665);$writer.Write([uint16]4);$writer.Write([uint32]1);$writer.Write([uint32]26)
                $writer.Write([uint32]0)
                $entries=if($Characterization -eq 'Duplicate'){3}else{2}
                $space=if($Characterization -eq 'AdobeRgb'){65535}else{1}
                $writer.Write([uint16]$entries)
                $writer.Write([uint16]40961);$writer.Write([uint16]3);$writer.Write([uint32]1);$writer.Write([uint16]$space);$writer.Write([uint16]0)
                $writer.Write([uint16]40965);$writer.Write([uint16]4);$writer.Write([uint32]1);$writer.Write([uint32](26+2+12*$entries+4))
                if($Characterization -eq 'Duplicate'){
                    $writer.Write([uint16]40961);$writer.Write([uint16]3);$writer.Write([uint32]1);$writer.Write([uint16]65535);$writer.Write([uint16]0)
                }
                $writer.Write([uint32]0);$writer.Write([uint16]1)
                $writer.Write([uint16]1);$writer.Write([uint16]2);$writer.Write([uint32]4)
                $index=if($Characterization -eq 'AdobeRgb'){"R03`0"}else{"R98`0"}
                $writer.Write([Text.Encoding]::ASCII.GetBytes($index));$writer.Write([uint32]0)
                return ,$stream.ToArray()
            } finally {$writer.Dispose();$stream.Dispose()}
        }
    }

    It 'T076 uses <Characterization> characterization for actual <Problem> ICC without an inaccurate sRGB assumption' -ForEach @(
        @{Characterization='Srgb';Problem='bad-signature';Policy='JpegExifSrgb';Calls=1;Managed=0},
        @{Characterization='AdobeRgb';Problem='bad-signature';Policy='JpegExifAdobeRgb';Calls=1;Managed=1},
        @{Characterization='AdobeRgb';Problem='invalid-curve';Policy='JpegExifAdobeRgb';Calls=2;Managed=1},
        @{Characterization='Duplicate';Problem='bad-signature';Policy='JpegIccAssumeSrgb';Calls=1;Managed=0}
    ) {
        $case=New-ColourCase 'exif-characterization' 'exif.jpg'
        $sourcePath=Join-Path $case.Source $case.Name
        New-ColourRgb -Path $sourcePath -Values $reference.rgb_source -Coder JPEG
        $profile=[IO.File]::ReadAllBytes($srgbProfile)
        if($Problem -eq 'bad-signature'){
            [Array]::Copy([Text.Encoding]::ASCII.GetBytes('bad!'),0,$profile,36,4)
        }else{
            $tags=Get-ColourBigEndian $profile 128 4;$curve=-1
            for($i=0;$i -lt $tags;$i++){
                $entry=132+12*$i
                if([Text.Encoding]::ASCII.GetString($profile,$entry,4) -eq 'rTRC'){$curve=Get-ColourBigEndian $profile ($entry+4) 4}
            }
            $curve | Should -BeGreaterThan 0
            [Array]::Copy([Text.Encoding]::ASCII.GetBytes('xxxx'),0,$profile,$curve,4)
        }
        Add-ColourJpegSegment -Path $sourcePath -Marker 226 -Payload ([Text.Encoding]::ASCII.GetBytes("ICC_PROFILE`0")+[byte[]]@(1,1)+$profile)
        Add-ColourJpegSegment -Path $sourcePath -Marker 225 -Payload (New-ColourExifCharacterization $Characterization)
        Set-ColourSourceTimes $sourcePath
        $before=Get-ColourSourceState $case.Source
        $trace=[pscustomobject]@{Calls=0}
        $runner={
            param([string]$Executable,[string[]]$Arguments)
            $trace.Calls++
            $nativeText=@(& $Executable @Arguments 2>&1)
            return [pscustomobject]@{ExitCode=$LASTEXITCODE;DiagnosticOutput=$nativeText -join "`n"}
        }
        $result=Invoke-ColourRun -Case $case -Runner $runner
        $result.Code | Should -Be 2 -Because $result.Text
        $trace.Calls | Should -Be $Calls
        $result.Log | Should -Match ([regex]::Escape('Policy='+$Policy))
        $result.Log | Should -Match ('ColourManaged='+$Managed+' ColourAssumed='+$(if($Characterization -eq 'Duplicate'){1}else{0})+' ColourWarnings=1')
        $expected=if($Characterization -eq 'AdobeRgb'){$reference.wide_rgb}else{$reference.rgb_source}
        $output=Join-Path $result.Run 'exif.jpeg'
        Assert-ColourPixels -Path $output -Width 128 -Height 96 -Expected $expected -Tolerance $reference.tolerance.jpeg_channel_units -Format JPEG
        Assert-ColourNoMetadata $output
        @(Get-ChildItem -LiteralPath (Join-Path $result.Run '.WinImgNormalizer/work') -Force).Count | Should -Be 0
        Get-ColourSourceState $case.Source | Should -Be $before
    }
}

Describe '1.0.3 damaged JPEG source scans remain genuine failures (T076)' {
    It 'T076 rejects damaged scan data with <ProfileState> through the actual native JPEG encoder' -ForEach @(
        @{ProfileState='no ICC'},@{ProfileState='valid ICC'}
    ) {
        $case=New-ColourCase 'jpeg-source-damage' 'damaged.jpg'
        $damaged=[IO.File]::ReadAllBytes((Join-Path $iccRepairAssets 'fixtures/negative/01-damaged-scan-and-bad-icc.jpg'))
        # Remove exactly the injected first APP2, preserving the damaged scan.
        $damaged[2] | Should -Be 255;$damaged[3] | Should -Be 226
        $segmentLength=Get-ColourBigEndian $damaged 4 2
        $bytes=[byte[]]@($damaged[0..1]+$damaged[(4+$segmentLength)..($damaged.Length-1)])
        $sourcePath=Join-Path $case.Source $case.Name
        [IO.File]::WriteAllBytes($sourcePath,$bytes)
        if($ProfileState -eq 'valid ICC'){
            Add-ColourJpegSegment -Path $sourcePath -Marker 226 -Payload ([Text.Encoding]::ASCII.GetBytes("ICC_PROFILE`0")+[byte[]]@(1,1)+[IO.File]::ReadAllBytes($srgbProfile))
        }
        $before=Get-ColourSourceState $case.Source
        $result=Invoke-ColourRun $case
        Assert-ColourFailed $result
        $result.Log | Should -Match 'NATIVE DETAILS:'
        $result.Log | Should -Match '(?i)(premature|corrupt|insufficient|end of data|end of file)'
        $result.Log | Should -Not -Match 'JPEG saved; embedded colour profile unusable'
        Get-ColourSourceState $case.Source | Should -Be $before
    }
}
