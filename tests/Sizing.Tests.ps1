BeforeAll {
    $repository=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    function Assert-SizingAncestors([string]$Path) {
        $current=[IO.DirectoryInfo]::new([IO.Path]::GetFullPath($Path))
        while ($current) {
            if ($current.Exists -and ($current.Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Sizing fixture crosses a reparse point.' }
            $current=$current.Parent
        }
    }
    $scratch=Join-Path $repository '.scratch'; Assert-SizingAncestors $scratch
    foreach ($pictures in @([Environment]::GetFolderPath('MyPictures'),(Join-Path $env:USERPROFILE 'Pictures'))) {
        if (-not $pictures) { continue }
        $full=[IO.Path]::GetFullPath($pictures).TrimEnd('\','/')
        if ($scratch.Equals($full,[StringComparison]::OrdinalIgnoreCase) -or $scratch.StartsWith($full+'\',[StringComparison]::OrdinalIgnoreCase) -or $full.StartsWith($scratch+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Sizing fixture overlaps real Pictures.' }
    }
    & git -C $repository check-ignore --quiet --no-index -- (Join-Path $scratch 'sizing-ignore-probe')
    if ($LASTEXITCODE -ne 0) { throw 'Sizing fixture must be ignored.' }
    $ownedRoot=Join-Path $scratch ('M2-T04-sizing-'+[Guid]::NewGuid().ToString('N'))
    if ([IO.Directory]::Exists($ownedRoot)) { throw 'Sizing ownership collision.' }
    $null=[IO.Directory]::CreateDirectory($ownedRoot)
    [IO.File]::WriteAllText((Join-Path $ownedRoot '.winimg-fixture-root'),'M2-T04 owned synthetic sizing')
    function New-SizingDirectory([string]$Label) {
        $path=[IO.Path]::GetFullPath((Join-Path $ownedRoot ($Label+'-'+[Guid]::NewGuid().ToString('N').Substring(0,8))))
        if (-not $path.StartsWith($ownedRoot+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Sizing fixture escaped ownership.' }
        Assert-SizingAncestors $path; $null=[IO.Directory]::CreateDirectory($path); return $path
    }
    if (-not $env:WINIMG_TEST_MAGICK -or -not [IO.Path]::IsPathRooted($env:WINIMG_TEST_MAGICK)) { throw 'Explicit verified ImageMagick required.' }
    $magick=[IO.Path]::GetFullPath($env:WINIMG_TEST_MAGICK); Assert-SizingAncestors $magick
    . (Join-Path $repository 'WinImgNormalizer.ps1')
    function Invoke-SizingMagick([string[]]$Arguments) {
        $messages=@(& $magick @Arguments 2>&1); $code=$LASTEXITCODE
        if ($code -ne 0 -or ($messages.Count -and $Arguments[0] -ne 'identify' -and $Arguments[-1] -ne 'info:')) { throw ('Sizing native fixture/probe failed: '+($messages -join "`n")) }
        return $messages -join "`n"
    }
    $templates=New-SizingDirectory 'templates'
    $plain=Join-Path $templates 'plain.png'; $jpeg=Join-Path $templates 'valid.jpg'; $tagged=Join-Path $templates 'tagged-alpha.png'
    $null=Invoke-SizingMagick @('-size','64x48','xc:rgb(224,32,32)',('PNG:'+$plain))
    $null=Invoke-SizingMagick @($plain,('JPEG:'+$jpeg))
    $null=Invoke-SizingMagick @('-size','64x48','xc:none','-fill','rgb(224,32,32)','-draw','rectangle 16,8 47,39','-profile',(Join-Path $PSScriptRoot 'fixtures/colour/sRGB-v4.icc'),('PNG:'+$tagged))
    $plainBytes=[IO.File]::ReadAllBytes($plain); $jpegBytes=[IO.File]::ReadAllBytes($jpeg); $taggedBytes=[IO.File]::ReadAllBytes($tagged)
    function Write-SizingSource([string]$Root,[string]$Relative,[byte[]]$Bytes) {
        $path=Join-Path $Root $Relative; $null=[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($path))
        [IO.File]::WriteAllBytes($path,$Bytes)
        $fixed=[DateTime]::Parse('2022-04-05T06:07:08Z').ToUniversalTime()
        [IO.File]::SetCreationTimeUtc($path,$fixed); [IO.File]::SetLastWriteTimeUtc($path,$fixed)
        return $path
    }
    function Get-SizingSourceState([string]$Root) {
        return @(Get-ChildItem -LiteralPath $Root -Recurse -Force -File | Sort-Object FullName | ForEach-Object {
            [pscustomobject]@{Path=$_.FullName.Substring($Root.Length+1);Hash=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash;Length=$_.Length;Creation=$_.CreationTimeUtc.Ticks;Modified=$_.LastWriteTimeUtc.Ticks}
        }) | ConvertTo-Json -Compress
    }
    function Invoke-SizingRun([string]$Source,[string]$Parent,[scriptblock]$Runner,[object]$Cap) {
        $parameters=@{Source=$Source;OutputParent=$Parent;MagickPath=$magick}
        if ($Runner) { $parameters.ProcessRunner=$Runner }
        if ($null -ne $Cap) { $parameters.MaxBytes=$Cap }
        $observed=@(& Invoke-WinImgNormalizer @parameters 6>&1 3>&1 2>&1)
        $codes=@($observed | Where-Object {$_ -is [int] -or $_ -is [long]}); $codes.Count | Should -Be 1
        $runs=@(Get-ChildItem -LiteralPath $Parent -Directory); $runs.Count | Should -Be 1
        $logs=@(Get-ChildItem -LiteralPath (Join-Path $runs[0].FullName '.WinImgNormalizer/reports') -File); $logs.Count | Should -Be 1
        return [pscustomobject]@{Code=$codes[0];Run=$runs[0].FullName;Log=[IO.File]::ReadAllText($logs[0].FullName);Text=$observed -join "`n"}
    }
    function Write-SizingCandidate([string]$Operand,[byte[]]$Bytes) {
        if (-not $Operand.StartsWith('JPEG:',[StringComparison]::Ordinal)) { throw 'Expected explicit JPEG output operand.' }
        $path=$Operand.Substring(5); $ordinary=$path
        if ($ordinary.StartsWith('\\?\',[StringComparison]::Ordinal)) { $ordinary=$ordinary.Substring(4) }
        if (-not $ordinary.StartsWith($ownedRoot+'\',[StringComparison]::OrdinalIgnoreCase) -or $ordinary -notlike '*\.WinImgNormalizer\work\*\image.jpeg') { throw 'Controlled sizing encoder escaped owned candidate.' }
        [IO.File]::WriteAllBytes($path,$Bytes)
    }
    function Assert-SizingJpeg([string]$Path,[int]$Width=64,[int]$Height=48) {
        Invoke-SizingMagick @('identify','+ping','-regard-warnings','-define','registry:filename:literal=true','-format','%m|%w|%h|%n',(Get-WinImgNativeOutputPath $Path)) | Should -Be ('JPEG|'+$Width+'|'+$Height+'|1')
    }
    function Assert-SizingCleanup([object]$Result,[string[]]$Files) {
        $actual=@(Get-ChildItem -LiteralPath $Result.Run -Recurse -Force -File | Where-Object Extension -ne '.log' | ForEach-Object {$_.FullName.Substring($Result.Run.Length+1)} | Sort-Object)
        ($actual -join '|') | Should -Be (@($Files | Sort-Object) -join '|')
        @(Get-ChildItem -LiteralPath (Join-Path $Result.Run '.WinImgNormalizer/work') -Force).Count | Should -Be 0
    }
    function New-SizingPaddedJpeg([int]$Length) {
        # Genuine COM segments control the encoder-result length without fake
        # validator sizes, trailing junk or a privacy/quality claim for this seam.
        $remaining=$Length-$jpegBytes.Length
        if ($remaining -lt 4) { throw 'Sizing padding needs room for a JPEG segment.' }
        $stream=New-Object IO.MemoryStream
        try {
            $stream.Write($jpegBytes,0,2)
            while ($remaining -gt 0) {
                $segment=[Math]::Min(65537,$remaining)
                if ($remaining-$segment -gt 0 -and $remaining-$segment -lt 4) { $segment-=4 }
                $payload=New-Object byte[] ($segment-4)
                $declared=$payload.Length+2
                $header=[byte[]]@(255,254,[byte]([Math]::Floor($declared/256)),[byte]($declared%256))
                $stream.Write($header,0,4); $stream.Write($payload,0,$payload.Length); $remaining-=$segment
            }
            $stream.Write($jpegBytes,2,$jpegBytes.Length-2)
            return ,$stream.ToArray()
        } finally { $stream.Dispose() }
    }
    $boundaryBytes=@{}
    foreach ($delta in @(-1,0,1)) {
        $bytes=New-SizingPaddedJpeg (1048576+$delta)
        $path=Join-Path $templates ('boundary-'+$delta+'.jpg'); [IO.File]::WriteAllBytes($path,$bytes)
        Assert-SizingJpeg $path
        $bytes.Length | Should -Be (1048576+$delta)
        $boundaryBytes[$delta]=$bytes
    }
    function Assert-SizingScaleContract([object[]]$Calls,[string]$Extent,[switch]$Tagged) {
        $scales=@()
        foreach ($arguments in $Calls) {
            $arguments | Should -Contain '-quiet'
            $arguments | Should -Contain '-regard-warnings'
            $arguments | Should -Contain 'registry:filename:literal=true'
            $arguments | Should -Contain ('jpeg:extent='+$Extent)
            $arguments | Should -Contain '-auto-orient'
            $arguments | Should -Not -Contain '-quality'
            $arguments | Should -Not -Contain '-sharpen'
            $arguments[[Array]::IndexOf($arguments,'-sampling-factor')+1] | Should -Be '4:2:0'
            $arguments[[Array]::IndexOf($arguments,'-interlace')+1] | Should -Be 'Line'
            $background=[Array]::IndexOf($arguments,'-background'); $strip=[Array]::IndexOf($arguments,'-strip')
            $arguments[$background+1] | Should -Be 'white'
            (@($arguments[($background+2)..($strip-1)]) -join '|') | Should -Be '-alpha|remove|-alpha|off'
            if ($Tagged) {
                $profile=[Array]::IndexOf($arguments,'-profile'); $intent=[Array]::IndexOf($arguments,'-intent')
                $arguments | Should -Contain '+black-point-compensation'
                $arguments[$intent+1] | Should -Be 'Relative'
                $profile | Should -BeGreaterThan $intent; $background | Should -BeGreaterThan $profile
                $hash=[Security.Cryptography.SHA256]::Create()
                try { [BitConverter]::ToString($hash.ComputeHash([IO.File]::ReadAllBytes($arguments[$profile+1]))).Replace('-','').ToLowerInvariant() | Should -Be 'c56e1685d888f5edb92fe07f2750f387f8fe8e91b32ff8fb0b56bfbbb9458353' } finally {$hash.Dispose()}
            }
            $scales+=$arguments[[Array]::IndexOf($arguments,'-resize')+1]
        }
        return $scales -join ','
    }
}

Describe 'M2-T04 exact extent and final-byte decisions (T040)' {
    It 'T040 uses exact invariant byte argv for <Label>' -ForEach @(
        @{Label='default1MiB';Cap=[long]1048576;Default=$true},
        @{Label='custom1024';Cap=[long]1024;Default=$false},
        @{Label='nonKiB1537';Cap=[long]1537;Default=$false},
        @{Label='nonKiB1048577';Cap=[long]1048577;Default=$false},
        @{Label='Int64max hostile culture';Cap=[long]::MaxValue;Default=$false}
    ) {
        $source=New-SizingDirectory 'extent-source'; $parent=New-SizingDirectory 'extent-output'
        $null=Write-SizingSource $source 'image.png' $plainBytes; $before=Get-SizingSourceState $source
        $trace=[pscustomobject]@{Calls=New-Object 'Collections.Generic.List[object]'}
        $runner={param($Executable,[string[]]$Arguments);$trace.Calls.Add([string[]]$Arguments.Clone());Write-SizingCandidate $Arguments[-1] $jpegBytes;return 0}
        $oldCulture=[Threading.Thread]::CurrentThread.CurrentCulture
        try {
            $hostile=[Globalization.CultureInfo]::GetCultureInfo('ar-SA').Clone()
            $hostile.NumberFormat.NumberGroupSeparator='_'; $hostile.NumberFormat.NumberDecimalSeparator=','
            [Threading.Thread]::CurrentThread.CurrentCulture=$hostile
            $capArgument=if ($Default) {$null} else {$Cap}
            $result=Invoke-SizingRun $source $parent $runner $capArgument
        } finally { [Threading.Thread]::CurrentThread.CurrentCulture=$oldCulture }
        $result.Code | Should -Be 0
        $trace.Calls.Count | Should -Be 1
        Assert-SizingScaleContract $trace.Calls.ToArray() ($Cap.ToString([Globalization.CultureInfo]::InvariantCulture)+'B') | Should -Be '100%'
        $result.Log | Should -Match ('MaxBytes='+$Cap.ToString([Globalization.CultureInfo]::InvariantCulture))
        Assert-SizingJpeg (Join-Path $result.Run 'image.jpeg'); Assert-SizingCleanup $result @('image.jpeg')
        Get-SizingSourceState $source | Should -Be $before
    }
    It 'T040 classifies actual final bytes at cap<Delta> independently of rounded1MiB display' -ForEach @(@{Delta=-1},@{Delta=0},@{Delta=1}) {
        $source=New-SizingDirectory 'boundary-source'; $parent=New-SizingDirectory 'boundary-output'
        $null=Write-SizingSource $source 'image.png' $plainBytes; $before=Get-SizingSourceState $source
        $trace=[pscustomobject]@{Calls=0;Bytes=$boundaryBytes[[int]$Delta]}
        $runner={param($Executable,$Arguments);$trace.Calls++;Write-SizingCandidate $Arguments[-1] $trace.Bytes;return 0}
        $result=Invoke-SizingRun $source $parent $runner 1048576
        $warning=[int]($Delta -gt 0); $result.Code | Should -Be (2*$warning)
        $trace.Calls | Should -Be $(if ($warning) {6} else {1})
        $output=Join-Path $result.Run 'image.jpeg'; [IO.File]::ReadAllBytes((Get-WinImgNativeOutputPath $output)).Length | Should -Be (1048576+$Delta)
        Assert-SizingJpeg $output
        $scale=if ($warning) {50} else {100}
        $result.Log | Should -Match ([regex]::Escape(('['+(1048576+$Delta)+' bytes, MaxBytes=1048576, Width=64, Height=48, Scale='+$scale+'%]')))
        $result.Log | Should -Match ('SUMMARY ConvertedImages=1 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=0 SizeWarnings='+$warning)
        if ($warning) {$result.Log | Should -Match 'WARN IMG:';$result.Log | Should -Not -Match 'OK IMG:'} else {$result.Log | Should -Match 'OK IMG:';$result.Log | Should -Not -Match 'WARN IMG:'}
        [Math]::Round((1048576+$Delta)/1MB,2) | Should -Be 1
        Assert-SizingCleanup $result @('image.jpeg'); Get-SizingSourceState $source | Should -Be $before
    }
}

Describe 'M2-T04 valid above-cap retention and unchanged real scale policy (T041-T042)' {
    It 'T040/T041 supersedes an above-cap trial with a compliant next-scale JPEG without retaining a warning' {
        $source=New-SizingDirectory 'compliant-source'; $parent=New-SizingDirectory 'compliant-output'
        $null=Write-SizingSource $source 'image.png' $plainBytes; $before=Get-SizingSourceState $source
        $trace=[pscustomobject]@{Calls=New-Object 'Collections.Generic.List[object]'}
        $runner={
            param($Executable,[string[]]$Arguments)
            $trace.Calls.Add([string[]]$Arguments.Clone())
            $bytes=if ($trace.Calls.Count -eq 1) {$boundaryBytes[-1]} else {$jpegBytes}
            Write-SizingCandidate $Arguments[-1] $bytes; return 0
        }
        $result=Invoke-SizingRun $source $parent $runner 1024
        $result.Code | Should -Be 0; $trace.Calls.Count | Should -Be 2
        Assert-SizingScaleContract $trace.Calls.ToArray() '1024B' | Should -Be '100%,90%'
        $output=Join-Path $result.Run 'image.jpeg'; Assert-SizingJpeg $output
        [Convert]::ToBase64String([IO.File]::ReadAllBytes((Get-WinImgNativeOutputPath $output))) | Should -Be ([Convert]::ToBase64String($jpegBytes))
        $jpegBytes.Length | Should -BeLessOrEqual 1024
        $result.Log | Should -Match ([regex]::Escape(('['+$jpegBytes.Length+' bytes, MaxBytes=1024, Width=64, Height=48, Scale=90%]')))
        $result.Log | Should -Match 'OK IMG:'; $result.Log | Should -Not -Match 'WARN IMG:'
        $result.Log | Should -Match 'SUMMARY ConvertedImages=1 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=0 SizeWarnings=0'
        Assert-SizingCleanup $result @('image.jpeg'); Get-SizingSourceState $source | Should -Be $before
    }
    It 'T041/T042 retains a real tiny-cap tagged alpha JPEG only with warning and preserves all six native attempts' {
        $source=New-SizingDirectory 'tiny-source'; $parent=New-SizingDirectory 'tiny-output'
        $null=Write-SizingSource $source 'image.png' $taggedBytes; $before=Get-SizingSourceState $source
        $trace=[pscustomobject]@{Calls=New-Object 'Collections.Generic.List[object]'}
        $runner={
            param($Executable,[string[]]$Arguments)
            $trace.Calls.Add([string[]]$Arguments.Clone())
            Assert-SizingScaleContract -Calls (,([string[]]$Arguments.Clone())) -Extent '1B' -Tagged | Out-Null
            # At this deliberately impossible cap the JPEG encoder heavily
            # quantizes pixels. Verify white composition independently before
            # JPEG coding with the same live profile/alpha/resize arguments.
            $lossless=@($Arguments.Clone() | Where-Object {$_ -ne 'jpeg:extent=1B'})
            $lossless=$lossless[0..($lossless.Length-3)]
            $probe=Join-Path $templates ('tiny-alpha-'+$trace.Calls.Count+'.png')
            $lossless+=('PNG:'+$probe)
            $null=Invoke-SizingMagick $lossless
            Invoke-SizingMagick @($probe,'-format','%[fx:p{1,1}.r==1&&p{1,1}.g==1&&p{1,1}.b==1]','info:') | Should -Be '1'
            $messages=@(& $Executable @Arguments 2>&1)
            return [pscustomobject]@{ExitCode=$LASTEXITCODE;DiagnosticOutput=$messages -join "`n"}
        }
        $result=Invoke-SizingRun $source $parent $runner 1
        $result.Code | Should -Be 2 -Because $result.Text
        $trace.Calls.Count | Should -Be 6
        (@($trace.Calls | ForEach-Object {$_[[Array]::IndexOf($_,'-resize')+1]}) -join ',') | Should -Be '100%,90%,80%,70%,60%,50%'
        $output=Join-Path $result.Run 'image.jpeg'; Assert-SizingJpeg $output 32 24
        $actualBytes=[IO.File]::ReadAllBytes((Get-WinImgNativeOutputPath $output)).Length; $actualBytes | Should -BeGreaterThan 1
        $result.Log | Should -Match ([regex]::Escape(('['+$actualBytes+' bytes, MaxBytes=1, Width=32, Height=24, Scale=50%]')))
        $result.Log | Should -Match 'WARN IMG:'; $result.Log | Should -Not -Match 'OK IMG:'
        $result.Log | Should -Match 'SUMMARY ConvertedImages=1 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=0 SizeWarnings=1'
        # Only this one-byte stress case allows 24/255 corner error after JPEG;
        # the lossless oracle above requires exact white on every scale.
        Invoke-SizingMagick @((Get-WinImgNativeOutputPath $output),'-format','%[fx:p{1,1}.r>=231/255&&p{1,1}.g>=231/255&&p{1,1}.b>=231/255&&p{16,12}.r>p{16,12}.g+0.3]','info:') | Should -Be '1'
        Assert-SizingCleanup $result @('image.jpeg'); Get-SizingSourceState $source | Should -Be $before
    }
    It 'T041 links a later duplicate to the retained warning and counts that output once' {
        $source=New-SizingDirectory 'duplicate-source'; $parent=New-SizingDirectory 'duplicate-output'
        $null=Write-SizingSource $source 'a\image.png' $plainBytes; $null=Write-SizingSource $source 'b\image.png' $plainBytes
        $before=Get-SizingSourceState $source
        $trace=[pscustomobject]@{Calls=0}
        $runner={param($Executable,$Arguments);$trace.Calls++;Write-SizingCandidate $Arguments[-1] $jpegBytes;return 0}
        $result=Invoke-SizingRun $source $parent $runner 1
        $result.Code | Should -Be 2; $trace.Calls | Should -Be 6
        $result.Log | Should -Match ([regex]::Escape('Heuristic duplicate skipped: b\image.png (retained source: a\image.png; retained output: a\image.jpeg; retained status: ConvertedWithWarning)'))
        $result.Log | Should -Match 'SUMMARY ConvertedImages=1 CopiedVideos=0 Duplicates=1 Unsupported=0 Errors=0 SizeWarnings=1'
        Assert-SizingJpeg (Join-Path $result.Run 'a\image.jpeg'); Assert-SizingCleanup $result @('a\image.jpeg')
        Get-SizingSourceState $source | Should -Be $before
    }
    It 'T041 refuses stale best effort after the last attempt is <Failure>' -ForEach @(@{Failure='truncated'},@{Failure='wrong-format'},@{Failure='nonzero'}) {
        $source=New-SizingDirectory 'last-source'; $parent=New-SizingDirectory 'last-output'
        $null=Write-SizingSource $source 'image.png' $plainBytes; $before=Get-SizingSourceState $source
        $trace=[pscustomobject]@{Calls=0;Paths=New-Object 'Collections.Generic.List[string]'}
        $runner={
            param($Executable,$Arguments)
            $trace.Calls++;$trace.Paths.Add($Arguments[-1])
            $bytes=$jpegBytes
            if ($trace.Calls -eq 6 -and $Failure -eq 'truncated') {$bytes=[byte[]]@(255,216,255)}
            if ($trace.Calls -eq 6 -and $Failure -eq 'wrong-format') {$bytes=$plainBytes}
            Write-SizingCandidate $Arguments[-1] $bytes
            if ($trace.Calls -eq 6 -and $Failure -eq 'nonzero') {return 8}; return 0
        }
        $result=Invoke-SizingRun $source $parent $runner 1
        $result.Code | Should -Be 2; $trace.Calls | Should -Be 6
        @($trace.Paths | Select-Object -Unique).Count | Should -Be 6
        $result.Log | Should -Match 'SUMMARY ConvertedImages=0 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=1 SizeWarnings=0'
        $result.Log | Should -Not -Match '(OK|WARN) IMG:'
        Assert-SizingCleanup $result @(); Get-SizingSourceState $source | Should -Be $before
    }
    It 'T042 keeps the real default single100percent geometry without new quality sharpening or upscaling flags' {
        $source=New-SizingDirectory 'default-source';$parent=New-SizingDirectory 'default-output'
        $null=Write-SizingSource $source 'image.png' $plainBytes; $before=Get-SizingSourceState $source
        $trace=[pscustomobject]@{Calls=New-Object 'Collections.Generic.List[object]'}
        $runner={param($Executable,[string[]]$Arguments);$trace.Calls.Add([string[]]$Arguments.Clone());$messages=@(& $Executable @Arguments 2>&1);return [pscustomobject]@{ExitCode=$LASTEXITCODE;DiagnosticOutput=$messages -join "`n"}}
        $result=Invoke-SizingRun $source $parent $runner
        $result.Code | Should -Be 0; $trace.Calls.Count | Should -Be 1
        Assert-SizingScaleContract $trace.Calls.ToArray() '1048576B' | Should -Be '100%'
        Assert-SizingJpeg (Join-Path $result.Run 'image.jpeg'); Assert-SizingCleanup $result @('image.jpeg')
        Get-SizingSourceState $source | Should -Be $before
    }
}
