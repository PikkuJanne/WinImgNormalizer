BeforeAll {
    $repository=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    function Assert-DiagnosticsAncestors([string]$Path) {
        $current=[IO.DirectoryInfo]::new([IO.Path]::GetFullPath($Path))
        while ($current) {
            if ($current.Exists -and ($current.Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Diagnostics fixture crosses a reparse point.' }
            $current=$current.Parent
        }
    }
    $scratch=Join-Path $repository '.scratch'; Assert-DiagnosticsAncestors $scratch
    foreach ($pictures in @([Environment]::GetFolderPath('MyPictures'),(Join-Path $env:USERPROFILE 'Pictures'))) {
        if (-not $pictures) { continue }
        $full=[IO.Path]::GetFullPath($pictures).TrimEnd('\','/')
        if ($scratch.Equals($full,[StringComparison]::OrdinalIgnoreCase) -or $scratch.StartsWith($full+'\',[StringComparison]::OrdinalIgnoreCase) -or $full.StartsWith($scratch+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Diagnostics fixture overlaps real Pictures.' }
    }
    & git -C $repository check-ignore --quiet --no-index -- (Join-Path $scratch 'diagnostics-ignore-probe')
    if ($LASTEXITCODE -ne 0) { throw 'Diagnostics fixtures must be ignored.' }
    $ownedRoot=Join-Path $scratch ('M2-T05-diagnostics-'+[Guid]::NewGuid().ToString('N'))
    if ([IO.Directory]::Exists($ownedRoot)) { throw 'Diagnostics ownership collision.' }
    $null=[IO.Directory]::CreateDirectory($ownedRoot)
    [IO.File]::WriteAllText((Join-Path $ownedRoot '.winimg-fixture-root'),'M2-T05 owned synthetic diagnostics')
    function New-DiagnosticsDirectory([string]$Label) {
        $path=[IO.Path]::GetFullPath((Join-Path $ownedRoot ($Label+'-'+[Guid]::NewGuid().ToString('N').Substring(0,8))))
        if (-not $path.StartsWith($ownedRoot+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Diagnostics fixture escaped ownership.' }
        Assert-DiagnosticsAncestors $path; $null=[IO.Directory]::CreateDirectory($path); return $path
    }
    if (-not $env:WINIMG_TEST_MAGICK -or -not [IO.Path]::IsPathRooted($env:WINIMG_TEST_MAGICK)) { throw 'Explicit verified ImageMagick required.' }
    $magick=[IO.Path]::GetFullPath($env:WINIMG_TEST_MAGICK); Assert-DiagnosticsAncestors $magick
    . (Join-Path $repository 'WinImgNormalizer.ps1')
    function Invoke-DiagnosticsMagick([string[]]$Arguments) {
        $result=Invoke-WinImgNativeProcess -Executable $magick -Arguments $Arguments
        if ($result.ExitCode -ne 0 -or $result.StartError -or -not $result.StreamsComplete -or $result.StdErr) { throw ('Diagnostics fixture/probe failed: '+$result.StdErr+' '+$result.StdOut) }
        return $result.StdOut.TrimEnd("`r","`n")
    }
    $templates=New-DiagnosticsDirectory 'templates'
    $png=Join-Path $templates 'tagged-alpha.png'; $jpeg=Join-Path $templates 'valid.jpg'; $truncated=Join-Path $templates 'truncated.jpg'
    $null=Invoke-DiagnosticsMagick @('-size','64x48','xc:none','-fill','rgb(224,32,32)','-draw','rectangle 16,8 47,39','-profile',(Join-Path $PSScriptRoot 'fixtures/colour/sRGB-v4.icc'),('PNG:'+$png))
    $null=Invoke-DiagnosticsMagick @($png,'-background','white','-alpha','remove','-alpha','off','-strip',('JPEG:'+$jpeg))
    $pngBytes=[IO.File]::ReadAllBytes($png); $jpegBytes=[IO.File]::ReadAllBytes($jpeg)
    # A genuine COM-segment JPEG controls encoder-result length for the combined
    # retry/sizing state machine. It does not claim real resizing or metadata
    # stripping by that controlled encoder call; final actual calls do both.
    $padded=Join-Path $templates 'valid-above-cap.jpg'
    $stream=[IO.MemoryStream]::new()
    try {
        $stream.Write($jpegBytes,0,2); $remaining=1048577-$jpegBytes.Length
        while ($remaining -gt 0) {
            $segment=[Math]::Min(65537,$remaining)
            if ($remaining-$segment -gt 0 -and $remaining-$segment -lt 4) { $segment-=4 }
            $payload=New-Object byte[] ($segment-4); $declared=$payload.Length+2
            $header=[byte[]]@(255,254,[byte]([Math]::Floor($declared/256)),[byte]($declared%256))
            $stream.Write($header,0,4); $stream.Write($payload,0,$payload.Length); $remaining-=$segment
        }
        $stream.Write($jpegBytes,2,$jpegBytes.Length-2); [IO.File]::WriteAllBytes($padded,$stream.ToArray())
    } finally {$stream.Dispose()}
    [IO.File]::WriteAllBytes($truncated,[byte[]]$jpegBytes[0..([int]($jpegBytes.Length*0.6))])
    $videoBytes=New-Object byte[] 65537
    for ($i=0;$i -lt $videoBytes.Length;$i++) { $videoBytes[$i]=[byte](($i*31+13)%256) }
    $observations=New-Object 'Collections.Generic.List[object]'

    # A managed native child is compiled only into marked ignored scratch. Two
    # threads write independently; no PowerShell error-stream behavior decides
    # its native exit. The recipe is bound by the suite's source hash.
    $childRoot=New-DiagnosticsDirectory "child & O'Brien [pipes]"
    $childSource=Join-Path $childRoot 'DiagnosticFixture.cs'; $child=Join-Path $childRoot 'DiagnosticFixture.exe'
    $recipe=@'
using System;
using System.IO;
using System.Threading;
public static class DiagnosticFixture {
 public static int Main(string[] args) {
  if(args.Length!=8) return 97;
  int code=int.Parse(args[0]),lines=int.Parse(args[1]);
  if(args[4]!="-") File.WriteAllBytes(args[5],File.ReadAllBytes(args[4]));
  Thread stdout=new Thread(delegate() {
   if(args[2]!="-") Console.Out.WriteLine(args[2]);
   for(int i=0;i<lines;i++) Console.Out.Write(new string('O',4096));
   if(args[6]!="-") Console.Out.WriteLine(args[6]);
  });
  Thread stderr=new Thread(delegate() {
   if(args[3]!="-") Console.Error.WriteLine(args[3]);
   for(int i=0;i<lines;i++) Console.Error.Write(new string('E',4096));
   if(args[7]!="-") Console.Error.WriteLine(args[7]);
  });
  stdout.Start();stderr.Start();stdout.Join();stderr.Join();return code;
 }
}
'@
    [IO.File]::WriteAllText($childSource,$recipe)
    $compiler=Join-Path $env:SystemRoot 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
    if (-not [IO.File]::Exists($compiler)) { throw 'Windows .NET Framework compiler required for the owned native-process fixture.' }
    $compile=@(& $compiler '/nologo' '/target:exe' ('/out:'+$child) $childSource 2>&1); $compileCode=$LASTEXITCODE
    [IO.File]::WriteAllText((Join-Path $childRoot 'compile.log'),($compile -join "`n"))
    if ($compileCode -ne 0 -or -not [IO.File]::Exists($child)) { throw 'Owned diagnostics fixture compilation failed.' }
    # The Framework compiler targets legacy path behavior by default. This
    # executable-local setting permits the owned extended-length candidate
    # operands; no machine configuration or application settings are changed.
    [IO.File]::WriteAllText(($child+'.config'),'<configuration><runtime><AppContextSwitchOverrides value="Switch.System.IO.UseLegacyPathHandling=false;Switch.System.IO.BlockLongPaths=false" /></runtime></configuration>')
    $observations.Add([pscustomobject]@{Case='fixture compiler provenance';CompilerSha256=(Get-FileHash -LiteralPath $compiler).Hash;CompilerVersion=([IO.FileInfo]::new($compiler)).VersionInfo.FileVersion;RecipeSha256=(Get-FileHash -LiteralPath $childSource).Hash;ChildSha256=(Get-FileHash -LiteralPath $child).Hash;ChildConfigSha256=(Get-FileHash -LiteralPath ($child+'.config')).Hash;NativeExit=$compileCode})
    function Invoke-DiagnosticsChild([int]$Code=0,[string]$StdOut='-',[string]$StdErr='-',[int]$Lines=0,[string]$CopyFrom='-',[string]$Candidate='-',[string]$OutTail='-',[string]$ErrTail='-',[int]$Limit=16384) {
        return Invoke-WinImgNativeProcess -Executable $child -Arguments @([string]$Code,[string]$Lines,$StdOut,$StdErr,$CopyFrom,$Candidate,$OutTail,$ErrTail) -OutputLimit $Limit -TimeoutMilliseconds 15000
    }
    $harmless="magick.exe: lossless to lossy JPEG conversion " + [char]96 + "owned.jpg' @ warning/jpeg.c/WriteJPEGImage_/3195."
    $sharing="magick.exe: unable to open image 'owned.jpg': The process cannot access the file because it is being used by another process. @ error/blob.c/OpenBlob/3596."
    function Write-DiagnosticsSource([string]$Root,[string]$Relative,[byte[]]$Bytes) {
        $path=Join-Path $Root $Relative; $null=[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($path))
        [IO.File]::WriteAllBytes($path,$Bytes)
        $fixed=[DateTime]::Parse('2022-04-05T06:07:08Z').ToUniversalTime()
        [IO.File]::SetCreationTimeUtc($path,$fixed); [IO.File]::SetLastWriteTimeUtc($path,$fixed)
        return $path
    }
    function Get-DiagnosticsState([string]$Root) {
        return @(Get-ChildItem -LiteralPath $Root -Recurse -Force -File | Sort-Object FullName | ForEach-Object {
            [pscustomobject]@{Path=$_.FullName.Substring($Root.Length+1);Hash=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash;Length=$_.Length;Creation=$_.CreationTimeUtc.Ticks;Modified=$_.LastWriteTimeUtc.Ticks}
        }) | ConvertTo-Json -Compress
    }
    function New-DiagnosticsCase([string]$Label) {
        $source=New-DiagnosticsDirectory ($Label+'-source'); $parent=New-DiagnosticsDirectory ($Label+'-output')
        $null=Write-DiagnosticsSource $source 'single.png' $pngBytes
        $video=Write-DiagnosticsSource $source 'nested/unchanged.MP4' $videoBytes
        $sentinel=Write-DiagnosticsSource $parent 'parent-sentinel.dat' ([Text.Encoding]::UTF8.GetBytes('pre-existing output-parent bytes'))
        return [pscustomobject]@{Source=$source;Parent=$parent;Before=(Get-DiagnosticsState $source);Video=$video;Sentinel=$sentinel;SentinelBefore=(Get-DiagnosticsState $parent)}
    }
    function Get-DiagnosticsCandidate([string[]]$Arguments) {
        $operand=$Arguments[-1]
        if (-not $operand.StartsWith('JPEG:',[StringComparison]::Ordinal)) { throw 'Expected explicit JPEG output.' }
        $path=$operand.Substring(5); $ordinary=$path
        if ($ordinary.StartsWith('\\?\',[StringComparison]::Ordinal)) { $ordinary=$ordinary.Substring(4) }
        if (-not $ordinary.StartsWith($ownedRoot+'\',[StringComparison]::OrdinalIgnoreCase) -or $ordinary -notlike '*\.WinImgNormalizer\work\*\image.jpeg') { throw 'Diagnostics runner escaped owned scratch.' }
        [IO.File]::Exists($path) | Should -BeTrue
        ([IO.FileInfo]::new($path)).Length | Should -Be 0
        return $path
    }
    function Invoke-DiagnosticsRun([object]$Case,[scriptblock]$Runner,[object]$Cap=1048576) {
        $parameters=@{Source=$Case.Source;OutputParent=$Case.Parent;MagickPath=$magick;MaxBytes=$Cap}
        if ($Runner) { $parameters.ProcessRunner=$Runner }
        $native=@(& Invoke-WinImgNormalizer @parameters 6>&1 3>&1 2>&1)
        $codes=@($native | Where-Object {$_ -is [int] -or $_ -is [long]}); $codes.Count | Should -Be 1
        $runs=@(Get-ChildItem -LiteralPath $Case.Parent -Directory); $runs.Count | Should -Be 1
        $logs=@(Get-ChildItem -LiteralPath (Join-Path $runs[0].FullName '.WinImgNormalizer/reports') -File -Filter '*.log'); $logs.Count | Should -Be 1
        return [pscustomobject]@{Code=$codes[0];Run=$runs[0].FullName;Log=[IO.File]::ReadAllText($logs[0].FullName);Text=$native -join "`n"}
    }
    function Assert-DiagnosticsPreserved([object]$Case,[object]$Result,[switch]$Image) {
        Get-DiagnosticsState $Case.Source | Should -Be $Case.Before
        $sentinel=[IO.FileInfo]::new($Case.Sentinel)
        [IO.File]::ReadAllText($sentinel.FullName) | Should -Be 'pre-existing output-parent bytes'
        ($Case.SentinelBefore | ConvertFrom-Json).Creation | Should -Be $sentinel.CreationTimeUtc.Ticks
        ($Case.SentinelBefore | ConvertFrom-Json).Modified | Should -Be $sentinel.LastWriteTimeUtc.Ticks
        $video=Join-Path $Result.Run 'nested/unchanged.MP4'
        (Get-FileHash -LiteralPath $video).Hash | Should -Be (Get-FileHash -LiteralPath $Case.Video).Hash
        ([IO.FileInfo]::new($video)).LastWriteTimeUtc.Ticks | Should -Be ([IO.FileInfo]::new($Case.Video)).LastWriteTimeUtc.Ticks
        $files=@(Get-ChildItem -LiteralPath $Result.Run -Recurse -Force -File | Where-Object { $_.Extension -notin @('.log','.csv') } | ForEach-Object {$_.FullName.Substring($Result.Run.Length+1)} | Sort-Object)
        $expected=@('nested\unchanged.MP4'); if ($Image) { $expected+='single.jpeg' }
        ($files -join '|') | Should -Be (($expected | Sort-Object) -join '|')
        @(Get-ChildItem -LiteralPath (Join-Path $Result.Run '.WinImgNormalizer/work') -Force).Count | Should -Be 0
    }
    function Assert-DiagnosticsJpeg([string]$Path,[int]$Width=64,[int]$Height=48) {
        $native=Get-WinImgNativeOutputPath $Path
        Invoke-DiagnosticsMagick @('identify','+ping','-regard-warnings','-define','registry:filename:literal=true','-format','%m|%w|%h|%n',$native) | Should -Be ('JPEG|'+$Width+'|'+$Height+'|1')
        $rgb=(Invoke-DiagnosticsMagick @($native,'-format','%[fx:p{2,2}.r]|%[fx:p{2,2}.g]|%[fx:p{2,2}.b]|%[fx:p{32,24}.r]|%[fx:p{32,24}.g]|%[fx:p{32,24}.b]','info:')).Split('|')
        foreach ($index in 0,1,2) { [double]::Parse($rgb[$index],[Globalization.CultureInfo]::InvariantCulture) | Should -BeGreaterThan 0.95 }
        [double]::Parse($rgb[3],[Globalization.CultureInfo]::InvariantCulture) | Should -BeGreaterThan 0.8
        foreach ($index in 4,5) { [double]::Parse($rgb[$index],[Globalization.CultureInfo]::InvariantCulture) | Should -BeLessThan 0.2 }
    }
    function Assert-DiagnosticsPolicy([object[]]$Calls,[string[]]$Scales,[string]$Extent='1048576B',[switch]$LiveProfile) {
        $actual=New-Object 'Collections.Generic.List[string]'
        $targets=New-Object 'Collections.Generic.List[string]'
        foreach ($arguments in $Calls) {
            $arguments | Should -Contain '-quiet'; $arguments | Should -Contain '-regard-warnings'
            $arguments | Should -Contain 'registry:filename:literal=true'; $arguments | Should -Contain '-auto-orient'
            $arguments | Should -Contain '+black-point-compensation'; $arguments | Should -Contain ('jpeg:extent='+$Extent)
            $profile=[Array]::IndexOf($arguments,'-profile'); $intent=[Array]::IndexOf($arguments,'-intent'); $background=[Array]::IndexOf($arguments,'-background'); $strip=[Array]::IndexOf($arguments,'-strip')
            $arguments[$intent+1] | Should -Be 'Relative'; $profile | Should -BeGreaterThan $intent
            $arguments[$background+1] | Should -Be 'white'; $background | Should -BeGreaterThan $profile
            (@($arguments[($background+2)..($strip-1)]) -join '|') | Should -Be '-alpha|remove|-alpha|off'
            $arguments[[Array]::IndexOf($arguments,'-sampling-factor')+1] | Should -Be '4:2:0'
            $arguments[[Array]::IndexOf($arguments,'-interlace')+1] | Should -Be 'Line'
            $arguments | Should -Not -Contain '-quality'; $arguments | Should -Not -Contain '-sharpen'
            if ($LiveProfile) {
                $hash=[Security.Cryptography.SHA256]::Create()
                try { [BitConverter]::ToString($hash.ComputeHash([IO.File]::ReadAllBytes($arguments[$profile+1]))).Replace('-','').ToLowerInvariant() | Should -Be 'c56e1685d888f5edb92fe07f2750f387f8fe8e91b32ff8fb0b56bfbbb9458353' } finally {$hash.Dispose()}
            }
            $targets.Add($arguments[$profile+1]); $actual.Add($arguments[[Array]::IndexOf($arguments,'-resize')+1])
        }
        ($actual.ToArray() -join '|') | Should -Be ($Scales -join '|')
        @($targets.ToArray() | Select-Object -Unique).Count | Should -Be 1
    }
}

AfterAll {
    if ($ownedRoot -and $observations) {
        [IO.File]::WriteAllText((Join-Path $ownedRoot 'native-observations.json'),($observations.ToArray() | ConvertTo-Json -Depth 12),[Text.UTF8Encoding]::new($false))
        Write-Host ('Diagnostics owned observations: '+$ownedRoot)
    }
}

Describe 'M2-T05 native streams and bounded retention (T043, T045)' {
    It 'T043 captures stderr and native exit independently of PowerShell error preferences' {
        $prior=$ErrorActionPreference
        try {
            $ErrorActionPreference='Stop'
            $result=Invoke-DiagnosticsChild -Code 7 -StdOut 'stdout payload with spaces, quotes " and a trailing \' -StdErr 'stderr payload is native data'
        } finally { $ErrorActionPreference=$prior }
        $result.ExitCode | Should -Be 7; $result.StreamsComplete | Should -BeTrue
        $result.StdOut | Should -Be "stdout payload with spaces, quotes `" and a trailing \`r`n"
        $result.StdErr | Should -Be "stderr payload is native data`r`n"
        $result.StdOutTruncated | Should -BeFalse; $result.StdErrTruncated | Should -BeFalse
        $result.TimedOut | Should -BeFalse; $result.StartError | Should -BeNullOrEmpty
        $result.StdOutCharacters | Should -Be $result.StdOut.Length; $result.StdErrCharacters | Should -Be $result.StdErr.Length
        $observations.Add([pscustomobject]@{Case='T043 separate streams';Result=$result})
    }

    It 'T043 records a failed native launch separately from an exit or stderr diagnostic' {
        $absent=Join-Path $childRoot 'owned-absent.exe'
        $result=Invoke-WinImgNativeProcess -Executable $absent -Arguments @()
        $result.ExitCode | Should -BeNullOrEmpty; $result.StartError | Should -Not -BeNullOrEmpty
        $result.Win32ErrorCode | Should -Be 2
        $outcome=Get-WinImgNativeOutcome -Result $result
        $outcome.Category | Should -Be 'StartFailure'; $outcome.Acceptable | Should -BeFalse; $outcome.Retryable | Should -BeFalse
        $observations.Add([pscustomobject]@{Case='T043 actual absent executable';Result=$result;Outcome=$outcome})
    }

    It 'T045 drains concurrent newline-free four-MiB stdout and stderr without retaining unbounded data' {
        $watch=[Diagnostics.Stopwatch]::StartNew()
        $result=Invoke-DiagnosticsChild -Code 9 -StdOut 'STDOUT-HEAD' -StdErr 'STDERR-HEAD' -Lines 1024 -OutTail 'STDOUT-TAIL' -ErrTail 'STDERR-TAIL' -Limit 1024
        $watch.Stop()
        $result.ExitCode | Should -Be 9; $result.StreamsComplete | Should -BeTrue; $result.TimedOut | Should -BeFalse
        $watch.Elapsed.TotalSeconds | Should -BeLessThan 15
        $result.StdOutCharacters | Should -Be (4194304+26); $result.StdErrCharacters | Should -Be (4194304+26)
        $result.StdOut.Length | Should -BeLessOrEqual 1024; $result.StdErr.Length | Should -BeLessOrEqual 1024
        $result.StdOutTruncated | Should -BeTrue; $result.StdErrTruncated | Should -BeTrue
        $result.StdOut | Should -Match '^STDOUT-HEAD'; $result.StdOut | Should -Match 'STDOUT-TAIL\r\n$'
        $result.StdErr | Should -Match '^STDERR-HEAD'; $result.StdErr | Should -Match 'STDERR-TAIL\r\n$'
        (Get-WinImgNativeOutcome -Result $result).Category | Should -Be 'OutputLimit'
        $observations.Add([pscustomobject]@{Case='T045 concurrent four-MiB streams';ElapsedMs=$watch.ElapsedMilliseconds;Result=$result})
    }

    It 'T043 verifies the allowlisted warning with an actual pinned lossless-to-lossy native JPEG command' {
        $target=Join-Path $templates 'actual-lossless-to-lossy.jpg'
        $result=Invoke-WinImgNativeProcess -Executable $magick -Arguments @($png,'-background','white','-alpha','remove','-alpha','off','-compress','LosslessJPEG','-quality','95',('JPEG:'+$target))
        $result.ExitCode | Should -Be 0; $result.StdErr | Should -Match 'lossless to lossy JPEG conversion.*@ warning/jpeg\.c/WriteJPEGImage_/'
        $outcome=Get-WinImgNativeOutcome -Result $result
        $outcome.Category | Should -Be 'Warning'; $outcome.Acceptable | Should -BeTrue; $outcome.Retryable | Should -BeFalse
        Assert-DiagnosticsJpeg $target
        $observations.Add([pscustomobject]@{Case='T043 genuine benign warning semantics';Result=$result;Outcome=$outcome})
        $nonzeroTarget=Join-Path $templates 'actual-regarded-lossless-to-lossy.jpg'
        $nonzero=Invoke-WinImgNativeProcess -Executable $magick -Arguments @('-regard-warnings',$png,'-background','white','-alpha','remove','-alpha','off','-compress','LosslessJPEG','-quality','95',('JPEG:'+$nonzeroTarget))
        $nonzero.ExitCode | Should -Not -Be 0
        $nonzero.StdErr | Should -Match 'lossless to lossy JPEG conversion'
        (Get-WinImgNativeOutcome -Result $nonzero).Acceptable | Should -BeFalse
        Assert-DiagnosticsJpeg $nonzeroTarget
        $observations.Add([pscustomobject]@{Case='T043 same genuine warning with regard-warnings is nonzero and rejected';Result=$nonzero})
    }
}

Describe 'M2-T05 classified failure and fresh candidate boundaries (T043, T044)' {
    It 'T044 classifies genuine unknown-coder and truncated-JPEG native failures without treating either as transient' {
        $unknown=Join-Path $templates 'not-an-image.xyz'
        [IO.File]::WriteAllText($unknown,'owned synthetic bytes without any image signature')
        $missing=Invoke-WinImgNativeProcess -Executable $magick -Arguments @('-regard-warnings',$unknown,('JPEG:'+(Join-Path $templates 'unknown-coder-output.jpg')))
        $missing.ExitCode | Should -Not -Be 0; $missing.StdErr | Should -Match 'no decode delegate'
        $missingOutcome=Get-WinImgNativeOutcome -Result $missing
        $missingOutcome.Category | Should -Be 'MissingCodec'; $missingOutcome.Acceptable | Should -BeFalse; $missingOutcome.Retryable | Should -BeFalse
        $damaged=Invoke-WinImgNativeProcess -Executable $magick -Arguments @('-regard-warnings',$truncated,('JPEG:'+(Join-Path $templates 'damaged-output.jpg')))
        $damaged.ExitCode | Should -Not -Be 0
        $damagedOutcome=Get-WinImgNativeOutcome -Result $damaged
        $damagedOutcome.Category | Should -Be 'DamagedInput'; $damagedOutcome.Acceptable | Should -BeFalse; $damagedOutcome.Retryable | Should -BeFalse
        $observations.Add([pscustomobject]@{Case='T044 genuine undecodable coder';Result=$missing;Outcome=$missingOutcome})
        $observations.Add([pscustomobject]@{Case='T044 genuine truncated JPEG';Result=$damaged;Outcome=$damagedOutcome})
    }

    It 'T044 treats generic native permission-denied text as permanent even when an owned exclusive lock caused it' {
        $lock=[IO.File]::Open($png,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::None)
        try { $result=Invoke-WinImgNativeProcess -Executable $magick -Arguments @('identify',$png) }
        finally {$lock.Dispose()}
        $result.ExitCode | Should -Not -Be 0; $result.StdErr | Should -Match 'Permission denied'
        $outcome=Get-WinImgNativeOutcome -Result $result
        $outcome.Category | Should -Be 'AccessDenied'; $outcome.Acceptable | Should -BeFalse; $outcome.Retryable | Should -BeFalse
        $observations.Add([pscustomobject]@{Case='T044 owned exclusive lock, ambiguous native Permission denied';Result=$result;Outcome=$outcome})
    }

    It 'T044 stops a <Category> native outcome after one attempt despite a decodable JPEG' -ForEach @(
        @{Category='MissingCodec';Message="magick.exe: no decode delegate for this image format 'XYZ' @ error/constitute.c/ReadImage/746."},
        @{Category='DamagedInput';Message="magick.exe: insufficient image data in file 'owned.png' @ error/png.c/ReadPNGImage/4260."},
        @{Category='AccessDenied';Message="magick.exe: unable to open image 'owned.jpg': Permission denied @ error/blob.c/OpenBlob/3596."},
        @{Category='ResourceExhaustion';Message="magick.exe: cache resources exhausted 'owned.png' @ error/cache.c/OpenPixelCache/4095."},
        @{Category='NativeFailure';Message='controlled unknown native failure without a retry justification'}
    ) {
        $case=New-DiagnosticsCase $Category
        $trace=[pscustomobject]@{Calls=0;Paths=New-Object 'Collections.Generic.List[string]';Results=New-Object 'Collections.Generic.List[object]'}
        $runner={
            param([string]$Executable,[string[]]$Arguments)
            $trace.Calls++; $candidate=Get-DiagnosticsCandidate $Arguments; $trace.Paths.Add($candidate)
            $result=Invoke-DiagnosticsChild -Code 1 -StdErr $Message -CopyFrom $jpeg -Candidate $candidate
            $result.ExitCode | Should -Be 1; $result.StreamsComplete | Should -BeTrue
            ([IO.FileInfo]::new($candidate)).Length | Should -Be $jpegBytes.Length
            $trace.Results.Add($result); return $result
        }
        $result=Invoke-DiagnosticsRun $case $runner
        $result.Code | Should -Be 2; $trace.Calls | Should -Be 1
        $result.Log | Should -Match ('Category='+$Category); $result.Log | Should -Match 'Exit=1'
        $result.Log | Should -Not -Match 'DelayMs='; $result.Log | Should -Not -Match 'OK IMG:|WARN IMG:'
        foreach ($path in $trace.Paths) { [IO.File]::Exists($path) | Should -BeFalse }
        Assert-DiagnosticsPreserved $case $result
        $observations.Add([pscustomobject]@{Case=('T044 controlled native '+$Category);ControlledDiagnostic=$Message;NativeResults=$trace.Results.ToArray();ApplicationCode=$result.Code})
    }

    It 'T044 retries only explicit sharing failures, keeping full ICC and white policy through success on call <SuccessCall>' -ForEach @(@{SuccessCall=2},@{SuccessCall=3}) {
        $case=New-DiagnosticsCase ('sharing-'+$SuccessCall)
        $trace=[pscustomobject]@{Calls=0;Paths=New-Object 'Collections.Generic.List[string]';Arguments=New-Object 'Collections.Generic.List[object]'}
        $runner={
            param([string]$Executable,[string[]]$Arguments)
            $trace.Calls++; $candidate=Get-DiagnosticsCandidate $Arguments; $trace.Paths.Add($candidate); $trace.Arguments.Add($Arguments.Clone())
            Assert-DiagnosticsPolicy -Calls (,($Arguments.Clone())) -Scales @('100%') -LiveProfile
            if ($trace.Calls -lt $SuccessCall) { return Invoke-DiagnosticsChild -Code 1 -StdErr $sharing -CopyFrom $truncated -Candidate $candidate }
            return Invoke-WinImgNativeProcess -Executable $Executable -Arguments $Arguments
        }
        $result=Invoke-DiagnosticsRun $case $runner
        $result.Code | Should -Be 0 -Because $result.Text; $trace.Calls | Should -Be $SuccessCall
        @($trace.Paths.ToArray() | Select-Object -Unique).Count | Should -Be $SuccessCall
        foreach ($path in $trace.Paths) { [IO.File]::Exists($path) | Should -BeFalse }
        Assert-DiagnosticsPolicy $trace.Arguments.ToArray() (@(1..$SuccessCall | ForEach-Object {'100%'}))
        $result.Log | Should -Match 'Category=TransientIO'; $result.Log | Should -Match 'DelayMs=100'
        if ($SuccessCall -eq 3) { $result.Log | Should -Match 'DelayMs=200' }
        $result.Log | Should -Not -Match 'WARN IMG:'
        Assert-DiagnosticsJpeg (Join-Path $result.Run 'single.jpeg')
        Assert-DiagnosticsPreserved $case $result -Image
        $observations.Add([pscustomobject]@{Case=('T044 sharing then actual conversion call '+$SuccessCall);Calls=$trace.Calls;Arguments=$trace.Arguments.ToArray();ApplicationCode=$result.Code})
    }

    It 'T044 ends persistent sharing failure at three fresh same-scale attempts and keeps no earlier JPEG' {
        $case=New-DiagnosticsCase 'sharing-exhausted'
        $trace=[pscustomobject]@{Calls=0;Paths=New-Object 'Collections.Generic.List[string]';Arguments=New-Object 'Collections.Generic.List[object]'}
        $runner={
            param([string]$Executable,[string[]]$Arguments)
            $trace.Calls++; $candidate=Get-DiagnosticsCandidate $Arguments; $trace.Paths.Add($candidate); $trace.Arguments.Add($Arguments.Clone())
            Assert-DiagnosticsPolicy -Calls (,($Arguments.Clone())) -Scales @('100%') -LiveProfile
            return Invoke-DiagnosticsChild -Code 1 -StdErr $sharing -CopyFrom $jpeg -Candidate $candidate
        }
        $result=Invoke-DiagnosticsRun $case $runner
        $result.Code | Should -Be 2; $trace.Calls | Should -Be 3
        @($trace.Paths.ToArray() | Select-Object -Unique).Count | Should -Be 3
        foreach ($path in $trace.Paths) { [IO.File]::Exists($path) | Should -BeFalse }
        Assert-DiagnosticsPolicy $trace.Arguments.ToArray() @('100%','100%','100%')
        $result.Log | Should -Match 'DelayMs=100'; $result.Log | Should -Match 'DelayMs=200'; $result.Log | Should -Not -Match 'DelayMs=300'
        Assert-DiagnosticsPreserved $case $result
    }

    It 'T044 executes at most eight calls when two sharing retries precede six real valid above-cap size attempts' {
        $case=New-DiagnosticsCase 'eight-call-bound'
        $trace=[pscustomobject]@{Calls=0;Paths=New-Object 'Collections.Generic.List[string]';Arguments=New-Object 'Collections.Generic.List[object]';NativeResults=New-Object 'Collections.Generic.List[object]'}
        $runner={
            param([string]$Executable,[string[]]$Arguments)
            if ($trace.Paths.Count) {
                $previous=$trace.Paths[$trace.Paths.Count-1]
                [IO.File]::Exists($previous) | Should -BeFalse
                [IO.Directory]::Exists([IO.Path]::GetDirectoryName($previous)) | Should -BeFalse
            }
            $trace.Calls++; $candidate=Get-DiagnosticsCandidate $Arguments; $trace.Paths.Add($candidate); $trace.Arguments.Add($Arguments.Clone())
            $scale=$Arguments[[Array]::IndexOf($Arguments,'-resize')+1]
            Assert-DiagnosticsPolicy -Calls (,($Arguments.Clone())) -Scales @($scale) -Extent '1B' -LiveProfile
            if ($trace.Calls -le 2) { return Invoke-DiagnosticsChild -Code 1 -StdErr $sharing -CopyFrom $truncated -Candidate $candidate }
            $native=Invoke-WinImgNativeProcess -Executable $Executable -Arguments $Arguments
            $native.ExitCode | Should -Be 0; $native.StdErr | Should -BeNullOrEmpty
            ([IO.FileInfo]::new($candidate)).Length | Should -BeGreaterThan 1
            $trace.NativeResults.Add($native); return $native
        }
        $result=Invoke-DiagnosticsRun $case $runner 1
        $result.Code | Should -Be 2; $trace.Calls | Should -Be 8; $trace.NativeResults.Count | Should -Be 6
        Assert-DiagnosticsPolicy -Calls $trace.Arguments.ToArray() -Scales @('100%','100%','100%','90%','80%','70%','60%','50%') -Extent '1B'
        @($trace.Paths.ToArray() | Select-Object -Unique).Count | Should -Be 8
        foreach ($path in $trace.Paths) { [IO.File]::Exists($path) | Should -BeFalse }
        ([regex]::Matches($result.Log,'RETRY IMG:')).Count | Should -Be 2
        $result.Log | Should -Match 'NATIVE IMG: Attempt=8; Scale=50%; Category=Success'
        $result.Log | Should -Match 'WARN IMG: single.png'; $result.Log | Should -Match 'SizeWarnings=1 NativeWarnings=0'
        $target=Join-Path $result.Run 'single.jpeg'
        ([IO.FileInfo]::new($target)).Length | Should -BeGreaterThan 1
        Invoke-DiagnosticsMagick @('identify','+ping','-regard-warnings','-define','registry:filename:literal=true','-format','%m|%w|%h|%n',(Get-WinImgNativeOutputPath $target)) | Should -Be 'JPEG|32|24|1'
        # T041's separate exact pre-JPEG white and calibrated JPEG checks cover
        # alpha quality. This stress case asserts validity and retry/policy bounds,
        # without inventing a high-quality floor for the one-byte target.
        Assert-DiagnosticsPreserved $case $result -Image
        $observations.Add([pscustomobject]@{Case='T044 actual eight-call bound';ApplicationCode=$result.Code;Calls=$trace.Calls;Arguments=$trace.Arguments.ToArray();NativeSizeAttemptResults=$trace.NativeResults.ToArray();RetainedBytes=([IO.FileInfo]::new($target)).Length;RetainedWidth=32;RetainedHeight=24})
    }

    It 'T044 trusts only the Win32 HRESULT facility for controlled IOException <Bits>' -ForEach @(
        @{Bits='80070020';Retry=$true},@{Bits='80070021';Retry=$true},@{Bits='80050020';Retry=$false}
    ) {
        $case=New-DiagnosticsCase ('hresult-'+$Bits)
        $trace=[pscustomobject]@{Calls=0;Paths=New-Object 'Collections.Generic.List[string]';Arguments=New-Object 'Collections.Generic.List[object]'}
        $runner={
            param([string]$Executable,[string[]]$Arguments)
            $trace.Calls++; $candidate=Get-DiagnosticsCandidate $Arguments; $trace.Paths.Add($candidate); $trace.Arguments.Add($Arguments.Clone())
            Assert-DiagnosticsPolicy -Calls (,($Arguments.Clone())) -Scales @('100%') -LiveProfile
            if ($trace.Calls -eq 1) { throw [IO.IOException]::new('Controlled IOException; retry depends on trusted HRESULT only.',[Convert]::ToInt32($Bits,16)) }
            return Invoke-WinImgNativeProcess -Executable $Executable -Arguments $Arguments
        }
        $result=Invoke-DiagnosticsRun $case $runner
        if ($Retry) {
            $result.Code | Should -Be 0; $trace.Calls | Should -Be 2
            $result.Log | Should -Match 'Category=TransientIO'; $result.Log | Should -Match 'DelayMs=100'
            Assert-DiagnosticsPolicy $trace.Arguments.ToArray() @('100%','100%')
            Assert-DiagnosticsJpeg (Join-Path $result.Run 'single.jpeg')
            Assert-DiagnosticsPreserved $case $result -Image
        } else {
            $result.Code | Should -Be 2; $trace.Calls | Should -Be 1
            $result.Log | Should -Not -Match 'Category=TransientIO|DelayMs='
            Assert-DiagnosticsPreserved $case $result
        }
        @($trace.Paths.ToArray() | Select-Object -Unique).Count | Should -Be $trace.Calls
        foreach ($path in $trace.Paths) { [IO.File]::Exists($path) | Should -BeFalse }
        $observations.Add([pscustomobject]@{Case=('T044 controlled HRESULT '+$Bits);ExpectedRetry=$Retry;Calls=$trace.Calls;ApplicationCode=$result.Code})
    }

    It 'T044 separates a same-scale transient retry from the next valid sizing scale' {
        $case=New-DiagnosticsCase 'retry-then-sizing'
        $trace=[pscustomobject]@{Calls=0;Paths=New-Object 'Collections.Generic.List[string]';Arguments=New-Object 'Collections.Generic.List[object]'}
        $runner={
            param([string]$Executable,[string[]]$Arguments)
            $trace.Calls++; $candidate=Get-DiagnosticsCandidate $Arguments; $trace.Paths.Add($candidate); $trace.Arguments.Add($Arguments.Clone())
            $scale=$Arguments[[Array]::IndexOf($Arguments,'-resize')+1]
            Assert-DiagnosticsPolicy -Calls (,($Arguments.Clone())) -Scales @($scale) -LiveProfile
            if ($trace.Calls -eq 1) { return Invoke-DiagnosticsChild -Code 1 -StdErr $sharing -CopyFrom $truncated -Candidate $candidate }
            if ($trace.Calls -eq 2) { return Invoke-DiagnosticsChild -CopyFrom $padded -Candidate $candidate }
            return Invoke-WinImgNativeProcess -Executable $Executable -Arguments $Arguments
        }
        $result=Invoke-DiagnosticsRun $case $runner
        $result.Code | Should -Be 0; $trace.Calls | Should -Be 3
        Assert-DiagnosticsPolicy $trace.Arguments.ToArray() @('100%','100%','90%')
        @($trace.Paths.ToArray() | Select-Object -Unique).Count | Should -Be 3
        $result.Log | Should -Match 'Scale=90%'; $result.Log | Should -Match 'SizeWarnings=0'; $result.Log | Should -Match 'NativeWarnings=0'
        Assert-DiagnosticsJpeg (Join-Path $result.Run 'single.jpeg') 58 43
        Assert-DiagnosticsPreserved $case $result -Image
    }

    It 'T044 keeps the two-sharing-retry budget per file when a later sizing scale also fails' {
        $case=New-DiagnosticsCase 'file-retry-budget'
        $trace=[pscustomobject]@{Calls=0;Paths=New-Object 'Collections.Generic.List[string]';Arguments=New-Object 'Collections.Generic.List[object]'}
        $runner={
            param([string]$Executable,[string[]]$Arguments)
            $trace.Calls++; $candidate=Get-DiagnosticsCandidate $Arguments; $trace.Paths.Add($candidate); $trace.Arguments.Add($Arguments.Clone())
            $scale=$Arguments[[Array]::IndexOf($Arguments,'-resize')+1]
            Assert-DiagnosticsPolicy -Calls (,($Arguments.Clone())) -Scales @($scale) -LiveProfile
            if ($trace.Calls -eq 3) { return Invoke-DiagnosticsChild -CopyFrom $padded -Candidate $candidate }
            return Invoke-DiagnosticsChild -Code 1 -StdErr $sharing -CopyFrom $jpeg -Candidate $candidate
        }
        $result=Invoke-DiagnosticsRun $case $runner
        $result.Code | Should -Be 2; $trace.Calls | Should -Be 4
        Assert-DiagnosticsPolicy $trace.Arguments.ToArray() @('100%','100%','100%','90%')
        @($trace.Paths.ToArray() | Select-Object -Unique).Count | Should -Be 4
        ([regex]::Matches($result.Log,'DelayMs=')).Count | Should -Be 2
        $result.Log | Should -Match 'DelayMs=100'; $result.Log | Should -Match 'DelayMs=200'
        $result.Log | Should -Not -Match 'OK IMG:|WARN IMG:'
        Assert-DiagnosticsPreserved $case $result
    }
}

Describe 'M2-T05 acceptable warnings require validated media (T043)' {
    It 'T043 finalizes an actual converted JPEG with only the known lossy warning, and links its skipped duplicate as warned' {
        $case=New-DiagnosticsCase 'harmless-warning'
        $null=Write-DiagnosticsSource $case.Source 'z-later/single.png' $pngBytes
        $case.Before=Get-DiagnosticsState $case.Source
        $trace=[pscustomobject]@{Calls=0;Paths=New-Object 'Collections.Generic.List[string]'}
        $runner={
            param([string]$Executable,[string[]]$Arguments)
            $trace.Calls++; $candidate=Get-DiagnosticsCandidate $Arguments; $trace.Paths.Add($candidate)
            Assert-DiagnosticsPolicy -Calls (,($Arguments.Clone())) -Scales @('100%') -LiveProfile
            $conversion=Invoke-WinImgNativeProcess -Executable $Executable -Arguments $Arguments
            $conversion.ExitCode | Should -Be 0; $conversion.StdErr | Should -BeNullOrEmpty
            # Production -quiet suppresses the genuine warning proved above.
            # Re-emission is a controlled native outcome, after real conversion.
            return Invoke-DiagnosticsChild -StdErr $harmless
        }
        $result=Invoke-DiagnosticsRun $case $runner
        $result.Code | Should -Be 2; $trace.Calls | Should -Be 1
        $result.Log | Should -Match 'WARN IMG: single.png'; $result.Log | Should -Not -Match 'OK IMG:|ERR IMG:'
        $result.Log | Should -Match 'lossless to lossy JPEG conversion'
        $result.Log | Should -Match 'retained status: ConvertedWithWarning'
        $result.Log | Should -Match 'ConvertedImages=1 CopiedVideos=1 Duplicates=1 Unsupported=0 Errors=0'
        $result.Log | Should -Match 'NativeWarnings=1'; $result.Log | Should -Match 'SizeWarnings=0'
        Assert-DiagnosticsJpeg (Join-Path $result.Run 'single.jpeg')
        Assert-DiagnosticsPreserved $case $result -Image
        $observations.Add([pscustomobject]@{Case='T043 actual conversion then controlled harmless native warning';Calls=$trace.Calls;ApplicationCode=$result.Code;Warning=$harmless})
    }

    It 'T043 rejects <Failure> with a known warning instead of treating any warning as permission to finalize' -ForEach @(
        @{Failure='nonzero exit with valid JPEG';Code=1;Copy='valid'},
        @{Failure='zero exit with truncated JPEG';Code=0;Copy='truncated'}
    ) {
        $case=New-DiagnosticsCase 'unacceptable-warning'
        $trace=[pscustomobject]@{Calls=0}
        $runner={
            param([string]$Executable,[string[]]$Arguments)
            $trace.Calls++; $candidate=Get-DiagnosticsCandidate $Arguments
            $fixture=if ($Copy -eq 'valid') {$jpeg} else {$truncated}
            $native=Invoke-DiagnosticsChild -Code $Code -StdErr $harmless -CopyFrom $fixture -Candidate $candidate
            $native.ExitCode | Should -Be $Code; $native.StdErr | Should -Match 'lossless to lossy JPEG conversion'
            return $native
        }
        $result=Invoke-DiagnosticsRun $case $runner
        $result.Code | Should -Be 2; $trace.Calls | Should -Be 1
        $result.Log | Should -Not -Match 'OK IMG:|WARN IMG:'
        $result.Log | Should -Match 'NativeWarnings=0'; $result.Log | Should -Match 'ConvertedImages=0 CopiedVideos=1 Duplicates=0 Unsupported=0 Errors=1'
        Assert-DiagnosticsPreserved $case $result
    }

    It 'T043 rejects a zero exit with <Stream> diagnostics and a valid JPEG when the complete diagnostic set is unsafe' -ForEach @(
        @{Stream='stdout';Output='controlled unrecognized stdout diagnostic';Error='-'},
        @{Stream='stderr';Output='-';Error='controlled unrecognized stderr diagnostic'},
        @{Stream='mixed warning and ICC error';Output='-';Error="magick.exe: lossless to lossy JPEG conversion 'owned.jpg' @ warning/jpeg.c/WriteJPEGImage_/3195.`nmagick.exe: ColorspaceColorProfileMismatch 'owned.jpg' @ error/profile.c/ProfileImage/1049."}
    ) {
        $case=New-DiagnosticsCase 'unsafe-zero'
        $trace=[pscustomobject]@{Calls=0}
        $runner={
            param([string]$Executable,[string[]]$Arguments)
            $trace.Calls++; $candidate=Get-DiagnosticsCandidate $Arguments
            $native=Invoke-DiagnosticsChild -StdOut $Output -StdErr $Error -CopyFrom $jpeg -Candidate $candidate
            $native.ExitCode | Should -Be 0; $native.StreamsComplete | Should -BeTrue
            ([IO.FileInfo]::new($candidate)).Length | Should -Be $jpegBytes.Length
            return $native
        }
        $result=Invoke-DiagnosticsRun $case $runner
        $result.Code | Should -Be 2; $trace.Calls | Should -Be 1
        $result.Log | Should -Not -Match 'OK IMG:|WARN IMG:'
        $result.Log | Should -Match 'NativeWarnings=0'; $result.Log | Should -Not -Match 'DelayMs='
        Assert-DiagnosticsPreserved $case $result
    }
}

Describe 'M2-T05 bounded output cannot conceal errors (T045)' {
    It 'T045 refuses a decodable JPEG after large native streams, retaining usable details without console flooding' {
        $case=New-DiagnosticsCase 'overflow'
        $trace=[pscustomobject]@{Calls=0;Result=$null}
        $runner={
            param([string]$Executable,[string[]]$Arguments)
            $trace.Calls++; $candidate=Get-DiagnosticsCandidate $Arguments
            $trace.Result=Invoke-DiagnosticsChild -StdOut 'stdout early detail' -StdErr "magick.exe: ColorspaceColorProfileMismatch 'owned.jpg' @ error/profile.c/ProfileImage/1049." -Lines 1024 -OutTail 'stdout useful last detail' -ErrTail $harmless -CopyFrom $jpeg -Candidate $candidate
            return $trace.Result
        }
        $result=Invoke-DiagnosticsRun $case $runner
        $result.Code | Should -Be 2; $trace.Calls | Should -Be 1
        $trace.Result.ExitCode | Should -Be 0; $trace.Result.StdOutTruncated | Should -BeTrue; $trace.Result.StdErrTruncated | Should -BeTrue
        $result.Log | Should -Match 'Category=OutputLimit'; $result.Log | Should -Match 'stdout useful last detail'
        $result.Log | Should -Not -Match 'OK IMG:|WARN IMG:|DelayMs='
        # Keep the native-diagnostic console budget; bound the two new report lines separately.
        $consoleLines=@($result.Text -split '\r?\n')
        $reportLines=@($consoleLines | Where-Object { $_ -match '^(ACCOUNTING Discovered=|Final report outcome: State=)' })
        $reportLines.Count | Should -Be 2
        @($reportLines | Where-Object { $_ -match '^ACCOUNTING ' })[0].Length | Should -BeLessOrEqual 2048
        @($reportLines | Where-Object { $_ -match '^Final report outcome: ' })[0].Length | Should -BeLessOrEqual 1024
        (@($consoleLines | Where-Object { $_ -notmatch '^(ACCOUNTING Discovered=|Final report outcome: State=)' }) -join [Environment]::NewLine).Length | Should -BeLessThan 4096
        $result.Text | Should -Not -Match 'O{64}|E{64}'
        $result.Log.Length | Should -BeLessThan 40000
        Assert-DiagnosticsPreserved $case $result
        $observations.Add([pscustomobject]@{Case='T045 controlled overflow cannot conceal an ICC error';ApplicationCode=$result.Code;NativeResult=$trace.Result;ConsoleCharacters=$result.Text.Length;LogCharacters=$result.Log.Length})
    }
}
