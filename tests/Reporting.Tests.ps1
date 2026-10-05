BeforeAll {
    $repository=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    $scratch=Join-Path $repository '.scratch'
    function Assert-ReportingAncestors([string]$Path) {
        $entry=[IO.DirectoryInfo]::new([IO.Path]::GetFullPath($Path))
        while($entry){if($entry.Exists -and ($entry.Attributes -band [IO.FileAttributes]::ReparsePoint)){throw 'Reporting fixture crosses a reparse point.'};$entry=$entry.Parent}
    }
    Assert-ReportingAncestors $scratch
    foreach($pictures in @([Environment]::GetFolderPath('MyPictures'),(Join-Path $env:USERPROFILE 'Pictures'))){
        if(-not $pictures){continue};$p=[IO.Path]::GetFullPath($pictures).TrimEnd('\','/')
        if($scratch.Equals($p,[StringComparison]::OrdinalIgnoreCase) -or $scratch.StartsWith($p+'\',[StringComparison]::OrdinalIgnoreCase) -or $p.StartsWith($scratch+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Reporting fixtures overlap real Pictures.'}
    }
    & git -C $repository check-ignore --quiet --no-index -- (Join-Path $scratch 'reporting-ignore-probe')
    if($LASTEXITCODE -ne 0){throw 'Reporting fixtures must be ignored.'}
    $ownedRoot=Join-Path $scratch ('M3-T04-reporting-'+[Guid]::NewGuid().ToString('N'))
    if([IO.Directory]::Exists($ownedRoot)){throw 'Reporting fixture ownership collision.'}
    $null=[IO.Directory]::CreateDirectory($ownedRoot)
    [IO.File]::WriteAllText((Join-Path $ownedRoot '.winimg-fixture-root'),'M3-T04 owned synthetic reporting fixtures',[Text.UTF8Encoding]::new($false))
    $observations=New-Object 'Collections.Generic.List[object]'
    $script:reportingStateSerial=0
    if(-not $env:WINIMG_TEST_MAGICK){throw 'Explicit verified ImageMagick required.'}
    $magick=[IO.Path]::GetFullPath($env:WINIMG_TEST_MAGICK);Assert-ReportingAncestors $magick
    function Get-ReportingBindings {
        foreach($relative in @('WinImgNormalizer.ps1','WinImgNormalizer.bat','tests/Reporting.Tests.ps1','tests/Invoke-Tests.ps1','tests/Initialize-TestDependencies.ps1','tests/dependencies.json')){
            [pscustomobject]@{Path=$relative;Sha256=(Get-FileHash -LiteralPath (Join-Path $repository $relative)).Hash.ToLowerInvariant()}
        }
        [pscustomobject]@{Path='verified ImageMagick executable';Sha256=(Get-FileHash -LiteralPath $magick).Hash.ToLowerInvariant()}
    }
    $bindingsBefore=@(Get-ReportingBindings)
    . (Join-Path $repository 'WinImgNormalizer.ps1')
    $timestampImplementation=(Get-Command Set-WinImgOutputTimestampValue).ScriptBlock
    $reportCreateImplementation=(Get-Command New-WinImgRunReportFile).ScriptBlock
    $reportAppendImplementation=(Get-Command Add-WinImgRunReportLine).ScriptBlock
    $reportCloseImplementation=(Get-Command Complete-WinImgRunReportFile).ScriptBlock
    $candidateCleanupImplementation=(Get-Command Remove-WinImgOwnedCandidate).ScriptBlock
    $ancestorImplementation=(Get-Command Assert-WinImgNoReparseAncestors).ScriptBlock
    $writeHostCommand=Get-Command Write-Host -CommandType Cmdlet
    $nativeVersion=Invoke-WinImgNativeProcess -Executable $magick -Arguments @('-version') -TimeoutMilliseconds 15000
    if($nativeVersion.ExitCode -ne 0 -or $nativeVersion.StdErr -or -not $nativeVersion.StreamsComplete){throw 'Actual native version probe failed.'}
    $fixed=[DateTime]::Parse('2018-06-07T08:09:10Z').ToUniversalTime()
    function New-ReportingDirectory([string]$Relative) {
        $path=[IO.Path]::GetFullPath((Join-Path $ownedRoot $Relative))
        if(-not $path.StartsWith($ownedRoot+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Reporting fixture escaped ownership.'}
        Assert-ReportingAncestors $path;$null=[IO.Directory]::CreateDirectory($path);return $path
    }
    function Write-ReportingBytes([string]$Path,[byte[]]$Bytes) {
        if(-not ([IO.Path]::GetFullPath($Path)).StartsWith($ownedRoot+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Reporting bytes escaped ownership.'}
        Assert-ReportingAncestors ([IO.Path]::GetDirectoryName($Path));$null=[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($Path))
        $stream=[IO.FileStream]::new($Path,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
        try{$stream.Write($Bytes,0,$Bytes.Length)}finally{$stream.Dispose()}
        [IO.File]::SetCreationTimeUtc($Path,$fixed);[IO.File]::SetLastWriteTimeUtc($Path,$fixed)
    }
    function Invoke-ReportingMagick([string[]]$Arguments) {
        $literal=@('-define','registry:filename:literal=true')
        $tokens=if($Arguments[0] -eq 'identify'){@('identify')+$literal+$Arguments[1..($Arguments.Length-1)]}else{$literal+$Arguments}
        $native=Invoke-WinImgNativeProcess -Executable $magick -Arguments $tokens -TimeoutMilliseconds 15000
        $observations.Add([pscustomobject]@{Kind='actual synthetic fixture or full-decode native command';Arguments=$tokens;Result=$native})
        if($native.ExitCode -ne 0 -or $native.StdErr -or -not $native.StreamsComplete -or $native.StdOutTruncated -or $native.StdErrTruncated){throw ('Reporting native fixture/decode failed: '+$native.StdErr)}
        return ($native.StdOut -replace "`r`n","`n").TrimEnd("`r","`n")
    }
    function New-ReportingImage([string]$Path,[string]$Coder='PNG',[string]$Colour='#E02020') {
        Write-ReportingBytes $Path ([byte[]]@())
        $tokens=@('-size','48x32',('xc:'+$Colour))
        if($Coder -eq 'JPEG'){$tokens=@('-size','8x8',('xc:'+$Colour),'-quality','1','-interlace','None','-sampling-factor','4:2:0','-strip')}
        $null=Invoke-ReportingMagick ($tokens+@($Coder+':'+(Get-WinImgNativeOutputPath $Path)))
        [IO.File]::SetCreationTimeUtc($Path,$fixed);[IO.File]::SetLastWriteTimeUtc($Path,$fixed)
    }
    function New-ReportingGif([string]$Path) {
        Write-ReportingBytes $Path ([byte[]]@())
        $null=Invoke-ReportingMagick @('-size','48x32','xc:#E02020','-size','48x32','xc:#2040E0','-delay','10','-loop','0','-adjoin',('GIF:'+(Get-WinImgNativeOutputPath $Path)))
        $frames=Invoke-ReportingMagick @('identify','-format',"%m|%w|%h`n",(Get-WinImgNativeOutputPath $Path))
        @($frames.Split("`n")).Count | Should -Be 2
        $frames | Should -Be "GIF|48|32`nGIF|48|32"
        [IO.File]::SetCreationTimeUtc($Path,$fixed);[IO.File]::SetLastWriteTimeUtc($Path,$fixed)
    }
    function New-ReportingVideo([string]$Path,[int]$Length=4097) {
        $bytes=New-Object byte[] $Length;[Random]::new(4319).NextBytes($bytes);Write-ReportingBytes $Path $bytes
    }
    function Copy-ReportingFixture([string]$Source,[string]$Destination) {Write-ReportingBytes $Destination ([IO.File]::ReadAllBytes($Source))}
    function Set-ReportingDirectoryTimes([string]$Root) {
        foreach($directory in @(@(Get-ChildItem -LiteralPath $Root -Directory -Recurse -Force)|Sort-Object FullName -Descending)+@([IO.DirectoryInfo]::new($Root))){
            [IO.Directory]::SetCreationTimeUtc($directory.FullName,$fixed);[IO.Directory]::SetLastWriteTimeUtc($directory.FullName,$fixed)
        }
    }
    function Get-ReportingSourceState([string]$Root) {
        $rows=@(@([IO.DirectoryInfo]::new($Root))+@(Get-ChildItem -LiteralPath $Root -Recurse -Force)|Sort-Object FullName|ForEach-Object{
            $_.Refresh();$isDirectory=$_ -is [IO.DirectoryInfo]
            [pscustomobject]@{Relative=$_.FullName.Substring($Root.Length).Replace('\','/');Directory=$isDirectory;Length=if($isDirectory){$null}else{$_.Length};Sha256=if($isDirectory){$null}else{(Get-FileHash -LiteralPath $_.FullName).Hash.ToLowerInvariant()};CreationTicks=$_.CreationTimeUtc.Ticks;ModifiedTicks=$_.LastWriteTimeUtc.Ticks;Attributes=[int]$_.Attributes}
        })
        foreach($row in $rows){if(-not $row.Directory){$row.Sha256 | Should -Match '^[a-f0-9]{64}$';$row.Length | Should -Not -BeNullOrEmpty}}
        $json=$rows|ConvertTo-Json -Depth 8 -Compress;$script:reportingStateSerial++
        $artifact=Join-Path $ownedRoot ('source-state-{0:d4}.json' -f $script:reportingStateSerial)
        [IO.File]::WriteAllText($artifact,$json,[Text.UTF8Encoding]::new($false))
        $observations.Add([pscustomobject]@{Kind='all regular source bytes/timestamps and directory state';Artifact=[pscustomobject]@{Path=$artifact.Substring($repository.Length+1).Replace('\','/');Sha256=(Get-FileHash -LiteralPath $artifact).Hash.ToLowerInvariant()}})
        return $json
    }
    function Assert-ReportingJpeg([string]$Path,[int]$Width=48,[int]$Height=32) {
        Invoke-ReportingMagick @('identify','+ping','-regard-warnings','-format','%m|%w|%h|%n',(Get-WinImgNativeOutputPath $Path)) | Should -Be ('JPEG|'+$Width+'|'+$Height+'|1')
        $pixel=Invoke-ReportingMagick @((Get-WinImgNativeOutputPath $Path),'-format','%[fx:round(255*p{0,0}.r)]|%[fx:round(255*p{0,0}.g)]|%[fx:round(255*p{0,0}.b)]','info:')
        $rgb=@($pixel.Split('|')|ForEach-Object{[int]$_});[Math]::Abs($rgb[0]-224) | Should -BeLessOrEqual 12;[Math]::Abs($rgb[1]-32) | Should -BeLessOrEqual 12;[Math]::Abs($rgb[2]-32) | Should -BeLessOrEqual 12
    }
    function Read-ReportingText([string]$Value) {
        # Independent consumer for the documented text: / UTF-16 wire format.
        if(-not $Value.StartsWith('text:',[StringComparison]::Ordinal)){throw 'CSV text cell has no fixed inert prefix.'}
        $body=$Value.Substring(5);$decoded=[Text.StringBuilder]::new()
        for($i=0;$i -lt $body.Length;$i++){
            if($body[$i] -ne [char]92){$null=$decoded.Append($body[$i]);continue}
            if($i+1 -lt $body.Length -and $body[$i+1] -eq [char]92){$null=$decoded.Append([char]92);$i++;continue}
            if($i+5 -ge $body.Length -or $body[$i+1] -ne 'u' -or $body.Substring($i+2,4) -notmatch '^[a-fA-F0-9]{4}$'){throw 'CSV text escape is invalid.'}
            $null=$decoded.Append([char][Convert]::ToInt32($body.Substring($i+2,4),16));$i+=5
        }
        return $decoded.ToString()
    }
    function Invoke-ReportingRun([string]$Source,[string]$Parent,[object]$Cap,[scriptblock]$Runner,[object]$Cancellation,[scriptblock]$Stage,[scriptblock]$Copy) {
        $script:reportingCaptured=$null;$script:reportingObserverCalls=0
        $parameters=@{Source=$Source;OutputParent=$Parent;MagickPath=$magick;NativeTemporaryRoot=$ownedRoot;ReportObserver={param($State)$script:reportingCaptured=$State;$script:reportingObserverCalls++}}
        if($null -ne $Cap){$parameters.MaxBytes=$Cap};if($Runner){$parameters.ProcessRunner=$Runner};if($Cancellation){$parameters.CancellationState=$Cancellation};if($Stage){$parameters.RunStageObserver=$Stage};if($Copy){$parameters.CopyProgressObserver=$Copy}
        $savedError=[Console]::Error;$emergency=[IO.StringWriter]::new([Globalization.CultureInfo]::InvariantCulture)
        try{[Console]::SetError($emergency);$output=@(& Invoke-WinImgNormalizer @parameters 6>&1 3>&1 2>&1)}finally{[Console]::SetError($savedError)}
        $codes=@($output|Where-Object{$_ -is [int] -or $_ -is [long]});$codes.Count | Should -Be 1
        $script:reportingObserverCalls | Should -Be 1;$script:reportingCaptured | Should -Not -BeNullOrEmpty
        $runs=@(Get-ChildItem -LiteralPath $Parent -Directory);$runs.Count | Should -Be 1;$run=$runs[0].FullName
        $logs=@(Get-ChildItem -LiteralPath $run -Recurse -File -Force|Where-Object Extension -eq '.log')
        $diskLog=if($logs.Count -eq 1){[IO.File]::ReadAllText($logs[0].FullName)}else{''}
        $state=$script:reportingCaptured
        $csv=if($state.CsvPath -and [IO.File]::Exists($state.CsvPath) -and $state.ReportComplete){@(Import-Csv -LiteralPath $state.CsvPath -Delimiter ',' -Encoding UTF8)}else{@()}
        $result=[pscustomobject]@{Code=$codes[0];Run=$run;State=$state;CsvRows=$csv;Log=$diskLog;Text=($output -join "`n");EmergencyText=$emergency.ToString()};$emergency.Dispose()
        $observations.Add([pscustomobject]@{Kind='actual application and reporting outcome';Result=$result});return $result
    }
    function Assert-ReportingPartition([object]$Result,[int]$Discovered,[hashtable]$Expected) {
        $counts=$Result.State.Summary.Counts;$counts.Discovered | Should -Be $Discovered
        $Result.State.Rows.Count | Should -Be $Discovered
        $paths=@($Result.State.Rows|ForEach-Object{$_.SourceRelativePath});@($paths|Select-Object -Unique).Count | Should -Be $Discovered
        $total=0
        foreach($status in @('Converted','CopiedVideo','SkippedDuplicate','Ignored','Error','Cancelled','NotStarted')){
            $expectedCount=if($Expected.ContainsKey($status)){$Expected[$status]}else{0};$counts.$status | Should -Be $expectedCount
            @($Result.State.Rows|Where-Object Status -eq $status).Count | Should -Be $expectedCount;$total+=$counts.$status
        }
        $total | Should -Be $Discovered
        foreach($row in $Result.State.Rows){
            if($row.Status -in @('Converted','CopiedVideo')){[IO.File]::Exists((Join-Path $Result.Run $row.OutputRelativePath)) | Should -BeTrue;$row.OutputRelativePath | Should -Be $row.PlannedOutputRelativePath}
            else{[string]$row.OutputRelativePath | Should -BeNullOrEmpty}
        }
        $summary=$Result.State.Summary;$summary.ElapsedSeconds | Should -BeGreaterOrEqual 0;$summary.FilesPerSecond | Should -BeGreaterOrEqual 0
        if($summary.ElapsedSeconds -gt 0){[Math]::Abs([double]$summary.FilesPerSecond-(($counts.Converted+$counts.CopiedVideo)/[double]$summary.ElapsedSeconds)) | Should -BeLessOrEqual 0.0001}
    }
    function Assert-ReportingCsv([object]$Result) {
        $Result.State.ReportComplete | Should -BeTrue;$Result.State.ReportWarnings | Should -Be 0
        $bytes=[IO.File]::ReadAllBytes($Result.State.CsvPath);([BitConverter]::ToString($bytes,0,3)) | Should -Be 'EF-BB-BF'
        $fileRows=@($Result.CsvRows|Where-Object{(Read-ReportingText $_.RecordType) -eq 'File'});$fileRows.Count | Should -Be $Result.State.Rows.Count
        foreach($row in $Result.State.Rows){
            $matching=@($fileRows|Where-Object{(Read-ReportingText $_.SourceRelativePath) -ceq $row.SourceRelativePath});$matching.Count | Should -Be 1;$csv=$matching[0]
            foreach($name in @('SourceRelativePath','PlannedOutputRelativePath','OutputRelativePath','Kind','Status','Reason','NamingReason','RetainedSourceRelativePath','RetainedOutputRelativePath','RetainedStatus')){
                $csv.$name | Should -Match '^text:';(Read-ReportingText $csv.$name) | Should -BeExactly ([string]$row.$name)
                $csv.$name | Should -Not -Match '[\x00-\x1f\x7f]'
            }
            if($row.Status -eq 'Converted'){[decimal]::Parse($csv.ImageSavingsBytes,[Globalization.CultureInfo]::InvariantCulture) | Should -Be ([decimal]$row.InputBytes-[decimal]$row.OutputBytes)}
            else{$csv.ImageSavingsBytes | Should -BeNullOrEmpty}
            foreach($name in @('InputBytes','OutputBytes','ImageSavingsBytes','ScalePercent','Width','Height','Attempts','SizeWarning','NativeWarning','TimestampWarning','FramesOmitted','AncillaryWarning','Started')){
                if($csv.$name -ne ''){$csv.$name | Should -Match '^-?[0-9]+(?:\.[0-9]+)?$'}
            }
        }
    }
    function Assert-ReportingRetained([object]$Result,[string]$Source) {
        foreach($row in $Result.State.Rows){
            if($row.Status -notin @('Converted','CopiedVideo')){continue}
            $original=[IO.FileInfo]::new((Join-Path $Source $row.SourceRelativePath));$final=[IO.FileInfo]::new((Join-Path $Result.Run $row.OutputRelativePath))
            $final.CreationTimeUtc.Ticks | Should -Be $original.CreationTimeUtc.Ticks;$final.LastWriteTimeUtc.Ticks | Should -Be $original.LastWriteTimeUtc.Ticks
            $row.OutputBytes | Should -Be $final.Length
            if($row.Status -eq 'CopiedVideo'){(Get-FileHash -LiteralPath $final.FullName).Hash | Should -Be (Get-FileHash -LiteralPath $original.FullName).Hash}
            else{Assert-ReportingJpeg $final.FullName $row.Width $row.Height}
        }
        @(Get-ChildItem -LiteralPath (Join-Path $Result.Run '.WinImgNormalizer/work') -Force).Count | Should -Be 0
    }
    function Get-ReportingAcl([string]$Path) {
        $descriptor=[WinImgReportingAcl]::ReadDescriptor((Get-WinImgNativeOutputPath $Path))
        $raw=[Security.AccessControl.RawSecurityDescriptor]::new($descriptor,0)
        $dacl=New-Object byte[] $raw.DiscretionaryAcl.BinaryLength;$raw.DiscretionaryAcl.GetBinaryForm($dacl,0)
        return [pscustomobject]@{Descriptor=$descriptor;Dacl=[Convert]::ToBase64String($dacl);Control=[int]$raw.ControlFlags;Sddl=$raw.GetSddlForm([Security.AccessControl.AccessControlSections]::All);AclObject=(Get-Acl -LiteralPath $Path)}
    }
    function Set-ReportingAcl([string]$Path,[byte[]]$Descriptor) {
        if(-not ([IO.Path]::GetFullPath($Path)).StartsWith($ownedRoot+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Reporting ACL escaped fixture ownership.'}
        Assert-ReportingAncestors $Path;[WinImgReportingAcl]::WriteDacl((Get-WinImgNativeOutputPath $Path),$Descriptor)
    }
    if(-not ('WinImgReportingAcl' -as [type])){Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
public static class WinImgReportingAcl {
    [DllImport("advapi32.dll",CharSet=CharSet.Unicode,SetLastError=true)]
    static extern bool GetFileSecurityW(string path,uint information,byte[] descriptor,uint length,out uint needed);
    [DllImport("advapi32.dll",CharSet=CharSet.Unicode,SetLastError=true)]
    static extern bool SetFileSecurityW(string path,uint information,byte[] descriptor);
    public static byte[] ReadDescriptor(string path) {
        uint needed;GetFileSecurityW(path,7,null,0,out needed);
        if(needed==0)throw new Win32Exception(Marshal.GetLastWin32Error());
        byte[] descriptor=new byte[needed];
        if(!GetFileSecurityW(path,7,descriptor,needed,out needed))throw new Win32Exception(Marshal.GetLastWin32Error());
        return descriptor;
    }
    public static void WriteDacl(string path,byte[] descriptor) {
        if(!SetFileSecurityW(path,4,descriptor))throw new Win32Exception(Marshal.GetLastWin32Error());
    }
}
'@}
}

AfterAll {
    if($ownedRoot -and $observations){
        $artifacts=@(Get-ChildItem -LiteralPath $ownedRoot -Recurse -File -Force|ForEach-Object{[pscustomobject]@{Path=$_.FullName.Substring($repository.Length+1).Replace('\','/');Sha256=(Get-FileHash -LiteralPath $_.FullName).Hash.ToLowerInvariant()}})
        $bindingsAfter=@(Get-ReportingBindings)
        $evidence=[pscustomobject]@{Task='M3-T04';PowerShell=$PSVersionTable.PSVersion.ToString();Edition=$PSVersionTable.PSEdition;BindingsBefore=$bindingsBefore;BindingsAfter=$bindingsAfter;BindingsUnchanged=(($bindingsBefore|ConvertTo-Json -Compress)-eq($bindingsAfter|ConvertTo-Json -Compress));NativeVersion=$nativeVersion;Observations=$observations.ToArray();Artifacts=$artifacts;Qualification='Owned synthetic reporting fixtures only. Native normalization and report/import observations are distinguished from controlled ancillary failures; no spreadsheet execution claim.'}
        [IO.File]::WriteAllText((Join-Path $ownedRoot 'reporting-observations.json'),($evidence|ConvertTo-Json -Depth 25),[Text.UTF8Encoding]::new($false))
        Write-Host ('Reporting owned artifacts: '+$ownedRoot)
    }
}

Describe 'M3-T04 inert CSV text and invariant generated numbers (T059)' {
    It 'T059 round-trips <Label> through an actual UTF-8 CSV text consumer' -ForEach @(
        @{Label='equals';Value='=1+2'},@{Label='plus';Value='+SUM(1,2)'},@{Label='minus';Value='-1+2'},@{Label='at';Value='@SUM(1,2)'},
        @{Label='leading spaces';Value='  =1+2'},@{Label='tab and line breaks';Value="`t=1+2`r`nnext"},
        @{Label='quotes and separators';Value='=1+2";,@SUM(1,2)'},@{Label='apostrophe and literal prefix';Value="'text:=1+2"},
        @{Label='literal escape spelling';Value='\u000A\n\\end'},
        @{Label='Unicode formatting';Value=([string][char]0xFEFF+[char]0x200B+[char]0x00A0+'=1+2')},
        @{Label='full-width formula introducers';Value=([string][char]0xFF1D+[char]0xFF0B+[char]0xFF0D+[char]0xFF20+'SUM(1,2)')},
        @{Label='emoji and surrogate units';Value=([char]::ConvertFromUtf32(0x1F600)+[char]0xD800+'x'+[char]0xDC00+[char]0xDFFF)},
        @{Label='NUL C1 and Unicode line separators';Value=([string][char]0+[char]0x85+[char]0x2028+[char]0x2029+'@sum')}
    ) {
        $encoded=ConvertTo-WinImgReportText $Value;$encoded | Should -Match '^text:'
        (Read-ReportingText $encoded) | Should -BeExactly $Value
        $encoded | Should -Not -Match '[\x00-\x1f\x7f]'
        $field=ConvertTo-WinImgCsvField -Value $Value
        $directory=New-ReportingDirectory ('cells/'+[Guid]::NewGuid().ToString('N'));$path=Join-Path $directory 'import.csv'
        [IO.File]::WriteAllText($path,("Payload,Number`r`n"+$field+",17`r`n"),[Text.UTF8Encoding]::new($true))
        $rows=@(Import-Csv -LiteralPath $path -Encoding UTF8 -Delimiter ',');$rows.Count | Should -Be 1
        $rows[0].Payload | Should -BeExactly $encoded;(Read-ReportingText $rows[0].Payload) | Should -BeExactly $Value;$rows[0].Number | Should -Be '17'
        $observations.Add([pscustomobject]@{Kind='controlled hostile text with actual Import-Csv round-trip';Label=$Label;Encoded=$encoded;Import='PowerShell Import-Csv -Encoding UTF8 -Delimiter comma';SpreadsheetExecutionClaim=$false})
    }

    It 'T059 emits exact invariant signed generated numbers under a comma-decimal culture' {
        $previous=[Globalization.CultureInfo]::CurrentCulture
        try{
            [Globalization.CultureInfo]::CurrentCulture=[Globalization.CultureInfo]::GetCultureInfo('fr-FR')
            ConvertTo-WinImgCsvField -Value ([decimal]::Parse('-12.25',[Globalization.CultureInfo]::InvariantCulture)) -Number | Should -Be '"-12.25"'
            ConvertTo-WinImgCsvField -Value ([long]::MaxValue) -Number | Should -Be '"9223372036854775807"'
        }finally{[Globalization.CultureInfo]::CurrentCulture=$previous}
    }
}

Describe 'M3-T04 finalized image-only byte accounting (T058)' {
    It 'T058 reports genuine negative savings when a small low-quality JPEG grows after actual normalization' {
        $source=New-ReportingDirectory 'growth/source';$parent=New-ReportingDirectory 'growth/output';$input=Join-Path $source 'growing.jpg'
        New-ReportingImage $input 'JPEG';Set-ReportingDirectoryTimes $source;$before=Get-ReportingSourceState $source
        $inputLength=[IO.FileInfo]::new($input).Length;$result=Invoke-ReportingRun $source $parent
        $result.Code | Should -Be 0;Assert-ReportingPartition $result 1 @{Converted=1};Assert-ReportingCsv $result
        $output=Join-Path $result.Run 'growing.jpeg';Assert-ReportingJpeg $output 8 8;$outputLength=[IO.FileInfo]::new($output).Length
        $outputLength | Should -BeGreaterThan $inputLength
        $result.State.Summary.ImageInputBytes | Should -Be ([decimal]$inputLength);$result.State.Summary.ImageOutputBytes | Should -Be ([decimal]$outputLength)
        $result.State.Summary.ImageSavingsBytes | Should -Be ([decimal]$inputLength-$outputLength);$result.State.Summary.ImageSavingsBytes | Should -BeLessThan 0
        $result.State.Summary.ImageSavingsPercent | Should -BeLessThan 0;$result.State.Summary.VideoBytes | Should -Be 0
        Get-ReportingSourceState $source | Should -BeExactly $before
        $observations.Add([pscustomobject]@{Kind='actual growing low-quality JPEG';InputBytes=$inputLength;OutputBytes=$outputLength;SignedSavings=([decimal]$inputLength-$outputLength);NoPerceptualQualityClaim=$true})
    }

    It 'T058 pairs only finalized images while a large video, duplicate, ignored and failed input remain outside image savings' {
        $source=New-ReportingDirectory 'byte-cohort/source';$parent=New-ReportingDirectory 'byte-cohort/output'
        $bmp=Join-Path $source 'a/large.bmp';New-ReportingImage $bmp 'BMP';Copy-ReportingFixture $bmp (Join-Path $source 'b/large.bmp')
        $jpeg=Join-Path $source 'growing.jpg';New-ReportingImage $jpeg 'JPEG'
        $video=Join-Path $source 'video.mp4';New-ReportingVideo $video (1024*1024)
        Write-ReportingBytes (Join-Path $source 'ignored.dat') (New-Object byte[] 16384)
        Write-ReportingBytes (Join-Path $source 'damaged.png') ([byte[]]@(137,80,78,71,13,10,26,10))
        Set-ReportingDirectoryTimes $source;$before=Get-ReportingSourceState $source;$result=Invoke-ReportingRun $source $parent
        $result.Code | Should -Be 2;Assert-ReportingPartition $result 6 @{Converted=2;CopiedVideo=1;SkippedDuplicate=1;Ignored=1;Error=1};Assert-ReportingCsv $result;Assert-ReportingRetained $result $source
        $input=[decimal]([IO.FileInfo]::new($bmp).Length)+[decimal]([IO.FileInfo]::new($jpeg).Length)
        $output=[decimal]([IO.FileInfo]::new((Join-Path $result.Run 'a/large.jpeg')).Length)+[decimal]([IO.FileInfo]::new((Join-Path $result.Run 'growing.jpeg')).Length)
        [IO.FileInfo]::new((Join-Path $result.Run 'a/large.jpeg')).Length | Should -BeLessThan ([IO.FileInfo]::new($bmp).Length)
        $summary=$result.State.Summary;$summary.ImageInputBytes | Should -Be $input;$summary.ImageOutputBytes | Should -Be $output;$summary.ImageSavingsBytes | Should -Be ($input-$output)
        $expectedPercent=[decimal]100*($input-$output)/$input
        [Math]::Abs([decimal]$summary.ImageSavingsPercent-$expectedPercent) | Should -BeLessOrEqual ([decimal]::Parse('0.000000000001',[Globalization.CultureInfo]::InvariantCulture))
        [double]::IsNaN([double]$summary.ImageSavingsPercent) | Should -BeFalse;[double]::IsInfinity([double]$summary.ImageSavingsPercent) | Should -BeFalse
        $summary.VideoBytes | Should -Be (1024*1024)
        Get-ReportingSourceState $source | Should -BeExactly $before
    }
}

Describe 'M3-T04 balanced discovered-file outcomes and warning attributes (T057)' {
    It 'T057 partitions mixed real media, duplicates, ignored files and a damaged image while warnings remain attributes' {
        $source=New-ReportingDirectory 'mixed/source';$parent=New-ReportingDirectory 'mixed/output'
        $image=Join-Path $source 'a/same.png';New-ReportingImage $image
        Copy-ReportingFixture $image (Join-Path $source 'b/same.png')
        New-ReportingGif (Join-Path $source 'frame.gif')
        $video=Join-Path $source 'v/video.mp4';New-ReportingVideo $video
        Copy-ReportingFixture $video (Join-Path $source 'w/video.mp4')
        Write-ReportingBytes (Join-Path $source 'broken.png') ([byte[]]@(137,80,78,71,13,10,26,10))
        Write-ReportingBytes (Join-Path $source 'ignored.txt') ([byte[]]@(1,2,3))
        $hidden=Join-Path $source 'hidden.bin';Write-ReportingBytes $hidden ([byte[]]@(4,5,6));[IO.File]::SetAttributes($hidden,[IO.FileAttributes]::Hidden)
        Set-ReportingDirectoryTimes $source;$before=Get-ReportingSourceState $source
        Mock Set-WinImgOutputTimestampValue {
            param($Path,$Name,$Value)
            if($Name -eq 'CreationTimeUtc' -and $Path.EndsWith('\a\same.jpeg',[StringComparison]::OrdinalIgnoreCase)){throw [IO.IOException]::new('Controlled creation-time failure after finalization')}
            & $timestampImplementation -Path $Path -Name $Name -Value $Value
        }
        $runner={param($Executable,$Arguments)
            $actual=Invoke-WinImgNativeProcess -Executable $Executable -Arguments $Arguments
            $observations.Add([pscustomobject]@{Kind='actual conversion with controlled benign-warning replay';Result=$actual;Arguments=$Arguments})
            if($actual.ExitCode -eq 0 -and $Arguments -contains 'PNG') {throw 'Unexpected coder token in converter.'}
            if($actual.ExitCode -eq 0 -and (($Arguments -join '|') -match 'source\.png')){
                return [pscustomobject]@{ExitCode=0;DiagnosticOutput=('magick.exe: lossless to lossy JPEG conversion '+[char]96+"owned.jpg' @ warning/jpeg.c/WriteJPEGImage_/3195.")}
            }
            return $actual
        }
        $result=Invoke-ReportingRun $source $parent 1 $runner
        $result.Code | Should -Be 2;Assert-ReportingPartition $result 8 @{Converted=2;CopiedVideo=1;SkippedDuplicate=2;Ignored=2;Error=1};Assert-ReportingCsv $result
        $warnings=$result.State.Summary.WarningCounts;$warnings.SizeWarnings | Should -Be 2;$warnings.NativeWarnings | Should -Be 1;$warnings.TimestampWarnings | Should -Be 1;$warnings.FramesOmitted | Should -Be 1
        $retained=@($result.State.Rows|Where-Object SourceRelativePath -eq 'a\same.png')[0]
        $retained.Status | Should -Be 'Converted';$retained.SizeWarning | Should -BeTrue;$retained.NativeWarning | Should -BeTrue;$retained.TimestampWarning | Should -BeTrue
        $duplicate=@($result.State.Rows|Where-Object SourceRelativePath -eq 'b\same.png')[0]
        $duplicate.RetainedSourceRelativePath | Should -Be 'a\same.png';$duplicate.RetainedOutputRelativePath | Should -Be 'a\same.jpeg';$duplicate.RetainedStatus | Should -Be 'ConvertedWithWarning'
        $videoDuplicate=@($result.State.Rows|Where-Object SourceRelativePath -eq 'w\video.mp4')[0]
        $videoDuplicate.RetainedSourceRelativePath | Should -Be 'v\video.mp4';$videoDuplicate.RetainedStatus | Should -Be 'CopiedVideo'
        foreach($row in @($result.State.Rows|Where-Object Status -eq 'Converted')){$row.Attempts | Should -Be 6;$row.ScalePercent | Should -Be 50;Assert-ReportingJpeg (Join-Path $result.Run $row.OutputRelativePath) 24 16}
        $inputs=[decimal]([IO.FileInfo]::new($image).Length)+[decimal]([IO.FileInfo]::new((Join-Path $source 'frame.gif')).Length)
        $outputs=[decimal]([IO.FileInfo]::new((Join-Path $result.Run 'a/same.jpeg')).Length)+[decimal]([IO.FileInfo]::new((Join-Path $result.Run 'frame.jpeg')).Length)
        $result.State.Summary.ImageInputBytes | Should -Be $inputs;$result.State.Summary.ImageOutputBytes | Should -Be $outputs;$result.State.Summary.ImageSavingsBytes | Should -Be ($inputs-$outputs)
        $result.State.Summary.VideoBytes | Should -Be ([IO.FileInfo]::new($video).Length)
        (Get-FileHash -LiteralPath (Join-Path $result.Run 'v/video.mp4')).Hash | Should -Be (Get-FileHash -LiteralPath $video).Hash
        $result.State.RunState | Should -Be 'Partial';$result.State.ScanComplete | Should -BeTrue
        Get-ReportingSourceState $source | Should -BeExactly $before
    }

    It 'T057 records <Label> as a complete known partition without inventing image savings' -ForEach @(@{Label='an empty tree';Files=0},@{Label='only regular ignored files';Files=2}) {
        $name=if($Files){'ignored-only'}else{'empty'};$source=New-ReportingDirectory ($name+'/source');$parent=New-ReportingDirectory ($name+'/output')
        for($i=0;$i -lt $Files;$i++){Write-ReportingBytes (Join-Path $source ('ignored-'+$i+'.txt')) ([byte[]]@(1,2,3))}
        Set-ReportingDirectoryTimes $source;$before=Get-ReportingSourceState $source;$result=Invoke-ReportingRun $source $parent
        $result.Code | Should -Be 0;Assert-ReportingPartition $result $Files @{Ignored=$Files};Assert-ReportingCsv $result
        $result.State.RunState | Should -Be 'Completed';$result.State.Summary.ImageInputBytes | Should -Be 0;$result.State.Summary.ImageOutputBytes | Should -Be 0;$result.State.Summary.ImageSavingsBytes | Should -Be 0
        $result.State.Summary.ImageSavingsPercent | Should -BeNullOrEmpty;$result.State.Summary.VideoBytes | Should -Be 0
        Get-ReportingSourceState $source | Should -BeExactly $before
    }

    It 'T057 retains finalized <Kind> after an ordinary observer failure without also counting an error' -ForEach @(@{Kind='Image';Extension='png'},@{Kind='Video';Extension='mp4'}) {
        $base='post-final-'+$Kind;$source=New-ReportingDirectory ($base+'/source');$parent=New-ReportingDirectory ($base+'/output')
        $first=Join-Path $source ('01-first.'+$Extension);if($Kind -eq 'Image'){New-ReportingImage $first}else{New-ReportingVideo $first}
        New-ReportingImage (Join-Path $source '02-later.png');Set-ReportingDirectoryTimes $source;$before=Get-ReportingSourceState $source
        $stage={param($Stage,$State,$Relative,$Path)if($Stage -eq 'AfterFinalization' -and $Relative.StartsWith('01-first.')){throw [IO.IOException]::new("Controlled ancillary failure, with `"quotes`" and`r`nlinebreak")}}
        $result=Invoke-ReportingRun $source $parent $null $null $null $stage
        $result.Code | Should -Be 2;$expected=if($Kind -eq 'Image'){@{Converted=2}}else{@{Converted=1;CopiedVideo=1}}
        Assert-ReportingPartition $result 2 $expected;Assert-ReportingCsv $result;Assert-ReportingRetained $result $source
        $result.State.Summary.WarningCounts.AncillaryWarnings | Should -Be 1
        @($result.State.Rows|Where-Object AncillaryWarning).Count | Should -Be 1
        $result.State.Summary.Counts.Error | Should -Be 0;$result.State.RunState | Should -Be 'Partial'
        Get-ReportingSourceState $source | Should -BeExactly $before
    }
}

Describe 'M3-T04 reversible real filename mappings (T059)' {
    It 'T059 preserves literal hostile relative names and retained duplicate links through actual CSV import' {
        $unicode=[string][char]0x00E4+[char]0x00F6+[char]0x00DF+[char]::ConvertFromUtf32(0x1F600)
        $source=New-ReportingDirectory ('names/%d[ab] '+$unicode+'/source');$parent=New-ReportingDirectory 'names/output'
        $names=@('=formula.png','+formula.png','-formula.png','@formula.png','  =space.png',"comma, O'Brien.png",([string][char]0xFEFF+[char]0x200B+[char]0xFF1D+'formula.png'),'collision.png','collision.jpg')
        foreach($name in $names){New-ReportingImage (Join-Path $source $name)}
        $original=Join-Path $source 'a/=same.png';New-ReportingImage $original;Copy-ReportingFixture $original (Join-Path $source 'b/=same.png')
        Set-ReportingDirectoryTimes $source;$before=Get-ReportingSourceState $source;$result=Invoke-ReportingRun $source $parent
        $result.Code | Should -Be 0;Assert-ReportingPartition $result 11 @{Converted=10;SkippedDuplicate=1};Assert-ReportingCsv $result;Assert-ReportingRetained $result $source
        foreach($row in $result.State.Rows){$row.SourceRelativePath | Should -Not -Match '^[A-Za-z]:|^\\\\'}
        foreach($name in @('collision__png.jpeg','collision__jpg.jpeg')){[IO.File]::Exists((Join-Path $result.Run $name)) | Should -BeTrue}
        @($result.State.Rows|Where-Object{$_.SourceRelativePath -in @('collision.png','collision.jpg')}|Where-Object{$_.NamingReason}).Count | Should -Be 2
        $duplicate=@($result.State.Rows|Where-Object Status -eq 'SkippedDuplicate')[0]
        $duplicate.RetainedSourceRelativePath | Should -Be 'a\=same.png';$duplicate.RetainedOutputRelativePath | Should -Be 'a\=same.jpeg'
        Get-ReportingSourceState $source | Should -BeExactly $before
        $observations.Add([pscustomobject]@{Kind='actual literal filenames and CSV importer';ImportRecipe='UTF-8 BOM; comma delimiter; all columns imported as Text; retain text: prefix until explicitly decoding backslash/UTF-16 escapes';ActualExcelExecution=$false})
    }
}

Describe 'M3-T04 degraded CSV and cancellation outcomes (T060)' {
    It 'T060 records a mirror-only output failure as partial without fabricating a known file error or incomplete source scan' {
        $source=New-ReportingDirectory 'mirror-only/source';$parent=New-ReportingDirectory 'mirror-only/output';$null=New-ReportingDirectory 'mirror-only/source/empty'
        Set-ReportingDirectoryTimes $source;$before=Get-ReportingSourceState $source
        Mock Assert-WinImgNoReparseAncestors {param($Path)
            if($Path.Contains('\mirror-only\output\') -and $Path.EndsWith('\empty')){throw [IO.IOException]::new('Controlled destination mirror-only failure')}
            & $ancestorImplementation -Path $Path
        }
        $result=Invoke-ReportingRun $source $parent
        $result.Code | Should -Be 2;Assert-ReportingPartition $result 0 @{};Assert-ReportingCsv $result
        $result.State.ScanComplete | Should -BeTrue;$result.State.IncompleteDirectories | Should -Be 0;$result.State.UninspectableEntries | Should -Be 0
        $result.State.ProcessingWarnings | Should -BeTrue;$result.State.RunState | Should -Be 'Partial';$result.State.Summary.Counts.Error | Should -Be 0
        $result.Log | Should -Match 'Could not mirror directory:'
        Get-ReportingSourceState $source | Should -BeExactly $before
        $observations.Add([pscustomobject]@{Kind='controlled destination directory mirror failure after actual complete source inventory';InventedFileOutcomes=$false;SourceScanStillComplete=$true;ResultCode=$result.Code})
    }

    It 'T060 retains complete CSV and media while an ordinary final console-report failure degrades the run once' {
        $source=New-ReportingDirectory 'final-console/source';$parent=New-ReportingDirectory 'final-console/output'
        New-ReportingImage (Join-Path $source 'image.png');Set-ReportingDirectoryTimes $source;$before=Get-ReportingSourceState $source
        Mock Write-Host {param($Object,$ForegroundColor,$NoNewline)
            if([string]$Object -match '^ACCOUNTING Discovered='){throw [IO.IOException]::new('Controlled final console sink unavailable')}
            $arguments=@{Object=$Object};if($null -ne $ForegroundColor){$arguments.ForegroundColor=$ForegroundColor};if($NoNewline){$arguments.NoNewline=$true}
            & $writeHostCommand @arguments
        }
        $result=Invoke-ReportingRun $source $parent
        $result.Code | Should -Be 2;Assert-ReportingPartition $result 1 @{Converted=1};Assert-ReportingCsv $result;Assert-ReportingRetained $result $source
        $result.State.ConsoleWarnings | Should -Be 1;$result.State.ReportWarnings | Should -Be 0;$result.State.ReportComplete | Should -BeTrue;$result.State.RunState | Should -Be 'Partial'
        $result.EmergencyText | Should -Match 'Console report warning:';$result.State.Summary.Counts.Error | Should -Be 0
        Get-ReportingSourceState $source | Should -BeExactly $before
    }

    It 'T060 keeps media and balances outcomes when report <Phase> fails' -ForEach @(@{Phase='Creation'},@{Phase='Write'},@{Phase='Close'}) {
        $base='report-'+$Phase;$source=New-ReportingDirectory ($base+'/source');$parent=New-ReportingDirectory ($base+'/output')
        New-ReportingImage (Join-Path $source 'image.png');New-ReportingVideo (Join-Path $source 'video.mp4');Set-ReportingDirectoryTimes $source;$before=Get-ReportingSourceState $source
        # A mock executes below the production function's dynamic scope, whose
        # own $phase variable must not replace the requested failure target.
        $script:reportingFailureInjected=$false;$script:reportingFailureTarget=$Phase
        Mock New-WinImgRunReportFile {param($Path,$Header)if($script:reportingFailureTarget -eq 'Creation'){throw [IO.IOException]::new('Controlled report CreateNew failure')}; & $reportCreateImplementation -Path $Path -Header $Header}
        Mock Add-WinImgRunReportLine {param($Writer,$Line)if($script:reportingFailureTarget -eq 'Write' -and -not $script:reportingFailureInjected){$script:reportingFailureInjected=$true;$Writer.Dispose()}; & $reportAppendImplementation -Writer $Writer -Line $Line}
        Mock Complete-WinImgRunReportFile {param($Writer)& $reportCloseImplementation -Writer $Writer;if($script:reportingFailureTarget -eq 'Close'){throw [IO.IOException]::new('Controlled report close failure')}}
        $result=Invoke-ReportingRun $source $parent
        $result.Code | Should -Be 2;Assert-ReportingPartition $result 2 @{Converted=1;CopiedVideo=1};Assert-ReportingRetained $result $source
        $result.State.ReportComplete | Should -BeFalse;$result.State.ReportWarnings | Should -Be 1;$result.State.ReportFailurePhase | Should -Be $Phase;$result.State.RunState | Should -Be 'Partial'
        ($result.Text+$result.EmergencyText+$result.Log) | Should -Match 'CSV report warning:'
        Should -Invoke New-WinImgRunReportFile -Times 1 -Exactly
        Get-ReportingSourceState $source | Should -BeExactly $before
    }

    It 'T060 preserves a foreign report arrival and reports incomplete CSV instead of adopting or overwriting it' {
        $source=New-ReportingDirectory 'report-arrival/source';$parent=New-ReportingDirectory 'report-arrival/output';New-ReportingImage (Join-Path $source 'image.png')
        Set-ReportingDirectoryTimes $source;$before=Get-ReportingSourceState $source;$script:reportingForeignPath=$null;$script:reportingForeignHash=$null
        Mock New-WinImgRunReportFile {param($Path,$Header)
            Write-ReportingBytes $Path ([Text.Encoding]::UTF8.GetBytes('foreign report arrival, =do-not-adopt'))
            $script:reportingForeignPath=$Path;$script:reportingForeignHash=(Get-FileHash -LiteralPath $Path).Hash
            & $reportCreateImplementation -Path $Path -Header $Header
        }
        $result=Invoke-ReportingRun $source $parent
        $result.Code | Should -Be 2;Assert-ReportingPartition $result 1 @{Converted=1};Assert-ReportingRetained $result $source
        $result.State.ReportWarnings | Should -Be 1;$result.State.ReportComplete | Should -BeFalse
        (Get-FileHash -LiteralPath $script:reportingForeignPath).Hash | Should -Be $script:reportingForeignHash
        [IO.File]::ReadAllText($script:reportingForeignPath) | Should -Be 'foreign report arrival, =do-not-adopt'
        [IO.File]::GetCreationTimeUtc($script:reportingForeignPath).Ticks | Should -Be $fixed.Ticks;[IO.File]::GetLastWriteTimeUtc($script:reportingForeignPath).Ticks | Should -Be $fixed.Ticks
        Get-ReportingSourceState $source | Should -BeExactly $before
    }

    It 'T060 emits bounded fallback accounting when both initial disk logging and report creation fail' {
        $source=New-ReportingDirectory 'report-and-log/source';$parent=New-ReportingDirectory 'report-and-log/output'
        New-ReportingImage (Join-Path $source 'image.png');New-ReportingVideo (Join-Path $source 'video.mp4');Set-ReportingDirectoryTimes $source;$before=Get-ReportingSourceState $source
        Mock New-WinImgRunLogFile {throw [IO.IOException]::new('Controlled unavailable log sink')}
        Mock Add-WinImgRunLogLine {throw 'Disabled disk log must not append'}
        Mock New-WinImgRunReportFile {throw [IO.IOException]::new('Controlled unavailable CSV sink')}
        $result=Invoke-ReportingRun $source $parent
        $result.Code | Should -Be 2;Assert-ReportingPartition $result 2 @{Converted=1;CopiedVideo=1};Assert-ReportingRetained $result $source
        $result.State.ReportWarnings | Should -Be 1;$result.State.ReportComplete | Should -BeFalse;$result.State.RunState | Should -Be 'Partial'
        $result.EmergencyText | Should -Match 'ACCOUNTING Discovered=2';$result.EmergencyText | Should -Match 'CSV report warning:'
        $result.EmergencyText | Should -Match 'LogWarnings=1';$result.EmergencyText.Length | Should -BeLessOrEqual 13000
        Should -Invoke Add-WinImgRunLogLine -Times 0 -Exactly
        Get-ReportingSourceState $source | Should -BeExactly $before
    }

    It 'T060 reports completed, current-cancelled, unstarted and ignored files without inventing video savings' {
        $source=New-ReportingDirectory 'interrupted/source';$parent=New-ReportingDirectory 'interrupted/output'
        New-ReportingImage (Join-Path $source '01-first.png');New-ReportingVideo (Join-Path $source '02-current.mp4') (1024*1024)
        New-ReportingImage (Join-Path $source '03-later.png');Write-ReportingBytes (Join-Path $source 'ignored.txt') ([byte[]]@(7,8));Write-ReportingBytes (Join-Path $source 'ignored.bin') ([byte[]]@(9,10))
        Set-ReportingDirectoryTimes $source;$before=Get-ReportingSourceState $source;$state=[WinImgNormalizer.CancellationState]::new();$script:reportingCopied=0
        $copy={param($SourcePath,$CandidatePath,$Copied,$State)
            if($SourcePath.EndsWith('\02-current.mp4') -and $CandidatePath.EndsWith('\video.partial')){$script:reportingCopied=$Copied;$State.Request()}
        }
        $result=Invoke-ReportingRun $source $parent $null $null $state $null $copy
        $result.Code | Should -Be 130;$script:reportingCopied | Should -BeGreaterThan 0;$script:reportingCopied | Should -BeLessThan (1024*1024)
        Assert-ReportingPartition $result 5 @{Converted=1;Cancelled=1;NotStarted=1;Ignored=2};Assert-ReportingCsv $result;Assert-ReportingRetained $result $source
        $result.State.RunState | Should -Be 'Interrupted';$result.State.Summary.VideoBytes | Should -Be 0
        $result.State.Summary.ImageInputBytes | Should -Be ([IO.FileInfo]::new((Join-Path $source '01-first.png')).Length)
        $current=@($result.State.Rows|Where-Object Status -eq 'Cancelled')[0];$current.SourceRelativePath | Should -Be '02-current.mp4';$current.Started | Should -BeTrue
        $later=@($result.State.Rows|Where-Object Status -eq 'NotStarted')[0];$later.SourceRelativePath | Should -Be '03-later.png';$later.Started | Should -BeFalse
        $result.Log | Should -Match 'INTERRUPTED Mode=Cooperative';$result.Log | Should -Match 'ACCOUNTING Discovered=5'
        Get-ReportingSourceState $source | Should -BeExactly $before
        $observations.Add([pscustomobject]@{Kind='controlled cooperative request during actual video copy';CopiedBeforeRequest=$script:reportingCopied;ActualConsoleEvent=$false;ResultCode=$result.Code})
    }

    It 'T060 retains the committed video outcome when cancellation is requested during post-move owned cleanup' {
        $source=New-ReportingDirectory 'video-final-cleanup/source';$parent=New-ReportingDirectory 'video-final-cleanup/output'
        $video=Join-Path $source '01-video.mp4';New-ReportingVideo $video;New-ReportingImage (Join-Path $source '02-later.png')
        Set-ReportingDirectoryTimes $source;$before=Get-ReportingSourceState $source;$state=[WinImgNormalizer.CancellationState]::new();$script:reportingPostMoveObserved=$false
        Mock Remove-WinImgOwnedCandidate {param($CandidatePath,$FileOwned)
            if($CandidatePath.EndsWith('\video.partial') -and -not [IO.File]::Exists($CandidatePath)){
                $script:reportingPostMoveObserved=$true;& $candidateCleanupImplementation -CandidatePath $CandidatePath -FileOwned $FileOwned
                $state.Request();$state.ThrowIfRequested()
            }
            else{& $candidateCleanupImplementation -CandidatePath $CandidatePath -FileOwned $FileOwned}
        }
        $result=Invoke-ReportingRun $source $parent $null $null $state
        $result.Code | Should -Be 130;$script:reportingPostMoveObserved | Should -BeTrue
        Assert-ReportingPartition $result 2 @{CopiedVideo=1;NotStarted=1};Assert-ReportingCsv $result
        $row=@($result.State.Rows|Where-Object Status -eq 'CopiedVideo')[0];$row.SourceRelativePath | Should -Be '01-video.mp4';$row.Attempts | Should -Be 1;$row.OutputBytes | Should -Be ([IO.FileInfo]::new($video).Length)
        (Get-FileHash -LiteralPath (Join-Path $result.Run $row.OutputRelativePath)).Hash | Should -Be (Get-FileHash -LiteralPath $video).Hash
        $result.State.Summary.VideoBytes | Should -Be ([IO.FileInfo]::new($video).Length);$result.State.Summary.ImageInputBytes | Should -Be 0;$result.State.Summary.ImageOutputBytes | Should -Be 0
        $result.State.RunState | Should -Be 'Interrupted';Get-ReportingSourceState $source | Should -BeExactly $before
        $observations.Add([pscustomobject]@{Kind='controlled cancellation after actual video final move';CommittedOutcomeRetained=$true;ActualConsoleEvent=$false;TimestampCompletionClaim=$false})
    }

    It 'T060 reports an actual inaccessible subtree separately from the known regular-file partition' {
        $source=New-ReportingDirectory 'inaccessible/source';$parent=New-ReportingDirectory 'inaccessible/output';$denied=New-ReportingDirectory 'inaccessible/source/denied'
        New-ReportingImage (Join-Path $source 'visible.png');Write-ReportingBytes (Join-Path $source 'ignored.txt') ([byte[]]@(1,2));New-ReportingVideo (Join-Path $denied 'unknown.mp4')
        Set-ReportingDirectoryTimes $source;$before=Get-ReportingSourceState $source;$natural=Get-ReportingAcl $denied
        try{
            $baseline=[Security.AccessControl.RawSecurityDescriptor]::new($natural.Descriptor,0)
            $baseline.SetFlags([Security.AccessControl.ControlFlags]([int]$baseline.ControlFlags -band (-bnot 0x400)))
            $bytes=New-Object byte[] $baseline.BinaryLength;$baseline.GetBinaryForm($bytes,0);Set-ReportingAcl $denied $bytes
            $original=Get-ReportingAcl $denied;($original.Control -band 0x400) | Should -Be 0
            $acl=Get-Acl -LiteralPath $denied;$sid=[Security.Principal.WindowsIdentity]::GetCurrent().User
            $rule=[Security.AccessControl.FileSystemAccessRule]::new($sid,[Security.AccessControl.FileSystemRights]::ListDirectory,[Security.AccessControl.InheritanceFlags]::None,[Security.AccessControl.PropagationFlags]::None,[Security.AccessControl.AccessControlType]::Deny)
            $acl.AddAccessRule($rule);Set-Acl -LiteralPath $denied -AclObject $acl -ErrorAction Stop
            $applied=Get-ReportingAcl $denied;$accessDenied=$false
            try{$null=Get-ChildItem -LiteralPath $denied -Force -ErrorAction Stop}catch{$accessDenied=$true};$accessDenied | Should -BeTrue
            $result=Invoke-ReportingRun $source $parent
            $result.Code | Should -Be 2;Assert-ReportingPartition $result 2 @{Converted=1;Ignored=1};Assert-ReportingCsv $result;Assert-ReportingRetained $result $source
            $result.State.ScanComplete | Should -BeFalse;$result.State.IncompleteDirectories | Should -Be 1;$result.State.UninspectableEntries | Should -Be 0
            $result.State.ScanIssues.Count | Should -Be 1;$result.State.ScanIssues[0].Kind | Should -Be 'DirectoryEnumerationFailed'
            $scan=@($result.CsvRows|Where-Object{(Read-ReportingText $_.RecordType) -eq 'ScanIssue'});$scan.Count | Should -Be 1;(Read-ReportingText $scan[0].SourceRelativePath) | Should -Be 'denied'
            @($result.State.Rows|Where-Object SourceRelativePath -eq 'denied\unknown.mp4').Count | Should -Be 0
            $after=Get-ReportingAcl $denied;[Convert]::ToBase64String($after.Descriptor) | Should -Be ([Convert]::ToBase64String($applied.Descriptor))
            Set-ReportingAcl $denied $original.Descriptor;$restored=Get-ReportingAcl $denied
            [Convert]::ToBase64String($restored.Descriptor) | Should -Be ([Convert]::ToBase64String($original.Descriptor));$restored.Dacl | Should -Be $original.Dacl;$restored.Control | Should -Be $original.Control;$restored.Sddl | Should -Be $original.Sddl
            $observations.Add([pscustomobject]@{Kind='actual ACL-denied subtree with exact fixture restoration';KnownRegularFiles=2;UnknownSubtreeFilesNotInventoried=$true;OriginalControl=$original.Control;RestoredControl=$restored.Control;DescriptorEqual=$true;DaclEqual=$true})
        }finally{if($natural.Control -band 0x400){Set-Acl -LiteralPath $denied -AclObject $natural.AclObject -ErrorAction Stop}else{Set-ReportingAcl $denied $natural.Descriptor}}
        $restoredNatural=Get-ReportingAcl $denied;[Convert]::ToBase64String($restoredNatural.Descriptor) | Should -Be ([Convert]::ToBase64String($natural.Descriptor))
        Get-ReportingSourceState $source | Should -BeExactly $before
    }
}
