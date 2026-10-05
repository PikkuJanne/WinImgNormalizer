BeforeAll {
    $repository=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    $scratch=Join-Path $repository '.scratch'
    function Assert-RecoveryAncestors([string]$Path){
        $item=[IO.DirectoryInfo]::new([IO.Path]::GetFullPath($Path))
        while($item){if($item.Exists -and ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)){throw 'Recovery fixture crosses a reparse point.'};$item=$item.Parent}
    }
    Assert-RecoveryAncestors $scratch
    foreach($pictures in @([Environment]::GetFolderPath('MyPictures'),(Join-Path $env:USERPROFILE 'Pictures'))){
        if(-not $pictures){continue};$p=[IO.Path]::GetFullPath($pictures).TrimEnd('\','/')
        if($scratch.Equals($p,[StringComparison]::OrdinalIgnoreCase) -or $scratch.StartsWith($p+'\',[StringComparison]::OrdinalIgnoreCase) -or $p.StartsWith($scratch+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Recovery fixture overlaps real Pictures.'}
    }
    & git -C $repository check-ignore --quiet --no-index -- (Join-Path $scratch 'recovery-ignore-probe')
    if($LASTEXITCODE -ne 0){throw 'Recovery fixtures must be ignored.'}
    $ownedRoot=Join-Path $scratch ('M3-T01-recovery-'+[Guid]::NewGuid().ToString('N'))
    if([IO.Directory]::Exists($ownedRoot)){throw 'Recovery ownership collision.'}
    $null=[IO.Directory]::CreateDirectory($ownedRoot)
    [IO.File]::WriteAllText((Join-Path $ownedRoot '.winimg-fixture-root'),'M3-T01 owned synthetic recovery fixtures')
    $observations=New-Object 'Collections.Generic.List[object]'
    $script:recoveryStateSerial=0
    function Get-RecoveryBindings {
        $paths=@('WinImgNormalizer.ps1','WinImgNormalizer.bat','tests/RecoveryReporting.Tests.ps1','tests/Invoke-Tests.ps1','tests/Initialize-TestDependencies.ps1','tests/dependencies.json')
        foreach($relative in $paths){$p=Join-Path $repository $relative;[pscustomobject]@{Path=$relative;Sha256=(Get-FileHash -LiteralPath $p).Hash.ToLowerInvariant()}}
        [pscustomobject]@{Path='verified ImageMagick executable';Sha256=(Get-FileHash -LiteralPath $env:WINIMG_TEST_MAGICK).Hash.ToLowerInvariant()}
    }
    $bindingsBefore=@(Get-RecoveryBindings)
    $environment=[pscustomobject]@{PowerShell=$PSVersionTable.PSVersion.ToString();Edition=$PSVersionTable.PSEdition;Culture=[Globalization.CultureInfo]::CurrentCulture.Name;ArchitectureBits=[IntPtr]::Size*8}

    if(-not $env:WINIMG_TEST_MAGICK){throw 'Explicit verified ImageMagick required.'}
    $magick=[IO.Path]::GetFullPath($env:WINIMG_TEST_MAGICK);Assert-RecoveryAncestors $magick
    . (Join-Path $repository 'WinImgNormalizer.ps1')
    $timestampImplementation=(Get-Command Set-WinImgOutputTimestampValue).ScriptBlock
    $appendImplementation=(Get-Command Add-WinImgRunLogLine).ScriptBlock
    $nativeVersion=Invoke-WinImgNativeProcess -Executable $magick -Arguments @('-version')
    if($nativeVersion.ExitCode -ne 0 -or $nativeVersion.StdErr -or -not $nativeVersion.StreamsComplete){throw 'Actual native version probe failed.'}
    $fixed=[DateTime]::Parse('2019-06-07T08:09:10Z').ToUniversalTime()
    function New-RecoveryDirectory([string]$Relative){
        $p=[IO.Path]::GetFullPath((Join-Path $ownedRoot $Relative))
        if(-not $p.StartsWith($ownedRoot+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Recovery fixture escaped ownership.'}
        Assert-RecoveryAncestors $p;$null=[IO.Directory]::CreateDirectory($p);return $p
    }
    function Invoke-RecoveryMagick([string[]]$Arguments){
        $literal=@('-define','registry:filename:literal=true')
        $args=if($Arguments[0] -eq 'identify'){@('identify')+$literal+$Arguments[1..($Arguments.Length-1)]}else{$literal+$Arguments}
        $r=Invoke-WinImgNativeProcess -Executable $magick -Arguments $args
        $observations.Add([pscustomobject]@{Kind='actual native fixture/validation';Arguments=$args;Result=$r})
        if($r.ExitCode -ne 0 -or $r.StdErr -or -not $r.StreamsComplete -or $r.StdOutTruncated -or $r.StdErrTruncated){throw ('Recovery native probe failed: '+$r.StdErr)}
        return ($r.StdOut -replace "`r`n","`n").TrimEnd("`r","`n")
    }
    function New-RecoveryImage([string]$Path,[string]$Colour='#E02020'){
        $null=[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($Path))
        $null=Invoke-RecoveryMagick @('-size','48x32',('xc:'+$Colour),('PNG:'+(Get-WinImgNativeOutputPath $Path)))
        [IO.File]::SetCreationTimeUtc($Path,$fixed);[IO.File]::SetLastWriteTimeUtc($Path,$fixed)
    }
    function New-RecoveryVideo([string]$Path){
        $null=[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($Path))
        $bytes=New-Object byte[] 4097;for($i=0;$i -lt $bytes.Length;$i++){$bytes[$i]=[byte](($i*29+7)%256)}
        [IO.File]::WriteAllBytes($Path,$bytes);[IO.File]::SetCreationTimeUtc($Path,$fixed);[IO.File]::SetLastWriteTimeUtc($Path,$fixed)
    }
    function Get-RecoveryState([string]$Root){
        # Inspect entries but never recurse into a fixture link or junction.
        $files=New-Object 'Collections.Generic.List[object]';$directories=New-Object 'Collections.Generic.List[object]'
        $pending=New-Object 'Collections.Generic.Stack[string]';$pending.Push($Root)
        while($pending.Count){$p=$pending.Pop();$dir=[IO.DirectoryInfo]::new($p);$directories.Add([pscustomobject]@{Path=$p.Substring($Root.Length);Creation=$dir.CreationTimeUtc.Ticks;Modified=$dir.LastWriteTimeUtc.Ticks;Attributes=[int]$dir.Attributes})
            foreach($e in @(Get-ChildItem -LiteralPath $p -Force -ErrorAction Stop)){
                $link=($e.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0
                if($e.PSIsContainer -and -not $link){$pending.Push($e.FullName)}
                else{$files.Add([pscustomobject]@{Path=$e.FullName.Substring($Root.Length);Creation=$e.CreationTimeUtc.Ticks;Modified=$e.LastWriteTimeUtc.Ticks;Attributes=[int]$e.Attributes;Length=if($link){$null}else{$e.Length};Sha256=if($link){$null}else{(Get-FileHash -LiteralPath $e.FullName).Hash.ToLowerInvariant()}})}
            }
        }
        $json=([pscustomobject]@{Files=@($files.ToArray()|Sort-Object Path);Directories=@($directories.ToArray()|Sort-Object Path)}|ConvertTo-Json -Depth 7 -Compress)
        $script:recoveryStateSerial++
        $statePath=Join-Path $ownedRoot ('source-state-{0:d4}.json' -f $script:recoveryStateSerial)
        [IO.File]::WriteAllText($statePath,$json,[Text.UTF8Encoding]::new($false))
        $observations.Add([pscustomobject]@{Kind='regular files/directories and untraversed reparse entries';Root=$Root;Artifact=[pscustomobject]@{Path=$statePath.Substring($repository.Length+1).Replace('\','/');Sha256=(Get-FileHash -LiteralPath $statePath).Hash.ToLowerInvariant()}})
        return $json
    }
    function Invoke-RecoveryRun([string]$Source,[string]$Parent,[scriptblock]$Runner){
        $parameters=@{Source=$Source;OutputParent=$Parent;MagickPath=$magick};if($Runner){$parameters.ProcessRunner=$Runner}
        $saved=[Console]::Error;$errorText=[IO.StringWriter]::new([Globalization.CultureInfo]::InvariantCulture)
        try{[Console]::SetError($errorText);$output=@(& Invoke-WinImgNormalizer @parameters 6>&1 3>&1 2>&1)}finally{[Console]::SetError($saved)}
        $codes=@($output|Where-Object{$_ -is [int] -or $_ -is [long]});$codes.Count|Should -Be 1
        $runs=@(Get-ChildItem -LiteralPath $Parent -Directory);$runs.Count|Should -BeLessOrEqual 1
        $run=if($runs.Count){$runs[0].FullName}else{$null}
        $logs=if($run){@(Get-ChildItem -LiteralPath $run -Recurse -Force -File|Where-Object Extension -eq '.log')}else{@()}
        $result=[pscustomobject]@{Code=$codes[0];Text=($output -join "`n");EmergencyText=$errorText.ToString();Run=$run;LogPath=if($logs.Count){$logs[0].FullName}else{$null}}
        $errorText.Dispose();$observations.Add([pscustomobject]@{Kind='application outcome';Source=$Source;Parent=$Parent;Result=$result});return $result
    }
    function Assert-RecoveryMedia([object]$Result,[string[]]$Jpegs,[string[]]$Videos,[string]$Source){
        foreach($relative in $Jpegs){
            $p=Join-Path $Result.Run $relative
            Invoke-RecoveryMagick @('identify','+ping','-regard-warnings','-format','%m|%w|%h|%n',(Get-WinImgNativeOutputPath $p))|Should -Be 'JPEG|48|32|1'
            $pixel=Invoke-RecoveryMagick @((Get-WinImgNativeOutputPath $p),'-format','%[fx:round(255*p{24,16}.r)]|%[fx:round(255*p{24,16}.g)]|%[fx:round(255*p{24,16}.b)]','info:')
            $rgb=@($pixel.Split('|')|ForEach-Object{[int]$_});[Math]::Abs($rgb[0]-224)|Should -BeLessOrEqual 12;[Math]::Abs($rgb[1]-32)|Should -BeLessOrEqual 12;[Math]::Abs($rgb[2]-32)|Should -BeLessOrEqual 12
        }
        foreach($relative in $Videos){(Get-FileHash -LiteralPath (Join-Path $Result.Run $relative)).Hash|Should -Be (Get-FileHash -LiteralPath (Join-Path $Source $relative)).Hash}
        @(Get-ChildItem -LiteralPath (Join-Path $Result.Run '.WinImgNormalizer/work') -Force).Count|Should -Be 0
    }
    function Assert-RecoverySummary([string]$Text,[int]$Images,[int]$Videos,[int]$Duplicates=0,[int]$TimestampWarnings=0,[bool]$ScanComplete=$true,[int]$Directories=0,[int]$Links=0,[int]$LogWarnings=0){
        $Text|Should -Match ('SUMMARY ConvertedImages='+$Images+' CopiedVideos='+$Videos+' Duplicates='+$Duplicates+' Unsupported=0 Errors=0 SizeWarnings=0 NativeWarnings=0')
        $Text|Should -Match ('TimestampWarnings='+$TimestampWarnings+';?\s+ScanComplete='+$ScanComplete+';?\s+IncompleteDirectories='+$Directories+';?\s+UninspectableEntries=0;?\s+SkippedLinks='+$Links+';?\s+LogWarnings='+$LogWarnings+';?\s+FallbackDropped=\d+')
    }
    function Get-RecoveryByteHash([byte[]]$Bytes){
        $sha=[Security.Cryptography.SHA256]::Create()
        try{return ([BitConverter]::ToString($sha.ComputeHash($Bytes))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
    }
    function Get-RecoveryAclState([string]$Path,[string]$Label){
        Assert-RecoveryAncestors $Path
        $descriptor=[WinImgRecoveryFixtures]::ReadDescriptor((Get-WinImgNativeOutputPath $Path))
        $raw=[Security.AccessControl.RawSecurityDescriptor]::new($descriptor,0)
        $dacl=New-Object byte[] $raw.DiscretionaryAcl.BinaryLength;$raw.DiscretionaryAcl.GetBinaryForm($dacl,0)
        $artifact=Join-Path $ownedRoot ($Label+'-descriptor.bin')
        if([IO.File]::Exists($artifact)){throw 'Recovery ACL evidence collision.'}
        [IO.File]::WriteAllBytes($artifact,$descriptor)
        $acl=Get-Acl -LiteralPath $Path -ErrorAction Stop
        return [pscustomobject]@{AclObject=$acl;DescriptorBytes=$descriptor;DescriptorSha256=(Get-RecoveryByteHash $descriptor);DaclSha256=(Get-RecoveryByteHash $dacl);Control=[int]$raw.ControlFlags;AutoInherited=([int]$raw.ControlFlags -band 0x400) -ne 0;Sddl=$acl.GetSecurityDescriptorSddlForm([Security.AccessControl.AccessControlSections]::All);Artifact=[pscustomobject]@{Path=$artifact.Substring($repository.Length+1).Replace('\','/');Sha256=(Get-FileHash -LiteralPath $artifact).Hash.ToLowerInvariant()}}
    }
    function Set-RecoveryFixtureDacl([string]$Path,[byte[]]$Descriptor){
        if(-not ([IO.Path]::GetFullPath($Path)).StartsWith($ownedRoot+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'ACL fixture escaped ownership.'}
        Assert-RecoveryAncestors $Path
        # Fixture only: DACL_SECURITY_INFORMATION does not write owner/group/SACL
        # or impose the auto-inheritance algorithm on existing child files.
        [WinImgRecoveryFixtures]::WriteDacl((Get-WinImgNativeOutputPath $Path),$Descriptor)
    }
    function Restore-RecoveryFixtureDacl([string]$Path,[object]$Original){
        # Set-Acl imposes current inheritance and can add D:AI. Preserve an
        # original legacy no-AI descriptor with the selected native DACL write.
        if($Original.AutoInherited){Set-Acl -LiteralPath $Path -AclObject $Original.AclObject -ErrorAction Stop}
        else{Set-RecoveryFixtureDacl $Path $Original.DescriptorBytes}
    }
    if(-not ('WinImgRecoveryFixtures' -as [type])){Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Runtime.InteropServices;
public static class WinImgRecoveryFixtures {
    [DllImport("kernel32.dll",CharSet=CharSet.Unicode,SetLastError=true)]
    [return:MarshalAs(UnmanagedType.I1)]
    public static extern bool CreateSymbolicLinkW(string link,string target,int flags);
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
public sealed class WinImgRecoveryBrokenWriter : TextWriter {
    public override Encoding Encoding { get { return Encoding.UTF8; } }
    public override void Write(char value) { throw new IOException("Controlled unavailable emergency stderr sink"); }
    public override void WriteLine(string value) { throw new IOException("Controlled unavailable emergency stderr sink"); }
}
'@}
}

AfterAll {
    if($ownedRoot -and $null -ne $observations){
        # Prune links and junctions when hashing retained owned evidence. Never
        # read a linked external target, and keep a setup failure observable.
        $artifacts=New-Object 'Collections.Generic.List[object]'
        $pending=New-Object 'Collections.Generic.Stack[string]';$pending.Push($ownedRoot)
        while($pending.Count){
            $dir=$pending.Pop()
            try{foreach($entry in @(Get-ChildItem -LiteralPath $dir -Force -ErrorAction Stop)){
                if($entry.Attributes -band [IO.FileAttributes]::ReparsePoint){continue}
                if($entry.PSIsContainer){$pending.Push($entry.FullName)}else{$artifacts.Add([pscustomobject]@{Path=$entry.FullName.Substring($repository.Length+1).Replace('\','/');Sha256=(Get-FileHash -LiteralPath $entry.FullName).Hash.ToLowerInvariant();Length=$entry.Length})}
            }}catch{$observations.Add([pscustomobject]@{Kind='evidence enumeration failure';Directory=$dir;Reason=$_.Exception.Message})}
        }
        $bindingsAfter=@(Get-RecoveryBindings)
        $evidence=[pscustomobject]@{Task='M3-T01';Environment=$environment;NativeVersion=$nativeVersion;BindingsBefore=$bindingsBefore;BindingsAfter=$bindingsAfter;BindingsUnchanged=(($bindingsBefore|ConvertTo-Json -Compress) -eq ($bindingsAfter|ConvertTo-Json -Compress));Observations=$observations.ToArray();Artifacts=$artifacts.ToArray()}
        [IO.File]::WriteAllText((Join-Path $ownedRoot 'recovery-observations.json'),($evidence|ConvertTo-Json -Depth 18),[Text.UTF8Encoding]::new($false))
        Write-Host ('Recovery reporting owned artifacts: '+$ownedRoot)
    }
}

Describe 'M3-T01 inaccessible traversal and ancillary reporting (T047-T050)' {
    It 'T047 continues real native image/video siblings when an owned directory ACL actually denies enumeration' -ForEach @(@{Eligible=$true},@{Eligible=$false}) {
        $suffix=if($Eligible){'siblings'}else{'denied-only'};$source=New-RecoveryDirectory ('acl-'+$suffix+'/source');$parent=New-RecoveryDirectory ('acl-'+$suffix+'/output')
        $denied=Join-Path $source 'denied[1]%d';$null=[IO.Directory]::CreateDirectory($denied);New-RecoveryImage (Join-Path $denied 'unknown.png')
        if($Eligible){New-RecoveryImage (Join-Path $source 'accessible/image.png');New-RecoveryVideo (Join-Path $source 'z sibling/video.mp4')}
        $natural=Get-RecoveryAclState $denied ('acl-'+$suffix+'-natural')
        if(-not $Eligible){
            # Establish the hosted legacy no-AI descriptor state explicitly,
            # before the source baseline, without normalizing comparison text.
            Set-RecoveryFixtureDacl $denied $natural.DescriptorBytes
        }
        $original=Get-RecoveryAclState $denied ('acl-'+$suffix+'-original')
        if(-not $Eligible){$original.AutoInherited|Should -BeFalse -Because 'The denied-only fixture must exercise original no-auto-inheritance restoration.'}
        $before=Get-RecoveryState $source
        $changed=[Security.AccessControl.DirectorySecurity]::new();$changed.SetSecurityDescriptorBinaryForm($original.AclObject.GetSecurityDescriptorBinaryForm())
        $sid=[Security.Principal.WindowsIdentity]::GetCurrent().User
        $rule=[Security.AccessControl.FileSystemAccessRule]::new($sid,[Security.AccessControl.FileSystemRights]::ListDirectory,[Security.AccessControl.InheritanceFlags]::None,[Security.AccessControl.PropagationFlags]::None,[Security.AccessControl.AccessControlType]::Deny)
        $null=$changed.AddAccessRule($rule);$deniedObserved=$false;$result=$null
        try{
            Set-Acl -LiteralPath $denied -AclObject $changed -ErrorAction Stop
            $applied=Get-RecoveryAclState $denied ('acl-'+$suffix+'-applied')
            try{$null=[IO.Directory]::GetFileSystemEntries($denied)}catch{if($_.Exception.InnerException -is [UnauthorizedAccessException] -or $_.Exception -is [UnauthorizedAccessException]){$deniedObserved=$true}else{throw}}
            $deniedObserved|Should -BeTrue -Because 'A mandatory real ACL test must establish actual denied access, never substitute a mock or skip.'
            $result=Invoke-RecoveryRun $source $parent;$result.Code|Should -Be 2
            $combined=$result.Text+"`n"+$result.EmergencyText
            $combined|Should -Match 'Incomplete source scan';$combined|Should -Match ([regex]::Escape('denied[1]%d'))
            Assert-RecoverySummary $combined $(if($Eligible){1}else{0}) $(if($Eligible){1}else{0}) -ScanComplete $false -Directories 1
            $afterApplication=Get-RecoveryAclState $denied ('acl-'+$suffix+'-after-application')
            $afterApplication.Sddl|Should -Be $applied.Sddl;$afterApplication.Control|Should -Be $applied.Control
            $afterApplication.DescriptorSha256|Should -Be $applied.DescriptorSha256;$afterApplication.DaclSha256|Should -Be $applied.DaclSha256
            [IO.File]::Exists((Join-Path $result.Run 'denied[1]%d/unknown.jpeg'))|Should -BeFalse
            if($Eligible){Assert-RecoveryMedia $result @('accessible/image.jpeg') @('z sibling/video.mp4') $source}
        }finally{Restore-RecoveryFixtureDacl $denied $original}
        $restored=Get-RecoveryAclState $denied ('acl-'+$suffix+'-restored')
        $restored.Sddl|Should -Be $original.Sddl;$restored.Control|Should -Be $original.Control
        $restored.DescriptorSha256|Should -Be $original.DescriptorSha256;$restored.DaclSha256|Should -Be $original.DaclSha256
        Get-RecoveryState $source|Should -Be $before
        $observations.Add([pscustomobject]@{Kind='actual ACL enumeration denial';EligibleSiblings=$Eligible;DeniedAccessEstablished=$deniedObserved;AclRestoredExactly=$true;OriginalAutoInherited=$original.AutoInherited;OriginalControl=$original.Control;RestoredControl=$restored.Control;OriginalDescriptorSha256=$original.DescriptorSha256;RestoredDescriptorSha256=$restored.DescriptorSha256;OriginalDaclSha256=$original.DaclSha256;RestoredDaclSha256=$restored.DaclSha256;DescriptorArtifacts=@($natural.Artifact,$original.Artifact,$applied.Artifact,$afterApplication.Artifact,$restored.Artifact);SourceStateBefore=$before;SourceStateAfter=(Get-RecoveryState $source)})
    }

    It 'T048 skips a real junction loop/outside junction and available file symlink while preserving native siblings and outside content' {
        $source=New-RecoveryDirectory 'links/source';$parent=New-RecoveryDirectory 'links/output';$outside=New-RecoveryDirectory 'links/outside'
        New-RecoveryImage (Join-Path $source 'inside.png');New-RecoveryVideo (Join-Path $source 'inside.mp4');New-RecoveryImage (Join-Path $outside 'outside.png') '#2020E0'
        New-Item -ItemType Junction -Path (Join-Path $source 'loop') -Target $source -ErrorAction Stop|Out-Null
        New-Item -ItemType Junction -Path (Join-Path $source 'outside-directory') -Target $outside -ErrorAction Stop|Out-Null
        $link=Join-Path $source 'outside-file.png';$target=Join-Path $outside 'outside.png'
        $created=[WinImgRecoveryFixtures]::CreateSymbolicLinkW($link,$target,2);$nativeError=if($created){0}else{[Runtime.InteropServices.Marshal]::GetLastWin32Error()}
        if(-not $created -and $nativeError -eq 87){$created=[WinImgRecoveryFixtures]::CreateSymbolicLinkW($link,$target,0);$nativeError=if($created){0}else{[Runtime.InteropServices.Marshal]::GetLastWin32Error()}}
        $observations.Add([pscustomobject]@{Kind='actual file-symlink capability';Created=$created;NativeError=$nativeError})
        # Junction traversal is mandatory. File symlink privilege is separately
        # recorded; absence never masquerades as actual file-link coverage.
        if($created){([IO.File]::GetAttributes($link) -band [IO.FileAttributes]::ReparsePoint)|Should -Not -Be 0}
        $before=Get-RecoveryState $source;$outsideBefore=Get-RecoveryState $outside
        $r=Invoke-RecoveryRun $source $parent;$r.Code|Should -Be 2
        Assert-RecoverySummary ($r.Text+$r.EmergencyText) 1 1 -Links $(if($created){3}else{2})
        foreach($name in (@('loop','outside-directory')+$(if($created){@('outside-file.png')}else{@()}))){($r.Text+$r.EmergencyText)|Should -Match ([regex]::Escape($name));[IO.File]::Exists((Join-Path $r.Run $name))|Should -BeFalse;[IO.Directory]::Exists((Join-Path $r.Run $name))|Should -BeFalse}
        Assert-RecoveryMedia $r @('inside.jpeg') @('inside.mp4') $source
        Get-RecoveryState $source|Should -Be $before;Get-RecoveryState $outside|Should -Be $outsideBefore
    }

    It 'T048 deliberately rejects an actual source-root junction alias before output creation' {
        $physical=New-RecoveryDirectory 'root-alias/physical';$parent=New-RecoveryDirectory 'root-alias/output';$link=Join-Path (New-RecoveryDirectory 'root-alias') 'alias'
        New-RecoveryImage (Join-Path $physical 'image.png');$before=Get-RecoveryState $physical
        New-Item -ItemType Junction -Path $link -Target $physical -ErrorAction Stop|Out-Null
        $r=Invoke-RecoveryRun $link $parent;$r.Code|Should -Be 1;$r.Text|Should -Match '(?i)(linked|alias|reparse|actual directory)';$r.Run|Should -BeNullOrEmpty
        @(Get-ChildItem -LiteralPath $parent -Force).Count|Should -Be 0;Get-RecoveryState $physical|Should -Be $before
    }

    It 'T049 visibly degrades after a controlled initial log creation failure and retains actual media or the honest empty outcome' -ForEach @(@{Eligible=$true},@{Eligible=$false}) {
        $source=New-RecoveryDirectory ('initial-log-'+$Eligible+'/source');$parent=New-RecoveryDirectory ('initial-log-'+$Eligible+'/output')
        if($Eligible){New-RecoveryImage (Join-Path $source 'image.png');New-RecoveryVideo (Join-Path $source 'video.mp4')}
        $before=Get-RecoveryState $source
        Mock New-WinImgRunLogFile {throw [IO.IOException]::new('Controlled initial run-log creation failure')}
        Mock Add-WinImgRunLogLine {throw 'A disabled initial log sink must never be appended.'}
        $r=Invoke-RecoveryRun $source $parent;$r.Code|Should -Be 2
        $r.EmergencyText|Should -Match '(?i)log';$r.EmergencyText|Should -Match 'Controlled initial run-log creation failure'
        Assert-RecoverySummary ($r.Text+$r.EmergencyText) $(if($Eligible){1}else{0}) $(if($Eligible){1}else{0}) -LogWarnings 1
        Should -Invoke New-WinImgRunLogFile -Times 1 -Exactly;Should -Invoke Add-WinImgRunLogLine -Times 0 -Exactly
        if($Eligible){Assert-RecoveryMedia $r @('image.jpeg') @('video.mp4') $source}
        Get-RecoveryState $source|Should -Be $before
    }

    It 'T049 a real exclusively held log fails its append while actual native conversion and later video copying complete with one degraded outcome' {
        $source=New-RecoveryDirectory 'locked-log/source';$parent=New-RecoveryDirectory 'locked-log/output';New-RecoveryImage (Join-Path $source 'image.png');New-RecoveryVideo (Join-Path $source 'video.mp4');$before=Get-RecoveryState $source
        $script:recoveryLogLock=$null
        $runner={param($Executable,$Arguments)
            if(-not $script:recoveryLogLock){$logs=@(Get-ChildItem -LiteralPath $parent -Recurse -Force -File|Where-Object Extension -eq '.log');$logs.Count|Should -Be 1;$script:recoveryLogLock=[IO.FileStream]::new($logs[0].FullName,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)}
            # Controlled log interference only; delegate unchanged production
            # argv to the actual native process and return its actual outcome.
            Invoke-WinImgNativeProcess -Executable $Executable -Arguments $Arguments
        }
        try{$r=Invoke-RecoveryRun $source $parent $runner}finally{if($script:recoveryLogLock){$script:recoveryLogLock.Dispose();$script:recoveryLogLock=$null}}
        $r.Code|Should -Be 2;$r.EmergencyText|Should -Match '(?i)(log|logging).*(failed|degrad|unavailable)'
        Assert-RecoverySummary ($r.Text+$r.EmergencyText) 1 1 -LogWarnings 1
        $r.EmergencyText|Should -Match 'NATIVE IMG: Attempt=1';$r.EmergencyText|Should -Match 'SUMMARY ConvertedImages=1'
        [IO.File]::ReadAllText($r.LogPath)|Should -Not -Match 'SUMMARY ConvertedImages='
        Assert-RecoveryMedia $r @('image.jpeg') @('video.mp4') $source;Get-RecoveryState $source|Should -Be $before
    }

    It 'T049 reports the final persisted degraded outcome when the <Phase> append fails after valid media completes' -ForEach @(@{Phase='SUMMARY';Needle='SUMMARY ConvertedImages='},@{Phase='Processing-ended';Needle='Processing ended '}) {
        $source=New-RecoveryDirectory ('late-log-'+$Phase+'/source');$parent=New-RecoveryDirectory ('late-log-'+$Phase+'/output')
        New-RecoveryImage (Join-Path $source 'image.png');New-RecoveryVideo (Join-Path $source 'video.mp4');$before=Get-RecoveryState $source
        $lateNeedle=$Needle
        Mock Add-WinImgRunLogLine {param($Path,$Line)
            if($Line.Contains($lateNeedle)){throw [IO.IOException]::new('Controlled late append failure: '+$lateNeedle)}
            & $appendImplementation -Path $Path -Line $Line
        }
        $r=Invoke-RecoveryRun $source $parent;$r.Code|Should -Be 2
        # SUMMARY may be formatted before its own sink fails. The subsequent
        # final state must report the actual shared mutable failure count.
        $r.Text|Should -Match 'Final reporting state: LogWarnings=1 DiskLogIncomplete=True FallbackDropped=0'
        $r.EmergencyText|Should -Match 'Controlled late append failure';$r.EmergencyText|Should -Match 'FailureCount=1|LogWarnings=1'
        Should -Invoke Add-WinImgRunLogLine -Times 1 -Exactly -ParameterFilter {$Line.Contains($lateNeedle)}
        Assert-RecoveryMedia $r @('image.jpeg') @('video.mp4') $source;Get-RecoveryState $source|Should -Be $before
        $observations.Add([pscustomobject]@{Kind='controlled late log primitive failure after actual media';Phase=$Phase;FinalCode=$r.Code;ValidMediaRetained=$true;FinalReportingState=$r.Text})
    }

    It 'T049 bounded fallback retains quiet diagnostics, drops excess lines and emits only once without retrying a failed disk sink' {
        $path=Join-Path (New-RecoveryDirectory 'bounded-log') 'owned.log';$state=New-WinImgLogState -Path $path;Initialize-WinImgRunLog -State $state
        Mock Add-WinImgRunLogLine {throw [IO.IOException]::new('Controlled ordinary append failure')}
        $saved=[Console]::Error;$stderr=[IO.StringWriter]::new()
        try{[Console]::SetError($stderr)
            Write-WinImgRunLog -State $state -Message 'quiet diagnostic triggers ordinary sink failure' -Quiet
            for($i=0;$i -lt 100;$i++){Write-WinImgRunLog -State $state -Message ('quiet sequence '+$i+' '+('x'*1200)) -Quiet}
            Write-WinImgRunLog -State $state -Message 'SUMMARY retained valid media and degraded logging' -Quiet
            $state.Fallback.Length|Should -BeLessOrEqual 8192;$state.FallbackDroppedLines|Should -BeGreaterThan 0;$state.FallbackTruncatedLines|Should -BeGreaterThan 0
            $state.Degraded|Should -BeTrue;$state.DiskEnabled|Should -BeFalse;$state.FailureCount|Should -Be 1
            Complete-WinImgRunLog -State $state;$first=$stderr.ToString();Complete-WinImgRunLog -State $state;$stderr.ToString()|Should -Be $first
            $first|Should -Match 'SUMMARY retained valid media';$first.Length|Should -BeLessThan 10000;$state.FallbackEmitted|Should -BeTrue
        }finally{[Console]::SetError($saved);$stderr.Dispose()}
        Should -Invoke Add-WinImgRunLogLine -Times 1 -Exactly
    }

    It 'T049 emergency stderr failure is caught and cannot recurse or erase the degraded state' {
        $path=Join-Path (New-RecoveryDirectory 'broken-stderr') 'owned.log';$state=New-WinImgLogState -Path $path
        Mock New-WinImgRunLogFile {throw [IO.IOException]::new('Controlled initial sink unavailable')}
        $saved=[Console]::Error;$broken=[WinImgRecoveryBrokenWriter]::new()
        try{[Console]::SetError($broken);{Initialize-WinImgRunLog -State $state;Write-WinImgRunLog -State $state -Message 'Retained data outcome';Complete-WinImgRunLog -State $state}|Should -Not -Throw}
        finally{[Console]::SetError($saved);$broken.Dispose()}
        $state.Degraded|Should -BeTrue;$state.FailureCount|Should -Be 1;$state.Fallback.Length|Should -BeLessOrEqual 8192
    }

    It 'T049 isolates controlled PipelineStoppedException from the Pester pipeline and preserves its ordinary-log-failure classification' {
        $dir=New-RecoveryDirectory 'stopped-pipeline';$child=Join-Path $dir 'controlled-stop.ps1';$record=Join-Path $dir 'control-state.json';$log=Join-Path $dir 'owned.log'
        $childSource=@'
param([string]$RuntimePath,[string]$LogPath,[string]$RecordPath)
$ErrorActionPreference='Stop'
. $RuntimePath
$state=New-WinImgLogState -Path $LogPath
Initialize-WinImgRunLog -State $state
$script:triggered=$false;$afterCall=$false
function Add-WinImgRunLogLine {
    param([string]$Path,[string]$Line)
    $script:triggered=$true
    throw [Management.Automation.PipelineStoppedException]::new()
}
try { Write-WinImgRunLog -State $state -Message 'Controlled helper pipeline stop';$afterCall=$true }
finally {
    $observed=[pscustomobject]@{BeforeCall=$true;Triggered=$script:triggered;ReachedAfterCall=$afterCall;Degraded=$state.Degraded;FailureCount=$state.FailureCount;DiskEnabled=$state.DiskEnabled}
    [IO.File]::WriteAllText($RecordPath,($observed|ConvertTo-Json),[Text.UTF8Encoding]::new($false))
}
'@
        [IO.File]::WriteAllText($child,$childSource,[Text.UTF8Encoding]::new($false))
        $hostExecutable=(Get-Process -Id $PID).Path
        $actual=Invoke-WinImgNativeProcess -Executable $hostExecutable -Arguments @('-NoLogo','-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',$child,(Join-Path $repository 'WinImgNormalizer.ps1'),$log,$record) -TimeoutMilliseconds 15000
        $actual.StreamsComplete|Should -BeTrue;$actual.TimedOut|Should -BeFalse
        [IO.File]::Exists($record)|Should -BeTrue -Because 'A silently stopped child with native zero alone is not a completed test observation.'
        $state=[IO.File]::ReadAllText($record)|ConvertFrom-Json
        $state.BeforeCall|Should -BeTrue;$state.Triggered|Should -BeTrue;$state.ReachedAfterCall|Should -BeFalse
        $state.Degraded|Should -BeFalse;$state.FailureCount|Should -Be 0;$state.DiskEnabled|Should -BeTrue
        $observations.Add([pscustomobject]@{Kind='isolated actual-host controlled PipelineStopped helper';NativeChild=$actual;State=$state;Qualification='Helper propagation/classification only; not real Ctrl+C or owned-process cancellation lifecycle.'})
    }

    It 'T050 retains validated <Kind> data and warning duplicate status when setting <Field> fails, attempts the other timestamp and continues siblings' -ForEach @(
        @{Kind='image';Field='LastWriteTimeUtc'},@{Kind='image';Field='CreationTimeUtc'},@{Kind='video';Field='LastWriteTimeUtc'},@{Kind='video';Field='CreationTimeUtc'}
    ) {
        $source=New-RecoveryDirectory ('timestamp-'+$Kind+'-'+$Field+'/source');$parent=New-RecoveryDirectory ('timestamp-'+$Kind+'-'+$Field+'/output')
        if($Kind -eq 'image'){
            $firstRel='a retained/same.png';$firstOut='a retained/same.jpeg';$duplicateRel='b duplicate/same.png'
            New-RecoveryImage (Join-Path $source $firstRel);New-RecoveryImage (Join-Path $source 'z later/later.png');New-RecoveryVideo (Join-Path $source 'later.mp4');$images=2;$videos=1;$duplicateStatus='ConvertedWithWarning'
            $jpegs=@($firstOut,'z later/later.jpeg');$videoPaths=@('later.mp4')
        }else{
            $firstRel='a retained/same.mp4';$firstOut=$firstRel;$duplicateRel='b duplicate/same.mp4'
            New-RecoveryVideo (Join-Path $source $firstRel);New-RecoveryVideo (Join-Path $source 'z later/later.mp4');New-RecoveryImage (Join-Path $source 'later.png');$images=1;$videos=2;$duplicateStatus='CopiedVideoWithWarning'
            $jpegs=@('later.jpeg');$videoPaths=@($firstOut,'z later/later.mp4')
        }
        $duplicatePath=Join-Path $source $duplicateRel;$null=[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($duplicatePath));[IO.File]::WriteAllBytes($duplicatePath,[IO.File]::ReadAllBytes((Join-Path $source $firstRel)));[IO.File]::SetCreationTimeUtc($duplicatePath,$fixed);[IO.File]::SetLastWriteTimeUtc($duplicatePath,$fixed)
        $before=Get-RecoveryState $source;$firstSuffix=$firstOut.Replace('/','\');$failedField=$Field
        Mock Set-WinImgOutputTimestampValue {param($Path,$Name,$Value)
            if($Path.EndsWith($firstSuffix,[StringComparison]::OrdinalIgnoreCase) -and $Name -eq $failedField){throw [IO.IOException]::new('Controlled timestamp setter failure: '+$Name)}
            & $timestampImplementation -Path $Path -Name $Name -Value $Value
        }
        $r=Invoke-RecoveryRun $source $parent;$r.Code|Should -Be 2
        $combined=$r.Text+$r.EmergencyText;$combined|Should -Match '(?i)timestamp';$combined|Should -Match $Field
        Assert-RecoverySummary $combined $images $videos -Duplicates 1 -TimestampWarnings 1
        $log=[IO.File]::ReadAllText($r.LogPath);$log|Should -Match ('retained status: '+$duplicateStatus)
        Should -Invoke Set-WinImgOutputTimestampValue -Times 1 -Exactly -ParameterFilter {$Path.EndsWith($firstSuffix,[StringComparison]::OrdinalIgnoreCase) -and $Name -eq 'LastWriteTimeUtc'}
        Should -Invoke Set-WinImgOutputTimestampValue -Times 1 -Exactly -ParameterFilter {$Path.EndsWith($firstSuffix,[StringComparison]::OrdinalIgnoreCase) -and $Name -eq 'CreationTimeUtc'}
        $other=if($Field -eq 'LastWriteTimeUtc'){'CreationTimeUtc'}else{'LastWriteTimeUtc'}
        ([IO.FileInfo]::new((Join-Path $r.Run $firstOut))).$other.Ticks|Should -Be $fixed.Ticks
        Assert-RecoveryMedia $r $jpegs $videoPaths $source;Get-RecoveryState $source|Should -Be $before
        [IO.File]::Exists((Join-Path $r.Run ([IO.Path]::ChangeExtension($duplicateRel, $(if($Kind -eq 'image'){'.jpeg'}else{'.mp4'})))))|Should -BeFalse
    }
}
