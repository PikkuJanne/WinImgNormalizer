BeforeAll {
    $repository=Split-Path $PSScriptRoot -Parent
    $application=Join-Path $repository 'WinImgNormalizer.ps1'
    $scratch=Join-Path $repository '.scratch'
    foreach($ancestor in @($repository,$scratch)){if(([IO.File]::GetAttributes($ancestor)-band[IO.FileAttributes]::ReparsePoint)-ne0){throw 'Cancellation fixture ancestor is a reparse point.'}}
    $ownedRoot=Join-Path $scratch ('M3-T03-cancel-'+[Guid]::NewGuid().ToString('N'))
    [IO.Directory]::CreateDirectory($ownedRoot)|Out-Null
    [IO.File]::WriteAllText((Join-Path $ownedRoot 'fixture-owner.txt'),'Owned synthetic cancellation fixtures. No real Pictures, global console/environment/policy change, or recursive cleanup.',[Text.UTF8Encoding]::new($false))
    $observations=New-Object 'Collections.Generic.List[object]'
    $suitePath=Join-Path $PSScriptRoot 'Cancellation.Tests.ps1'
    $bindingsBefore=@($application,$suitePath)|ForEach-Object{[pscustomobject]@{Path=$_.Substring($repository.Length+1).Replace('\','/');Sha256=(Get-FileHash -LiteralPath $_).Hash.ToLowerInvariant()}}
    . $application
    $magick=$env:WINIMG_TEST_MAGICK
    if(-not$magick -or -not[IO.File]::Exists($magick)){throw 'Verified native ImageMagick is required.'}
    $version=Invoke-WinImgNativeProcess -Executable $magick -Arguments @('-version') -TimeoutMilliseconds 15000
    if($version.ExitCode-ne0 -or -not$version.StreamsComplete){throw 'Verified native version query failed.'}
    $hostExecutable=if($PSVersionTable.PSEdition-eq'Core'){Join-Path $PSHOME 'pwsh.exe'}else{Join-Path $PSHOME 'powershell.exe'}
    $compiler=Join-Path $env:SystemRoot 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
    if(-not[IO.File]::Exists($compiler)){throw 'Windows Framework compiler is required for the synthetic child fixture.'}
    $fixtureSource=Join-Path $ownedRoot 'CancellationFixture.cs'
    $fixture=Join-Path $ownedRoot 'CancellationFixture.exe'
    $fixtureCode=@'
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
public static class CancellationFixture {
 [DllImport("kernel32.dll", SetLastError=true)] static extern bool IsProcessInJob(IntPtr process,IntPtr job,out bool answer);
 static string Q(string s){return "\""+s.Replace("\"","\\\"")+"\"";}
 public static int Main(string[] args){
  string mode=args[0],root=args[1],role=args.Length>2?args[2]:"root";
  Directory.CreateDirectory(root);Process self=Process.GetCurrentProcess();bool inJob;IsProcessInJob(self.Handle,IntPtr.Zero,out inJob);
  using(FileStream held=new FileStream(Path.Combine(root,role+".held"),FileMode.CreateNew,FileAccess.ReadWrite,FileShare.None)){
   File.WriteAllText(Path.Combine(root,role+".identity"),self.Id+"|"+self.StartTime.ToUniversalTime().Ticks+"|"+inJob);
   if(mode=="unrelated"){
    File.WriteAllText(Path.Combine(root,"ready.txt"),"independent worker ready");
    while(!File.Exists(Path.Combine(root,"release.txt")))Thread.Sleep(20);
    return 0;
   }
   if(role!="grand"){
    string next=role=="root"?"child":"grand";
    ProcessStartInfo childInfo=new ProcessStartInfo(self.MainModule.FileName,"hold "+Q(root)+" "+next);
    childInfo.UseShellExecute=false;childInfo.CreateNoWindow=true;Process child=Process.Start(childInfo);child.Dispose();
   }
   if(role=="root"){
    Stopwatch clock=Stopwatch.StartNew();
    while((!File.Exists(Path.Combine(root,"root.identity"))||!File.Exists(Path.Combine(root,"child.identity"))||!File.Exists(Path.Combine(root,"grand.identity")))&&clock.ElapsedMilliseconds<10000)Thread.Sleep(20);
    if(!File.Exists(Path.Combine(root,"root.identity"))||!File.Exists(Path.Combine(root,"child.identity"))||!File.Exists(Path.Combine(root,"grand.identity")))return 4;
    if(args.Length>4&&args[3]!="-"){
     using(FileStream input=new FileStream(args[3],FileMode.Open,FileAccess.Read,FileShare.Read))
     using(FileStream output=new FileStream(args[4],FileMode.Open,FileAccess.Write,FileShare.None)){input.CopyTo(output);output.Flush(true);}
    }
    File.WriteAllText(Path.Combine(root,"ready.txt"),"owned root-child-grandchild ready");
   }
   while(true)Thread.Sleep(1000);
  }
 }
}
'@
    [IO.File]::WriteAllText($fixtureSource,$fixtureCode,[Text.UTF8Encoding]::new($false))
    & $compiler /nologo /target:exe ("/out:"+$fixture) $fixtureSource *> (Join-Path $ownedRoot 'compiler-console.log')
    if($LASTEXITCODE-ne0){throw 'Synthetic cancellation fixture compilation failed.'}
    if(-not('WinImgCancellationFixtureWatch' -as[type])){
        Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Threading;
public sealed class WinImgCancellationFixtureWatch:IDisposable {
 readonly Thread thread;volatile bool stopped;public bool Requested;public string Error;
 public WinImgCancellationFixtureWatch(object state,string marker){
  thread=new Thread(delegate(){try{DateTime end=DateTime.UtcNow.AddSeconds(15);while(!stopped&&!File.Exists(marker)&&DateTime.UtcNow<end)Thread.Sleep(20);if(!stopped&&File.Exists(marker)){state.GetType().GetMethod("Request").Invoke(state,null);Requested=true;}}catch(Exception ex){Error=ex.ToString();}});
  thread.IsBackground=true;thread.Start();
 }
 public void Dispose(){stopped=true;if(!thread.Join(2000))throw new Exception("Owned readiness watcher did not join");}
}
'@
    }
    function New-CancelDirectory([string]$Relative){$path=Join-Path $ownedRoot $Relative;[IO.Directory]::CreateDirectory($path)|Out-Null;return $path}
    function Get-CancelTreeState([string]$Root){
        $rows=@([IO.DirectoryInfo]::new($Root))+@(Get-ChildItem -LiteralPath $Root -Force -Recurse)
        return @($rows|Sort-Object FullName|ForEach-Object{
            $_.Refresh();$directory=$_-is[IO.DirectoryInfo];[pscustomobject]@{Path=$_.FullName.Substring($Root.Length).Replace('\','/');Directory=$directory;Length=$(if($directory){$null}else{$_.Length});Sha256=$(if($directory){$null}else{(Get-FileHash -LiteralPath $_.FullName).Hash.ToLowerInvariant()});CreationTicks=$_.CreationTimeUtc.Ticks;ModifiedTicks=$_.LastWriteTimeUtc.Ticks;Attributes=[int]$_.Attributes}
        })|ConvertTo-Json -Depth 6 -Compress
    }
    function Set-CancelFixtureTimes([string]$Root){
        $stamp=[DateTime]::Parse('2020-06-07T08:09:10Z').ToUniversalTime()
        foreach($file in @(Get-ChildItem -LiteralPath $Root -File -Recurse -Force)){[IO.File]::SetCreationTimeUtc($file.FullName,$stamp);[IO.File]::SetLastWriteTimeUtc($file.FullName,$stamp)}
        foreach($directory in @(@(Get-ChildItem -LiteralPath $Root -Directory -Recurse -Force)|Sort-Object FullName -Descending)+@([IO.DirectoryInfo]::new($Root))){[IO.Directory]::SetCreationTimeUtc($directory.FullName,$stamp);[IO.Directory]::SetLastWriteTimeUtc($directory.FullName,$stamp)}
    }
    function Invoke-CancelMagick([string[]]$Arguments){
        $tokens=@('-define','registry:filename:literal=true')+$Arguments
        if($Arguments[0]-eq'identify'){$tokens=@('identify','-define','registry:filename:literal=true')+$Arguments[1..($Arguments.Count-1)]}
        $r=Invoke-WinImgNativeProcess -Executable $magick -Arguments $tokens -TimeoutMilliseconds 15000
        if($r.ExitCode-ne0 -or -not$r.StreamsComplete -or $r.StdErr){throw ('Native fixture/decode failed: '+$r.StdErr)}
        return $r.StdOut.Trim()
    }
    $red=Join-Path $ownedRoot 'red.png';$blue=Join-Path $ownedRoot 'blue.png';$jpeg=Join-Path $ownedRoot 'valid.jpeg'
    $null=Invoke-CancelMagick @('-size','16x12','xc:#E02020',('PNG:'+$red))
    $null=Invoke-CancelMagick @('-size','16x12','xc:#2040E0',('PNG:'+$blue))
    $null=Invoke-CancelMagick @($red,('JPEG:'+$jpeg))
    $videoBytes=New-Object byte[] (3*1024*1024+17);[Random]::new(4319).NextBytes($videoBytes)
    function Get-CancelFileState([string]$Path){$f=[IO.FileInfo]::new($Path);$f.Refresh();return [pscustomobject]@{Sha256=(Get-FileHash -LiteralPath $Path).Hash;Length=$f.Length;CreationTicks=$f.CreationTimeUtc.Ticks;ModifiedTicks=$f.LastWriteTimeUtc.Ticks;Attributes=[int]$f.Attributes}|ConvertTo-Json -Compress}
    function New-CancelCase([string]$Name){
        $root=New-CancelDirectory $Name;$source=New-CancelDirectory ($Name+'/source');$parent=New-CancelDirectory ($Name+'/output');$control=New-CancelDirectory ($Name+'/control')
        [IO.File]::WriteAllBytes((Join-Path $source '01-first.png'),[IO.File]::ReadAllBytes($red))
        [IO.File]::WriteAllBytes((Join-Path $source '02-inflight.png'),[IO.File]::ReadAllBytes($blue))
        [IO.File]::WriteAllBytes((Join-Path $source '03-later.png'),[IO.File]::ReadAllBytes($blue))
        [IO.File]::WriteAllBytes((Join-Path $source '04-video.mp4'),$videoBytes)
        [IO.Directory]::CreateDirectory((Join-Path $source 'empty mirrored'))|Out-Null
        [IO.File]::WriteAllBytes((Join-Path $source 'unsupported.bin'),[byte[]]@(1,2,3,4))
        [IO.File]::WriteAllText((Join-Path $parent 'external-sentinel.txt'),'owned external arrival',[Text.UTF8Encoding]::new($false))
        Set-CancelFixtureTimes $source;Set-CancelFixtureTimes $parent
        $case=[pscustomobject]@{Root=$root;Source=$source;Parent=$parent;Control=$control;Before=(Get-CancelTreeState $source);ParentSentinel=(Get-CancelFileState (Join-Path $parent 'external-sentinel.txt'));Run=$null}
        [IO.File]::WriteAllText((Join-Path $control 'source-before.json'),$case.Before,[Text.UTF8Encoding]::new($false));return $case
    }
    function Invoke-CancelRun($Case,$State,[scriptblock]$Stage,[scriptblock]$Copy,[scriptblock]$Native){
        $before=@(Get-ChildItem -LiteralPath $Case.Parent -Directory|ForEach-Object FullName)
        $parameters=@{Source=$Case.Source;OutputParent=$Case.Parent;MagickPath=$magick;NativeTemporaryRoot=$ownedRoot;CancellationState=$State;CaptureConsoleControl=$false;ExecutionPolicy=(New-WinImgExecutionPolicy -TimeoutMilliseconds 15000)}
        if($Stage){$parameters.RunStageObserver=$Stage};if($Copy){$parameters.CopyProgressObserver=$Copy};if($Native){$parameters.NativeProcessObserver=$Native}
        $captured=@(& Invoke-WinImgNormalizer @parameters 6>&1 3>&1 2>&1)
        $codes=@($captured|Where-Object{$_-is[int]});if($codes.Count-ne1){throw 'Application did not return exactly one integer exit outcome.'}
        $runs=@(Get-ChildItem -LiteralPath $Case.Parent -Directory|Where-Object{$before-notcontains$_.FullName})
        if($runs.Count-eq1){$Case.Run=$runs[0].FullName}
        $text=($captured|ForEach-Object{$_|Out-String})-join''
        $log='';if($Case.Run){$logs=@(Get-ChildItem -LiteralPath $Case.Run -File -Filter '*.log' -Recurse -Force);if($logs.Count-ne1){throw 'Created run did not have exactly one required disk log.'};$log=[IO.File]::ReadAllText($logs[0].FullName)}
        $r=[pscustomobject]@{Code=$codes[0];Text=$text;Log=$log;Run=$Case.Run;Requested=$State.Requested}
        $observations.Add([pscustomobject]@{Kind='controlled cooperative application outcome';Case=$Case.Source;Result=$r;ConsoleControlCount=$State.ConsoleControlCount})
        [IO.File]::WriteAllText((Join-Path $Case.Control ('application-'+[Guid]::NewGuid().ToString('N')+'.json')),($r|ConvertTo-Json -Depth 8),[Text.UTF8Encoding]::new($false));return $r
    }
    function Assert-CancelSource($Case){$after=Get-CancelTreeState $Case.Source;[IO.File]::WriteAllText((Join-Path $Case.Control 'source-after.json'),$after,[Text.UTF8Encoding]::new($false));$after|Should -Be $Case.Before;(Get-CancelFileState (Join-Path $Case.Parent 'external-sentinel.txt'))|Should -Be $Case.ParentSentinel}
    function Assert-CancelSummary($Result){$Result.Code|Should -Be 130;$summary=if($Result.Log){$Result.Log}else{$Result.Text};([regex]::Matches($summary,'(?m)INTERRUPTED Mode=Cooperative ExitCode=130 ')).Count|Should -Be 1;$summary|Should -Match 'CompletedOutputs=Retained Cleanup=BestEffortOwnedOnly';$Result.Text|Should -Not -Match 'Processing ended successfully'}
    function Assert-CancelJpeg([string]$Path){Invoke-CancelMagick @('identify','+ping','-regard-warnings','-format','%m|%w|%h',$Path)|Should -Be 'JPEG|16|12';Invoke-CancelMagick @($Path,'-format','%[fx:round(255*p{0,0}.r)]|%[fx:round(255*p{0,0}.g)]|%[fx:round(255*p{0,0}.b)]','info:')|Should -Match '^22[0-9]\|3[0-9]\|3[0-9]$';$f=[IO.FileInfo]::new($Path);$f.CreationTimeUtc|Should -Be ([DateTime]::Parse('2020-06-07T08:09:10Z').ToUniversalTime());$f.LastWriteTimeUtc|Should -Be ([DateTime]::Parse('2020-06-07T08:09:10Z').ToUniversalTime())}
    function Assert-CancelWorkEmpty($Result){if($Result.Run){@(Get-ChildItem -LiteralPath (Join-Path $Result.Run '.WinImgNormalizer/work') -Force).Count|Should -Be 0}}
    function Assert-CancelTreeGone([string]$Control){
        $identities=@(Get-ChildItem -LiteralPath $Control -File -Filter '*.identity'|Where-Object{$_.BaseName-in@('root','child','grand')});$identities.Count|Should -Be 3
        foreach($file in $identities){$fields=[IO.File]::ReadAllText($file.FullName).Split('|');$clock=[Diagnostics.Stopwatch]::StartNew();$same=$true
            while($same-and$clock.ElapsedMilliseconds-lt3500){try{$process=[Diagnostics.Process]::GetProcessById([int]$fields[0]);$same=$process.StartTime.ToUniversalTime().Ticks-eq[long]$fields[1]-and-not$process.HasExited;$process.Dispose()}catch [ArgumentException]{$same=$false};if($same){Start-Sleep -Milliseconds 20}}
            $same|Should -BeFalse;$held=[IO.File]::Open((Join-Path $Control ($file.BaseName+'.held')),[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None);$held.Dispose()
        }
    }
}

Describe 'M3-T03 cooperative cancellation and recognizable force-termination regression' {
    It 'T054 has an idempotent thread-safe request flag and interruptible wait without implying a console event' {
        $state=[WinImgNormalizer.CancellationState]::new();$state.Requested|Should -BeFalse;{$state.Wait(1)}|Should -Not -Throw
        $state.Request();$state.Request();$state.Requested|Should -BeTrue;{$state.Wait(0)}|Should -Throw '*cancel*';$state.ConsoleControlCount|Should -Be 0
        {$state.ThrowIfRequested()}|Should -Throw '*cancel*'
    }
    It 'T054 stops an already-requested run without starting or finalizing source items' {
        $case=New-CancelCase 'already-requested';$state=[WinImgNormalizer.CancellationState]::new();$state.Request()
        $r=Invoke-CancelRun $case $state;Assert-CancelSummary $r
        if($r.Run){@(Get-ChildItem -LiteralPath $r.Run -File -Filter '*.jpeg').Count|Should -Be 0;@(Get-ChildItem -LiteralPath $r.Run -File -Filter '*.mp4').Count|Should -Be 0}
        Assert-CancelSource $case;Assert-CancelWorkEmpty $r
    }
    It 'T054 cancels before the second item and retains the finalized first native JPEG' {
        $case=New-CancelCase 'before-next';$state=[WinImgNormalizer.CancellationState]::new();$started=New-Object 'Collections.Generic.List[string]'
        $stage={param($stage,$token,$relative,$path)if($stage-eq'BeforeItem'){$started.Add($relative);if($relative-eq'02-inflight.png'){$token.Request()}}}
        $r=Invoke-CancelRun $case $state $stage;Assert-CancelSummary $r;Assert-CancelJpeg (Join-Path $r.Run '01-first.jpeg')
        ($started.ToArray()-join'|')|Should -Be '01-first.png|02-inflight.png';[IO.File]::Exists((Join-Path $r.Run '02-inflight.jpeg'))|Should -BeFalse;[IO.File]::Exists((Join-Path $r.Run '04-video.mp4'))|Should -BeFalse
        Assert-CancelSource $case;Assert-CancelWorkEmpty $r
    }
    It 'T054 cancels an actual owned native tree during controlled <Phase> pacing and retains only the completed first image' -ForEach @(@{Phase='inspection'},@{Phase='conversion'},@{Phase='validation'}) {
        $case=New-CancelCase ('native-'+$Phase);$state=[WinImgNormalizer.CancellationState]::new();$routing=@{Relative='';Used=$false};$trace=New-Object 'Collections.Generic.List[object]'
        $stage={param($stage,$token,$relative,$path)if($stage-eq'BeforeItem'){$routing.Relative=$relative}}
        $native={param($executable,[string[]]$arguments,[int]$remaining,$environment)
            $convert=$arguments[-1].StartsWith('JPEG:',[StringComparison]::Ordinal);$validate=$arguments[0]-eq'identify'-and$arguments-contains'+ping'
            $selected=(-not$routing.Used -and $routing.Relative-eq'02-inflight.png' -and (($Phase-eq'inspection')-or($Phase-eq'conversion'-and$convert)-or($Phase-eq'validation'-and$validate)))
            if($selected){$routing.Used=$true;$copyFrom='-';$copyTo='-';if($convert){$copyFrom=$jpeg;$copyTo=$arguments[-1].Substring(5)}
                $watch=[WinImgCancellationFixtureWatch]::new($state,(Join-Path $case.Control 'ready.txt'))
                try{$result=Invoke-WinImgNativeProcess -Executable $fixture -Arguments @('hold',$case.Control,'root',$copyFrom,$copyTo) -TimeoutMilliseconds $remaining -Environment $environment}finally{$watch.Dispose()}
                $watch.Requested|Should -BeTrue;$watch.Error|Should -BeNullOrEmpty
            }else{$result=Invoke-WinImgNativeProcess -Executable $executable -Arguments $arguments -TimeoutMilliseconds $remaining -Environment $environment}
            $trace.Add([pscustomobject]@{Selected=$selected;Result=$result});return $result
        }
        $r=Invoke-CancelRun $case $state $stage $null $native;Assert-CancelSummary $r;$routing.Used|Should -BeTrue
        $cancelled=@($trace.ToArray()|Where-Object Selected);$cancelled.Count|Should -Be 1;$cancelled[0].Result.Cancelled|Should -BeTrue;$cancelled[0].Result.TimedOut|Should -BeFalse;$cancelled[0].Result.TreeTerminated|Should -BeTrue;$cancelled[0].Result.StreamsComplete|Should -BeTrue
        Assert-CancelTreeGone $case.Control;Assert-CancelJpeg (Join-Path $r.Run '01-first.jpeg');[IO.File]::Exists((Join-Path $r.Run '02-inflight.jpeg'))|Should -BeFalse
        Assert-CancelSource $case;Assert-CancelWorkEmpty $r
        $observations.Add([pscustomobject]@{Kind='actual owned tree with controlled readiness cancellation';Phase=$Phase;Calls=$trace.ToArray();ActualConsoleEvent=$false})
    }
    It 'T055 cancels after a real video chunk and removes the partial without presenting a completed video' {
        $case=New-CancelCase 'video-chunk';$state=[WinImgNormalizer.CancellationState]::new();$copyTrace=New-Object 'Collections.Generic.List[object]'
        $copy={param($source,$candidate,[long]$copied,$token)if([IO.Path]::GetExtension($source)-eq'.mp4'-and[IO.Path]::GetFileName($candidate)-eq'video.partial'){$copyTrace.Add([pscustomobject]@{Source=$source;Candidate=$candidate;Copied=$copied;ActualLength=[IO.FileInfo]::new($candidate).Length});$token.Request()}}
        $r=Invoke-CancelRun $case $state $null $copy;Assert-CancelSummary $r;$copyTrace.Count|Should -Be 1;$copyTrace[0].Copied|Should -Be 262144;$copyTrace[0].ActualLength|Should -BeGreaterThan 0
        [IO.File]::Exists((Join-Path $r.Run '04-video.mp4'))|Should -BeFalse;[IO.File]::Exists($copyTrace[0].Candidate)|Should -BeFalse;Assert-CancelJpeg (Join-Path $r.Run '01-first.jpeg')
        Assert-CancelSource $case;Assert-CancelWorkEmpty $r;$observations.Add([pscustomobject]@{Kind='actual video chunk with controlled Request';Chunks=$copyTrace.ToArray();ActualConsoleEvent=$false})
    }
    It 'T055 preserves an unknown sibling arrival when cancelling the actual video-copy partial' {
        $case=New-CancelCase 'video-arrival';$state=[WinImgNormalizer.CancellationState]::new();$arrival=@{Path=$null;Before=$null}
        $copy={param($source,$candidate,[long]$copied,$token)if([IO.Path]::GetExtension($source)-eq'.mp4'-and[IO.Path]::GetFileName($candidate)-eq'video.partial'){$arrival.Path=Join-Path ([IO.Path]::GetDirectoryName($candidate)) 'foreign-arrival.bin';[IO.File]::WriteAllBytes($arrival.Path,[byte[]]@(5,7,11,13));$arrival.Before=(Get-FileHash -LiteralPath $arrival.Path).Hash;$token.Request()}}
        $r=Invoke-CancelRun $case $state $null $copy;Assert-CancelSummary $r;(Get-FileHash -LiteralPath $arrival.Path).Hash|Should -Be $arrival.Before;[IO.File]::Exists((Join-Path $r.Run '04-video.mp4'))|Should -BeFalse
        Assert-CancelSource $case;@(Get-ChildItem -LiteralPath ([IO.Path]::GetDirectoryName($arrival.Path)) -File).Count|Should -Be 1
    }
    It 'T054 refuses finalization when cancellation arrives after full native candidate validation' {
        $case=New-CancelCase 'validated-not-final';$state=[WinImgNormalizer.CancellationState]::new();$seen=New-Object 'Collections.Generic.List[string]'
        $stage={param($stage,$token,$relative,$path)if($stage-eq'CandidateValidated'){$seen.Add($path);$token.Request()}}
        $r=Invoke-CancelRun $case $state $stage;Assert-CancelSummary $r;$seen.Count|Should -Be 1;[IO.File]::Exists((Join-Path $r.Run '01-first.jpeg'))|Should -BeFalse
        Assert-CancelSource $case;Assert-CancelWorkEmpty $r
    }
    It 'T054 retains a finalized image when a repeated request follows its completed bookkeeping' {
        $case=New-CancelCase 'after-final';$state=[WinImgNormalizer.CancellationState]::new();$stage={param($stage,$token,$relative,$path)if($stage-eq'AfterFinalization'){$token.Request();$token.Request()}}
        $r=Invoke-CancelRun $case $state $stage;Assert-CancelSummary $r;Assert-CancelJpeg (Join-Path $r.Run '01-first.jpeg');$r.Log|Should -Match 'ConvertedImages=1 CopiedVideos=0'
        Assert-CancelSource $case;Assert-CancelWorkEmpty $r
    }
    It 'T054 restores scoped ambient cancellation before a subsequent independent normal run' {
        $first=New-CancelCase 'scope-first';$firstState=[WinImgNormalizer.CancellationState]::new();$firstState.Request();$r=Invoke-CancelRun $first $firstState;Assert-CancelSummary $r
        [WinImgNormalizer.CancellationSession]::Current|Should -BeNullOrEmpty
        $second=New-CancelCase 'scope-second';$normal=Invoke-CancelRun $second ([WinImgNormalizer.CancellationState]::new());$normal.Code|Should -Be 0
        [WinImgNormalizer.CancellationSession]::Current|Should -BeNullOrEmpty;Assert-CancelJpeg (Join-Path $normal.Run '01-first.jpeg')
        (Get-FileHash -LiteralPath (Join-Path $normal.Run '04-video.mp4')).Hash|Should -Be (Get-FileHash -LiteralPath (Join-Path $second.Source '04-video.mp4')).Hash
        Assert-CancelSource $first;Assert-CancelSource $second;Assert-CancelWorkEmpty $normal
    }
    It 'T056 preserves recognizable nonfinal scratch and completed media after actual forced worker termination, without reusing it in a fresh run' {
        $case=New-CancelCase 'force';$driver=Join-Path $case.Control 'force-driver.ps1';$config=Join-Path $case.Control 'force-config.json'
        [IO.File]::WriteAllText($config,([pscustomobject]@{Runtime=$application;Source=$case.Source;Output=$case.Parent;Magick=$magick;Temporary=$ownedRoot;Fixture=$fixture;Control=$case.Control;Jpeg=$jpeg}|ConvertTo-Json),[Text.UTF8Encoding]::new($false))
        $driverCode=@'
$ErrorActionPreference='Stop'
$config=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'force-config.json'))|ConvertFrom-Json
. $config.Runtime
[IO.File]::WriteAllText((Join-Path $config.Control 'host.identity'),($PID.ToString()+'|'+[Diagnostics.Process]::GetCurrentProcess().StartTime.ToUniversalTime().Ticks),[Text.UTF8Encoding]::new($false))
$script:forceRelative=''
$stage={param($stage,$state,$relative,$path)if($stage-eq'BeforeItem'){$script:forceRelative=$relative}}
$native={param($executable,[string[]]$arguments,[int]$remaining,$environment)
 if($script:forceRelative-eq'02-inflight.png'-and$arguments[-1].StartsWith('JPEG:',[StringComparison]::Ordinal)){
  return Invoke-WinImgNativeProcess -Executable $config.Fixture -Arguments @('hold',$config.Control,'root',$config.Jpeg,$arguments[-1].Substring(5)) -TimeoutMilliseconds $remaining -Environment $environment
 }
 return Invoke-WinImgNativeProcess -Executable $executable -Arguments $arguments -TimeoutMilliseconds $remaining -Environment $environment
}
$code=Invoke-WinImgNormalizerCommand -Arguments @($config.Source) -OutputParent $config.Output -MagickPath $config.Magick -NativeTemporaryRoot $config.Temporary -RunStageObserver $stage -NativeProcessObserver $native -CaptureConsoleControl $false
[IO.File]::WriteAllText((Join-Path $config.Control 'normal-return.txt'),$code.ToString(),[Text.UTF8Encoding]::new($false))
exit $code
'@
        [IO.File]::WriteAllText($driver,$driverCode,[Text.UTF8Encoding]::new($false))
        $info=[Diagnostics.ProcessStartInfo]::new($hostExecutable);$info.UseShellExecute=$false;$info.CreateNoWindow=$true;$info.WindowStyle=[Diagnostics.ProcessWindowStyle]::Hidden;$info.Arguments='-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "'+$driver+'"'
        $previous=[Environment]::GetEnvironmentVariable('PSModulePath','Process')
        try{[Environment]::SetEnvironmentVariable('PSModulePath',$null,'Process');$child=[Diagnostics.Process]::Start($info)}finally{[Environment]::SetEnvironmentVariable('PSModulePath',$previous,'Process')}
        $childId=$child.Id;$childTicks=$child.StartTime.ToUniversalTime().Ticks;$clock=[Diagnostics.Stopwatch]::StartNew()
        try{
            while(-not[IO.File]::Exists((Join-Path $case.Control 'ready.txt'))-and-not$child.HasExited-and$clock.ElapsedMilliseconds-lt20000){Start-Sleep -Milliseconds 20}
            [IO.File]::Exists((Join-Path $case.Control 'ready.txt'))|Should -BeTrue;$child.HasExited|Should -BeFalse
            [IO.File]::ReadAllText((Join-Path $case.Control 'host.identity'))|Should -Be ($childId.ToString()+'|'+$childTicks)
            $child.StartTime.ToUniversalTime().Ticks|Should -Be $childTicks;$child.Kill();$child.WaitForExit(5000)|Should -BeTrue;$forcedExit=$child.ExitCode
        }finally{if(-not$child.HasExited){$child.Kill();$child.WaitForExit(5000)|Out-Null};$child.Dispose()}
        [IO.File]::Exists((Join-Path $case.Control 'normal-return.txt'))|Should -BeFalse
        $runs=@(Get-ChildItem -LiteralPath $case.Parent -Directory);$runs.Count|Should -Be 1;$case.Run=$runs[0].FullName
        Assert-CancelTreeGone $case.Control;Assert-CancelJpeg (Join-Path $case.Run '01-first.jpeg');[IO.File]::Exists((Join-Path $case.Run '02-inflight.jpeg'))|Should -BeFalse
        $work=Join-Path $case.Run '.WinImgNormalizer/work';@(Get-ChildItem -LiteralPath $work -File -Recurse -Force).Count|Should -BeGreaterThan 0
        $firstRun=Get-CancelTreeState $case.Run;$firstPath=$case.Run;Assert-CancelSource $case
        $r=Invoke-CancelRun $case ([WinImgNormalizer.CancellationState]::new());$r.Code|Should -Be 0;$r.Run|Should -Not -Be $firstPath
        Get-CancelTreeState $firstPath|Should -Be $firstRun;Assert-CancelJpeg (Join-Path $r.Run '01-first.jpeg');@(Get-ChildItem -LiteralPath $r.Run -File -Filter '*.jpeg').Count|Should -Be 3
        Assert-CancelSource $case;Assert-CancelWorkEmpty $r
        $observations.Add([pscustomobject]@{Kind='actual force-terminated owned application child';HostPid=$childId;HostStartTicks=$childTicks;ObservedNativeExit=$forcedExit;NormalApplicationReturnAbsent=$true;LeftoverOwnedScratchRetained=$true;PriorRunExactAfterFreshRun=$true;Qualification='Exact owned PID/start identity force-kill, not cooperative Request or Ctrl+C. No exit130 or complete cleanup promise; the actual kernel-owned job tree was checked gone.'})
    }
}

AfterAll {
    if($ownedRoot -and [IO.Directory]::Exists($ownedRoot)){
        $bindingsAfter=@($application,$suitePath)|Where-Object{$_ -and [IO.File]::Exists($_)}|ForEach-Object{[pscustomobject]@{Path=$_.Substring($repository.Length+1).Replace('\','/');Sha256=(Get-FileHash -LiteralPath $_).Hash.ToLowerInvariant()}}
        $artifacts=@(Get-ChildItem -LiteralPath $ownedRoot -File -Recurse -Force|ForEach-Object{[pscustomobject]@{Path=$_.FullName.Substring($repository.Length+1).Replace('\','/');Sha256=(Get-FileHash -LiteralPath $_.FullName).Hash.ToLowerInvariant()}})
        $evidence=[pscustomobject]@{Task='M3-T03';PowerShell=$PSVersionTable.PSVersion.ToString();Edition=$PSVersionTable.PSEdition;BindingsBefore=$bindingsBefore;BindingsAfter=$bindingsAfter;BindingsUnchanged=(($bindingsBefore|ConvertTo-Json -Compress)-eq($bindingsAfter|ConvertTo-Json -Compress));NativeVersion=$version;Observations=$(if($observations){$observations.ToArray()}else{@()});Artifacts=$artifacts;Qualification='Synthetic cooperative Request controls. Actual Windows console events and force-termination observations are separately recorded; native zero alone is not cancellation acceptance.'}
        [IO.File]::WriteAllText((Join-Path $ownedRoot 'cancellation-observations.json'),($evidence|ConvertTo-Json -Depth 25),[Text.UTF8Encoding]::new($false))
        Write-Host ('Cancellation owned artifacts: '+$ownedRoot)
    }
}
