BeforeAll {
    $repository=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    $scratch=Join-Path $repository '.scratch'
    function Assert-LifetimeAncestors([string]$Path){
        $current=[IO.DirectoryInfo]::new([IO.Path]::GetFullPath($Path))
        while($current){if($current.Exists -and ($current.Attributes -band [IO.FileAttributes]::ReparsePoint)){throw 'Lifetime fixture crosses a reparse point.'};$current=$current.Parent}
    }
    Assert-LifetimeAncestors $scratch
    foreach($pictures in @([Environment]::GetFolderPath('MyPictures'),(Join-Path $env:USERPROFILE 'Pictures'))){
        if(-not $pictures){continue};$p=[IO.Path]::GetFullPath($pictures).TrimEnd('\','/')
        if($scratch.Equals($p,[StringComparison]::OrdinalIgnoreCase) -or $scratch.StartsWith($p+'\',[StringComparison]::OrdinalIgnoreCase) -or $p.StartsWith($scratch+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Lifetime fixture overlaps real Pictures.'}
    }
    & git -C $repository check-ignore --quiet --no-index -- (Join-Path $scratch 'native-lifetime-ignore-probe')
    if($LASTEXITCODE -ne 0){throw 'Lifetime fixture must be ignored.'}
    $ownedRoot=Join-Path $scratch ('M3-T02-lifetime-'+[Guid]::NewGuid().ToString('N'))
    if([IO.Directory]::Exists($ownedRoot)){throw 'Lifetime ownership collision.'}
    $null=[IO.Directory]::CreateDirectory($ownedRoot)
    [IO.File]::WriteAllText((Join-Path $ownedRoot '.winimg-fixture-root'),'M3-T02 owned benign native lifetime fixtures')
    $observations=New-Object 'Collections.Generic.List[object]'
    $outsideProcesses=New-Object 'Collections.Generic.List[object]'
    $script:stateSerial=0
    function New-LifetimeDirectory([string]$Relative){
        $path=[IO.Path]::GetFullPath((Join-Path $ownedRoot $Relative))
        if(-not $path.StartsWith($ownedRoot+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Lifetime fixture escaped ownership.'}
        Assert-LifetimeAncestors $path;$null=[IO.Directory]::CreateDirectory($path);return $path
    }
    if(-not $env:WINIMG_TEST_MAGICK){throw 'Explicit verified ImageMagick required.'}
    $magick=[IO.Path]::GetFullPath($env:WINIMG_TEST_MAGICK);Assert-LifetimeAncestors $magick
    . (Join-Path $repository 'WinImgNormalizer.ps1')
    function Get-LifetimeBindings {
        foreach($relative in @('WinImgNormalizer.ps1','WinImgNormalizer.bat','tests/NativeLifetime.Tests.ps1','tests/Invoke-Tests.ps1','tests/Initialize-TestDependencies.ps1','tests/dependencies.json')){
            $p=Join-Path $repository $relative;[pscustomobject]@{Path=$relative;Sha256=(Get-FileHash -LiteralPath $p).Hash.ToLowerInvariant()}
        }
        [pscustomobject]@{Path='verified ImageMagick executable';Sha256=(Get-FileHash -LiteralPath $magick).Hash.ToLowerInvariant()}
    }
    $bindingsBefore=@(Get-LifetimeBindings)
    $environmentBefore=@{};foreach($key in @('TEMP','TMP','MAGICK_TEMPORARY_PATH','MAGICK_CONFIGURE_PATH','MAGICK_MEMORY_LIMIT','MAGICK_MAP_LIMIT','MAGICK_DISK_LIMIT','MAGICK_THREAD_LIMIT')){$environmentBefore[$key]=[Environment]::GetEnvironmentVariable($key,'Process')}
    $nativeVersion=Invoke-WinImgNativeProcess -Executable $magick -Arguments @('-version') -TimeoutMilliseconds 15000
    if($nativeVersion.ExitCode -ne 0 -or $nativeVersion.StdErr -or -not $nativeVersion.StreamsComplete){throw 'Actual native version probe failed.'}
    $policyFiles=@(Get-ChildItem -LiteralPath ([IO.Path]::GetDirectoryName($magick)) -Filter '*policy*' -File -Force|ForEach-Object{[pscustomobject]@{Path=$_.FullName;Sha256=(Get-FileHash -LiteralPath $_.FullName).Hash.ToLowerInvariant()}})

    # Compiled only into owned ignored scratch. The observer runs OUTSIDE the
    # wrapper's job and obtains parent IDs with Toolhelp32, independently of the
    # root's reported hierarchy. Descendants inherit actual stdout/stderr pipes.
    $fixtureRoot=New-LifetimeDirectory "native child & O'Brien [tree]"
    $fixtureSource=Join-Path $fixtureRoot 'LifetimeFixture.cs';$fixture=Join-Path $fixtureRoot 'LifetimeFixture.exe'
    $recipe=@'
using System;
using System.IO;
using System.Text;
using System.Diagnostics;
using System.Threading;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public static class LifetimeFixture {
 [StructLayout(LayoutKind.Sequential,CharSet=CharSet.Unicode)] struct PROCESSENTRY32 {
  public uint size,usage,pid; public IntPtr heap; public uint module,threads,parent; public int priority; public uint flags;
  [MarshalAs(UnmanagedType.ByValTStr,SizeConst=260)] public string exe;
 }
 [DllImport("kernel32.dll",SetLastError=true)] static extern IntPtr CreateToolhelp32Snapshot(uint flags,uint id);
 [DllImport("kernel32.dll",CharSet=CharSet.Unicode,SetLastError=true)] static extern bool Process32FirstW(IntPtr snapshot,ref PROCESSENTRY32 entry);
 [DllImport("kernel32.dll",CharSet=CharSet.Unicode,SetLastError=true)] static extern bool Process32NextW(IntPtr snapshot,ref PROCESSENTRY32 entry);
 [DllImport("kernel32.dll",SetLastError=true)] static extern bool CloseHandle(IntPtr handle);
 [DllImport("kernel32.dll",SetLastError=true)] static extern bool IsProcessInJob(IntPtr process,IntPtr job,out bool answer);
 [DllImport("kernel32.dll")] static extern IntPtr GetCurrentProcess();
 public static string Quote(string s) {
  StringBuilder b=new StringBuilder("\"");int slash=0;
  foreach(char c in s){if(c=='\\'){slash++;continue;}if(c=='\"'){b.Append('\\',slash*2+1);b.Append(c);}else{b.Append('\\',slash);b.Append(c);}slash=0;}
  b.Append('\\',slash*2);return b.Append('"').ToString();
 }
 static Process Spawn(string[] args) {
  ProcessStartInfo p=new ProcessStartInfo();p.FileName=Process.GetCurrentProcess().MainModule.FileName;
  p.UseShellExecute=false;p.CreateNoWindow=true;StringBuilder b=new StringBuilder();
  foreach(string a in args){if(b.Length!=0)b.Append(' ');b.Append(Quote(a));}p.Arguments=b.ToString();return Process.Start(p);
 }
 static Dictionary<int,int> Parents(){
  Dictionary<int,int> rows=new Dictionary<int,int>();IntPtr h=CreateToolhelp32Snapshot(2,0);
  if(h==new IntPtr(-1))throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
  try{PROCESSENTRY32 e=new PROCESSENTRY32();e.size=(uint)Marshal.SizeOf(typeof(PROCESSENTRY32));
   if(Process32FirstW(h,ref e)){do{rows[(int)e.pid]=(int)e.parent;}while(Process32NextW(h,ref e));}
  }finally{CloseHandle(h);}return rows;
 }
 static int Observe(string root){
  string output=Path.Combine(root,"independent-parents.tsv");HashSet<string> seen=new HashSet<string>();
  File.WriteAllText(Path.Combine(root,"observer.ready"),"ready");Stopwatch clock=Stopwatch.StartNew();
  while(!File.Exists(Path.Combine(root,"observer.stop"))&&clock.ElapsedMilliseconds<30000){
   Dictionary<int,int> parents=Parents();
   foreach(string file in Directory.GetFiles(root,"*.identity")){
    try{string[] row=File.ReadAllText(file).Split('|');int pid=int.Parse(row[0]);long ticks=long.Parse(row[1]);
     using(Process p=Process.GetProcessById(pid)){
      long actual=p.StartTime.ToUniversalTime().Ticks;
      if(actual==ticks&&parents.ContainsKey(pid)&&seen.Add(file)){File.AppendAllText(output,Path.GetFileNameWithoutExtension(file)+"|"+pid+"|"+ticks+"|"+parents[pid]+"\n");if(seen.Count==3)File.WriteAllText(Path.Combine(root,"tree.observed"),"observed");}
     }
    }catch(ArgumentException){}catch(IOException){}catch(InvalidOperationException){}
   }
   Thread.Sleep(20);
  }return 0;
 }
 static int Tree(string[] args){
  string root=args[1],mode=args[2];int stage=int.Parse(args[3]),blocks=int.Parse(args[4]);bool job=false;
  IsProcessInJob(GetCurrentProcess(),IntPtr.Zero,out job);
  // Publish complete identity bytes before the independent observer can see them.
  string identity=Path.Combine(root,"stage"+stage+".identity");
  using(Process self=Process.GetCurrentProcess())File.WriteAllText(identity+".pending",self.Id+"|"+self.StartTime.ToUniversalTime().Ticks+"|"+job);
  File.Move(identity+".pending",identity);
  using(FileStream held=new FileStream(Path.Combine(root,"stage"+stage+".held"),FileMode.CreateNew,FileAccess.Write,FileShare.None)){
   if(stage<3)Spawn(new string[]{"tree",root,mode,(stage+1).ToString(),"0","-","-"}).Dispose();
   if(stage==1){
    Stopwatch wait=Stopwatch.StartNew();while(!File.Exists(Path.Combine(root,"stage3.identity"))&&wait.ElapsedMilliseconds<5000)Thread.Sleep(10);
    if(args[5]!="-")File.WriteAllBytes(args[6],File.ReadAllBytes(args[5]));
    Thread stdout=new Thread(delegate(){for(int i=0;i<blocks;i++)Console.Out.Write(new string('O',4096));});
    Thread stderr=new Thread(delegate(){for(int i=0;i<blocks;i++)Console.Error.Write(new string('E',4096));});
    stdout.Start();stderr.Start();stdout.Join();stderr.Join();
    if(mode=="exitroot"){Stopwatch observed=Stopwatch.StartNew();while(!File.Exists(Path.Combine(root,"tree.observed"))&&observed.ElapsedMilliseconds<5000)Thread.Sleep(10);return 0;}
   }
   Thread.Sleep(20000);
  }return 0;
 }
 public static int Main(string[] args){
  if(args[0]=="observe")return Observe(args[1]);
  if(args[0]=="tree")return Tree(args);
  if(args[0]=="finite"){Thread.Sleep(int.Parse(args[1]));return 0;}
  if(args[0]=="environment"){foreach(string key in new string[]{"TEMP","TMP","MAGICK_TEMPORARY_PATH"})Console.Out.WriteLine(key+"="+Environment.GetEnvironmentVariable(key));return 0;}
  return 97;
 }
}
'@
    [IO.File]::WriteAllText($fixtureSource,$recipe,[Text.UTF8Encoding]::new($false))
    $compiler=Join-Path $env:SystemRoot 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
    if(-not [IO.File]::Exists($compiler)){throw 'Windows .NET Framework compiler required for benign owned lifetime fixture.'}
    $compile=@(& $compiler '/nologo' '/target:exe' ('/out:'+$fixture) $fixtureSource 2>&1);$compileExit=$LASTEXITCODE
    [IO.File]::WriteAllText((Join-Path $fixtureRoot 'compile.log'),($compile -join "`n"))
    if($compileExit -ne 0 -or -not [IO.File]::Exists($fixture)){throw 'Lifetime fixture compilation failed.'}
    [IO.File]::WriteAllText(($fixture+'.config'),'<configuration><runtime><AppContextSwitchOverrides value="Switch.System.IO.UseLegacyPathHandling=false;Switch.System.IO.BlockLongPaths=false" /></runtime></configuration>')
    $observations.Add([pscustomobject]@{Kind='benign native fixture provenance';CompilerSha256=(Get-FileHash -LiteralPath $compiler).Hash.ToLowerInvariant();CompilerVersion=([IO.FileInfo]::new($compiler)).VersionInfo.FileVersion;RecipeSha256=(Get-FileHash -LiteralPath $fixtureSource).Hash.ToLowerInvariant();ExecutableSha256=(Get-FileHash -LiteralPath $fixture).Hash.ToLowerInvariant();CompileExit=$compileExit})
    function Start-LifetimeOutside([string]$Executable,[string[]]$Arguments,[switch]$RedirectInput){
        $info=[Diagnostics.ProcessStartInfo]::new();$info.FileName=$Executable;$info.UseShellExecute=$false;$info.CreateNoWindow=$true
        $info.RedirectStandardInput= [bool]$RedirectInput
        $info.Arguments=(@($Arguments|ForEach-Object{[WinImgNormalizer.NativeProcess]::Quote($_)}) -join ' ')
        $p=[Diagnostics.Process]::new();$p.StartInfo=$info;if(-not $p.Start()){throw 'Owned outside fixture did not start.'}
        $entry=[pscustomobject]@{Process=$p;Id=$p.Id;StartTicks=$p.StartTime.ToUniversalTime().Ticks;Executable=$Executable}
        $outsideProcesses.Add($entry);return $entry
    }
    function Stop-LifetimeOutside([object]$Entry){
        # Exact owned Process object and creation identity; never name-wide kill.
        $p=$Entry.Process
        if(-not $p.HasExited -and $p.Id -eq $Entry.Id -and $p.StartTime.ToUniversalTime().Ticks -eq $Entry.StartTicks){$p.Kill();$null=$p.WaitForExit(5000)}
    }
    function Wait-LifetimeFile([string]$Path,[int]$Milliseconds=5000){
        $watch=[Diagnostics.Stopwatch]::StartNew();while(-not [IO.File]::Exists($Path) -and $watch.ElapsedMilliseconds -lt $Milliseconds){Start-Sleep -Milliseconds 20}
        if(-not [IO.File]::Exists($Path)){throw ('Owned fixture did not write '+[IO.Path]::GetFileName($Path))}
    }
    function Start-LifetimeObserver([string]$Root){$entry=Start-LifetimeOutside $fixture @('observe',$Root);Wait-LifetimeFile (Join-Path $Root 'observer.ready');return $entry}
    function Complete-LifetimeObserver([object]$Entry,[string]$Root){
        [IO.File]::WriteAllText((Join-Path $Root 'observer.stop'),'stop')
        if(-not $Entry.Process.WaitForExit(5000)){Stop-LifetimeOutside $Entry;throw 'Owned independent observer did not finish.'}
        $Entry.Process.ExitCode|Should -Be 0
    }
    function Assert-LifetimeTreeGone([string]$Root,[object]$Result){
        $records=@(Get-ChildItem -LiteralPath $Root -Filter '*.identity' -File|Sort-Object Name)
        $records.Count|Should -Be 3
        $parents=@([IO.File]::ReadAllLines((Join-Path $Root 'independent-parents.tsv'))|ForEach-Object{
            $row=$_.Split('|');[pscustomobject]@{Role=$row[0];Id=[int]$row[1];Ticks=[long]$row[2];ParentId=[int]$row[3]}
        })
        $parents.Count|Should -Be 3
        ($parents|Where-Object Role -eq 'stage1').Id|Should -Be $Result.ProcessId
        ($parents|Where-Object Role -eq 'stage2').ParentId|Should -Be ($parents|Where-Object Role -eq 'stage1').Id
        ($parents|Where-Object Role -eq 'stage3').ParentId|Should -Be ($parents|Where-Object Role -eq 'stage2').Id
        foreach($record in $records){
            $row=[IO.File]::ReadAllText($record.FullName).Split('|');$id=[int]$row[0];$ticks=[long]$row[1]
            $alive=$false;try{$p=[Diagnostics.Process]::GetProcessById($id);try{$alive=(-not $p.HasExited -and $p.StartTime.ToUniversalTime().Ticks -eq $ticks)}finally{$p.Dispose()}}catch [ArgumentException]{}
            $alive|Should -BeFalse -Because 'Only the recorded PID plus creation identity defines the owned process.'
            $handle=[IO.FileStream]::new((Join-Path $Root ($record.BaseName+'.held')),[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None);$handle.Dispose()
        }
        $Result.JobAssigned|Should -BeTrue;$Result.TreeTerminated|Should -BeTrue
        $observations.Add([pscustomobject]@{Kind='independent direct-parent/creation-identity tree proof';Root=$Root;Parents=$parents;Result=$Result;HeldHandlesReleased=$true})
    }
    function Invoke-LifetimeMagick([string[]]$Arguments){
        $literal=@('-define','registry:filename:literal=true')
        $tokens=if($Arguments[0] -eq 'identify'){@('identify')+$literal+$Arguments[1..($Arguments.Length-1)]}else{$literal+$Arguments}
        $r=Invoke-WinImgNativeProcess -Executable $magick -Arguments $tokens -TimeoutMilliseconds 15000
        $observations.Add([pscustomobject]@{Kind='actual native fixture/validation';Arguments=$tokens;Result=$r})
        if($r.ExitCode -ne 0 -or $r.StdErr -or -not $r.StreamsComplete -or $r.TimedOut){throw ('Lifetime native fixture/validation failed: '+$r.StdErr+' '+$r.StartError)}
        return $r.StdOut.TrimEnd("`r","`n")
    }
    $templates=New-LifetimeDirectory 'templates';$small=Join-Path $templates 'small.png';$large=Join-Path $templates 'large.png';$jpeg=Join-Path $templates 'valid.jpg'
    $null=Invoke-LifetimeMagick @('-size','16x12','xc:#E02020',('PNG:'+$small))
    $null=Invoke-LifetimeMagick @('-size','1024x768','gradient:#E02020-#2060E0',('PNG:'+$large))
    $null=Invoke-LifetimeMagick @($small,'-strip',('JPEG:'+$jpeg))
    $fixed=[DateTime]::Parse('2020-06-07T08:09:10Z').ToUniversalTime()
    function Write-LifetimeSource([string]$Path,[byte[]]$Bytes){
        $null=[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($Path));[IO.File]::WriteAllBytes($Path,$Bytes)
        [IO.File]::SetCreationTimeUtc($Path,$fixed);[IO.File]::SetLastWriteTimeUtc($Path,$fixed)
    }
    function Set-LifetimeFixtureDirectoryTimes([string]$Root){
        # Complete all fixture writes before taking a stable source baseline.
        # Set children first, then their parents; keep exact directory checks.
        foreach($path in @(@(Get-ChildItem -LiteralPath $Root -Recurse -Force -Directory|Sort-Object {$_.FullName.Length} -Descending|ForEach-Object{$_.FullName})+$Root)){
            [IO.Directory]::SetCreationTimeUtc($path,$fixed);[IO.Directory]::SetLastWriteTimeUtc($path,$fixed)
        }
    }
    function Get-LifetimeState([string]$Root){
        $items=@(Get-ChildItem -LiteralPath $Root -Recurse -Force|Sort-Object FullName|ForEach-Object{
            if($_.Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Lifetime state must not traverse a link.'}
            $_.Refresh()
            [pscustomobject]@{Path=$_.FullName.Substring($Root.Length);Directory=$_.PSIsContainer;Creation=$_.CreationTimeUtc.Ticks;Modified=$_.LastWriteTimeUtc.Ticks;Attributes=[int]$_.Attributes;Length=if($_.PSIsContainer){$null}else{$_.Length};Sha256=if($_.PSIsContainer){$null}else{(Get-FileHash -LiteralPath $_.FullName).Hash.ToLowerInvariant()}}
        })
        $rootItem=[IO.DirectoryInfo]::new($Root);$rootItem.Refresh()
        $json=([pscustomobject]@{RootCreation=$rootItem.CreationTimeUtc.Ticks;RootModified=$rootItem.LastWriteTimeUtc.Ticks;Entries=$items}|ConvertTo-Json -Depth 8 -Compress)
        $script:stateSerial++;$path=Join-Path $ownedRoot ('source-state-{0:d3}.json' -f $script:stateSerial)
        [IO.File]::WriteAllText($path,$json,[Text.UTF8Encoding]::new($false));return $json
    }
    function New-LifetimeCase([string]$Label,[switch]$LargeSource){
        $source=New-LifetimeDirectory ($Label+'/source');$parent=New-LifetimeDirectory ($Label+'/output');$control=New-LifetimeDirectory ($Label+'/control')
        Write-LifetimeSource (Join-Path $source 'a first.png') ([IO.File]::ReadAllBytes($(if($LargeSource){$large}else{$small})))
        Write-LifetimeSource (Join-Path $source 'z later/later.png') ([IO.File]::ReadAllBytes($small))
        $opaque=New-Object byte[] 4097;for($i=0;$i -lt $opaque.Length;$i++){$opaque[$i]=[byte](($i*31+19)%256)}
        Write-LifetimeSource (Join-Path $source 'z later/video.MP4') $opaque
        Write-LifetimeSource (Join-Path $source 'unsupported.dat') ([Text.Encoding]::UTF8.GetBytes('preserved synthetic unsupported source'))
        $null=[IO.Directory]::CreateDirectory((Join-Path $source 'empty'))
        Write-LifetimeSource (Join-Path $parent 'parent-sentinel.dat') ([Text.Encoding]::UTF8.GetBytes('pre-existing output parent'))
        Set-LifetimeFixtureDirectoryTimes $source;Set-LifetimeFixtureDirectoryTimes $parent
        return [pscustomobject]@{Source=$source;Parent=$parent;Control=$control;Before=(Get-LifetimeState $source);ParentBefore=(Get-LifetimeState $parent)}
    }
    function Invoke-LifetimeRun([object]$Case,[object]$Policy,[scriptblock]$Observer){
        $parameters=@{Source=$Case.Source;OutputParent=$Case.Parent;MagickPath=$magick;ExecutionPolicy=$Policy;NativeTemporaryRoot=$ownedRoot};if($Observer){$parameters.NativeProcessObserver=$Observer}
        $output=@(& Invoke-WinImgNormalizer @parameters 6>&1 3>&1 2>&1)
        $codes=@($output|Where-Object{$_ -is [int] -or $_ -is [long]});$codes.Count|Should -Be 1
        $runs=@(Get-ChildItem -LiteralPath $Case.Parent -Directory);$runs.Count|Should -Be 1
        $logs=@(Get-ChildItem -LiteralPath $runs[0].FullName -Recurse -Force -File|Where-Object Extension -eq '.log');$logs.Count|Should -Be 1
        $result=[pscustomobject]@{Code=$codes[0];Run=$runs[0].FullName;LogPath=$logs[0].FullName;Log=[IO.File]::ReadAllText($logs[0].FullName);Text=$output -join "`n"}
        $observations.Add([pscustomobject]@{Kind='application outcome';Source=$Case.Source;Policy=$Policy;Result=$result});return $result
    }
    function Assert-LifetimeMedia([object]$Case,[object]$Result,[int]$Errors=1,[switch]$RetainedScratch){
        $Result.Code|Should -Be 2
        $Result.Log|Should -Match ('SUMMARY ConvertedImages=1 CopiedVideos=1 Duplicates=0 Unsupported=1 Errors='+$Errors+' SizeWarnings=0 NativeWarnings=0')
        [IO.File]::Exists((Join-Path $Result.Run 'a first.jpeg'))|Should -BeFalse
        $later=Join-Path $Result.Run 'z later/later.jpeg'
        Invoke-LifetimeMagick @('identify','+ping','-regard-warnings','-format','%m|%w|%h|%n',(Get-WinImgNativeOutputPath $later))|Should -Be 'JPEG|16|12|1'
        $pixels=Invoke-LifetimeMagick @((Get-WinImgNativeOutputPath $later),'-format','%[fx:round(255*p{8,6}.r)]|%[fx:round(255*p{8,6}.g)]|%[fx:round(255*p{8,6}.b)]','info:')
        $rgb=@($pixels.Split('|')|ForEach-Object{[int]$_});[Math]::Abs($rgb[0]-224)|Should -BeLessOrEqual 12;foreach($i in 1,2){[Math]::Abs($rgb[$i]-32)|Should -BeLessOrEqual 12}
        $video=Join-Path $Result.Run 'z later/video.MP4';(Get-FileHash -LiteralPath $video).Hash|Should -Be (Get-FileHash -LiteralPath (Join-Path $Case.Source 'z later/video.MP4')).Hash
        [IO.Directory]::Exists((Join-Path $Result.Run 'empty'))|Should -BeTrue
        foreach($p in @($later,$video)){([IO.FileInfo]::new($p)).CreationTimeUtc.Ticks|Should -Be $fixed.Ticks;([IO.FileInfo]::new($p)).LastWriteTimeUtc.Ticks|Should -Be $fixed.Ticks}
        Get-LifetimeState $Case.Source|Should -Be $Case.Before
        [IO.File]::ReadAllText((Join-Path $Case.Parent 'parent-sentinel.dat'))|Should -Be 'pre-existing output parent'
        $parentEntry=($Case.ParentBefore|ConvertFrom-Json).Entries[0];$sentinel=[IO.FileInfo]::new((Join-Path $Case.Parent 'parent-sentinel.dat'))
        $sentinel.CreationTimeUtc.Ticks|Should -Be $parentEntry.Creation;$sentinel.LastWriteTimeUtc.Ticks|Should -Be $parentEntry.Modified
        $media=@(Get-ChildItem -LiteralPath $Result.Run -Recurse -Force -File|Where-Object{$_.Extension -notin @('.log','.csv') -and (-not $RetainedScratch -or -not $_.FullName.StartsWith((Join-Path $Result.Run '.WinImgNormalizer')+'\',[StringComparison]::OrdinalIgnoreCase))}|ForEach-Object{$_.FullName.Substring($Result.Run.Length+1)}|Sort-Object)
        ($media -join '|')|Should -Be 'z later\later.jpeg|z later\video.MP4'
        if(-not $RetainedScratch){@(Get-ChildItem -LiteralPath (Join-Path $Result.Run '.WinImgNormalizer/work') -Force).Count|Should -Be 0}
    }
}

AfterAll {
    # Emergency fixture cleanup is scoped to exact outside Process objects. An
    # assertion failure must not leave the observer or unrelated test probe alive.
    if($outsideProcesses){foreach($entry in $outsideProcesses){try{Stop-LifetimeOutside $entry}catch{};try{$entry.Process.Dispose()}catch{}}}
    if($ownedRoot -and $observations){
        foreach($record in @(Get-ChildItem -LiteralPath $ownedRoot -Recurse -Filter '*.identity' -File)){
            try{$row=[IO.File]::ReadAllText($record.FullName).Split('|');$p=[Diagnostics.Process]::GetProcessById([int]$row[0])
                try{if(-not $p.HasExited -and $p.StartTime.ToUniversalTime().Ticks -eq [long]$row[1]){$observations.Add([pscustomobject]@{Kind='emergency cleanup after failed owned-tree assertion';Pid=$p.Id;StartTicks=[long]$row[1]});$p.Kill();$null=$p.WaitForExit(5000)}}finally{$p.Dispose()}
            }catch [ArgumentException]{}
        }
    }
    if($ownedRoot -and $observations){
        $bindingsAfter=@(Get-LifetimeBindings);$environmentAfter=@{};foreach($key in $environmentBefore.Keys){$environmentAfter[$key]=[Environment]::GetEnvironmentVariable($key,'Process')}
        $artifacts=@(Get-ChildItem -LiteralPath $ownedRoot -Recurse -Force -File|ForEach-Object{[pscustomobject]@{Path=$_.FullName.Substring($repository.Length+1).Replace('\','/');Sha256=(Get-FileHash -LiteralPath $_.FullName).Hash.ToLowerInvariant();Length=$_.Length}})
        $evidence=[pscustomobject]@{Task='M3-T02';PowerShell=$PSVersionTable.PSVersion.ToString();Edition=$PSVersionTable.PSEdition;NativeVersion=$nativeVersion;BindingsBefore=$bindingsBefore;BindingsAfter=$bindingsAfter;BindingsUnchanged=(($bindingsBefore|ConvertTo-Json -Compress) -eq ($bindingsAfter|ConvertTo-Json -Compress));EnvironmentBefore=$environmentBefore;EnvironmentAfter=$environmentAfter;PolicyFilesBefore=$policyFiles;Observations=$observations.ToArray();Artifacts=$artifacts;Qualification='Synthetic process lifetime/resource controls; no hostile media, global policy changes, or real Ctrl+C claim.'}
        [IO.File]::WriteAllText((Join-Path $ownedRoot 'native-lifetime-observations.json'),($evidence|ConvertTo-Json -Depth 20),[Text.UTF8Encoding]::new($false))
        Write-Host ('Native lifetime owned artifacts: '+$ownedRoot)
    }
}

Describe 'M3-T02 owned native lifetime and process-local resource regression' {
    It 'T051 creates a finite readonly policy and refuses a caller attempting to raise a ceiling' {
        $policy=New-WinImgExecutionPolicy
        $policy.TimeoutMilliseconds|Should -Be 120000;$policy.MemoryBytes|Should -Be 536870912;$policy.MapBytes|Should -Be 1073741824;$policy.DiskBytes|Should -Be 2147483648;$policy.Threads|Should -Be 2
        {$policy['Threads']=1}|Should -Throw
        {New-WinImgExecutionPolicy -TimeoutMilliseconds 120001}|Should -Throw
        {New-WinImgExecutionPolicy -MemoryBytes 536870913}|Should -Throw
        {New-WinImgExecutionPolicy -MapBytes 1073741825}|Should -Throw
        {New-WinImgExecutionPolicy -DiskBytes 2147483649}|Should -Throw
        {New-WinImgExecutionPolicy -Threads 3}|Should -Throw
        {New-WinImgExecutionPolicy -TimeoutMilliseconds 0}|Should -Throw
    }

    It 'T051 terminates the independently observed three-process owned tree after actual <Mode> and releases inherited pipes within the budget' -ForEach @(@{Mode='sleep'},@{Mode='exitroot'}) {
        $root=New-LifetimeDirectory ('wrapper-'+$Mode);$observer=Start-LifetimeObserver $root
        try{
            $clock=[Diagnostics.Stopwatch]::StartNew()
            $r=Invoke-WinImgNativeProcess -Executable $fixture -Arguments @('tree',$root,$Mode,'1','0','-','-') -TimeoutMilliseconds 2500
            $clock.Stop();$observations.Add([pscustomobject]@{Kind='actual wrapper lifetime result';Mode=$Mode;Result=$r;ObservedElapsedMilliseconds=$clock.ElapsedMilliseconds})
            if($Mode -eq 'sleep'){$r.TimedOut|Should -BeTrue}else{$r.TimedOut|Should -BeFalse;$r.ExitCode|Should -Be 0;$r.StreamsComplete|Should -BeTrue}
            $clock.ElapsedMilliseconds|Should -BeLessThan 6500
            $r.ElapsedMilliseconds|Should -BeLessThan 6500
        }finally{Complete-LifetimeObserver $observer $root}
        Assert-LifetimeTreeGone $root $r
        (Get-WinImgNativeOutcome -Result $r).Category|Should -Be $(if($Mode -eq 'sleep'){'Timeout'}else{'Success'})
    }

    It 'T051 bounds both streams and rejects timeout despite a genuine JPEG written before the benign child stalls' {
        $root=New-LifetimeDirectory 'flood-timeout';$candidate=Join-Path $root 'already-valid.jpeg';$observer=Start-LifetimeObserver $root
        try{$r=Invoke-WinImgNativeProcess -Executable $fixture -Arguments @('tree',$root,'sleep','1','512',$jpeg,$candidate) -OutputLimit 4096 -TimeoutMilliseconds 2500}finally{Complete-LifetimeObserver $observer $root}
        $observations.Add([pscustomobject]@{Kind='actual flood timeout result';Result=$r;WrittenJpeg=$candidate})
        $r.TimedOut|Should -BeTrue;$r.StdOut.Length|Should -BeLessOrEqual 4096;$r.StdErr.Length|Should -BeLessOrEqual 4096
        $r.StdOutCharacters|Should -BeGreaterThan 4096;$r.StdErrCharacters|Should -BeGreaterThan 4096;$r.StdOutTruncated|Should -BeTrue;$r.StdErrTruncated|Should -BeTrue
        (Get-WinImgNativeOutcome -Result $r).Acceptable|Should -BeFalse
        Invoke-LifetimeMagick @('identify','+ping','-regard-warnings','-format','%m|%w|%h',$candidate)|Should -Be 'JPEG|16|12'
        Assert-LifetimeTreeGone $root $r
    }

    It 'T053 leaves an actual unrelated ImageMagick stdin process alive and able to finish its own synthetic image after owned-tree timeout' {
        $root=New-LifetimeDirectory 'unrelated-isolation';$unrelatedOutput=Join-Path $root 'unrelated.png'
        # The first image is the output oracle. RGB stdin only blocks until the
        # synchronization payload arrives; discard every decoded stdin image,
        # including one produced by a host StreamWriter encoding preamble.
        $outside=Start-LifetimeOutside $magick @('-define','registry:filename:literal=true','-size','1x1','xc:#E02020','-depth','8','RGB:-','-delete','1--1',('PNG:'+$unrelatedOutput)) -RedirectInput
        $outside.Process.WaitForExit(150)|Should -BeFalse
        $observer=Start-LifetimeObserver $root
        try{
            $r=Invoke-WinImgNativeProcess -Executable $fixture -Arguments @('tree',$root,'sleep','1','0','-','-') -TimeoutMilliseconds 2500
            $outside.Process.HasExited|Should -BeFalse;$outside.Process.StartTime.ToUniversalTime().Ticks|Should -Be $outside.StartTicks
            $stream=$outside.Process.StandardInput.BaseStream;$stream.Write([byte[]]@(224,32,32),0,3);$stream.Flush();$stream.Close()
            $outside.Process.WaitForExit(5000)|Should -BeTrue;$outside.Process.ExitCode|Should -Be 0
        }finally{Complete-LifetimeObserver $observer $root;Stop-LifetimeOutside $outside}
        $r.TimedOut|Should -BeTrue;Assert-LifetimeTreeGone $root $r
        $unrelatedPngs=@(Get-ChildItem -LiteralPath $root -File -Filter '*.png');$unrelatedPngs.Count|Should -Be 1;$unrelatedPngs[0].Name|Should -Be 'unrelated.png'
        Invoke-LifetimeMagick @('identify','+ping','-regard-warnings','-format','%m|%w|%h',$unrelatedOutput)|Should -Be 'PNG|1|1'
        Invoke-LifetimeMagick @($unrelatedOutput,'-format','%[fx:round(255*p{0,0}.r)]|%[fx:round(255*p{0,0}.g)]|%[fx:round(255*p{0,0}.b)]','info:')|Should -Be '224|32|32'
        $observations.Add([pscustomobject]@{Kind='actual unrelated pinned ImageMagick isolation';Pid=$outside.Id;StartTicks=$outside.StartTicks;SynchronizationPayloadBytes=3;InputQualification='Raw stdin payload is synchronization only; host encoding preamble is not measured here. Output retains the independent first xc image.';Survived=$true;CompletedExit=0;IntendedPngCount=$unrelatedPngs.Count;Output=$unrelatedOutput})
    }

    It 'T052 passes owned temporary environment only to its native child and leaves the host and policy bytes unchanged' {
        $temporary=New-LifetimeDirectory 'child-environment';$childEnvironment=@{TEMP=$temporary;TMP=$temporary;MAGICK_TEMPORARY_PATH=$temporary}
        $r=Invoke-WinImgNativeProcess -Executable $fixture -Arguments @('environment') -Environment $childEnvironment -TimeoutMilliseconds 2500
        $r.ExitCode|Should -Be 0;$r.JobAssigned|Should -BeTrue;$r.StreamsComplete|Should -BeTrue
        foreach($key in $childEnvironment.Keys){$r.StdOut|Should -Match ([regex]::Escape($key+'='+$temporary));[Environment]::GetEnvironmentVariable($key,'Process')|Should -Be $environmentBefore[$key]}
        foreach($file in $policyFiles){(Get-FileHash -LiteralPath $file.Path).Hash.ToLowerInvariant()|Should -Be $file.Sha256}
        $observations.Add([pscustomobject]@{Kind='actual child-local temporary environment';ChildResult=$r;HostUnchanged=$true;InstalledPolicyFilesUnchanged=$true})
    }

    It 'T051 uses one shared stopwatch across phases and refuses a new phase after it expires' {
        $work=New-LifetimeDirectory 'shared-helper/work';$context=New-WinImgImageContext -WorkRoot $work -TemporaryRoot $ownedRoot -Policy (New-WinImgExecutionPolicy -TimeoutMilliseconds 700)
        try{
            $first=Get-WinImgRemainingTime $context;$first|Should -BeGreaterThan 0;$first|Should -BeLessOrEqual 700
            $r=Invoke-WinImgNativeProcess -Executable $fixture -Arguments @('finite','300') -TimeoutMilliseconds $first;$r.ExitCode|Should -Be 0
            $second=Get-WinImgRemainingTime $context;$second|Should -BeLessThan $first
            $r=Invoke-WinImgNativeProcess -Executable $fixture -Arguments @('finite','1000') -TimeoutMilliseconds $second;$r.TimedOut|Should -BeTrue
            {Get-WinImgRemainingTime $context}|Should -Throw '*runtime budget exhausted*'
            $observations.Add([pscustomobject]@{Kind='actual shared helper deadline';FirstRemaining=$first;SecondRemaining=$second;FinalNative=$r})
        }finally{Remove-WinImgImageContext $context}
        @(Get-ChildItem -LiteralPath $work -Force).Count|Should -Be 0
    }

    It 'T051 rejects an actual stalled <Phase> and still normalizes the later JPEG/video with source and namespace preservation' -ForEach @(@{Phase='inspection'},@{Phase='conversion'},@{Phase='validation'}) {
        $case=New-LifetimeCase ('phase-'+$Phase);$targetPhase=$Phase;$seen=New-Object 'Collections.Generic.List[object]';$routing=@{Done=$false};$observer=Start-LifetimeObserver $case.Control
        $delegate={param($executable,[string[]]$arguments,[int]$remaining,$environment)
            $isConvert=$arguments[-1].StartsWith('JPEG:',[StringComparison]::Ordinal)
            $isValidation=($arguments[0] -eq 'identify' -and $arguments -contains '+ping')
            $route=(-not $routing.Done -and (($targetPhase -eq 'inspection' -and $seen.Count -eq 0) -or ($targetPhase -eq 'conversion' -and $isConvert) -or ($targetPhase -eq 'validation' -and $isValidation)))
            $call=[pscustomobject]@{Arguments=$arguments.Clone();Remaining=$remaining;Environment=$environment;Routed=$route;Result=$null};$seen.Add($call)
            if($route){$routing.Done=$true;$copyFrom='-';$copyTo='-';if($isConvert){$copyFrom=$jpeg;$copyTo=$arguments[-1].Substring(5)}
                $call.Result=Invoke-WinImgNativeProcess -Executable $fixture -Arguments @('tree',$case.Control,'sleep','1','0',$copyFrom,$copyTo) -TimeoutMilliseconds $remaining -Environment $environment
            }else{$call.Result=Invoke-WinImgNativeProcess -Executable $executable -Arguments $arguments -TimeoutMilliseconds $remaining -Environment $environment}
            return $call.Result
        }
        try{$r=Invoke-LifetimeRun $case (New-WinImgExecutionPolicy -TimeoutMilliseconds 4000) $delegate}finally{Complete-LifetimeObserver $observer $case.Control}
        $r.Log|Should -Match 'Category=Timeout';@($seen|Where-Object Routed).Count|Should -Be 1
        foreach($call in $seen){$call.Arguments|Should -Contain '-limit';$call.Arguments|Should -Contain '536870912B';$call.Arguments|Should -Contain '1073741824B';$call.Arguments|Should -Contain '2147483648B';$call.Remaining|Should -BeGreaterThan 0;$call.Remaining|Should -BeLessOrEqual 4000
            foreach($key in @('TEMP','TMP','MAGICK_TEMPORARY_PATH')){$call.Environment[$key]|Should -Not -BeNullOrEmpty;$call.Environment[$key].StartsWith($ownedRoot+'\WinImgNormalizer-cache-',[StringComparison]::OrdinalIgnoreCase)|Should -BeTrue;$call.Environment[$key].StartsWith('\\?\',[StringComparison]::Ordinal)|Should -BeFalse;$call.Environment[$key].Length|Should -BeLessOrEqual 215}
            $call.Environment.TEMP|Should -Be $call.Environment.TMP;$call.Environment.TEMP|Should -Be $call.Environment.MAGICK_TEMPORARY_PATH
        }
        Assert-LifetimeTreeGone $case.Control ($seen|Where-Object Routed).Result
        $observations.Add([pscustomobject]@{Kind='controlled phase routing to actual benign sleeper, actual later media';Phase=$targetPhase;Calls=$seen.ToArray();Control=$case.Control})
        Assert-LifetimeMedia $case $r
    }

    It 'T052 fails a benign large cache allocation under tiny process-local ceilings, then completes a tiny image and opaque video without global changes' {
        $case=New-LifetimeCase 'resource-exhaustion' -LargeSource;$calls=New-Object 'Collections.Generic.List[object]'
        $delegate={param($executable,[string[]]$arguments,[int]$remaining,$environment)
            $native=Invoke-WinImgNativeProcess -Executable $executable -Arguments $arguments -TimeoutMilliseconds $remaining -Environment $environment
            $calls.Add([pscustomobject]@{Arguments=$arguments.Clone();Remaining=$remaining;Environment=$environment;Result=$native});return $native
        }
        $r=Invoke-LifetimeRun $case (New-WinImgExecutionPolicy -TimeoutMilliseconds 15000 -MemoryBytes 131072 -MapBytes 0 -DiskBytes 0 -Threads 1) $delegate
        $r.Log|Should -Match 'Category=ResourceExhaustion'
        @($calls|Where-Object{(Get-WinImgNativeOutcome -Result $_.Result -AllowStdOut).Category -eq 'ResourceExhaustion'}).Count|Should -BeGreaterThan 0
        foreach($call in $calls){$call.Arguments|Should -Contain '131072B';$call.Arguments|Should -Contain '0B';$call.Environment.MAGICK_TEMPORARY_PATH|Should -Be $call.Environment.TEMP;$call.Environment.TMP|Should -Be $call.Environment.TEMP}
        foreach($key in $environmentBefore.Keys){[Environment]::GetEnvironmentVariable($key,'Process')|Should -Be $environmentBefore[$key]}
        foreach($file in $policyFiles){(Get-FileHash -LiteralPath $file.Path).Hash.ToLowerInvariant()|Should -Be $file.Sha256}
        $observations.Add([pscustomobject]@{Kind='actual benign ImageMagick cache exhaustion with later siblings';Calls=$calls.ToArray();NoGlobalEnvironmentOrPolicyMutation=$true})
        Assert-LifetimeMedia $case $r
    }

    It 'T051 exhausts one cumulative deadline across real inspection/profile phases instead of renewing the allowance for conversion' {
        $case=New-LifetimeCase 'cumulative-phases'
        # A genuine tagged input makes colour inspection, ICC extraction and
        # conversion distinct native phases within the SAME image context.
        $first=Join-Path $case.Source 'a first.png'
        $null=Invoke-LifetimeMagick @($small,'-profile',(Join-Path $PSScriptRoot 'fixtures/colour/sRGB-v4.icc'),('PNG:'+$first))
        [IO.File]::SetCreationTimeUtc($first,$fixed);[IO.File]::SetLastWriteTimeUtc($first,$fixed);Set-LifetimeFixtureDirectoryTimes $case.Source;$case.Before=Get-LifetimeState $case.Source
        $calls=New-Object 'Collections.Generic.List[object]';$routing=@{TemporaryPath=$null;Ordinal=0}
        $delegate={param($executable,[string[]]$arguments,[int]$remaining,$environment)
            if(-not $routing.TemporaryPath){$routing.TemporaryPath=$environment.TEMP}
            $target=($environment.TEMP -eq $routing.TemporaryPath)
            $call=[pscustomobject]@{Arguments=$arguments.Clone();Remaining=$remaining;Environment=$environment;Target=$target;DelayResult=$null;Result=$null};$calls.Add($call)
            if($target){
                $routing.Ordinal++;$delay=if($routing.Ordinal -le 2){350}else{5000};$clock=[Diagnostics.Stopwatch]::StartNew()
                $call.DelayResult=Invoke-WinImgNativeProcess -Executable $fixture -Arguments @('finite',[string]$delay) -TimeoutMilliseconds $remaining -Environment $environment
                if($call.DelayResult.TimedOut){$call.Result=$call.DelayResult;return $call.Result}
                $remaining=[Math]::Max(1,$remaining-[int]$clock.ElapsedMilliseconds)
            }
            $call.Result=Invoke-WinImgNativeProcess -Executable $executable -Arguments $arguments -TimeoutMilliseconds $remaining -Environment $environment;return $call.Result
        }
        $r=Invoke-LifetimeRun $case (New-WinImgExecutionPolicy -TimeoutMilliseconds 3000) $delegate
        $firstCalls=@($calls|Where-Object Target);$firstCalls.Count|Should -Be 3
        $firstCalls[0].Remaining|Should -BeLessOrEqual 3000
        $firstCalls[1].Remaining|Should -BeLessThan ($firstCalls[0].Remaining-250)
        $firstCalls[2].Remaining|Should -BeLessThan ($firstCalls[1].Remaining-250)
        $firstCalls[0].Arguments[0]|Should -Be 'identify';$firstCalls[1].Arguments[-1]|Should -Match '^ICC:'
        $firstCalls[2].Arguments[-1]|Should -Match '^JPEG:';$firstCalls[2].Result.TimedOut|Should -BeTrue
        $r.Log|Should -Match 'Category=Timeout'
        $observations.Add([pscustomobject]@{Kind='controlled benign process delays, actual queries and cumulative deadline outcome';Calls=$calls.ToArray();Qualification='The observer deliberately adds native finite delays; actual wrapper outcomes and unchanged native queries are used, with no fabricated success/timeout result.'})
        Assert-LifetimeMedia $case $r
    }

    It 'T052 preserves an unexpected neighboring file while exact owned native cache cleanup refuses the occupied directory' {
        $work=New-LifetimeDirectory 'foreign-cache/work';$context=New-WinImgImageContext -WorkRoot $work -TemporaryRoot $ownedRoot
        $foreign=Join-Path $context.TemporaryPath 'unrelated-arrival.dat';[IO.File]::WriteAllText($foreign,'foreign arrival retained')
        $hash=(Get-FileHash -LiteralPath $foreign).Hash
        {Remove-WinImgImageContext $context}|Should -Throw '*unexpected entry*'
        [IO.Directory]::Exists($context.TemporaryPath)|Should -BeTrue
        (Get-FileHash -LiteralPath $foreign).Hash|Should -Be $hash
        $observations.Add([pscustomobject]@{Kind='controlled foreign cache neighbor';Preserved=$foreign;Sha256=$hash;NoRecursiveCleanup=$true})
    }

    It 'T052 respects a stricter owned child-only resource policy despite a larger command-line memory limit' {
        $private=New-LifetimeDirectory 'restrictive-policy';$portable=[IO.Path]::GetDirectoryName($magick)
        foreach($file in @(Get-ChildItem -LiteralPath $portable -Filter '*.xml' -File -Force)){
            [IO.File]::WriteAllBytes((Join-Path $private $file.Name),[IO.File]::ReadAllBytes($file.FullName))
        }
        $policy=Join-Path $private 'policy.xml';$original=[IO.File]::ReadAllText($policy)
        $addition='  <policy domain="resource" name="memory" value="1048576B"/>'
        $original.Contains('</policymap>')|Should -BeTrue
        # Preserve every original domain/rights restriction; add only a stricter
        # memory ceiling to a PRIVATE copy, never replace the installed policy.
        $stricter=$original.Replace('</policymap>',($addition+"`n"+'</policymap>'))
        $stricter.Replace(($addition+"`n"),'')|Should -Be $original
        [IO.File]::WriteAllText($policy,$stricter,[Text.UTF8Encoding]::new($false))
        $before=Invoke-WinImgNativeProcess -Executable $magick -Arguments @('-list','policy') -TimeoutMilliseconds 15000
        $childPath=$private;if($environmentBefore.MAGICK_CONFIGURE_PATH){$childPath+=';'+$environmentBefore.MAGICK_CONFIGURE_PATH}
        $childEnvironment=@{MAGICK_CONFIGURE_PATH=$childPath}
        $r=Invoke-WinImgNativeProcess -Executable $magick -Arguments @('identify','-limit','memory','536870912B','-list','resource') -Environment $childEnvironment -TimeoutMilliseconds 15000
        $r.ExitCode|Should -Be 0;$r.StdErr|Should -BeNullOrEmpty;$r.StreamsComplete|Should -BeTrue
        $r.StdOut|Should -Match '(?m)^\s*Memory:\s+1MiB\s*$'
        $applied=Invoke-WinImgNativeProcess -Executable $magick -Arguments @('-list','policy') -Environment $childEnvironment -TimeoutMilliseconds 15000
        $applied.ExitCode|Should -Be 0;$applied.StdOut|Should -Match '1048576B'
        $after=Invoke-WinImgNativeProcess -Executable $magick -Arguments @('-list','policy') -TimeoutMilliseconds 15000
        $after.ExitCode|Should -Be 0;$after.StdOut|Should -Be $before.StdOut
        foreach($file in $policyFiles){(Get-FileHash -LiteralPath $file.Path).Hash.ToLowerInvariant()|Should -Be $file.Sha256}
        [Environment]::GetEnvironmentVariable('MAGICK_CONFIGURE_PATH','Process')|Should -Be $environmentBefore.MAGICK_CONFIGURE_PATH
        $observations.Add([pscustomobject]@{Kind='actual stricter child-only policy limit';PrivateConfiguration=$private;NativeLimit=$r;AppliedPolicy=$applied;InstalledPolicyBefore=$before;InstalledPolicyAfter=$after;RequestedMemoryBytes=536870912;ActualPolicyMemoryBytes=1048576;Qualification='Owned portable configuration copies preserve original restrictions and add only resource memory. This does not alter or certify arbitrary machine policies.'})
    }

    It 'T052 explicitly refuses an overlong owned cache root before allocating a native context' {
        $longRoot=New-LifetimeDirectory ('overlong-cache-'+('x'*100));$work=New-LifetimeDirectory 'overlong-cache-work'
        $longRoot.Length|Should -BeGreaterThan 160;$before=Get-LifetimeState $longRoot
        {New-WinImgImageContext -WorkRoot $work -TemporaryRoot $longRoot}|Should -Throw '*temporary root is too long*'
        Get-LifetimeState $longRoot|Should -Be $before
        @(Get-ChildItem -LiteralPath $work -Force).Count|Should -Be 0
        $observations.Add([pscustomobject]@{Kind='owned cache path budget refusal';RootLength=$longRoot.Length;NativeLaunchAttempted=$false;SourceOrPolicyMutation=$false})
    }

    It 'T051 refuses finalization and preserves affected scratch for a controlled unconfirmed tree result while later real siblings complete' {
        $case=New-LifetimeCase 'unconfirmed-tree';$routing=@{Done=$false;Candidate=$null;TemporaryPath=$null};$calls=New-Object 'Collections.Generic.List[object]'
        $delegate={param($executable,[string[]]$arguments,[int]$remaining,$environment)
            $native=Invoke-WinImgNativeProcess -Executable $executable -Arguments $arguments -TimeoutMilliseconds $remaining -Environment $environment
            $route=(-not $routing.Done -and $arguments[-1].StartsWith('JPEG:',[StringComparison]::Ordinal))
            $originalTreeTerminated=$native.TreeTerminated
            if($route){$routing.Done=$true;$routing.Candidate=$arguments[-1].Substring(5);$routing.TemporaryPath=$environment.TEMP;$native.TreeTerminated=$false}
            $calls.Add([pscustomobject]@{Arguments=$arguments.Clone();Result=$native;ControlledUnconfirmedFlag=$route;ActualTreeTerminatedBeforeControl=$originalTreeTerminated});return $native
        }
        $r=Invoke-LifetimeRun $case (New-WinImgExecutionPolicy -TimeoutMilliseconds 15000) $delegate
        $r.Log|Should -Match 'tree termination is unconfirmed';$r.Log|Should -Match 'Category=NativeFailure'
        @($calls|Where-Object ControlledUnconfirmedFlag).Count|Should -Be 1
        ($calls|Where-Object ControlledUnconfirmedFlag).ActualTreeTerminatedBeforeControl|Should -BeTrue
        [IO.File]::Exists($routing.Candidate)|Should -BeTrue;[IO.Directory]::Exists($routing.TemporaryPath)|Should -BeTrue
        Invoke-LifetimeMagick @('identify','+ping','-regard-warnings','-format','%m|%w|%h',$routing.Candidate)|Should -Be 'JPEG|16|12'
        $retained=@(Get-ChildItem -LiteralPath (Join-Path $r.Run '.WinImgNormalizer/work') -Recurse -File -Force);$retained.Count|Should -Be 2
        @($retained|Where-Object Name -eq 'source.png').Count|Should -Be 1;@($retained|Where-Object Name -eq 'image.jpeg').Count|Should -Be 1
        $observations.Add([pscustomobject]@{Kind='controlled unconfirmed termination guard';ActualKillFailed=$false;Calls=$calls.ToArray();PreservedCandidate=$routing.Candidate;PreservedTemporaryPath=$routing.TemporaryPath;Qualification='The actual native process exited and its tree was confirmed gone; only the returned TreeTerminated flag is changed to exercise fail-closed finalization and cleanup.'})
        Assert-LifetimeMedia $case $r -RetainedScratch
    }

    It 'T052 rejects both a source-equal and source-contained native cache root before creating any output' {
        $case=New-LifetimeCase 'source-cache-boundary'
        foreach($temporary in @($case.Source,(Join-Path $case.Source 'z later'))){
            $output=@(& Invoke-WinImgNormalizer -Source $case.Source -OutputParent $case.Parent -MagickPath $magick -NativeTemporaryRoot $temporary 6>&1 3>&1 2>&1)
            $codes=@($output|Where-Object{$_ -is [int] -or $_ -is [long]});$codes.Count|Should -Be 1;$codes[0]|Should -Be 1
            ($output -join "`n")|Should -Match '(?i)(inside|equal).*source'
            Get-LifetimeState $case.Source|Should -Be $case.Before;Get-LifetimeState $case.Parent|Should -Be $case.ParentBefore
            $observations.Add([pscustomobject]@{Kind='actual source/cache physical containment rejection';TemporaryRoot=$temporary;ApplicationExit=$codes[0];OutputCreated=$false;SourcePreserved=$true})
        }
    }
}
