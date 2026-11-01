# Development-only real distribution/Windows-shell acceptance. No desktop settings
# are changed. Ordinary account/firewall controls exist only on disposable CI VMs.
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ArchivePath,
    [string]$SourceArchivePath,
    [Parameter(Mandatory)][string]$LegacyMagickPath,
    [string]$Revision,
    [string]$PublicTag,
    [string]$ResultDirectory = '.scratch/distribution-acceptance',
    [switch]$DisposableRunner,
    [switch]$Worker
)
$ErrorActionPreference = 'Stop'
if($Worker){
    # A PS7 parent can pass incompatible modules to Windows PowerShell 5.1.
    $moduleRoots=@([IO.Path]::Combine($PSHOME,'Modules'),[IO.Path]::Combine([Environment]::GetFolderPath('ProgramFiles'),'WindowsPowerShell/Modules'),[IO.Path]::Combine([Environment]::GetFolderPath('Windows'),'System32/WindowsPowerShell/v1.0/Modules'))
    [Environment]::SetEnvironmentVariable('PSModulePath',($moduleRoots -join ';'),'Process')
}
$repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$utf8 = [Text.UTF8Encoding]::new($false)
function Assert-Distribution([bool]$Value,[string]$Message) { if (-not $Value) { throw $Message } }
function Quote-Distribution([string]$Value) { '"' + [regex]::Replace([regex]::Replace($Value,'(\\*)"','$1$1\"'),'(\\+)$','$1$1') + '"' }
function Invoke-DistributionProcess([string]$Executable,[string[]]$Tokens,[string]$Work,[hashtable]$Environment,[int]$Timeout=180000) {
    $start=New-Object Diagnostics.ProcessStartInfo
    $start.FileName=$Executable; $start.Arguments=($Tokens | ForEach-Object { Quote-Distribution $_ }) -join ' '
    $start.WorkingDirectory=$Work; $start.UseShellExecute=$false; $start.CreateNoWindow=$true
    $start.WindowStyle=[Diagnostics.ProcessWindowStyle]::Hidden
    $start.RedirectStandardOutput=$true; $start.RedirectStandardError=$true; $start.RedirectStandardInput=$true
    $start.StandardOutputEncoding=$utf8; $start.StandardErrorEncoding=$utf8
    foreach($key in @($start.EnvironmentVariables.Keys)) {
        if([string]$key -ieq 'PSModulePath' -or [string]$key -like 'MAGICK_*' -or [string]$key -eq 'LIBHEIF_PLUGIN_PATH') { $start.EnvironmentVariables.Remove([string]$key) }
    }
    foreach($key in $Environment.Keys) { if($null -eq $Environment[$key]) {$start.EnvironmentVariables.Remove([string]$key)} else {$start.EnvironmentVariables[$key]=[string]$Environment[$key]} }
    $process=New-Object Diagnostics.Process; $process.StartInfo=$start
    try {
        Assert-Distribution ($process.Start()) 'Owned process failed to start.'
        $out=$process.StandardOutput.ReadToEndAsync(); $err=$process.StandardError.ReadToEndAsync(); $process.StandardInput.Close()
        if(-not $process.WaitForExit($Timeout)) {$process.Kill();throw 'Owned process deadline exceeded.'}
        Assert-Distribution ($out.Wait(5000) -and $err.Wait(5000)) 'Owned process streams incomplete.'
        [pscustomobject]@{Exit=$process.ExitCode;Out=$out.GetAwaiter().GetResult();Error=$err.GetAwaiter().GetResult()}
    } finally {$process.Dispose()}
}
function Invoke-DistributionGit([string[]]$Tokens,[string]$Work=$repository) {
    $result=Invoke-DistributionProcess $script:git (@('-c',('safe.directory='+$repository))+$Tokens) $Work @{}
    Assert-Distribution ($result.Exit -eq 0) ('Git fixture preparation failed: '+$result.Error)
    $result.Out.Trim()
}
function Test-DistributionOutbound {
    $socket=New-Object Net.Sockets.TcpClient
    try {$attempt=$socket.BeginConnect('github.com',443,$null,$null);if(-not $attempt.AsyncWaitHandle.WaitOne(4000)){return $false};try{$socket.EndConnect($attempt);return $socket.Connected}catch{return $false}}
    finally{$socket.Close()}
}
function Get-DistributionHash([string]$Path) { (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() }
function Write-DistributionJson([string]$Path,$Value) {[IO.File]::WriteAllText($Path,($Value|ConvertTo-Json -Depth 14),$utf8)}
function Get-DistributionCases {
    foreach($layout in @('portable','clone','source-archive','pull-upgrade')){
        foreach($pathMode in @('old','none')){
            foreach($entry in @('BAT-shortcut','PS7','BAT-exit')){[pscustomobject]@{Layout=$layout;PathMode=$pathMode;Entry=$entry;Negative=''}}
        }
    }
    [pscustomobject]@{Layout='bare';PathMode='old';Entry='BAT-shortcut';Negative=''}
    [pscustomobject]@{Layout='bare';PathMode='none';Entry='BAT-exit';Negative=''}
    [pscustomobject]@{Layout='bare';PathMode='old';Entry='BAT-shortcut';Negative='trusted-cache'}
    [pscustomobject]@{Layout='portable';PathMode='old';Entry='BAT-shortcut';Negative='corrupt-runtime'}
    [pscustomobject]@{Layout='source-archive';PathMode='old';Entry='BAT-shortcut';Negative='corrupt-payload'}
}
function Get-DistributionCaseLabel($Case) {$Case.Layout+'-'+$Case.PathMode+'-'+$Case.Entry+$(if($Case.Negative){'-'+$Case.Negative}else{''})}

Assert-Distribution ($env:OS -eq 'Windows_NT') 'Actual Windows is required.'
$root=[IO.Path]::GetFullPath((Join-Path $repository $ResultDirectory))
$scratch=[IO.Path]::GetFullPath((Join-Path $repository '.scratch'))
Assert-Distribution ($root.StartsWith($scratch+'\',[StringComparison]::OrdinalIgnoreCase)) 'Evidence must be in ignored owned scratch.'
$ancestor=[IO.DirectoryInfo]::new($root)
while($ancestor){Assert-Distribution (-not ($ancestor.Exists -and ($ancestor.Attributes -band [IO.FileAttributes]::ReparsePoint))) 'Evidence traverses a reparse point.';$ancestor=$ancestor.Parent}
[IO.Directory]::CreateDirectory($root)|Out-Null
$ArchivePath=[IO.Path]::GetFullPath($ArchivePath)
$LegacyMagickPath=[IO.Path]::GetFullPath($LegacyMagickPath)
$git=(Get-Command git.exe -CommandType Application -ErrorAction Stop|Select-Object -First 1).Source
if(-not $Revision){$Revision=Invoke-DistributionGit @('rev-parse','HEAD')}
Assert-Distribution ($Revision -match '^[0-9a-f]{40}$') 'Exact source revision is required.'
$ps5=Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe'
$ps7=(Get-Command pwsh.exe -CommandType Application -ErrorAction Stop|Select-Object -First 1).Source
$cmd=Join-Path $env:SystemRoot 'System32/cmd.exe'

if(-not $Worker) {
    if($PublicTag) {
        Assert-Distribution ($PublicTag -match '^v\d+\.\d+\.\d+$') 'Only a named stable public tag can be tested.'
        $SourceArchivePath=Join-Path $root 'actual-public-tag.zip'
        Invoke-WebRequest -UseBasicParsing -Uri ('https://github.com/PikkuJanne/WinImgNormalizer/archive/refs/tags/'+$PublicTag+'.zip') -OutFile $SourceArchivePath
    } elseif(-not $SourceArchivePath) {
        $SourceArchivePath=Join-Path $root 'candidate-tag-equivalent.zip'
        $null=Invoke-DistributionGit @('archive','--format=zip',('--output='+$SourceArchivePath),$Revision)
    }
    $SourceArchivePath=[IO.Path]::GetFullPath($SourceArchivePath)
    Assert-Distribution ([IO.File]::Exists($ArchivePath) -and [IO.File]::Exists($SourceArchivePath)) 'Fetch exact asset/source bytes before isolation.'
    Assert-Distribution ($DisposableRunner -and $env:GITHUB_ACTIONS -eq 'true' -and $env:RUNNER_ENVIRONMENT -eq 'github-hosted') 'Disposable account/firewall setup is restricted to an ephemeral hosted VM.'
    $userName='winimgdist'+[guid]::NewGuid().ToString('N').Substring(0,9)
    $password=New-Object Security.SecureString
    $random=New-Object byte[] 32;$rng=[Security.Cryptography.RandomNumberGenerator]::Create()
    try{$rng.GetBytes($random)}finally{$rng.Dispose()}
    foreach($character in ('W!9x'+[Convert]::ToBase64String($random)).ToCharArray()){$password.AppendChar($character)}
    $password.MakeReadOnly()
    $rules=New-Object 'Collections.Generic.List[string]';$created=$false;$originalAcl=$null
    try {
        Assert-Distribution (Test-DistributionOutbound) 'Pre-isolation credibility probe failed.'
        $user=New-LocalUser -Name $userName -Password $password -AccountNeverExpires -UserMayNotChangePassword;$created=$true
        Add-LocalGroupMember -SID 'S-1-5-32-545' -Member ([Environment]::MachineName+'\'+$userName) -ErrorAction SilentlyContinue
        $originalAcl=Get-Acl -LiteralPath $repository;$readAcl=Get-Acl -LiteralPath $repository
        $readAcl.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new($user.SID,'ReadAndExecute','ContainerInherit,ObjectInherit','None','Allow'));Set-Acl -LiteralPath $repository -AclObject $readAcl
        $writeAcl=Get-Acl -LiteralPath $root;$writeAcl.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new($user.SID,'FullControl','ContainerInherit,ObjectInherit','None','Allow'));Set-Acl -LiteralPath $root -AclObject $writeAcl
        # Future cache paths are fixed by the authenticated payload digest. Deny
        # these exact executable paths before any first run creates the cache;
        # no expanded runtime is staged for the source-layout acceptance tests.
        $payloadPin=[IO.File]::ReadAllText((Join-Path $repository 'resources/private-runtime.json'))|ConvertFrom-Json
        $canonical=[IO.File]::ReadAllText((Join-Path $repository 'release-metadata.json'))|ConvertFrom-Json
        $programs=@($ps5,$ps7,$cmd,$LegacyMagickPath,(Join-Path $env:SystemRoot 'System32/conhost.exe'),(Join-Path $root 'DistributionShellFixture.exe'))
        # PS5.1 Add-Type invokes the inbox C# compiler for the application's
        # native directory APIs. Its possible resource compiler is a child too.
        foreach($framework in @('Framework','Framework64')){
            foreach($compilerName in @('csc.exe','cvtres.exe')){
                $inboxCompiler=Join-Path $env:SystemRoot ('Microsoft.NET/'+$framework+'/v4.0.30319/'+$compilerName)
                if([IO.File]::Exists($inboxCompiler)){$programs+=$inboxCompiler}
            }
        }
        foreach($case in @(Get-DistributionCases)){
            $caseRoot=Join-Path $root (Get-DistributionCaseLabel $case)
            $programs+=Join-Path $caseRoot 'Application Ä 2026/runtime/imagemagick/magick.exe'
            $programs+=Join-Path $caseRoot ('Cold LocalAppData/WinImgNormalizer/runtimes/'+$canonical.version+'-'+$payloadPin.sha256+'/runtime/imagemagick/magick.exe')
        }
        foreach($program in @($programs|Select-Object -Unique)){
            $rule='WinImgDistribution-'+[guid]::NewGuid().ToString('N')
            New-NetFirewallRule -Name $rule -DisplayName $rule -Direction Outbound -Action Block -Program $program -Profile Any|Out-Null;$rules.Add($rule)
        }
        Write-DistributionJson (Join-Path $root 'isolation-control.json') @{WorkerSid=$user.SID.Value;HostedParentVerified=$true;FirewallRules=$rules.Count;ProgramPaths=@($programs|Select-Object -Unique);EveryTestedHostAndEngineOutboundBlocked=$true;CacheRulesAppliedBeforeFirstRun=$true;DownloadsCompletedBeforeIsolation=$true}
        $tokens=@('-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',$PSCommandPath,'-Worker','-ArchivePath',$ArchivePath,'-SourceArchivePath',$SourceArchivePath,'-LegacyMagickPath',$LegacyMagickPath,'-Revision',$Revision,'-ResultDirectory',$ResultDirectory)
        if($PublicTag){$tokens+=@('-PublicTag',$PublicTag)}
        $workerProcess=Start-Process -FilePath $ps5 -ArgumentList (($tokens|ForEach-Object {Quote-Distribution $_})-join ' ') -Credential ([Management.Automation.PSCredential]::new(([Environment]::MachineName+'\'+$userName),$password)) -LoadUserProfile -WindowStyle Hidden -WorkingDirectory $repository -RedirectStandardOutput (Join-Path $root 'worker.stdout.txt') -RedirectStandardError (Join-Path $root 'worker.stderr.txt') -PassThru
        $null=$workerProcess.Handle
        if(-not $workerProcess.WaitForExit(900000)){$workerProcess.Kill();throw 'Distribution worker exceeded deadline.'}
        if($workerProcess.ExitCode -ne 0){
            $summaryPath=Join-Path $root 'distribution-summary.json';$detail=''
            if([IO.File]::Exists($summaryPath)){
                $failure=[IO.File]::ReadAllText($summaryPath)|ConvertFrom-Json
                $detail='Line='+$failure.ErrorLine+'; completed cases='+@($failure.Cases).Count+'; '+$failure.Error
            }elseif([IO.File]::Exists((Join-Path $root 'worker.stderr.txt'))){$detail=[IO.File]::ReadAllText((Join-Path $root 'worker.stderr.txt'))}
            $detail=$detail.Replace($root,'<acceptance>').Replace($repository,'<checkout>').Replace($env:USERPROFILE,'<runner-profile>')
            $detail=[regex]::Replace($detail,'(?i)[A-Z]:[\\/]Users[\\/][^\\/\r\n" ]+','<disposable-profile>')
            if($detail.Length -gt 1800){$detail=$detail.Substring(0,700)+"`n[bounded]`n"+$detail.Substring($detail.Length-1000)}
            Write-Host ('Distribution worker cause: '+$detail)
        }
        Assert-Distribution ($workerProcess.ExitCode -eq 0) 'Distribution worker failed; inspect owned raw evidence.'
    } finally {
        foreach($rule in $rules){Remove-NetFirewallRule -Name $rule -ErrorAction Continue}
        if($originalAcl){Set-Acl -LiteralPath $repository -AclObject $originalAcl}
        if($created){Remove-LocalUser -Name $userName}
        $password.Dispose()
    }
    exit 0
}

Assert-Distribution ([IO.File]::Exists((Join-Path $root 'isolation-control.json'))) 'Exact positive BAT/source-archive acceptance requires its parent-created disposable isolation control; desktop Pictures are never used.'
$summary=[ordered]@{Schema=1;Passed=$false;SourceRevision=$Revision;PublicTag=$PublicTag;Cases=@();Windows10Executed=$false;Limitation='No authorized automated Windows 10 environment is available. Actual Windows Server build is recorded; desktop settings and output collections are untouched.'}
try {
    $os=Get-CimInstance Win32_OperatingSystem
    $identity=[Security.Principal.WindowsIdentity]::GetCurrent();$admin=[Security.Principal.WindowsPrincipal]::new($identity).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    $summary.Environment=@{OS=$os.Caption;Build=$os.BuildNumber;Shell=$PSVersionTable.PSVersion.ToString();Administrator=$admin}
    # PS5.1 CreateProcessWithLogon rebuilds the user's environment and may drop
    # GITHUB_ACTIONS/RUNNER_ENVIRONMENT. The parent verified both before creating
    # the account/rules; authenticate its actual disposable token here instead.
    $isolation=[IO.File]::ReadAllText((Join-Path $root 'isolation-control.json'))|ConvertFrom-Json
    Assert-Distribution ($isolation.HostedParentVerified -and $identity.User.Value -ceq $isolation.WorkerSid) 'Worker token does not match the disposable account created by the hosted parent.'
    Assert-Distribution (-not $admin -and $identity.Name -like '*\winimgdist*') 'Worker must be the ordinary disposable account.'
    $ownProfile=[Environment]::GetFolderPath('UserProfile')
    Assert-Distribution ($ownProfile -and $ownProfile -match '(?i)\\winimgdist[^\\]+$') 'The loaded profile must belong to this disposable account.'
    $env:USERPROFILE=$ownProfile;$env:USERNAME=($identity.Name -split '\\')[-1]
    # CreateProcessWithLogon loads this user's hive but can retain runneradmin's
    # TEMP/APPDATA environment. In particular csc cannot use that private TEMP.
    # Rebuild only this worker's process environment from its actual profile.
    $env:APPDATA=Join-Path $ownProfile 'AppData/Roaming'
    $env:LOCALAPPDATA=Join-Path $ownProfile 'AppData/Local'
    foreach($directory in @($env:APPDATA,$env:LOCALAPPDATA)){[IO.Directory]::CreateDirectory($directory)|Out-Null}
    $env:TEMP=$root;$env:TMP=$root
    Assert-Distribution (-not (Test-DistributionOutbound)) 'Outbound credibility probe unexpectedly succeeded.';$summary.OutboundBlocked=$true
    Assert-Distribution ($isolation.EveryTestedHostAndEngineOutboundBlocked -and $isolation.CacheRulesAppliedBeforeFirstRun -and $isolation.FirewallRules -gt 10) 'Disposable outbound isolation controls are absent.'
    $summary.Offline=@{FirewallRules=$isolation.FirewallRules;DownloadsCompletedBeforeIsolation=[bool]$isolation.DownloadsCompletedBeforeIsolation;ColdCachePathsBlockedBeforeCreation=$true;ActualApplicationChildrenCheckedAgainstBlockedPaths=$true}
    $summary.PortableSha256=Get-DistributionHash $ArchivePath;$summary.SourceArchiveSha256=Get-DistributionHash $SourceArchivePath
    $fixture=Join-Path $PSScriptRoot 'fixtures/distribution/Sample Kesä 2026';$expected=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'fixtures/distribution/expected.json'))|ConvertFrom-Json
    foreach($file in $expected.files){Assert-Distribution ((Get-DistributionHash (Join-Path $fixture $file.input)) -ceq $file.input_sha256) 'Original fixture does not match expected.json.'}
    Add-Type -AssemblyName System.Drawing
    $helper=Join-Path $root 'DistributionShellFixture.exe';$compiler=Join-Path $env:SystemRoot 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
    $compile=Invoke-DistributionProcess $compiler @('/nologo','/target:winexe','/reference:System.Web.Extensions.dll',('/out:'+$helper),(Join-Path $PSScriptRoot 'fixtures/DistributionShellFixture.cs')) $root @{}
    Assert-Distribution ($compile.Exit -eq 0) ('Shell helper compile failed: '+$compile.Out+$compile.Error)
    $summary.HelperSourceSha256=Get-DistributionHash (Join-Path $PSScriptRoot 'fixtures/DistributionShellFixture.cs')
    $systemPath=(Split-Path -Parent $ps5)+';'+(Join-Path $env:SystemRoot 'System32')
    $legacyHash=Get-DistributionHash $LegacyMagickPath
    Assert-Distribution ($legacyHash -ceq 'b98d9708e98f84625e308f3dcd3daedfb7399a2e6a2ddb2171b2070397b8ffb2') 'Real old module-enabled fixture executable differs from authenticated pin.'
    $old=Invoke-DistributionProcess $LegacyMagickPath @('-version') $root @{PATH=((Split-Path -Parent $LegacyMagickPath)+';'+$systemPath)}
    Assert-Distribution ($old.Exit -eq 0 -and $old.Out -match '7\.1\.2-5 Q16-HDRI x64' -and $old.Out -match 'Modules') 'Actual old DLL engine did not execute.'
    $summary.Legacy=@{Banner=$old.Out.Trim();Sha256=$legacyHash}
    $portableManifestRoot=Join-Path $root 'manifest-control';Expand-Archive -LiteralPath $ArchivePath -DestinationPath $portableManifestRoot
    $runtimeManifest=[IO.File]::ReadAllText((Join-Path $portableManifestRoot 'runtime/runtime-manifest.json'))|ConvertFrom-Json
    $enginePin=@($runtimeManifest.files|Where-Object path -CEQ $runtimeManifest.engine_relative_path)[0]
    $pictures=[Environment]::GetFolderPath('MyPictures');if(-not $pictures){$pictures=Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Pictures'}
    Assert-Distribution ($pictures.StartsWith($ownProfile+'\',[StringComparison]::OrdinalIgnoreCase)) 'Pictures does not belong to the disposable user.'
    [IO.Directory]::CreateDirectory($pictures)|Out-Null
    $desktop=[Environment]::GetFolderPath('Desktop')
    if(-not $desktop){$desktop=Join-Path $ownProfile 'Desktop'}
    Assert-Distribution ($desktop.StartsWith($ownProfile+'\',[StringComparison]::OrdinalIgnoreCase)) 'Desktop does not belong to the disposable user.'
    [IO.Directory]::CreateDirectory($desktop)|Out-Null
    function New-DistributionLayout([string]$Name,[string]$CaseRoot) {
        $app=Join-Path $CaseRoot 'Application Ä 2026'
        switch($Name){
            'portable'{Expand-Archive -LiteralPath $ArchivePath -DestinationPath $app}
            'source-archive'{
                $extract=Join-Path $CaseRoot 'source unzip';Expand-Archive -LiteralPath $SourceArchivePath -DestinationPath $extract
                if([IO.File]::Exists((Join-Path $extract 'WinImgNormalizer.bat'))){[IO.Directory]::Move($extract,$app)}
                else{$children=@(Get-ChildItem -LiteralPath $extract -Directory);Assert-Distribution ($children.Count -eq 1) 'Tagged source archive root is ambiguous.';[IO.Directory]::Move($children[0].FullName,$app)}
            }
            'clone'{$null=Invoke-DistributionGit @('clone','--no-local','--no-checkout',$repository,$app);$null=Invoke-DistributionGit @('-C',$app,'checkout','--detach',$Revision)}
            'pull-upgrade'{
                $null=Invoke-DistributionGit @('clone','--no-local','--no-checkout',$repository,$app)
                $null=Invoke-DistributionGit @('-C',$app,'checkout','-b','distribution-upgrade','v1.0.1')
            }
            'bare'{[IO.Directory]::CreateDirectory($app)|Out-Null;foreach($name in @('WinImgNormalizer.bat','WinImgNormalizer.ps1')){Copy-Item -LiteralPath (Join-Path $portableManifestRoot $name) -Destination (Join-Path $app $name)}}
            default{throw 'Unknown layout.'}
        }
        if($Name -in @('clone','source-archive','pull-upgrade')){Assert-Distribution (-not [IO.Directory]::Exists((Join-Path $app 'runtime')) -and -not [IO.File]::Exists((Join-Path $app 'portable-runtime.json'))) 'Source layout was improperly prepared with expanded runtime.'}
        $app
    }
    function Assert-DistributionOutputs([string]$Source,[string]$Output,[int]$Exit,[string]$Console) {
        Assert-Distribution ($Exit -eq $expected.expected_exit) 'Actual launcher exit differs from expected 0.'
        Assert-Distribution (@(Get-ChildItem -LiteralPath $Source -File).Count -eq 3) 'Synthetic source tree changed.'
        $media=@(Get-ChildItem -LiteralPath $Output -Recurse -File|Where-Object Extension -in @('.jpeg','.jpg','.png','.mp4'))
        Assert-Distribution ($media.Count -eq 3) 'Expected exactly two JPEGs and one MP4.'
        foreach($file in $expected.files){
            Assert-Distribution ((Get-DistributionHash (Join-Path $Source $file.input)) -ceq $file.input_sha256) 'Original fixture bytes changed.'
            $result=Join-Path $Output $file.output
            if($file.action -eq 'copy_bytes'){Assert-Distribution ((Get-DistributionHash $result) -ceq $file.input_sha256) 'MP4 bytes changed.'}
            else{
                $item=Get-Item -LiteralPath $result;Assert-Distribution ($item.Length -gt 0 -and $item.Length -le $expected.max_jpeg_bytes) 'Unexpected JPEG byte count.'
                $image=[Drawing.Bitmap]::new($result)
                try{Assert-Distribution ($image.RawFormat.Guid -eq [Drawing.Imaging.ImageFormat]::Jpeg.Guid -and $image.Width -eq $file.dimensions[0] -and $image.Height -eq $file.dimensions[1]) 'Independent GDI+ JPEG decode dimensions/format failed.'
                    # Force every raster pixel through the independent decoder,
                    # including the camera JPEG that has no white-corner check.
                    for($pixelY=0;$pixelY -lt $image.Height;$pixelY++){for($pixelX=0;$pixelX -lt $image.Width;$pixelX++){$null=$image.GetPixel($pixelX,$pixelY)}}
                    if($file.PSObject.Properties['white_point']){$pixel=$image.GetPixel($file.white_point[0],$file.white_point[1]);Assert-Distribution ($pixel.R -ge 245 -and $pixel.G -ge 245 -and $pixel.B -ge 245) 'Transparent corner is not white.'}
                }finally{$image.Dispose()}
            }
        }
        $csv=@(Get-ChildItem -LiteralPath $Output -Recurse -File -Filter '*.csv');Assert-Distribution ($csv.Count -eq 1) 'CSV report missing.'
        $rows=@(Import-Csv -LiteralPath $csv[0].FullName)
        Assert-Distribution ($Console -match '2 images converted\s*\|\s*1 videos copied\s*\|\s*0 errors') 'Actual console counts differ.'
        Assert-Distribution ($rows.Count -eq 3) 'CSV must contain three input records.'
        Assert-Distribution (@($rows|Where-Object Status -CEQ 'text:Converted').Count -eq 2 -and @($rows|Where-Object Status -CEQ 'text:CopiedVideo').Count -eq 1) 'CSV counts/statuses differ from the actual expected outputs.'
    }
    function Invoke-DistributionCase([string]$Layout,[string]$PathMode,[string]$Entry,[string]$Negative='') {
        $label=$Layout+'-'+$PathMode+'-'+$Entry+$(if($Negative){'-'+$Negative}else{''})
        $caseRoot=Join-Path $root $label;[IO.Directory]::CreateDirectory($caseRoot)|Out-Null
        $app=New-DistributionLayout $Layout $caseRoot
        $work=Join-Path $caseRoot 'Unrelated Start ö';[IO.Directory]::CreateDirectory($work)|Out-Null
        $caseDesktop=Join-Path $desktop $label;[IO.Directory]::CreateDirectory($caseDesktop)|Out-Null
        Copy-Item -LiteralPath $fixture -Destination $caseDesktop -Recurse;$source=Join-Path $caseDesktop 'Sample Kesä 2026'
        $cache=Join-Path $caseRoot 'Cold LocalAppData'
        if($Negative -eq 'trusted-cache'){
            $cache=Join-Path $root 'source-archive-old-PS7/Cold LocalAppData'
            Assert-Distribution ([IO.Directory]::Exists((Join-Path $cache 'WinImgNormalizer/runtimes'))) 'Prior authenticated cache is absent for explicit bare-copy reuse case.'
        }else{
            Assert-Distribution (-not [IO.Directory]::Exists($cache)) 'Per-case cache is not cold.'
            [IO.Directory]::CreateDirectory($cache)|Out-Null
        }
        $target=Join-Path $app 'WinImgNormalizer.bat';$shortcut=Join-Path $caseRoot 'WinImgNormalizer.lnk'
        $wscript=New-Object -ComObject WScript.Shell;$link=$wscript.CreateShortcut($shortcut)
        $link.TargetPath=$target;$link.Arguments=Quote-Distribution $source;$link.WorkingDirectory=$work;$link.Save()
        $resolved=$wscript.CreateShortcut($shortcut)
        Assert-Distribution ($resolved.TargetPath -ieq $target -and $resolved.Arguments -ceq (Quote-Distribution $source) -and $resolved.WorkingDirectory -ieq $work) 'Actual .lnk target/arguments/Start in differ.'
        $oldTargetHash=Get-DistributionHash $target
        if($Layout -eq 'pull-upgrade'){
            $null=Invoke-DistributionGit @('-C',$app,'pull','--ff-only','origin',$Revision)
            $afterLink=$wscript.CreateShortcut($shortcut);Assert-Distribution ($afterLink.TargetPath -ieq $target -and $afterLink.Arguments -ceq $resolved.Arguments) 'Git pull changed the preexisting shortcut target.'
            Assert-Distribution ((Invoke-DistributionGit @('-C',$app,'rev-parse','HEAD')) -ceq $Revision) 'Git pull did not reach exact candidate revision.'
        }
        if($Negative -eq 'corrupt-runtime'){[IO.File]::WriteAllBytes((Join-Path $app 'runtime/imagemagick/magick.exe'),[byte[]](0,1,2,3))}
        if($Negative -eq 'corrupt-payload'){
            $payload=Join-Path $app 'resources/private-runtime.zip';$bytes=[IO.File]::ReadAllBytes($payload);$bytes[40]=$bytes[40] -bxor 1;[IO.File]::WriteAllBytes($payload,$bytes)
        }
        $pathValue=$systemPath;if($PathMode -eq 'old'){$pathValue=(Split-Path -Parent $LegacyMagickPath)+';'+$pathValue}
        $environment=@{PATH=$pathValue;LOCALAPPDATA=$cache;WINIMG_NO_PAUSE='1';NO_COLOR='1';TEMP=$caseRoot;TMP=$caseRoot}
        $pathProbe=Invoke-DistributionProcess (Join-Path $env:SystemRoot 'System32/where.exe') @('magick.exe') $work $environment
        if($PathMode -eq 'old'){Assert-Distribution ($pathProbe.Exit -eq 0 -and $pathProbe.Out.Trim() -ieq $LegacyMagickPath) 'Disposable process PATH did not resolve the genuine old executable.'}
        else{Assert-Distribution ($pathProbe.Exit -ne 0 -and -not $pathProbe.Out.Trim()) 'No-engine PATH unexpectedly resolves an ImageMagick executable.'}
        $record=Join-Path $caseRoot 'actual-children.json';$before=@(Get-ChildItem -LiteralPath $pictures -Directory|ForEach-Object FullName)
        if($Entry -eq 'BAT-shortcut'){$tokens=@($shortcut,$work,$record)}
        elseif($Entry -eq 'BAT-exit'){$arguments='/d /v:off /s /c ""'+$target+'" "'+$source+'""';$tokens=@($cmd,$arguments,$work,$record)}
        else{$arguments=(@('-NoLogo','-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',(Join-Path $app 'WinImgNormalizer.ps1'),$source)|ForEach-Object {Quote-Distribution $_})-join ' ';$tokens=@($ps7,$arguments,$work,$record)}
        $result=Invoke-DistributionProcess $helper $tokens $work $environment
        Assert-Distribution ([IO.File]::Exists($record)) ($label+' did not record actual child instrumentation.')
        $instrumentation=[IO.File]::ReadAllText($record)|ConvertFrom-Json
        Assert-Distribution (-not $instrumentation.HarnessError -and $instrumentation.PrivateJobAssignedBeforeShell) ($label+' instrumentation failed: '+$instrumentation.HarnessError)
        foreach($child in $instrumentation.Processes){Assert-Distribution ($child.Path -iin $isolation.ProgramPaths) ('An actual application child executable lacked a pre-applied outbound deny rule: '+$child.Path)}
        Assert-Distribution ($result.Exit -eq $instrumentation.ObservedChildExitCode) 'Actual child exit and helper exit disagree.'
        $sibling=[IO.Path]::ChangeExtension($target,'.ps1')
        $hosts=@($instrumentation.Processes|Where-Object {$_.CommandLine -and $_.CommandLine.Contains($sibling) -and $_.CommandLine.Contains($source)})
        $hostDiagnostic=@($instrumentation.Processes|Where-Object {[IO.Path]::GetFileName($_.Path) -in @('powershell.exe','pwsh.exe')}|Select-Object Path,CommandLine,ExitCode)|ConvertTo-Json -Compress
        Assert-Distribution ($hosts.Count -eq 1 -and $hosts[0].CommandLine -match '(?i)(?:^|\s)"?-File"?(?=\s|$)') ($label+' actual host did not execute the BAT sibling PS1 with the quoted synthetic source; matched='+$hosts.Count+'; observed hosts='+$hostDiagnostic+'; console='+$instrumentation.Console)
        $expectedHost=if($Entry -eq 'PS7'){$ps7}else{$ps5}
        Assert-Distribution ($hosts[0].Path -ieq $expectedHost -and $hosts[0].ExitCode -eq $result.Exit) 'The real production PS5.1 or direct PS7 child/exit differs from its entry contract.'
        $after=@(Get-ChildItem -LiteralPath $pictures -Directory|Where-Object FullName -notin $before)
        $console=[string]$instrumentation.Console
        Assert-Distribution ($console.Contains($app)) 'Actual application directory was not visible in the launcher console.'
        $engineChildren=@($instrumentation.Processes|Where-Object {[IO.Path]::GetFileName($_.Path) -ieq 'magick.exe'})
        $expectedExit=if($Negative -eq 'trusted-cache'){0}elseif($Negative -or $Layout -eq 'bare'){1}else{0}
        Assert-Distribution ($result.Exit -eq $expectedExit) ($label+' actual exit '+$result.Exit+' expected '+$expectedExit+'. Console: '+$console)
        if($expectedExit -eq 0){
            Assert-Distribution ($after.Count -eq 1) ($label+' expected exactly one actual run folder.')
            Assert-DistributionOutputs $source $after[0].FullName $result.Exit $console
            Assert-Distribution ($engineChildren.Count -gt 1) 'Real private decoder children not observed.'
            foreach($child in $engineChildren){Assert-Distribution ($child.Sha256 -ceq $enginePin.sha256 -and $child.Path -ine $LegacyMagickPath -and $child.ExitCode -eq 0) 'A decoder child was not the authenticated private payload or exited unsuccessfully.'
                if($Layout -eq 'portable'){Assert-Distribution ($child.Path.StartsWith($app+'\',[StringComparison]::OrdinalIgnoreCase)) 'Full asset did not use healthy app-relative runtime.'}
                else{Assert-Distribution ($child.Path.StartsWith($cache+'\WinImgNormalizer\runtimes\',[StringComparison]::OrdinalIgnoreCase)) 'Source layout did not execute authenticated app-owned versioned cache.'}
            }
        }else{
            Assert-Distribution ($after.Count -eq 0 -and $engineChildren.Count -eq 0) 'Negative control executed a decoder or created an output run.'
            $engineIndex=$console.IndexOf('[02 / ENGINE]');$resultIndex=$console.IndexOf('[04 / RESULT]');Assert-Distribution ($engineIndex -ge 0 -and $resultIndex -gt $engineIndex) 'Engine diagnostic order is wrong.'
            Assert-Distribution ($console -notmatch 'update the dependency setup|install ImageMagick|7\.1\.2-35 or newer') 'Negative control repeated the original misleading external-upgrade advice.'
            Assert-Distribution ($console -match 'Selected engine: none') 'Rejected runtime invented a selected external engine or a version.'
            if($Layout -eq 'bare'){Assert-Distribution ($console -match 'WinImgNormalizer-1\.0\.2-portable\.zip' -and $console -match '(?i)incomplete') 'Bare-copy diagnostic is not truthful and asset-specific.'}
        }
        foreach($file in $expected.files){Assert-Distribution ((Get-DistributionHash (Join-Path $source $file.input)) -ceq $file.input_sha256) 'Input bytes changed in the actual launcher run.'}
        $summary.Cases += [ordered]@{Name=$label;Exit=$result.Exit;Images=$(if($expectedExit -eq 0){2}else{0});Videos=$(if($expectedExit -eq 0){1}else{0});InputsUnchanged=$true;GdiDecodeVerified=($expectedExit -eq 0);ColdCache=($Negative -ne 'trusted-cache');ExplicitAuthenticatedCacheReuse=($Negative -eq 'trusted-cache');NoRuntimeStaging=$true;ShortcutTarget='Application Ä 2026/WinImgNormalizer.bat';ShortcutArguments='quoted disposable Desktop/Sample Kesä 2026';UnrelatedStartIn=$true;ActualSiblingScriptSha256=(Get-DistributionHash $sibling);ActualBatSha256=(Get-DistributionHash $target);ActualHostExit=$hosts[0].ExitCode;DropEquivalent=($Entry -eq 'BAT-shortcut');BatExitPropagated=($Entry -in @('BAT-shortcut','BAT-exit'));Entry=$Entry;EngineChildren=$engineChildren.Count;EngineSha256=$(if($engineChildren.Count){$engineChildren[0].Sha256}else{$null});EngineSource=$(if($engineChildren.Count){if($Layout -eq 'portable'){'bundled'}else{'cache'}}else{'none'});AllDecoderChildrenAuthenticated=$true;OldPathEngineUnused=$true;ExistingShortcutPreserved=($Layout -eq 'pull-upgrade')}
        Write-DistributionJson (Join-Path $root 'distribution-summary.json') $summary
        Write-Host ('PASS '+$label+' exit '+$result.Exit+' authenticated decoder children '+$engineChildren.Count)
    }
    foreach($case in @(Get-DistributionCases)){Invoke-DistributionCase $case.Layout $case.PathMode $case.Entry $case.Negative}
    Assert-Distribution ((Get-DistributionHash $LegacyMagickPath) -ceq $legacyHash) 'The old engine bytes changed.'
    $summary.Passed=$true
}catch{
    $summary.Error=$_.Exception.Message.Replace($root,'<acceptance>').Replace($repository,'<checkout>').Replace($env:USERPROFILE,'<profile>')
    $summary.ErrorLine=$_.InvocationInfo.ScriptLineNumber
    Write-DistributionJson (Join-Path $root 'distribution-summary.json') $summary
    throw
}
Write-DistributionJson (Join-Path $root 'distribution-summary.json') $summary
Write-Host ('Distribution acceptance passed: '+$summary.Cases.Count+' actual entry runs; exact asset '+$summary.PortableSha256)
exit 0
