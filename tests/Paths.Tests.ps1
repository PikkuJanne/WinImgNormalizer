BeforeAll {
    $repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    $application = Join-Path $repository 'WinImgNormalizer.ps1'
    $batchApplication = Join-Path $repository 'WinImgNormalizer.bat'
    function Assert-PathsAncestors([string]$Path) {
        $directory = [IO.DirectoryInfo]::new([IO.Path]::GetFullPath($Path))
        while ($directory) {
            if ($directory.Exists -and ($directory.Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Path fixture crosses a reparse point.' }
            $directory = $directory.Parent
        }
    }
    $scratch = Join-Path $repository '.scratch'
    Assert-PathsAncestors $scratch
    foreach ($pictures in @([Environment]::GetFolderPath('MyPictures'),(Join-Path $env:USERPROFILE 'Pictures'))) {
        if (-not $pictures) { continue }
        $full = [IO.Path]::GetFullPath($pictures).TrimEnd('\','/')
        if ($scratch.Equals($full,[StringComparison]::OrdinalIgnoreCase) -or $scratch.StartsWith($full+'\',[StringComparison]::OrdinalIgnoreCase) -or $full.StartsWith($scratch+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Path fixture overlaps real Pictures.' }
    }
    & git -C $repository check-ignore --quiet --no-index -- (Join-Path $scratch 'paths-ignore-probe')
    if ($LASTEXITCODE -ne 0) { throw 'Path fixture scratch must already be ignored.' }
    $ownedRoot = Join-Path $scratch ('M2-T03-paths-' + [Guid]::NewGuid().ToString('N'))
    if ([IO.Directory]::Exists($ownedRoot)) { throw 'Path fixture ownership collision.' }
    $null = [IO.Directory]::CreateDirectory($ownedRoot)
    [IO.File]::WriteAllText((Join-Path $ownedRoot '.winimg-fixture-root'),'M2-T03 owned synthetic paths')
    function New-PathsDirectory([string]$Label) {
        $path = [IO.Path]::GetFullPath((Join-Path $ownedRoot ($Label+'-'+[Guid]::NewGuid().ToString('N').Substring(0,8))))
        if (-not $path.StartsWith($ownedRoot+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Path fixture escaped ownership.' }
        Assert-PathsAncestors $path
        $null = [IO.Directory]::CreateDirectory($path)
        return $path
    }
    $magick = [IO.Path]::GetFullPath($env:WINIMG_TEST_MAGICK)
    if (-not [IO.Path]::IsPathRooted($env:WINIMG_TEST_MAGICK) -or [IO.Path]::GetFileName($magick) -ne 'magick.exe') { throw 'Explicit verified ImageMagick required.' }
    Assert-PathsAncestors $magick
    . $application
    $realPathsSnapshot = (Get-Item Function:\New-WinImgSourceSnapshot).ScriptBlock
    function Invoke-PathsMagick([string[]]$Arguments) {
        $messages = @(& $magick @Arguments 2>&1); $code = $LASTEXITCODE
        if ($code -ne 0) { throw ('Synthetic path native probe failed: '+($messages -join "`n")) }
        return $messages -join "`n"
    }
    function Get-PathsFileState([string]$Path) {
        $native = Get-WinImgNativeOutputPath $Path
        $hash = [Security.Cryptography.SHA256]::Create()
        try { $digest = [BitConverter]::ToString($hash.ComputeHash([IO.File]::ReadAllBytes($native))).Replace('-','') } finally { $hash.Dispose() }
        return [pscustomobject]@{ Hash=$digest; Length=[IO.File]::ReadAllBytes($native).Length; Creation=[IO.File]::GetCreationTimeUtc($native).Ticks; Modified=[IO.File]::GetLastWriteTimeUtc($native).Ticks } | ConvertTo-Json -Compress
    }
    function Set-PathsTimes([string]$Path) {
        $native = Get-WinImgNativeOutputPath $Path
        $time = [DateTime]::Parse('2021-03-04T05:06:07Z').ToUniversalTime()
        [IO.File]::SetCreationTimeUtc($native,$time); [IO.File]::SetLastWriteTimeUtc($native,$time)
    }
    $templates = New-PathsDirectory 'templates'
    $red = Join-Path $templates 'red.png'; $blue = Join-Path $templates 'blue.png'; $gif = Join-Path $templates 'frames.gif'
    $null = Invoke-PathsMagick @('-size','48x32','xc:rgb(224,32,32)',('PNG:'+$red))
    $null = Invoke-PathsMagick @('-size','48x32','xc:rgb(32,32,224)',('PNG:'+$blue))
    $null = Invoke-PathsMagick @('-delay','10',$red,$blue,'-loop','0',('GIF:'+$gif))
    $profiled = Join-Path $templates 'tagged.png'
    $null = Invoke-PathsMagick @($red,'-profile',(Join-Path $PSScriptRoot 'fixtures/colour/sRGB-v4.icc'),('PNG:'+$profiled))
    $spoofed = Join-Path $templates 'tagged-property.png'
    $null = Invoke-PathsMagick @($profiled,'-set','filename:literal','false',('MIFF:'+$spoofed))
    function Copy-PathsFixture([string]$Template,[string]$Path) {
        [IO.File]::WriteAllBytes((Get-WinImgNativeOutputPath $Path),[IO.File]::ReadAllBytes($Template))
        Set-PathsTimes $Path
    }
    function Get-PathsResult([string]$Parent,[int]$Code,[string]$Text) {
        $runs = @(Get-ChildItem -LiteralPath $Parent -Force -Directory)
        $runs.Count | Should -Be 1
        $logs = @(Get-ChildItem -LiteralPath (Join-Path $runs[0].FullName '.WinImgNormalizer/reports') -File)
        $logs.Count | Should -Be 1
        return [pscustomobject]@{ Code=$Code; Run=$runs[0].FullName; Log=Get-Content -LiteralPath $logs[0].FullName -Raw; Text=$Text }
    }
    function Invoke-PathsRun([string]$Source,[string]$Parent,[scriptblock]$Runner) {
        $parameters = @{Source=$Source;OutputParent=$Parent;MagickPath=$magick}
        if ($Runner) { $parameters.ProcessRunner=$Runner }
        $observed = @(& Invoke-WinImgNormalizer @parameters 6>&1 3>&1 2>&1)
        $codes = @($observed | Where-Object { $_ -is [int] -or $_ -is [long] })
        $codes.Count | Should -Be 1
        return Get-PathsResult $Parent $codes[0] ($observed -join "`n")
    }
    function Assert-PathsJpeg([string]$Path,[object[]]$Rgb) {
        # Inspection itself must not interpret the intended percent/bracket name.
        $probe = Join-Path $templates ('inspect-'+[Guid]::NewGuid().ToString('N')+'.jpg')
        [IO.File]::WriteAllBytes($probe,[IO.File]::ReadAllBytes((Get-WinImgNativeOutputPath $Path)))
        Invoke-PathsMagick @('identify','+ping','-regard-warnings','-format','%m|%w|%h|%n',$probe) | Should -Be 'JPEG|48|32|1'
        $actual = (Invoke-PathsMagick @($probe,'-format','%[fx:round(255*p{24,16}.r)]|%[fx:round(255*p{24,16}.g)]|%[fx:round(255*p{24,16}.b)]','info:')) -split '\|'
        for ($channel=0;$channel -lt 3;$channel++) { [Math]::Abs([int]$actual[$channel]-$Rgb[$channel]) | Should -BeLessOrEqual 12 }
    }
    function Assert-PathsSuccess([object]$Result,[string[]]$Expected) {
        $Result.Code | Should -Be 0 -Because $Result.Text
        $files = @(Get-ChildItem -LiteralPath $Result.Run -Recurse -Force -File | Where-Object Extension -ne '.log' | ForEach-Object { $_.FullName.Substring($Result.Run.Length+1) } | Sort-Object)
        ($files -join '|') | Should -Be (@($Expected | Sort-Object) -join '|')
        @(Get-ChildItem -LiteralPath (Join-Path $Result.Run '.WinImgNormalizer/work') -Force).Count | Should -Be 0
    }
    function ConvertTo-PathsWindowsArgument([string]$Value) {
        $escaped = [regex]::Replace($Value,'(\\*)"','$1$1\"')
        $escaped = [regex]::Replace($escaped,'(\\+)$','$1$1')
        return '"'+$escaped+'"'
    }
    function Invoke-PathsChild([string]$Executable,[string[]]$Tokens,[string]$Working,[hashtable]$Environment,[string]$RawArguments,[int]$Timeout=45000) {
        $start = New-Object Diagnostics.ProcessStartInfo
        $start.FileName=$Executable; $start.Arguments=($Tokens | ForEach-Object { ConvertTo-PathsWindowsArgument $_ }) -join ' '
        if ($RawArguments) { $start.Arguments=$RawArguments }
        $start.WorkingDirectory=$Working; $start.UseShellExecute=$false; $start.CreateNoWindow=$true
        $start.RedirectStandardOutput=$true; $start.RedirectStandardError=$true; $start.RedirectStandardInput=$true
        foreach ($key in @($start.EnvironmentVariables.Keys)) { if ([string]$key -ieq 'PSModulePath') { $start.EnvironmentVariables.Remove([string]$key) } }
        foreach ($key in $Environment.Keys) { $start.EnvironmentVariables[$key]=$Environment[$key] }
        $process = New-Object Diagnostics.Process; $process.StartInfo=$start
        try {
            if (-not $process.Start()) { throw 'Owned path child failed to start.' }
            $stdout=$process.StandardOutput.ReadToEndAsync(); $stderr=$process.StandardError.ReadToEndAsync()
            $process.StandardInput.WriteLine(''); $process.StandardInput.Close()
            if (-not $process.WaitForExit($Timeout)) {
                & (Join-Path $env:SystemRoot 'System32/taskkill.exe') /PID $process.Id /T /F 2>&1 | Out-Null
                throw 'Owned path child timed out.'
            }
            if (-not $stdout.Wait(5000) -or -not $stderr.Wait(5000)) { throw 'Owned path child output incomplete.' }
            return [pscustomobject]@{ Exit=$process.ExitCode; Out=$stdout.GetAwaiter().GetResult(); Error=$stderr.GetAwaiter().GetResult() }
        } finally { $process.Dispose() }
    }
    function New-PathsContainedScript([string]$Path,[string]$Parent,[string]$ArgvRecord) {
        $text=[IO.File]::ReadAllText($application); $entry='exit (Invoke-WinImgNormalizerCommand -Arguments $args)'
        $index=$text.IndexOf($entry,[StringComparison]::Ordinal)
        $index | Should -BeGreaterOrEqual 0
        $text.IndexOf($entry,$index+$entry.Length,[StringComparison]::Ordinal) | Should -Be -1
        $encodedParent=[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Parent))
        $encodedRecord=[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($ArgvRecord))
        $encodedMagick=[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($magick))
        $injection=@'
$pathsReceivedArguments=@($args)
$pathsApplicationExit=Invoke-WinImgNormalizerCommand -Arguments $args -OutputParent ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('__PARENT__'))) -MagickPath ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('__MAGICK__')))
[IO.File]::WriteAllText([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('__RECORD__')), (ConvertTo-Json -InputObject ([pscustomobject]@{ Arguments=$pathsReceivedArguments; Version=$PSVersionTable.PSVersion.ToString(); Edition=$PSVersionTable.PSEdition; ApplicationExit=$pathsApplicationExit })))
exit $pathsApplicationExit
'@
        $injection=$injection.Replace('__RECORD__',$encodedRecord).Replace('__PARENT__',$encodedParent).Replace('__MAGICK__',$encodedMagick)
        [IO.File]::WriteAllText($Path,$text.Substring(0,$index)+$injection+$text.Substring($index+$entry.Length),(New-Object Text.UTF8Encoding $true))
    }
    $unicode = [string][char]0x00E4+[char]0x00F6+' '+[char]0x00FC+[char]0x00DF+' '+[char]::ConvertFromUtf32(0x1F642)
}

Describe 'M2-T03 literal native names and exact final identity (T037)' {
    It 'T037 retains literal <Name> and selects its red source rather than a deceptive blue neighbor' -ForEach @(
        @{Name='percent%d.png'},@{Name='percent%03d.png'},@{Name='selector[0].png'},
        @{Name='hash#name.png'},@{Name='@name.png'},@{Name='-name.png'}
    ) {
        $source=New-PathsDirectory 'literal-source'; $parent=New-PathsDirectory 'literal-output'
        $path=Join-Path $source $Name; $neighbor=Join-Path $source 'deceptive.png'
        Copy-PathsFixture $red $path; Copy-PathsFixture $blue $neighbor
        $before=Get-PathsFileState $path; $neighborBefore=Get-PathsFileState $neighbor
        $result=Invoke-PathsRun $source $parent
        $intended=[IO.Path]::GetFileNameWithoutExtension($Name)+'.jpeg'
        Assert-PathsSuccess $result @($intended,'deceptive.jpeg')
        Assert-PathsJpeg (Join-Path $result.Run $intended) @(224,32,32)
        Assert-PathsJpeg (Join-Path $result.Run 'deceptive.jpeg') @(32,32,224)
        Get-PathsFileState $path | Should -Be $before
        Get-PathsFileState $neighbor | Should -Be $neighborBefore
    }
    It 'T037 selects the first actual GIF frame at bracket/percent names and preserves nested literal paths' {
        $source=New-PathsDirectory 'frames-source'; $parent=New-PathsDirectory 'frames-output'
        $nested=Join-Path $source 'nested %d [1]'; $null=[IO.Directory]::CreateDirectory($nested)
        $path=Join-Path $nested 'sequence[1]%03d.gif'; Copy-PathsFixture $gif $path
        $video=Join-Path $nested 'opaque%d [0].mp4'
        [IO.File]::WriteAllBytes($video,[Text.Encoding]::UTF8.GetBytes('Owned opaque video bytes; no codec claim.'))
        Set-PathsTimes $video
        $before=Get-PathsFileState $path
        $videoBefore=Get-PathsFileState $video
        $result=Invoke-PathsRun $source $parent
        $relative='nested %d [1]\sequence[1]%03d.jpeg'
        Assert-PathsSuccess $result @($relative,'nested %d [1]\opaque%d [0].mp4')
        Assert-PathsJpeg (Join-Path $result.Run $relative) @(224,32,32)
        $result.Log | Should -Match 'SourceCount=2; Selected=1; Omitted=1; Unit=Frames; Policy=FirstDisplayedFrame; Decoder=GIF'
        Get-PathsFileState $path | Should -Be $before
        Get-PathsFileState $video | Should -Be $videoBefore
        Get-PathsFileState (Join-Path $result.Run 'nested %d [1]\opaque%d [0].mp4') | Should -Be $videoBefore
    }
    It 'T037 keeps tagged ICC extraction and JPEG writing literal under percent source/run/output ancestors' {
        $source=New-PathsDirectory 'source%d [0]'; $parent=New-PathsDirectory 'parent%03d [1]'
        $path=Join-Path $source 'tagged%.png'; Copy-PathsFixture $spoofed $path
        $before=Get-PathsFileState $path
        $result=Invoke-PathsRun $source $parent
        Assert-PathsSuccess $result @('tagged%.jpeg')
        Assert-PathsJpeg (Join-Path $result.Run 'tagged%.jpeg') @(224,32,32)
        $result.Log | Should -Match 'SourceICC=True; Policy=ProfileToSrgb;'
        Get-PathsFileState $path | Should -Be $before
        @(Get-ChildItem -LiteralPath $ownedRoot -Directory | Where-Object Name -like 'parent000*').Count | Should -Be 0
    }
    It 'T037 reads the exact percent-parent snapshot despite an actual differently colored numeric-template counterpart' {
        $source=New-PathsDirectory 'selected%d'; $parent=New-PathsDirectory 'selected-output%d'
        $path=Join-Path $source 'sequence[0].gif'; Copy-PathsFixture $gif $path
        $before=Get-PathsFileState $path
        $observer=[pscustomobject]@{ Neighbors=New-Object 'Collections.Generic.List[string]' }
        Mock New-WinImgSourceSnapshot {
            param($SourcePath,$WorkRoot,$ExpectedLength,$ExpectedModified,$OwnedCandidates)
            $snapshot=& $realPathsSnapshot @PSBoundParameters
            $counterpart=$snapshot.Replace('%d','0')
            if ($counterpart -eq $snapshot -or -not $counterpart.StartsWith($ownedRoot+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Numeric counterpart must remain owned and distinct.' }
            $null=[IO.Directory]::CreateDirectory((Get-WinImgNativeOutputPath ([IO.Path]::GetDirectoryName($counterpart))))
            [IO.File]::WriteAllBytes((Get-WinImgNativeOutputPath $counterpart),[IO.File]::ReadAllBytes($blue))
            $observer.Neighbors.Add($counterpart)
            return $snapshot
        }
        $result=Invoke-PathsRun $source $parent
        Assert-PathsSuccess $result @('sequence[0].jpeg')
        Assert-PathsJpeg (Join-Path $result.Run 'sequence[0].jpeg') @(224,32,32)
        $result.Log | Should -Match 'SourceCount=2; Selected=1; Omitted=1; Unit=Frames; Policy=FirstDisplayedFrame; Decoder=GIF'
        $observer.Neighbors.Count | Should -Be 1
        [IO.File]::Exists((Get-WinImgNativeOutputPath $observer.Neighbors[0])) | Should -BeTrue
        Get-PathsFileState $path | Should -Be $before
    }
    It 'T037 preserves an external arrival at the intended literal final name without overwriting or leaking candidates' {
        $source=New-PathsDirectory 'arrival-source'; $parent=New-PathsDirectory 'arrival-output'
        $path=Join-Path $source 'arrival%d [0].png'; Copy-PathsFixture $red $path
        $before=Get-PathsFileState $path
        $trace=[pscustomobject]@{ Arrival=$null; State=$null; Bytes=[Text.Encoding]::UTF8.GetBytes('owned external literal arrival') }
        $runner={
            param($Executable,$Arguments)
            if (-not $trace.Arrival) {
                $run=@(Get-ChildItem -LiteralPath $parent -Directory)[0].FullName
                $trace.Arrival=Join-Path $run 'arrival%d [0].jpeg'
                [IO.File]::WriteAllBytes($trace.Arrival,$trace.Bytes); Set-PathsTimes $trace.Arrival
                $trace.State=Get-PathsFileState $trace.Arrival
            }
            $diagnostics=@(& $Executable @Arguments 2>&1)
            return [pscustomobject]@{ExitCode=$LASTEXITCODE;DiagnosticOutput=$diagnostics -join "`n"}
        }
        $result=Invoke-PathsRun $source $parent $runner
        $result.Code | Should -Be 2
        Get-PathsFileState $trace.Arrival | Should -Be $trace.State
        @(Get-ChildItem -LiteralPath (Join-Path $result.Run '.WinImgNormalizer/work') -Force).Count | Should -Be 0
        $result.Log | Should -Not -Match 'OK IMG:'
        Get-PathsFileState $path | Should -Be $before
    }
}

AfterAll {
    # Live UNC is capability evidence outside the mandatory Pester case count.
    # Probe only the localhost administrative alias of this newly owned fixture.
    $uncSource=New-PathsDirectory 'unc-source%d'
    $uncFile=Join-Path $uncSource 'unc[0]%.png'; Copy-PathsFixture $profiled $uncFile
    $sourceState=Get-PathsFileState $uncFile
    $driveRoot=[IO.Path]::GetPathRoot($uncSource)
    $observations=New-Object 'Collections.Generic.List[object]'
    if ($driveRoot -match '^[A-Za-z]:\\$') {
        $alias='\\localhost\'+$driveRoot.Substring(0,1)+'$'+$uncSource.Substring(2)
        $probeWork=New-PathsDirectory 'unc-capability'
        $probeScript=Join-Path $probeWork 'probe.ps1'
        $encoded=[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes((Join-Path $alias 'unc[0]%.png')))
        [IO.File]::WriteAllText($probeScript,"if ([IO.File]::Exists([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('$encoded')))) { exit 0 }; exit 3",[Text.Encoding]::ASCII)
        try {
            $capability=Invoke-PathsChild (Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe') @('-NoLogo','-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',$probeScript) $probeWork @{} '' 5000
            $available=$capability.Exit -eq 0
            $reason='The existing localhost administrative alias is unavailable to this host/process.'
        } catch { $available=$false; $reason='The owned UNC capability probe exceeded its bound or could not start.' }
        if ($available) {
            $inside=Join-Path $uncSource 'existing-output'; $null=[IO.Directory]::CreateDirectory($inside)
            foreach ($direction in @('local source / UNC destination','UNC source / local destination')) {
                $boundarySource=if ($direction -like 'local*') {$uncSource} else {$alias}
                $boundaryParent=if ($direction -like 'local*') {Join-Path $alias 'existing-output'} else {$inside}
                $trace=[pscustomobject]@{Calls=0}
                $forbidden={param($Executable,$Arguments);$trace.Calls++;throw 'Rejected alias must not invoke a native probe/conversion.'}
                $rejected=@(& Invoke-WinImgNormalizerCommand -Arguments @($boundarySource) -OutputParent $boundaryParent -MagickPath $magick -PreflightRunner $forbidden -ProcessRunner $forbidden 6>&1 3>&1 2>&1)
                $codes=@($rejected | Where-Object {$_ -is [int] -or $_ -is [long]})
                $codes.Count | Should -Be 1
                $codes[0] | Should -Be 1
                $trace.Calls | Should -Be 0
                ($rejected -join "`n") | Should -Match 'destination is inside or equal to the source, including a directory alias'
                @(Get-ChildItem -LiteralPath $inside -Force).Count | Should -Be 0
                @(Get-ChildItem -LiteralPath $uncSource -Recurse -Force -File).Count | Should -Be 1
                Get-PathsFileState $uncFile | Should -Be $sourceState
                $observations.Add([pscustomobject]@{Form=$direction;State='pass';Code=1;NativeCalls=0;SourceState=$sourceState;DestinationUnchanged=$true})
            }
            $parent=New-PathsDirectory 'unc-direct-output'
            $parentAlias='\\localhost\'+$driveRoot.Substring(0,1)+'$'+$parent.Substring(2)
            $result=Invoke-PathsRun $alias $parentAlias
            Assert-PathsSuccess $result @('unc[0]%.jpeg')
            Assert-PathsJpeg (Join-Path $result.Run 'unc[0]%.jpeg') @(224,32,32)
            Get-PathsFileState $uncFile | Should -Be $sourceState
            $observations.Add([pscustomobject]@{Form='direct UNC input and output';State='pass';Source=$alias;Output=$result.Run;Code=$result.Code;SourceState=$sourceState;Pixels='red224,32,32';JpegFullyDecoded=$true})
            $work=New-PathsDirectory 'unc-batch'; $parent=Join-Path $work 'output'; $record=Join-Path $work 'actual-argv.json'
            New-PathsContainedScript (Join-Path $work 'WinImgNormalizer.ps1') $parent $record
            $copiedBatch=Join-Path $work 'WinImgNormalizer.bat'; [IO.File]::WriteAllBytes($copiedBatch,[IO.File]::ReadAllBytes($batchApplication))
            $child=Invoke-PathsChild (Join-Path $env:SystemRoot 'System32/cmd.exe') @() $work @{} ('/d /v:on /s /c ""'+$copiedBatch+'" "'+$alias+'""')
            $receiver=[IO.File]::ReadAllText($record) | ConvertFrom-Json
            $receiver.Arguments[0] | Should -BeExactly $alias
            $receiver.Edition | Should -Be 'Desktop'
            $receiver.Version | Should -Match '^5\.1\.'
            $receiver.ApplicationExit | Should -Be 0
            $child.Error | Should -BeNullOrEmpty
            $result=Get-PathsResult $parent $receiver.ApplicationExit $child.Out
            $result.Log | Should -Match 'SUMMARY ConvertedImages=1 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=0'
            Assert-PathsSuccess $result @('unc[0]%.jpeg')
            Assert-PathsJpeg (Join-Path $result.Run 'unc[0]%.jpeg') @(224,32,32)
            Get-PathsFileState $uncFile | Should -Be $sourceState
            $observations.Add([pscustomobject]@{Form='BAT';State='pass';Source=$alias;Output=$result.Run;Receiver=$receiver;SourceState=$sourceState;Pixels='red224,32,32';JpegFullyDecoded=$true})
        } else {
            $observations.Add([pscustomobject]@{Form='direct and BAT';State='not_run';Reason=$reason;Source=$alias})
            Write-Host ('T039 live UNC not_run: '+$reason)
        }
    } else { $observations.Add([pscustomobject]@{Form='direct and BAT';State='not_run';Reason='Workspace has no local drive administrative alias; no alternate share is created or searched.'}) }
    [IO.File]::WriteAllText((Join-Path $templates 'unc-capability-observations.json'),(ConvertTo-Json -InputObject $observations.ToArray() -Depth 5))
}

Describe 'M2-T03 actual Windows PowerShell5.1 argv and unchanged BAT (T038)' {
    It 'T038 preserves <Form> source arguments with spaces apostrophes ampersands parentheses bangs percent brackets and Unicode' -ForEach @(
        @{Form='direct source-only';Batch=$false;Cap=$false;Trailing=$false},
        @{Form='direct explicit-cap trailing-separator';Batch=$false;Cap=$true;Trailing=$true},
        @{Form='BAT inherited delayed-expansion';Batch=$true;Cap=$false;Trailing=$false},
        @{Form='BAT trailing-separator';Batch=$true;Cap=$false;Trailing=$true}
    ) {
        $work=New-PathsDirectory 'launcher'
        $literalPercent=if ($Batch) {'%d'} else {'%WINIMG_LITERAL%'}
        $source=Join-Path $work ("source space & O'Brien (x) ! "+$literalPercent+' [0] '+$unicode)
        $launch=Join-Path $work ("launcher space & O'Brien (x) ! "+$unicode)
        $null=[IO.Directory]::CreateDirectory($source); $null=[IO.Directory]::CreateDirectory($launch)
        $path=Join-Path $source 'red.png'; Copy-PathsFixture $red $path
        $before=Get-PathsFileState $path
        $parent=Join-Path $work 'contained-output'; $record=Join-Path $work 'actual-argv.json'
        $contained=Join-Path $launch 'WinImgNormalizer.ps1'
        New-PathsContainedScript $contained $parent $record
        $sourceArgument=$source; if ($Trailing) { $sourceArgument+='\' }
        $environment=@{WINIMG_LITERAL='caller-expansion-must-not-select-this'}
        if ($Batch) {
            $copiedBatch=Join-Path $launch 'WinImgNormalizer.bat'
            [IO.File]::WriteAllBytes($copiedBatch,[IO.File]::ReadAllBytes($batchApplication))
            Get-PathsFileState $copiedBatch | Should -Match ([regex]::Escape((Get-FileHash -LiteralPath $batchApplication -Algorithm SHA256).Hash))
            # /s removes the outer pair; the inner quotes preserve native CMD
            # tokens. No CALL, helper BAT or UTF-8 batch decoding is involved.
            $rawArguments='/d /v:on /s /c ""'+$copiedBatch+'" "'+$sourceArgument+'""'
            $child=Invoke-PathsChild (Join-Path $env:SystemRoot 'System32/cmd.exe') @() $work $environment $rawArguments
        } else {
            $tokens=@('-NoLogo','-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',$contained,$sourceArgument)
            if ($Cap) { $tokens+='1048576' }
            $child=Invoke-PathsChild (Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe') $tokens $work $environment
            $child.Exit | Should -Be 0 -Because $child.Out
        }
        $child.Error | Should -BeNullOrEmpty
        $actualReceiver=[IO.File]::ReadAllText($record) | ConvertFrom-Json
        $actualReceiver.Edition | Should -Be 'Desktop'
        $actualReceiver.Version | Should -Match '^5\.1\.'
        $actualReceiver.ApplicationExit | Should -Be 0
        $received=@($actualReceiver.Arguments)
        if ($Batch -and $Trailing) {
            [IO.Path]::GetFullPath($received[0]).TrimEnd('\','/') | Should -BeExactly $source
        } else { $received[0] | Should -BeExactly $sourceArgument }
        $received.Count | Should -Be $(if ($Cap) {2} else {1})
        if ($Cap) { $received[1] | Should -BeExactly '1048576' }
        $result=Get-PathsResult $parent $actualReceiver.ApplicationExit $child.Out
        $result.Log | Should -Match 'SUMMARY ConvertedImages=1 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=0'
        Assert-PathsSuccess $result @('red.jpeg')
        Assert-PathsJpeg (Join-Path $result.Run 'red.jpeg') @(224,32,32)
        Get-PathsFileState $path | Should -Be $before
    }
}

Describe 'M2-T03 local long paths and lexical root boundaries (T039)' {
    It 'T039 processes a real owned source file beyond MAX_PATH or reports explicit preserved failure without truncation' {
        $source=New-PathsDirectory 'long-source'; $parent=New-PathsDirectory 'long-output'
        $relative=(('segment-'+('x'*42)+'\')*4)+'image%d [0].png'
        $path=Join-Path $source $relative
        $null=[IO.Directory]::CreateDirectory((Get-WinImgNativeOutputPath ([IO.Path]::GetDirectoryName($path))))
        Copy-PathsFixture $red $path
        $path.Length | Should -BeGreaterThan 260
        $before=Get-PathsFileState $path
        $observed=@(& Invoke-WinImgNormalizer -Source $source -OutputParent $parent -MagickPath $magick 6>&1 3>&1 2>&1)
        $codes=@($observed | Where-Object {$_ -is [int] -or $_ -is [long]}); $codes.Count | Should -Be 1
        $runs=@(Get-ChildItem -LiteralPath $parent -Directory)
        if ($codes[0] -eq 0) {
            $result=Get-PathsResult $parent 0 ($observed -join "`n")
            $output=Join-Path $result.Run ([IO.Path]::ChangeExtension($relative,'.jpeg'))
            Assert-PathsSuccess $result @([IO.Path]::ChangeExtension($relative,'.jpeg'))
            Assert-PathsJpeg $output @(224,32,32)
            $result.Log | Should -Match 'SUMMARY ConvertedImages=1 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=0'
        } else {
            $codes[0] | Should -BeIn @(1,2)
            ($observed -join "`n") | Should -Match '(?i)(PathTooLong|path.*too long|long.*not supported|cannot.*(path|directory)|failed.*(enumerat|travers|scan)|could not.*(enumerat|travers|scan))'
            foreach ($run in $runs) { @(Get-ChildItem -LiteralPath $run.FullName -Recurse -File | Where-Object Extension -eq '.jpeg').Count | Should -Be 0 }
        }
        Get-PathsFileState $path | Should -Be $before
        [IO.File]::WriteAllText((Join-Path $templates 'long-path-observation.json'),([pscustomobject]@{PathLength=$path.Length;Code=$codes[0];Supported=($codes[0] -eq 0);Text=$observed -join "`n"}|ConvertTo-Json))
    }
    It 'T039 preserves drive and UNC trailing root separators in lexical containment without enumerating an unapproved share' {
        Normalize-WinImgRootPath 'C:\' | Should -Be 'C:\'
        Normalize-WinImgRootPath '\\server\share\' | Should -Be '\\server\share\'
        Test-WinImgPathContained 'C:\' 'C:\owned' | Should -BeTrue
        Test-WinImgPathContained '\\server\share\' '\\server\share\owned' | Should -BeTrue
        Test-WinImgPathContained '\\server\share\' '\\server\share-neighbor\owned' | Should -BeFalse
    }
    It 'T039 identifies the same owned physical directory and rejects missing identity while a sibling differs' {
        $source=New-PathsDirectory 'identity-source'; $sibling=New-PathsDirectory 'identity-sibling'
        $path=Join-Path $source 'red.png'; Copy-PathsFixture $red $path
        $before=Get-PathsFileState $path
        Initialize-WinImgDirectoryApi
        $identity=[WinImgNormalizer.NativeDirectory]::Identity($source)
        $identity | Should -Match '^[0-9A-F]{16}:[0-9A-F]{32}$'
        [WinImgNormalizer.NativeDirectory]::Identity($source.ToUpperInvariant()) | Should -BeExactly $identity
        [WinImgNormalizer.NativeDirectory]::Identity($sibling) | Should -Not -Be $identity
        { [WinImgNormalizer.NativeDirectory]::Identity((Join-Path $source 'not-created')) } | Should -Throw
        Get-PathsFileState $path | Should -Be $before
    }
}
