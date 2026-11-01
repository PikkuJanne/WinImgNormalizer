# Actual launcher/public-asset replay. Uses fresh synthetic inputs and normal CLI.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$ApplicationDirectory,
    [Parameter(Mandatory=$true)][string]$ResultDirectory,
    [Parameter(Mandatory=$true)][string]$VerifierMagickPath,
    [ValidateSet('BAT','PS51','PS7')][string[]]$Routes=@('BAT','PS51','PS7'),
    [ValidateSet('views','negative','mixed')][string[]]$Groups=@('views','negative','mixed'),
    [string]$LegacyMagickDirectory,
    [string]$PowerShell7Path
)
$ErrorActionPreference='Stop'
if($env:OS -ne 'Windows_NT'){throw 'Actual Windows is required.'}
. (Join-Path $PSScriptRoot 'IccAcceptanceHelpers.ps1')
$assets=Join-Path $PSScriptRoot 'fixtures/icc-repair'
$ApplicationDirectory=[IO.Path]::GetFullPath($ApplicationDirectory)
$ResultDirectory=[IO.Path]::GetFullPath($ResultDirectory)
if(Test-Path -LiteralPath $ResultDirectory){throw 'Choose a fresh owned result directory.'}
$null=[IO.Directory]::CreateDirectory($ResultDirectory)
[IO.File]::WriteAllText((Join-Path $ResultDirectory '.winimg-fixture-root'),'Owned synthetic ICC launcher replay; no private media.')
$pictures=[Environment]::GetFolderPath('MyPictures')
if(-not $pictures){$pictures=Join-Path $env:USERPROFILE 'Pictures'}
$pictures=[IO.Path]::GetFullPath($pictures)
if(-not $PowerShell7Path -and 'PS7' -in $Routes){$PowerShell7Path=(Get-Command pwsh.exe -CommandType Application -ErrorAction Stop|Select-Object -First 1).Source}
$observations=New-Object 'Collections.Generic.List[object]'
$summary=[ordered]@{Passed=$false;Windows=[Environment]::OSVersion.VersionString;PowerShell=$PSVersionTable.PSVersion.ToString();Routes=$Routes;Groups=$Groups;OldPathVersion='';Observations=@()}
try {
    if($LegacyMagickDirectory){
        $old=@(& (Join-Path $LegacyMagickDirectory 'magick.exe') -version 2>&1)
        if($LASTEXITCODE -ne 0){throw 'Old PATH control engine did not execute.'}
        $summary.OldPathVersion=[regex]::Match(($old -join "`n"),'(?m)^Version: ImageMagick (7\.\d+\.\d+-\d+)').Groups[1].Value
        if(-not $summary.OldPathVersion){throw 'Old PATH control engine version is unrecognized.'}
    }
    foreach($route in $Routes){foreach($group in $Groups){
        $label='ICC-'+$group+'-'+$route+'-Kes'+[char]0x00e4+'-'+[Guid]::NewGuid().ToString('N')
        $source=Join-Path $ResultDirectory $label
        $null=Copy-IccAcceptanceInputs -Assets $assets -Group $group -Source $source
        $before=Get-IccAcceptanceSourceState $source
        $prefix=$label+'_WinImgNormalized_'
        if(@(Get-ChildItem -LiteralPath $pictures -Directory -ErrorAction SilentlyContinue|Where-Object {$_.Name.StartsWith($prefix)}).Count){throw 'Owned launcher output prefix already exists.'}
        $start=New-Object Diagnostics.ProcessStartInfo
        $start.UseShellExecute=$false;$start.CreateNoWindow=$true;$start.RedirectStandardOutput=$true;$start.RedirectStandardError=$true
        $start.WorkingDirectory=$ApplicationDirectory
        foreach($value in @($ApplicationDirectory,$source,$VerifierMagickPath)){if($value -match '["\x00-\x1f]'){throw 'Unsupported control character in replay path.'}}
        if($route -eq 'BAT'){
            $start.FileName=Join-Path $env:SystemRoot 'System32/cmd.exe'
            $start.Arguments='/d /s /c ""'+(Join-Path $ApplicationDirectory 'WinImgNormalizer.bat')+'" "'+$source+'""'
        } else {
            $start.FileName=if($route -eq 'PS51'){Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe'}else{$PowerShell7Path}
            $start.Arguments='-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "'+(Join-Path $ApplicationDirectory 'WinImgNormalizer.ps1')+'" "'+$source+'"'
        }
        foreach($key in @($start.EnvironmentVariables.Keys)){if([string]::Equals([string]$key,'PSModulePath',[StringComparison]::OrdinalIgnoreCase)){$start.EnvironmentVariables.Remove([string]$key)}}
        $start.EnvironmentVariables['WINIMG_NO_PAUSE']='1'
        if($LegacyMagickDirectory){$start.EnvironmentVariables['PATH']=$LegacyMagickDirectory+';'+$env:PATH}
        $process=New-Object Diagnostics.Process;$process.StartInfo=$start;$started=$false
        try {
            if(-not $process.Start()){throw 'Actual launcher replay did not start.'}
            $started=$true
            $stdout=$process.StandardOutput.ReadToEndAsync();$stderr=$process.StandardError.ReadToEndAsync()
            if(-not $process.WaitForExit(180000)){throw 'Actual launcher replay exceeded three minutes.'}
            if(-not $stdout.Wait(5000) -or -not $stderr.Wait(5000)){throw 'Actual launcher replay streams did not finish.'}
            $code=$process.ExitCode;$text=$stdout.GetAwaiter().GetResult()+$stderr.GetAwaiter().GetResult()
        } finally {if($started -and -not $process.HasExited){$process.Kill();$null=$process.WaitForExit(5000)};$process.Dispose()}
        [IO.File]::WriteAllText((Join-Path $ResultDirectory ($route+'-'+$group+'.log')),$text,[Text.UTF8Encoding]::new($false))
        $outputs=@(Get-ChildItem -LiteralPath $pictures -Directory|Where-Object {$_.Name.StartsWith($prefix)})
        if($outputs.Count -ne 1){throw 'Actual launcher did not create exactly one uniquely owned output.'}
        if((Get-IccAcceptanceSourceState $source) -ne $before){throw 'The actual processed input tree changed.'}
        if($text -notmatch '7\.1\.2-35'){throw 'The actual launcher did not report the pinned private engine.'}
        $result=Assert-IccAcceptanceOutput -Assets $assets -Group $group -Output $outputs[0].FullName -MagickPath $VerifierMagickPath -ExitCode $code -ConsoleText $text
        $observations.Add([pscustomobject]@{Route=$route;Group=$group;Result=$result;SourceUnchanged=$true;Output=$outputs[0].FullName})
        Write-Host ('PASS '+$route+' '+$group+': '+$result.InputCount+' inputs, '+$result.OutputCount+' JPEGs, '+$result.ColourManaged+' managed, '+$result.ColourAssumed+' assumed, '+$result.GenuineFailures+' genuine failures.')
    }}
    $summary.Passed=$true
} catch {$summary.Error=$_.Exception.Message;throw}
finally {
    $summary.Observations=$observations.ToArray()
    [IO.File]::WriteAllText((Join-Path $ResultDirectory 'icc-acceptance-summary.json'),($summary|ConvertTo-Json -Depth 12),[Text.UTF8Encoding]::new($false))
}
