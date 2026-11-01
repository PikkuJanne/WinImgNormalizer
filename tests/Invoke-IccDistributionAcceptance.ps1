# Replays exact portable/source archives in fresh paths; never rebuilds assets.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$ArchivePath,
    [Parameter(Mandatory=$true)][string]$SourceArchivePath,
    [Parameter(Mandatory=$true)][string]$LegacyMagickPath,
    [Parameter(Mandatory=$true)][string]$ResultDirectory
)
$ErrorActionPreference='Stop'
$ResultDirectory=[IO.Path]::GetFullPath($ResultDirectory)
if(Test-Path -LiteralPath $ResultDirectory){throw 'ICC distribution acceptance requires a fresh owned directory.'}
$null=[IO.Directory]::CreateDirectory($ResultDirectory)
$summary=[ordered]@{Passed=$false;PortableSha256=(Get-FileHash -LiteralPath $ArchivePath -Algorithm SHA256).Hash.ToLowerInvariant();SourceArchiveSha256=(Get-FileHash -LiteralPath $SourceArchivePath -Algorithm SHA256).Hash.ToLowerInvariant();OldPathVersion='';Cases=@()}
$cases=New-Object 'Collections.Generic.List[object]'
try {
    $portable=Join-Path $ResultDirectory 'portable';Expand-Archive -LiteralPath $ArchivePath -DestinationPath $portable
    $source=Join-Path $ResultDirectory 'source';Expand-Archive -LiteralPath $SourceArchivePath -DestinationPath $source
    $sourceRoots=@(@(Get-Item -LiteralPath $source)+@(Get-ChildItem -LiteralPath $source -Directory)|Where-Object {Test-Path -LiteralPath (Join-Path $_.FullName 'WinImgNormalizer.ps1') -PathType Leaf})
    if($sourceRoots.Count -ne 1){throw 'Source ZIP has no unique application root.'}
    $verifier=Join-Path $portable 'runtime/imagemagick/magick.exe'
    foreach($layout in @('portable','source')){
        $app=if($layout -eq 'portable'){$portable}else{$sourceRoots[0].FullName}
        $result=Join-Path $ResultDirectory ($layout+'-replay')
        & (Join-Path $PSScriptRoot 'Invoke-IccAcceptance.ps1') -ApplicationDirectory $app -ResultDirectory $result -VerifierMagickPath $verifier -LegacyMagickDirectory ([IO.Path]::GetDirectoryName($LegacyMagickPath))
        $observed=Get-Content -LiteralPath (Join-Path $result 'icc-acceptance-summary.json') -Raw|ConvertFrom-Json
        if(-not $observed.Passed -or $observed.PSObject.Properties['Error'] -or $observed.Observations.Count -ne 9){throw 'Exact ZIP ICC replay was incomplete.'}
        $summary.OldPathVersion=$observed.OldPathVersion
        foreach($entry in $observed.Observations){
            $cases.Add([pscustomobject]@{Layout=$layout;Route=$entry.Route;Group=$entry.Group;InputCount=$entry.Result.InputCount;OutputCount=$entry.Result.OutputCount;ColourManaged=$entry.Result.ColourManaged;ColourAssumed=$entry.Result.ColourAssumed;GenuineFailures=$entry.Result.GenuineFailures;ExitCode=$entry.Result.ExitCode;IndependentReferencesPassed=$entry.Result.IndependentReferencesPassed;SourceUnchanged=$entry.SourceUnchanged})
        }
    }
    $summary.Passed=$true
} catch {$summary.Error=$_.Exception.Message;throw}
finally {
    $summary.Cases=$cases.ToArray()
    [IO.File]::WriteAllText((Join-Path $ResultDirectory 'icc-distribution-summary.json'),($summary|ConvertTo-Json -Depth 10),[Text.UTF8Encoding]::new($false))
}
