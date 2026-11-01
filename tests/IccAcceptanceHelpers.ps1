# Development-only checks against independent, precomputed Pillow/LittleCMS PNGs.
function Get-IccAcceptanceSourceState {
    param([string]$Source)
    return (@(Get-ChildItem -LiteralPath $Source -Recurse -Force -File | Sort-Object FullName | ForEach-Object {
        [pscustomobject]@{ Path=$_.FullName.Substring($Source.Length+1); Hash=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash; Bytes=$_.Length; Modified=$_.LastWriteTimeUtc.Ticks }
    }) | ConvertTo-Json -Compress)
}

function Copy-IccAcceptanceInputs {
    param([string]$Assets, [ValidateSet('views','negative','mixed')][string]$Group, [string]$Source)
    if (Test-Path -LiteralPath $Source) { throw 'ICC acceptance source must be a fresh directory.' }
    $null=[IO.Directory]::CreateDirectory($Source)
    $manifest=Get-Content -LiteralPath (Join-Path $Assets 'cases.json') -Raw | ConvertFrom-Json
    $cases=@($manifest.files | Where-Object { $Group -eq 'mixed' -or $_.file.StartsWith($Group+'/') })
    foreach ($case in $cases) {
        $original=Join-Path (Join-Path $Assets 'fixtures') $case.file
        if ((Get-FileHash -LiteralPath $original -Algorithm SHA256).Hash.ToLowerInvariant() -ne $case.sha256) { throw 'Synthetic ICC fixture digest mismatch.' }
        $relative=if($Group -eq 'mixed'){$case.file}else{[IO.Path]::GetFileName($case.file)}
        $destination=Join-Path $Source $relative
        $null=[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($destination))
        [IO.File]::Copy($original,$destination,$false)
    }
    return $cases
}

function Assert-IccAcceptanceJpeg {
    param([string]$Path, [string]$Reference, [string]$MagickPath)
    if (-not [IO.File]::Exists($Path)) { throw ('Missing actual JPEG: '+[IO.Path]::GetFileName($Path)) }
    $info=Get-Item -LiteralPath $Path
    if ($info.Length -le 0 -or $info.Length -gt 1048576) { throw 'Actual JPEG violates the one MiB fixture cap.' }
    # Full decoder diagnostics must be clean, independently of application exit.
    $messages=@(& $MagickPath identify +ping -regard-warnings -format '%m|%[colorspace]|%w|%h|%n' $Path 2>&1)
    if ($LASTEXITCODE -ne 0 -or ($messages -join "`n") -ne 'JPEG|sRGB|160|240|1') { throw 'Actual JPEG did not fully decode as the oriented single sRGB frame.' }
    $bytes=[IO.File]::ReadAllBytes($Path);$offset=2
    if($bytes[0] -ne 255 -or $bytes[1] -ne 216){throw 'Actual file lacks JPEG signature.'}
    while($offset+3 -lt $bytes.Length){
        if($bytes[$offset] -ne 255){throw 'Invalid output JPEG marker boundary.'}
        $marker=$bytes[$offset+1]
        if($marker -in @(218,217)){break}
        $length=([int]$bytes[$offset+2] -shl 8)+$bytes[$offset+3]
        if($length -lt 2 -or $offset+2+$length -gt $bytes.Length){throw 'Invalid output JPEG segment extent.'}
        if($marker -in @(225,226,237,254)){throw 'Output retained ICC or source private metadata.'}
        $offset+=2+$length
    }
    Add-Type -AssemblyName System.Drawing
    $actual=[Drawing.Bitmap]::new($Path);$expected=[Drawing.Bitmap]::new($Reference)
    try {
        if($actual.Width -ne 160 -or $actual.Height -ne 240){throw 'Independent decoder found wrong orientation.'}
        foreach($y in @(30,90,150,210)) { foreach($x in @(40,120)) {
            $a=@(0.0,0.0,0.0);$b=@(0.0,0.0,0.0)
            for($py=$y-4;$py -le $y+4;$py++){for($px=$x-4;$px -le $x+4;$px++){
                $p=$actual.GetPixel($px,$py);$q=$expected.GetPixel($px,$py)
                $a[0]+=$p.R;$a[1]+=$p.G;$a[2]+=$p.B;$b[0]+=$q.R;$b[1]+=$q.G;$b[2]+=$q.B
            }}
            for($channel=0;$channel -lt 3;$channel++){if([Math]::Abs(($a[$channel]-$b[$channel])/81) -gt 12){throw ('Actual colour differs from independent reference at patch '+$x+','+$y)}}
        }}
    } finally {$actual.Dispose();$expected.Dispose()}
}

function Assert-IccAcceptanceOutput {
    param([string]$Assets, [ValidateSet('views','negative','mixed')][string]$Group, [string]$Output, [string]$MagickPath, [int]$ExitCode, [string]$ConsoleText)
    if($ExitCode -ne 2){throw ('Expected partial exit 2 for ICC corpus, got '+$ExitCode)}
    $manifest=Get-Content -LiteralPath (Join-Path $Assets 'cases.json') -Raw | ConvertFrom-Json
    $cases=@($manifest.files | Where-Object {$Group -eq 'mixed' -or $_.file.StartsWith($Group+'/')})
    $wanted=@($cases|Where-Object reference)
    $actual=@(Get-ChildItem -LiteralPath $Output -Recurse -Force -File | Where-Object {$_.Extension -in @('.jpg','.jpeg')})
    if($actual.Count -ne $wanted.Count){throw ('Actual output count differs: '+$actual.Count+' vs '+$wanted.Count)}
    foreach($case in $wanted){
        $relative=if($Group -eq 'mixed'){$case.file}else{[IO.Path]::GetFileName($case.file)}
        $relative=[IO.Path]::ChangeExtension($relative,'.jpeg')
        Assert-IccAcceptanceJpeg -Path (Join-Path $Output $relative) -Reference (Join-Path $Assets ('references/'+$case.reference+'.png')) -MagickPath $MagickPath
    }
    $reports=@(Get-ChildItem -LiteralPath (Join-Path $Output '.WinImgNormalizer/reports') -Filter '*.csv' -File)
    if($reports.Count -ne 1){throw 'Expected exactly one actual outcome CSV.'}
    $rows=@(Import-Csv -LiteralPath $reports[0].FullName | ForEach-Object {
        foreach($property in $_.PSObject.Properties){if([string]$property.Value -like 'text:*'){$property.Value=([string]$property.Value).Substring(5)}}
        $_
    } | Where-Object RecordType -eq 'File')
    $converted=@($rows|Where-Object Status -eq 'Converted');$errors=@($rows|Where-Object Status -eq 'Error')
    $assumed=@($converted|Where-Object ColourPath -eq 'JpegIccAssumeSrgb');$managed=@($converted|Where-Object ColourPath -eq 'ProfileToSrgb')
    $expectedAssumed=if($Group -eq 'negative'){0}else{4};$expectedManaged=if($Group -eq 'negative'){0}else{5}
    if($rows.Count -ne $cases.Count -or $converted.Count -ne $wanted.Count -or $errors.Count -ne ($cases.Count-$wanted.Count)){throw 'CSV actual outcomes do not reconcile with the fixture corpus.'}
    if($assumed.Count -ne $expectedAssumed -or $managed.Count -ne $expectedManaged){throw 'CSV colour-managed/assumed counts do not reconcile.'}
    foreach($row in $assumed){if([int]$row.ColourWarning -ne 1 -or -not $row.ColourReason){throw 'Approximate preview lacks its warning/reason in CSV.'}}
    $logs=@(Get-ChildItem -LiteralPath (Join-Path $Output '.WinImgNormalizer/reports') -Filter '*.log' -File)
    if($logs.Count -ne 1){throw 'Expected exactly one actual local run log.'}
    $log=Get-Content -LiteralPath $logs[0].FullName -Raw
    if($log -notmatch ('ACCOUNTING Discovered='+$cases.Count+' Converted='+$wanted.Count+' CopiedVideo=0 SkippedDuplicate=0 Ignored=0 Error='+$errors.Count+' Cancelled=0 NotStarted=0 Balanced=True') -or
       $log -notmatch ('ColourManaged='+$expectedManaged+' ColourAssumed='+$expectedAssumed+' ColourWarnings='+$expectedAssumed)) {throw 'Actual final log accounting differs from the independently checked files and CSV.'}
    foreach($row in $errors){
        if($row.SourceRelativePath -like '*01-damaged-scan*' -and $row.Reason -notmatch 'full pixel decode.*DamagedInput'){throw 'Damaged scan did not fail its full pixel decode.'}
        if($row.SourceRelativePath -like '*02-untagged-cmyk*' -and $row.Reason -notmatch 'Untagged CMYK'){throw 'Uncharacterized CMYK failure was not retained.'}
        if($row.SourceRelativePath -like '*03-cmyk-bad-profile*' -and $row.Reason -notmatch 'cannot characterize.*components'){throw 'Invalid-profile CMYK failure was not retained.'}
    }
    if($expectedAssumed -gt 0 -and $ConsoleText -notmatch [regex]::Escape('JPEG saved; embedded colour profile unusable; sRGB assumed, colours may differ.')){throw 'Console omitted the approximate preview notice.'}
    if(@(Get-ChildItem -LiteralPath (Join-Path $Output '.WinImgNormalizer/work') -Force).Count -ne 0){throw 'A partial candidate or source snapshot remained in work storage.'}
    return [pscustomobject]@{InputCount=$cases.Count;OutputCount=$actual.Count;ColourManaged=$managed.Count;ColourAssumed=$assumed.Count;GenuineFailures=$errors.Count;ExitCode=$ExitCode;IndependentReferencesPassed=$true;FailureReasons=@($errors|ForEach-Object Reason)}
}
