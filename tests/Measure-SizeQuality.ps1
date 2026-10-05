# Development-only synthetic benchmark; no application dependency or algorithm change.
# Run in a fresh -NoProfile PS 5.1/PS 7 process after a clean implementation commit:
# .\tests\Measure-SizeQuality.ps1 -ImplementationCommit <SHA> -ResultDirectory .scratch\new-benchmark
# Use -Exploratory for a copied uncommitted runtime; its results are not clean-I gates.
[CmdletBinding()]
param(
    [string]$ImplementationCommit,
    [ValidatePattern('^[0-9a-f]{40}$')][string]$BaselineCommit = 'f9b00d8befac93c4fa1116efcbc2e9d7144f4509',
    [string]$ResultDirectory,
    [string]$MagickPath,
    [long[]]$Caps = @(1048576, 262144, 65537, 1024),
    [ValidateRange(1,5)][int]$RepeatCount = 1,
    [ValidateRange(128,2400)][int]$FixtureWidth = 1600,
    [ValidateRange(128,2400)][int]$FixtureHeight = 1200,
    [switch]$Exploratory
)
$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
$scratch = [IO.Path]::GetFullPath((Join-Path $repository '.scratch'))
Set-Location -LiteralPath $repository
if ($env:OS -ne 'Windows_NT') { throw 'This benchmark requires actual Windows.' }
if ($PSVersionTable.PSVersion.Major -lt 5) { throw 'Use Windows PowerShell 5.1 or PowerShell 7.' }
if (@($Caps).Count -eq 0 -or @($Caps | Where-Object { $_ -lt 1 }).Count) { throw 'Caps must be positive exact byte counts.' }
if (-not $Exploratory) {
    if ($ImplementationCommit -notmatch '^[0-9a-f]{40}$') { throw 'Supply the exact clean implementation SHA, or explicitly select -Exploratory.' }
    if ((git rev-parse HEAD).Trim() -ne $ImplementationCommit -or (git status --porcelain=v1 --untracked-files=all)) { throw 'Exact-I benchmark requires a clean matching checkout.' }
}
if ([string]::IsNullOrWhiteSpace($ResultDirectory)) { $ResultDirectory = Join-Path $scratch ('M2-T04-benchmark-' + [guid]::NewGuid().ToString('N')) }
$ResultDirectory = [IO.Path]::GetFullPath($ResultDirectory)
if (-not $ResultDirectory.StartsWith($scratch + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw 'Benchmark output must be a fresh child of repository .scratch.' }
$cursor = [IO.DirectoryInfo]::new($ResultDirectory)
while ($cursor) {
    if ($cursor.Exists -and ($cursor.Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Benchmark ownership must not traverse a reparse point.' }
    $cursor = $cursor.Parent
}
if ([IO.Directory]::Exists($ResultDirectory) -or [IO.File]::Exists($ResultDirectory)) { throw 'Preserve existing benchmark results; select a new output directory.' }
$null = [IO.Directory]::CreateDirectory($ResultDirectory)
[IO.File]::WriteAllText((Join-Path $ResultDirectory '.benchmark-owned'), 'Synthetic M2-T04 benchmark; no recursive cleanup.')
$policyBefore = @(Get-ExecutionPolicy -List | Where-Object { $_.Scope -ne 'Process' } | ForEach-Object { [ordered]@{ scope=$_.Scope.ToString(); policy=$_.ExecutionPolicy.ToString() } })
function Get-BenchHash([string]$Path) { (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() }
function Get-BenchRelative([string]$Path) { $Path.Substring($repository.Length + 1).Replace('\','/') }
function Write-BenchJson([string]$Path, [object]$Value) { [IO.File]::WriteAllText($Path, (ConvertTo-Json -InputObject $Value -Depth 18), [Text.UTF8Encoding]::new($false)) }
function Read-BenchGitBlob([string]$Revision, [string]$Path) {
    $info = [Diagnostics.ProcessStartInfo]::new()
    $info.FileName = (Get-Command git -CommandType Application | Select-Object -First 1).Source
    $info.Arguments = 'show ' + $Revision + ':' + $Path
    $info.UseShellExecute = $false; $info.CreateNoWindow = $true
    $info.RedirectStandardOutput = $true; $info.RedirectStandardError = $true
    $process = [Diagnostics.Process]::Start($info)
    $memory = [IO.MemoryStream]::new()
    try {
        $process.StandardOutput.BaseStream.CopyTo($memory)
        $errorText = $process.StandardError.ReadToEnd(); $process.WaitForExit()
        if ($process.ExitCode -ne 0) { throw 'Could not read the requested baseline Git blob.' }
        return ,$memory.ToArray()
    } finally { $memory.Dispose(); $process.Dispose() }
}
$runtimePath = Join-Path $repository 'WinImgNormalizer.ps1'
$baselinePath = Join-Path $ResultDirectory 'baseline-runtime.ps1'
$candidatePath = Join-Path $ResultDirectory 'candidate-runtime.ps1'
[IO.File]::WriteAllBytes($baselinePath, (Read-BenchGitBlob $BaselineCommit 'WinImgNormalizer.ps1'))
[IO.File]::WriteAllBytes($candidatePath, [IO.File]::ReadAllBytes($runtimePath))
$sourceBindings = @()
foreach ($path in @('WinImgNormalizer.ps1','tests/Measure-SizeQuality.ps1','tests/Initialize-TestDependencies.ps1','tests/dependencies.json','tests/legacy/toolchain.json')) {
    $sourceBindings += [ordered]@{ path=$path; observed_checkout_sha256=Get-BenchHash (Join-Path $repository $path) }
}
$dependencies = & (Join-Path $PSScriptRoot 'Initialize-TestDependencies.ps1') -MagickPath $MagickPath
$MagickPath = $dependencies.MagickPath
$nativeRecords = [Collections.Generic.List[object]]::new()
function Invoke-BenchNative([string[]]$NativeArguments, [string]$Purpose) {
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $previous = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
    try { $messages = @(& $MagickPath @NativeArguments 2>&1); $code = $LASTEXITCODE }
    finally { $ErrorActionPreference = $previous; $clock.Stop() }
    $nativeRecords.Add([ordered]@{ purpose=$Purpose; arguments=$NativeArguments.Clone(); native_exit=$code; elapsed_ms=$clock.Elapsed.TotalMilliseconds; messages=($messages -join "`n") })
    if ($code -ne 0) { throw ('Benchmark native operation failed: ' + $Purpose) }
    if ($Purpose -notin @('tool version','normalized JPEG full decode') -and -not [string]::IsNullOrWhiteSpace(($messages -join "`n"))) { throw ('Unexpected native diagnostics: ' + $Purpose) }
    return ($messages -join "`n")
}
$version = Invoke-BenchNative @('-version') 'tool version'
if ($version -notmatch 'ImageMagick 7\.1\.2-32 ') { throw 'The benchmark requires the pinned ImageMagick 7.1.2-32.' }

# Integer recipes avoid fonts and version-specific random-image generators.
Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Text;
public static class WinImgSizeQuality {
  private static uint Step(ref uint value) { value ^= value << 13; value ^= value >> 17; value ^= value << 5; return value; }
  public static void Ppm(string path, int width, int height, int kind) {
    uint seed = 0x6d2b79f5u;
    using (FileStream file = new FileStream(path, FileMode.CreateNew)) {
      byte[] header = Encoding.ASCII.GetBytes("P6\n" + width + " " + height + "\n255\n"); file.Write(header,0,header.Length);
      byte[] row = new byte[width * 3];
      for (int y=0;y<height;y++) {
        for (int x=0;x<width;x++) {
          int offset=x*3;
          if (kind==0) { row[offset]=(byte)(255L*x/(width-1)); row[offset+1]=(byte)(255L*y/(height-1)); row[offset+2]=(byte)(255L*(x+y)/(width+height-2)); }
          else if (kind==1) {
            int band=((x/16)+(y/16))%6;
            row[offset]=(byte)((band==0||band==3||band==5)?255:0); row[offset+1]=(byte)((band==1||band==3||band==4)?255:0); row[offset+2]=(byte)((band==2||band==4||band==5)?255:0);
            if (x%31==0||y%29==0) { row[offset]=255;row[offset+1]=255;row[offset+2]=255; }
            if ((x+y)%43==0) { row[offset]=0;row[offset+1]=0;row[offset+2]=0; }
          } else { row[offset]=(byte)(Step(ref seed)>>24); row[offset+1]=(byte)(Step(ref seed)>>24); row[offset+2]=(byte)(Step(ref seed)>>24); }
        }
        file.Write(row,0,row.Length);
      }
    }
  }
  public static void Video(string path, int length) {
    uint seed=0x47c0ffeeu; byte[] data=new byte[length]; for(int i=0;i<length;i++) data[i]=(byte)(Step(ref seed)>>24); File.WriteAllBytes(path,data);
  }
  public static double[] Metrics(string referencePath,string actualPath) {
    byte[] reference=File.ReadAllBytes(referencePath),actual=File.ReadAllBytes(actualPath);
    if(reference.Length==0||reference.Length!=actual.Length) throw new InvalidOperationException("Metric buffers differ.");
    double absolute=0,squared=0; for(int i=0;i<reference.Length;i++) { double delta=(double)reference[i]-actual[i]; absolute+=Math.Abs(delta); squared+=delta*delta; }
    double mse=squared/reference.Length; return new double[]{absolute/reference.Length,mse,mse==0?Double.PositiveInfinity:10*Math.Log10(255*255/mse)};
  }
}
'@
$fixtures = Join-Path $ResultDirectory 'fixtures'; $null=[IO.Directory]::CreateDirectory($fixtures)
$images = Join-Path $fixtures 'images'; $null=[IO.Directory]::CreateDirectory($images)
$videoSource = Join-Path $fixtures 'video'; $null=[IO.Directory]::CreateDirectory($videoSource)
$definitions = @(
    [ordered]@{ name='smooth-landscape'; width=$FixtureWidth; height=$FixtureHeight; kind=0; recipe='integer RGB x/y/diagonal gradients' },
    [ordered]@{ name='smooth-portrait'; width=$FixtureHeight; height=$FixtureWidth; kind=0; recipe='portrait integer RGB x/y/diagonal gradients' },
    [ordered]@{ name='hard-edges-detail'; width=$FixtureWidth; height=$FixtureHeight; kind=1; recipe='16-pixel saturated tile bands; white 31/29-period lines; black 43-period diagonals' },
    [ordered]@{ name='seeded-noise'; width=$FixtureWidth; height=$FixtureHeight; kind=2; recipe='xorshift32 seed 0x6d2b79f5; high byte per RGB8 channel' }
)
$fixtureByHash = @{}
foreach ($fixture in $definitions) {
    $ppm = Join-Path $fixtures ($fixture.name + '.ppm'); [WinImgSizeQuality]::Ppm($ppm,$fixture.width,$fixture.height,$fixture.kind)
    $path = Join-Path $images ($fixture.name + '.png')
    $null=Invoke-BenchNative @('-regard-warnings','-define','registry:filename:literal=true',$ppm,'-colorspace','sRGB','-strip',('PNG:'+$path)) ('fixture '+$fixture.name)
    $fixture.path=$path; $fixture.sha256=Get-BenchHash $path; $fixture.bytes=([IO.FileInfo]::new($path)).Length
    $fixtureByHash[$fixture.sha256]=$fixture
}
$lossy = [ordered]@{ name='already-lossy-jpeg'; width=$FixtureWidth; height=$FixtureHeight; kind=3; recipe='hard-edges-detail encoded once at JPEG quality 55, 4:2:0, progressive; benchmark reference is its decoded existing JPEG' }
$lossy.path=Join-Path $images ($lossy.name+'.jpg')
$null=Invoke-BenchNative @('-regard-warnings','-define','registry:filename:literal=true',(Join-Path $images 'hard-edges-detail.png'),'-strip','-sampling-factor','4:2:0','-interlace','Line','-quality','55',('JPEG:'+$lossy.path)) 'already-lossy source'
$lossy.sha256=Get-BenchHash $lossy.path; $lossy.bytes=([IO.FileInfo]::new($lossy.path)).Length
$definitions += $lossy; $fixtureByHash[$lossy.sha256]=$lossy
$videoPath=Join-Path $videoSource 'opaque-copy.mp4'; [WinImgSizeQuality]::Video($videoPath,3145728)
foreach ($file in @($definitions.path)+@($videoPath)) {
    $item=[IO.FileInfo]::new($file); $item.CreationTimeUtc=[datetime]'2010-01-02T03:04:05Z'; $item.LastWriteTimeUtc=[datetime]'2011-02-03T04:05:06Z'
}
function Get-BenchSourceState {
    @(@($definitions.path)+@($videoPath) | ForEach-Object {
        $item=[IO.FileInfo]::new($_); [ordered]@{path=Get-BenchRelative $_; sha256=Get-BenchHash $_; bytes=$item.Length; creation_utc=$item.CreationTimeUtc.ToString('o'); modified_utc=$item.LastWriteTimeUtc.ToString('o'); attributes=$item.Attributes.ToString()}
    })
}
$before = Get-BenchSourceState
$beforeJson = ConvertTo-Json -InputObject $before -Depth 6 -Compress
$metricsRoot=Join-Path $ResultDirectory 'metrics'; $null=[IO.Directory]::CreateDirectory($metricsRoot)
function Export-BenchRgb([string]$Source, [string]$Destination, [int]$Width=0, [int]$Height=0) {
    $arguments=@('-regard-warnings','-define','registry:filename:literal=true',$Source,'-auto-orient','-colorspace','sRGB','-background','white','-alpha','remove','-alpha','off')
    if ($Width -gt 0) { $arguments+=@('-resize',($Width.ToString()+'x'+$Height.ToString()+'!')) }
    $arguments+=@('-depth','8',('RGB:'+$Destination)); $null=Invoke-BenchNative $arguments 'metric RGB8 full decode'
}
foreach ($fixture in $definitions) {
    $fixture.reference_rgb=Join-Path $metricsRoot ($fixture.name+'-original.rgb')
    Export-BenchRgb $fixture.path $fixture.reference_rgb
    if (([IO.FileInfo]::new($fixture.reference_rgb)).Length -ne $fixture.width*$fixture.height*3) { throw 'Unexpected original RGB8 dimensions.' }
}
# Historical/current snapshots share a host and cannot replace an Add-Type
# class. Initialize the current wrapper once and restore it after each import;
# the older algorithm/default conversion scriptblock still supplies its flags.
$currentTokens=$null; $currentErrors=$null
$currentAst=[Management.Automation.Language.Parser]::ParseFile($candidatePath,[ref]$currentTokens,[ref]$currentErrors)
if ($currentErrors.Count) { throw 'Current snapshot runtime did not parse.' }
$currentFunction=$currentAst.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Invoke-WinImgNormalizer'},$true)
$currentHasObserver=@($currentFunction.Body.ParamBlock.Parameters | Where-Object {$_.Name.VariablePath.UserPath -eq 'NativeProcessObserver'}).Count -eq 1
$boundedNativeWrapper=$null
if ($currentHasObserver) {
    . $candidatePath
    Initialize-WinImgProcessApi
    $boundedNativeWrapper=(Get-Command Invoke-WinImgNativeProcess -CommandType Function).ScriptBlock
}
$rows=[Collections.Generic.List[object]]::new(); $applicationRuns=[Collections.Generic.List[object]]::new()
$flagShapes=@{}; $baselineNonDivisible=@{}
for ($repeat=1; $repeat -le $RepeatCount; $repeat++) {
    $order=if($repeat%2){@('baseline','candidate')}else{@('candidate','baseline')}
    foreach ($runtimeLabel in $order) {
        $selectedPath=if($runtimeLabel -eq 'baseline'){$baselinePath}else{$candidatePath}
        . $selectedPath
        $tokens=$null; $parseErrors=$null
        $ast=[Management.Automation.Language.Parser]::ParseFile($selectedPath,[ref]$tokens,[ref]$parseErrors)
        if ($parseErrors.Count) { throw 'Snapshot runtime did not parse.' }
        $function=$ast.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Invoke-WinImgNormalizer'},$true)
        if ($boundedNativeWrapper) { Set-Item -LiteralPath Function:Invoke-WinImgNativeProcess -Value $boundedNativeWrapper }
        $useNativeObserver=@($function.Body.ParamBlock.Parameters | Where-Object {$_.Name.VariablePath.UserPath -eq 'NativeProcessObserver'}).Count -eq 1
        $parameter=@($function.Body.ParamBlock.Parameters | Where-Object {$_.Name.VariablePath.UserPath -eq 'ProcessRunner'})[0]
        $actualRunner=$null
        if (-not $useNativeObserver) {
            if (-not $parameter.DefaultValue) { throw 'Historical conversion runner is unavailable.' }
            $actualRunner=$parameter.DefaultValue.ScriptBlock.GetScriptBlock()
        }
        foreach ($cap in $Caps) {
            $attempts=[Collections.Generic.List[object]]::new(); $inputHashes=@{}
            $tracedRunner={
                param([string]$Executable,[string[]]$Arguments,[int]$RemainingMilliseconds,[hashtable]$ChildEnvironment)
                if ($useNativeObserver -and -not $Arguments[-1].StartsWith('JPEG:',[StringComparison]::Ordinal)) {
                    return & $boundedNativeWrapper -Executable $Executable -Arguments $Arguments -TimeoutMilliseconds $RemainingMilliseconds -Environment $ChildEnvironment
                }
                # Retain actual native limits in the trace. Compare only the
                # existing image operations after the leading limit triplets.
                $skip=0; $resourceLimits=@()
                while ($skip+2 -lt $Arguments.Length -and $Arguments[$skip] -eq '-limit') {
                    if ($Arguments[$skip+1] -notin @('memory','map','disk','thread','time')) { throw 'Unexpected native resource operand.' }
                    $resourceLimits += [ordered]@{name=$Arguments[$skip+1];value=$Arguments[$skip+2]}; $skip+=3
                }
                $coreArguments=[string[]]@($Arguments | Select-Object -Skip $skip)
                $nativeInputs=@($coreArguments | Where-Object { $_ -match '(?i)source\.(png|jpg)$' })
                if ($nativeInputs.Count -ne 1) { throw 'Cannot identify the actual owned benchmark source snapshot.' }
                $path=$nativeInputs[0]; if($path.StartsWith('\\?\')){$path=$path.Substring(4)}
                if (-not $inputHashes.ContainsKey($path)) {
                    $hasher=[Security.Cryptography.SHA256]::Create()
                    try { $inputHashes[$path]=[BitConverter]::ToString($hasher.ComputeHash([IO.File]::ReadAllBytes($path))).Replace('-','').ToLowerInvariant() }
                    finally { $hasher.Dispose() }
                }
                $fixture=$fixtureByHash[$inputHashes[$path]]; if(-not $fixture){throw 'Native source bytes do not match a benchmark fixture.'}
                $resize=[Array]::IndexOf($coreArguments,'-resize'); $scale=[int]$coreArguments[$resize+1].TrimEnd('%')
                $extent=@($coreArguments | Where-Object {$_ -like 'jpeg:extent=*'})[0]
                $shape=@($coreArguments | ForEach-Object {if($_ -eq $nativeInputs[0]){'<owned-input>'}elseif($_ -like 'JPEG:*'){'JPEG:<owned-output>'}elseif($_ -like 'jpeg:extent=*'){'jpeg:extent=<budget>'}else{$_}})
                $clock=[Diagnostics.Stopwatch]::StartNew()
                try {
                    $result=if ($useNativeObserver) { & $boundedNativeWrapper -Executable $Executable -Arguments $Arguments -TimeoutMilliseconds $RemainingMilliseconds -Environment $ChildEnvironment } else { & $actualRunner $Executable $Arguments }
                } finally {$clock.Stop()}
                # Preserve diagnostics from the separated result and older Git runtimes.
                $measurementDiagnostics = if ($result.PSObject.Properties['StdOut'] -and $result.PSObject.Properties['StdErr']) { [string]$result.StdOut + "`n" + [string]$result.StdErr } else { [string]$result.DiagnosticOutput }
                $attempts.Add([ordered]@{ fixture=$fixture.name; scale=$scale; extent=$extent; arguments=$Arguments.Clone(); normalized_flags=$shape; resource_limits=$resourceLimits; remaining_budget_ms=$(if($useNativeObserver){$RemainingMilliseconds}else{$null}); native_exit=$result.ExitCode; diagnostics=$measurementDiagnostics; native_elapsed_ms=$clock.Elapsed.TotalMilliseconds })
                return $result
            }.GetNewClosure()
            $parent=Join-Path $ResultDirectory ('outputs-'+$runtimeLabel+'-'+$repeat+'-'+$cap); $null=[IO.Directory]::CreateDirectory($parent)
            $clock=[Diagnostics.Stopwatch]::StartNew()
            $output=if ($useNativeObserver) {
                @(& Invoke-WinImgNormalizer -Source $images -MaxBytes $cap -OutputParent $parent -MagickPath $MagickPath -NativeProcessObserver $tracedRunner 6>&1 3>&1 2>&1)
            } else { @(& Invoke-WinImgNormalizerCommand -Arguments @($images,$cap) -OutputParent $parent -MagickPath $MagickPath -ProcessRunner $tracedRunner 6>&1 3>&1 2>&1) }
            $clock.Stop(); $codes=@($output | Where-Object {$_ -is [int] -or $_ -is [long]})
            if ($codes.Count -ne 1 -or $codes[0] -notin @(0,2)) { throw 'Benchmark application returned an unexpected result.' }
            $children=@([IO.Directory]::GetDirectories($parent)); if($children.Count -ne 1){throw 'Expected one isolated application run.'}
            $run=$children[0]; $log=@([IO.Directory]::GetFiles($run,'*.log',[IO.SearchOption]::AllDirectories)); if($log.Count -ne 1){throw 'Expected one actual application log.'}
            $logText=[IO.File]::ReadAllText($log[0]); if($logText -notmatch 'Errors=0'){throw 'Application benchmark reported source/encoder errors.'}
            $console=Join-Path $parent 'application-console.log'; [IO.File]::WriteAllText($console,($output -join "`n"),[Text.UTF8Encoding]::new($false))
            $applicationRuns.Add([ordered]@{runtime=$runtimeLabel;repeat=$repeat;cap_bytes=$cap;application_return=$codes[0];elapsed_ms=$clock.Elapsed.TotalMilliseconds;log=Get-BenchRelative $log[0];log_sha256=Get-BenchHash $log[0];console=Get-BenchRelative $console;console_sha256=Get-BenchHash $console;conversion_attempts=$attempts.ToArray()})
            Write-BenchJson (Join-Path $ResultDirectory 'benchmark-progress.json') ([ordered]@{
                classification='partial development observation; not a completed benchmark'
                baseline_commit=$BaselineCommit
                baseline_git_blob_sha256=Get-BenchHash $baselinePath
                candidate_tested_raw_sha256=Get-BenchHash $candidatePath
                source_bindings=$sourceBindings
                application_runs=$applicationRuns.ToArray()
                native_operations=$nativeRecords.ToArray()
            })
            foreach ($fixture in $definitions) {
                $actual=Join-Path $run ($fixture.name+'.jpeg'); if(-not [IO.File]::Exists($actual)){throw 'Missing actual normalized JPEG.'}
                $identified=Invoke-BenchNative @('identify','+ping','-regard-warnings','-define','registry:filename:literal=true','-format','%m|%w|%h|%n|%Q',$actual) 'normalized JPEG full decode'
                $parts=$identified.Split('|'); if($parts.Count -ne 5 -or $parts[0] -ne 'JPEG' -or $parts[3] -ne '1'){throw 'Output is not one fully decoded JPEG.'}
                $width=[int]$parts[1]; $height=[int]$parts[2]; $file=[IO.FileInfo]::new($actual)
                $actualAttempts=@($attempts | Where-Object {$_.fixture -eq $fixture.name})
                $scales=@($actualAttempts | ForEach-Object {$_.scale}); if($scales.Count -eq 0 -or $scales.Count -gt 6){throw 'Missing or unexpected actual conversion attempts.'}
                $allowed=@(100,90,80,70,60,50); for($n=0;$n -lt $scales.Count;$n++){if($scales[$n] -ne $allowed[$n]){throw 'Scale sequence changed.'}}
                foreach($attempt in $actualAttempts){
                    $key=$fixture.name+'-'+$attempt.scale
                    $shape=ConvertTo-Json -InputObject $attempt.normalized_flags -Compress
                    if($flagShapes.ContainsKey($key)){if($flagShapes[$key] -ne $shape){throw 'Conversion flags differ beyond the extent budget.'}}else{$flagShapes[$key]=$shape}
                    if($attempt.native_exit -ne 0 -or -not [string]::IsNullOrWhiteSpace($attempt.diagnostics)){throw 'A measured native conversion failed or emitted diagnostics.'}
                }
                $prefix=Join-Path $metricsRoot ($runtimeLabel+'-'+$repeat+'-'+$cap+'-'+$fixture.name)
                $decoded=$prefix+'-actual.rgb'; Export-BenchRgb $actual $decoded
                $reference=$prefix+'-reference-at-output.rgb'; Export-BenchRgb $fixture.path $reference $width $height
                $upscaled=$prefix+'-upscaled.rgb'; Export-BenchRgb $actual $upscaled $fixture.width $fixture.height
                $atOutput=[WinImgSizeQuality]::Metrics($reference,$decoded); $atSource=[WinImgSizeQuality]::Metrics($fixture.reference_rgb,$upscaled)
                $sha=Get-BenchHash $actual
                $nativeElapsed=($actualAttempts | ForEach-Object {[double]$_.native_elapsed_ms} | Measure-Object -Sum).Sum
                $rows.Add([ordered]@{
                    runtime=$runtimeLabel; repeat=$repeat; fixture=$fixture.name; cap_bytes=$cap
                    source_bytes=$fixture.bytes; output_bytes=$file.Length; compliant=($file.Length -le $cap)
                    computed_status=$(if($file.Length -le $cap){'compliant'}else{'valid_above_target'})
                    width=$width; height=$height; selected_scale=$scales[-1]; attempt_count=$scales.Count
                    encoder_quality_estimate=[int]$parts[4]; native_conversion_elapsed_ms=$nativeElapsed
                    output_mae_rgb8=$atOutput[0]
                    output_psnr_db=$(if([double]::IsInfinity($atOutput[2])){$null}else{$atOutput[2]})
                    output_psnr_infinite=[double]::IsInfinity($atOutput[2])
                    source_grid_mae_rgb8=$atSource[0]
                    source_grid_psnr_db=$(if([double]::IsInfinity($atSource[2])){$null}else{$atSource[2]})
                    source_grid_psnr_infinite=[double]::IsInfinity($atSource[2])
                    output=Get-BenchRelative $actual; output_sha256=$sha; application_return=$codes[0]
                })
                if($cap -eq 65537){$key=$repeat.ToString()+'-'+$fixture.name;if($baselineNonDivisible.ContainsKey($key)){if($baselineNonDivisible[$key] -ne $sha){throw 'Identical 65537B budgets produced different JPEG bytes.'}}else{$baselineNonDivisible[$key]=$sha}}
            }
            $capRows=@($rows | Where-Object {$_.runtime -eq $runtimeLabel -and $_.repeat -eq $repeat -and $_.cap_bytes -eq $cap})
            $aboveRows=@($capRows | Where-Object {$_.output_bytes -gt $cap})
            if($capRows.Count -ne $definitions.Count){throw 'A fixture is missing from the measured output set.'}
            if(@($aboveRows | Where-Object {$_.selected_scale -ne 50}).Count){throw 'A valid above-target result was not the final 50-percent attempt.'}
            if($runtimeLabel -eq 'candidate') {
                $expectedCode=if($aboveRows.Count){2}else{0}
                if($codes[0] -ne $expectedCode){throw 'Candidate application return did not match the exact above-target outcome.'}
                $summary=[regex]::Match($logText,'SUMMARY ConvertedImages=(\d+) CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=0 SizeWarnings=(\d+)')
                if(-not $summary.Success -or [int]$summary.Groups[1].Value -ne $definitions.Count -or [int]$summary.Groups[2].Value -ne $aboveRows.Count){throw 'Candidate SizeWarnings summary does not match actual retained JPEG byte lengths.'}
                foreach($row in $capRows) {
                    $fixture=@($definitions | Where-Object {$_.name -eq $row.fixture})[0]
                    $sourceName=[IO.Path]::GetFileName($fixture.path)
                    $level=if($row.output_bytes -gt $cap){'WARN'}else{'OK'}
                    $expected=('{0} IMG: {1} -> {2}.jpeg [{3} bytes, MaxBytes={4}, Width={5}, Height={6}, Scale={7}%]' -f $level,$sourceName,$row.fixture,$row.output_bytes.ToString([Globalization.CultureInfo]::InvariantCulture),$cap.ToString([Globalization.CultureInfo]::InvariantCulture),$row.width,$row.height,$row.selected_scale)
                    $line=@($logText -split '\r?\n' | Where-Object {$_.Contains($expected)})
                    if($line.Count -ne 1 -or -not $line[0].Contains('['+$level+'] '+$expected)){throw 'Candidate item log does not report its exact bytes, dimensions, scale and status.'}
                    if($level -eq 'WARN' -and -not $line[0].Contains('Could not reach target; best-effort saved')){throw 'Above-target log omitted the retained best-effort qualification.'}
                }
            } elseif($codes[0] -ne 0) { throw 'Baseline comparison unexpectedly returned a nonzero application result.' }
            if((ConvertTo-Json -InputObject (Get-BenchSourceState) -Depth 6 -Compress) -ne $beforeJson){throw 'Benchmark modified source bytes or creation/modification state.'}
        }
        $parent=Join-Path $ResultDirectory ('video-'+$runtimeLabel+'-'+$repeat);$null=[IO.Directory]::CreateDirectory($parent)
        $clock=[Diagnostics.Stopwatch]::StartNew();$output=@(& Invoke-WinImgNormalizerCommand -Arguments @($videoSource) -OutputParent $parent -MagickPath $MagickPath 6>&1 3>&1 2>&1);$clock.Stop()
        $codes=@($output | Where-Object {$_ -is [int] -or $_ -is [long]});if($codes.Count -ne 1 -or $codes[0] -ne 0){throw 'Actual opaque-video application copy failed.'}
        $videoRun=@([IO.Directory]::GetDirectories($parent));$copied=Join-Path $videoRun[0] 'opaque-copy.mp4'
        if((Get-BenchHash $copied) -ne (Get-BenchHash $videoPath)){throw 'Opaque video copy differs.'}
        $source=[IO.FileInfo]::new($videoPath);$target=[IO.FileInfo]::new($copied)
        if($source.CreationTimeUtc -ne $target.CreationTimeUtc -or $source.LastWriteTimeUtc -ne $target.LastWriteTimeUtc){throw 'Copied video timestamps differ.'}
        $rows.Add([ordered]@{runtime=$runtimeLabel;repeat=$repeat;fixture='opaque-video-copy';source_bytes=$source.Length;output_bytes=$target.Length;byte_identical=$true;timestamps_preserved=$true;application_return=$codes[0];application_elapsed_ms=$clock.Elapsed.TotalMilliseconds;output=Get-BenchRelative $copied;output_sha256=Get-BenchHash $copied})
    }
}
$policyAfter=@(Get-ExecutionPolicy -List | Where-Object {$_.Scope -ne 'Process'} | ForEach-Object {[ordered]@{scope=$_.Scope.ToString();policy=$_.ExecutionPolicy.ToString()}})
if((ConvertTo-Json -InputObject $policyBefore -Compress) -ne (ConvertTo-Json -InputObject $policyAfter -Compress)){throw 'Persistent execution policy changed.'}
if((ConvertTo-Json -InputObject (Get-BenchSourceState) -Depth 6 -Compress) -ne $beforeJson){throw 'Source preservation check failed.'}
foreach($binding in $sourceBindings) {
    $binding.observed_end_checkout_sha256=Get-BenchHash (Join-Path $repository $binding.path)
    if($binding.observed_end_checkout_sha256 -ne $binding.observed_checkout_sha256){throw 'A benchmark-bound source/harness file changed during execution.'}
}
if((Get-BenchHash $candidatePath) -ne $sourceBindings[0].observed_checkout_sha256){throw 'Candidate runtime snapshot no longer matches its starting raw hash.'}
if(-not $Exploratory -and ((git rev-parse HEAD).Trim() -ne $ImplementationCommit -or (git status --porcelain=v1 --untracked-files=all))){throw 'Clean-I checkpoint changed during the benchmark.'}
$operatingSystem=Get-CimInstance Win32_OperatingSystem
$record=[ordered]@{schema_version=1;task_id='M2-T04';observed_at_utc=[datetime]::UtcNow.ToString('o');classification=$(if($Exploratory){'uncommitted exploratory benchmark on owned copied runtime bytes'}else{'clean implementation benchmark'});implementation_commit=$ImplementationCommit;baseline_commit=$BaselineCommit;baseline_git_blob_sha256=Get-BenchHash $baselinePath;candidate_tested_raw_sha256=Get-BenchHash $candidatePath;candidate_checkout_snapshot_binding='Exact raw bytes copied at benchmark start; baseline is exact Git blob bytes, whose line endings can differ from checkout bytes.';source_bindings=$sourceBindings;environment=[ordered]@{powershell_version=$PSVersionTable.PSVersion.ToString();powershell_edition=$PSVersionTable.PSEdition;os_version=$operatingSystem.Version;os_caption=$operatingSystem.Caption;os_build=$operatingSystem.BuildNumber;host_executable=(Get-Process -Id $PID).Path;architecture_bits=[IntPtr]::Size*8;culture=[Globalization.CultureInfo]::CurrentCulture.Name;imagemagick_sha256=$dependencies.ImageMagickExecutableSha256;imagemagick_version=$version};caps=$Caps;repeat_count=$RepeatCount;fixture_dimensions=@($FixtureWidth,$FixtureHeight);fixture_seed='xorshift32 0x6d2b79f5';source_preservation=$before;source_preserved=$true;source_harness_start_end_hashes_equal=$true;candidate_warning_contract_verified=$true;persistent_policy_unchanged=$true;conversion_flags_equal_except_extent=$true;native_limit_trace_separate_from_core_image_flags=$true;historical_wrapper_compatibility=$(if($currentHasObserver){'Current bounded native wrapper restored after each runtime import; historical image algorithm and default runner flags preserved. Current observer calls retain actual resources, environment and shared remaining deadline.'}else{'Historical native wrapper/default runner unchanged.'});scale_sequence_prefix_verified=@(100,90,80,70,60,50);non_kib_divisible_65537B_comparison_identical=($Caps -contains 65537);metric_method='RGB8 MAE and PSNR calculated across every decoded channel sample. Output-grid reference is original resized to retained dimensions using pinned ImageMagick default resizing; source-grid metric upscales the actual output to original dimensions. Lossy-source reference is its existing decoded JPEG, not the earlier pristine source. These are synthetic encoded-sRGB metrics, not perceptual scores.';timing_method='Stopwatch native conversion latency and full application wall time; dependency bootstrap, fixture generation and metric decode are outside application timing. Application timing includes validation, source-hash tracing and log overhead. Fixed baseline/candidate order is reversed on even repeats. Single-machine synthetic runs are not statistically controlled performance or subjective owner-quality acceptance.';limitations=@('No new encoder/scale/chroma/default algorithm is proposed. Correcting decimal MB/KB to exact byte budgets can change default or binary-divisible output bytes/quality; 65537B provides an unchanged-budget comparison.','No arbitrary photography, print accuracy, owner aesthetic/default acceptance or cross-machine performance claim.','The public function return is captured; the invoking parent must separately preserve the fresh host native process exit. Raw paths, media and transcripts remain ignored.');application_runs=$applicationRuns.ToArray();rows=$rows.ToArray();native_operations=$nativeRecords.ToArray()}
Write-BenchJson (Join-Path $ResultDirectory 'benchmark.json') $record
$csv=Join-Path $ResultDirectory 'measurements.csv'
$columns=@('runtime','repeat','fixture','cap_bytes','source_bytes','output_bytes','computed_status','width','height','selected_scale','attempt_count','encoder_quality_estimate','native_conversion_elapsed_ms','application_elapsed_ms','output_mae_rgb8','output_psnr_db','source_grid_mae_rgb8','source_grid_psnr_db','byte_identical','timestamps_preserved','application_return')
$rows.ToArray() | ForEach-Object { [pscustomobject]$_ } | Select-Object $columns | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
Write-Output ('Benchmark saved: '+(Get-BenchRelative (Join-Path $ResultDirectory 'benchmark.json'))+'; '+$rows.Count+' measured image/video rows; source preserved; flag/scale checks passed.')
exit 0
