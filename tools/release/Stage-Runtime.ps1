# Development/build-time staging only. The application never calls this script.
[CmdletBinding()]
param(
    [switch]$Download,
    [string]$DestinationRoot,
    [string]$CacheDirectory,
    [switch]$IncludeDependencySource,
    [switch]$CacheOnly
)

$ErrorActionPreference = 'Stop'
$repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
if (-not $DestinationRoot) { $DestinationRoot = $repository }
if (-not $CacheDirectory) { $CacheDirectory = Join-Path $repository '.scratch/runtime-downloads' }
$destination = [IO.Path]::GetFullPath($DestinationRoot)
$cache = [IO.Path]::GetFullPath($CacheDirectory)
$lock = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'runtime-lock.json') -Raw | ConvertFrom-Json
$utf8 = [Text.UTF8Encoding]::new($false)

function Assert-StageAncestors([string]$Path) {
    $current = [IO.DirectoryInfo]::new([IO.Path]::GetFullPath($Path))
    while ($null -ne $current) {
        if ($current.Exists -and ($current.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
            throw 'Runtime staging may not traverse reparse points.'
        }
        $current = $current.Parent
    }
}

function Assert-StageFile([string]$Path, $Pin) {
    $item = Get-Item -LiteralPath $Path -ErrorAction Stop
    if ($item.PSIsContainer -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -or
        $item.Length -ne [long]$Pin.bytes -or
        (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() -cne [string]$Pin.sha256) {
        throw ('Runtime dependency length or SHA-256 mismatch: ' + [IO.Path]::GetFileName($Path))
    }
}

function Invoke-StageGitHubProcess {
    param([string]$Executable, [string[]]$Arguments, [int]$TimeoutMilliseconds = 900000)
    # Windows PowerShell 5.1 can promote native stderr to a terminating error
    # under ErrorActionPreference=Stop even for an expected nonzero fallback.
    # Capture both streams directly and make the actual exit code authoritative.
    $start = New-Object Diagnostics.ProcessStartInfo
    $start.FileName = $Executable
    $start.Arguments = ($Arguments | ForEach-Object {
        '"' + [regex]::Replace([regex]::Replace([string]$_, '(\\*)"', '$1$1\"'), '(\\+)$', '$1$1') + '"'
    }) -join ' '
    $start.UseShellExecute = $false; $start.CreateNoWindow = $true
    $start.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
    $start.RedirectStandardOutput = $true; $start.RedirectStandardError = $true
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $start
    $started = $false
    try {
        if (-not $process.Start()) { throw 'GitHub dependency request did not start.' }
        $started = $true
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit($TimeoutMilliseconds)) { $process.Kill(); throw 'GitHub dependency request timed out.' }
        if (-not $stdout.Wait(10000) -or -not $stderr.Wait(10000)) { throw 'GitHub dependency response capture did not complete.' }
        return [pscustomobject]@{ ExitCode=$process.ExitCode; StdOut=$stdout.GetAwaiter().GetResult(); StdErr=$stderr.GetAwaiter().GetResult() }
    } finally {
        if ($started -and -not $process.HasExited) { $process.Kill(); $null = $process.WaitForExit(5000) }
        $process.Dispose()
    }
}

function Get-StageAsset($Pin) {
    $target = Join-Path $cache $Pin.name
    Assert-StageAncestors $cache
    if (Test-Path -LiteralPath $target) { Assert-StageFile $target $Pin; return $target }
    if (-not $Download) { throw ('Runtime cache is absent; run tools/release/Stage-Runtime.ps1 -Download: ' + $Pin.name) }
    $null = [IO.Directory]::CreateDirectory($cache)
    $partial = $target + '.' + [Guid]::NewGuid().ToString('N') + '.partial'
    try {
        # The immutable mirrored asset survives upstream Actions retention. Its
        # reviewed hash is identical to the authenticated upstream artifact.
        $previousProtocol = [Net.ServicePointManager]::SecurityProtocol
        try {
            [Net.ServicePointManager]::SecurityProtocol = $previousProtocol -bor [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -UseBasicParsing -Uri $Pin.mirror_url -OutFile $partial -ErrorAction Stop
        } finally { [Net.ServicePointManager]::SecurityProtocol = $previousProtocol }
        Assert-StageFile $partial $Pin
        [IO.File]::Move($partial, $target)
        return $target
    } catch {
        # Keep any partial bytes for diagnosis. This is a developer-only fallback
        # to the original signed upstream build, never application discovery.
        $gh = (Get-Command gh -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
        # Before publication, the ordinary draft release holds only the exact
        # authenticated upstream assets. CI uses its existing read-only token
        # for this repository; no cross-repository credential is required.
        $draftDirectory = Join-Path $cache ('draft-assets-' + [Guid]::NewGuid().ToString('N'))
        $null = [IO.Directory]::CreateDirectory($draftDirectory)
        $draftResult = Invoke-StageGitHubProcess -Executable $gh -Arguments @('release','download','v1.0.1','--repo','PikkuJanne/WinImgNormalizer','--pattern',[string]$Pin.name,'--dir',$draftDirectory)
        if ($draftResult.ExitCode -eq 0) {
            $draftAsset = Join-Path $draftDirectory $Pin.name
            Assert-StageFile $draftAsset $Pin
            [IO.File]::Move($draftAsset, $target)
            return $target
        }
        $wrapper = $target + '.artifact.zip'
        if (-not (Test-Path -LiteralPath $wrapper)) {
            $start = New-Object Diagnostics.ProcessStartInfo
            $start.FileName = $gh
            $start.Arguments = 'api repos/ImageMagick/ImageMagick/actions/artifacts/' + $Pin.artifact_id + '/zip'
            $start.UseShellExecute = $false; $start.CreateNoWindow = $true
            $start.RedirectStandardOutput = $true; $start.RedirectStandardError = $true
            $process = New-Object Diagnostics.Process
            $process.StartInfo = $start
            $stream = [IO.File]::Open($wrapper, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
            try {
                if (-not $process.Start()) { throw 'Upstream artifact request did not start.' }
                $capture = $process.StandardOutput.BaseStream.CopyToAsync($stream)
                $errors = $process.StandardError.ReadToEndAsync()
                if (-not $process.WaitForExit(900000)) { $process.Kill(); throw 'Upstream artifact request timed out.' }
                if (-not $capture.Wait(10000) -or -not $errors.Wait(10000) -or $process.ExitCode -ne 0) {
                    throw 'Upstream artifact request failed; no runtime staged.'
                }
            } finally { $stream.Dispose(); $process.Dispose() }
        }
        Assert-StageFile $wrapper ([pscustomobject]@{bytes=$Pin.artifact_zip_bytes;sha256=$Pin.artifact_zip_sha256})
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [IO.Compression.ZipFile]::OpenRead($wrapper)
        try {
            if ($zip.Entries.Count -ne 1 -or $zip.Entries[0].FullName -cne $Pin.name) { throw 'Upstream artifact layout changed.' }
            $entryStream = $zip.Entries[0].Open()
            $output = [IO.File]::Open($target, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
            try { $entryStream.CopyTo($output) } finally { $output.Dispose(); $entryStream.Dispose() }
        } finally { $zip.Dispose() }
        Assert-StageFile $target $Pin
        return $target
    }
}

Assert-StageAncestors $destination
if (-not $CacheOnly -and -not (Test-Path -LiteralPath (Join-Path $destination 'WinImgNormalizer.ps1') -PathType Leaf)) {
    throw 'Stage into an existing WinImgNormalizer checkout or extraction only.'
}
$archive = Get-StageAsset $lock.portable_asset
$sbom = Get-StageAsset $lock.sbom_asset
$source = $null
if ($IncludeDependencySource) { $source = Get-StageAsset $lock.source_asset }
if ($CacheOnly) {
    return [pscustomobject]@{PortableArchive=$archive;Sbom=$sbom;DependencySource=$source}
}
$staging = Join-Path $cache ('staging-' + [Guid]::NewGuid().ToString('N'))
Assert-StageAncestors $staging
$null = [IO.Directory]::CreateDirectory((Join-Path $staging 'imagemagick'))
$tar = Join-Path $env:SystemRoot 'System32/tar.exe'
$names = @(& $tar -tf $archive)
if ($LASTEXITCODE -ne 0 -or $names.Count -eq 0) { throw 'Cannot inspect verified runtime archive.' }
foreach ($name in $names) {
    if ([string]::IsNullOrWhiteSpace($name) -or $name -match '[/\\:]' -or $name -in @('.', '..')) {
        throw 'Unexpected upstream portable archive layout.'
    }
}
$selected = @($lock.files | Where-Object path -ne 'imagemagick/upstream-sbom.cdx.json' | ForEach-Object { [IO.Path]::GetFileName($_.path) })
& $tar -xf $archive -C (Join-Path $staging 'imagemagick') @selected
if ($LASTEXITCODE -ne 0) { throw 'Runtime archive extraction failed.' }
[IO.File]::Copy($sbom, (Join-Path $staging 'imagemagick/upstream-sbom.cdx.json'), $false)
foreach ($file in $lock.files) { Assert-StageFile (Join-Path $staging $file.path) $file }
$engine = Join-Path $staging $lock.engine_relative_path
$signature = Get-AuthenticodeSignature -LiteralPath $engine
if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Thumbprint -cne $lock.authenticode_thumbprint -or
    $signature.SignerCertificate.Subject -cne $lock.authenticode_subject) {
    throw 'ImageMagick upstream publisher signature did not validate; no runtime staged.'
}
$sourceNotice = Join-Path $repository 'docs/release/IMAGEMAGICK_SOURCE.md'
[IO.File]::Copy($sourceNotice, (Join-Path $staging 'SOURCE_INFORMATION.md'), $false)
$notice = Get-Item -LiteralPath $sourceNotice
$files = @($lock.files) + @([ordered]@{path='SOURCE_INFORMATION.md';bytes=$notice.Length;sha256=(Get-FileHash -LiteralPath $sourceNotice).Hash.ToLowerInvariant()})
$manifest = [ordered]@{
    schema_version=1;engine_relative_path=$lock.engine_relative_path;engine_version=$lock.engine_version
    upstream_channel=$lock.upstream_channel;upstream_source_revision=$lock.upstream_source_revision
    upstream_workflow_url=$lock.upstream_workflow_url;portable_asset=$lock.portable_asset
    sbom_asset=$lock.sbom_asset;source_asset=$lock.source_asset;files=$files
}
$manifestPath = Join-Path $staging 'runtime-manifest.json'
[IO.File]::WriteAllText($manifestPath, ($manifest | ConvertTo-Json -Depth 10 -Compress) + "`n", $utf8)
$manifestHash = (Get-FileHash -LiteralPath $manifestPath).Hash.ToLowerInvariant()
$marker = [ordered]@{schema_version=1;runtime_required=$true;runtime_manifest_sha256=$manifestHash}
$markerBytes = $utf8.GetBytes(($marker | ConvertTo-Json -Compress) + "`n")
$runtime = Join-Path $destination 'runtime'
$markerPath = Join-Path $destination 'portable-runtime.json'
if (Test-Path -LiteralPath $runtime) {
    # Idempotent staging verifies every existing byte; it never repairs a damaged
    # runtime in place or overwrites a different checkout's files.
    Assert-StageAncestors $runtime
    $expectedFiles = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $expectedDirectories = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $null = $expectedFiles.Add('runtime-manifest.json')
    foreach ($file in $files) {
        $null = $expectedFiles.Add([string]$file.path)
        $parent = [IO.Path]::GetDirectoryName([string]$file.path)
        while ($parent) {
            $null = $expectedDirectories.Add($parent.Replace('\','/'))
            $parent = [IO.Path]::GetDirectoryName($parent)
        }
    }
    foreach ($item in Get-ChildItem -LiteralPath $runtime -Recurse -Force -ErrorAction Stop) {
        if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Existing runtime contains a reparse point; no files replaced.' }
        $relative = $item.FullName.Substring($runtime.Length + 1).Replace('\','/')
        if (($item.PSIsContainer -and -not $expectedDirectories.Contains($relative)) -or
            (-not $item.PSIsContainer -and -not $expectedFiles.Contains($relative))) {
            throw ('Existing runtime contains an unexpected entry; no files replaced: ' + $relative)
        }
    }
    foreach ($file in $files) { Assert-StageFile (Join-Path $runtime $file.path) $file }
    $existingManifest = Join-Path $runtime 'runtime-manifest.json'
    if (Test-Path -LiteralPath $existingManifest) {
        Assert-StageFile $existingManifest ([pscustomobject]@{bytes=(Get-Item -LiteralPath $manifestPath).Length;sha256=$manifestHash})
    } else { [IO.File]::Copy($manifestPath, $existingManifest, $false) }
} else { [IO.Directory]::Move($staging, $runtime) }
if (Test-Path -LiteralPath $markerPath) {
    if ([Convert]::ToBase64String([IO.File]::ReadAllBytes($markerPath)) -cne [Convert]::ToBase64String($markerBytes)) { throw 'Existing portable marker differs; no replacement permitted.' }
} else { [IO.File]::WriteAllBytes($markerPath, $markerBytes) }
[pscustomobject]@{RuntimeDirectory=$runtime;EnginePath=(Join-Path $runtime $lock.engine_relative_path);EngineVersion=$lock.engine_version;RuntimeManifestSha256=$manifestHash;PortableArchive=$archive;Sbom=$sbom;DependencySource=$source}
