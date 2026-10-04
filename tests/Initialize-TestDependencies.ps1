# Development-only bootstrap. No installation, persistent PATH or policy changes.
[CmdletBinding()]
param(
    [switch]$Download,
    [string]$MagickPath
)

$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path -Parent $PSScriptRoot
$scratchRoot = Join-Path $repositoryRoot '.scratch'
$specification = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'dependencies.json') -Raw | ConvertFrom-Json
$imageSpecification = (Get-Content -LiteralPath (Join-Path $PSScriptRoot $specification.imagemagick_specification) -Raw | ConvertFrom-Json).portable_imagemagick

function Assert-NoReparseDirectory {
    param([string]$Directory)
    $current = [IO.DirectoryInfo]::new([IO.Path]::GetFullPath($Directory))
    while ($null -ne $current) {
        if ($current.Exists -and (($current.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0)) {
            throw "Development tool directory must not traverse a reparse point: $($current.FullName)"
        }
        $current = $current.Parent
    }
}

function Assert-VerifiedFile {
    param([string]$File, [string]$Sha256, [long]$Bytes = 0)
    $item = Get-Item -LiteralPath $File
    if ($item.PSIsContainer -or (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0)) {
        throw "Expected a regular development dependency file: $File"
    }
    if ($Bytes -gt 0 -and $item.Length -ne $Bytes) { throw "Dependency length mismatch: $File" }
    if ((Get-FileHash -LiteralPath $File -Algorithm SHA256).Hash -ne $Sha256) {
        throw "Dependency SHA-256 mismatch: $File"
    }
}

function Get-VerifiedArchive {
    param([string]$File, [string]$Url, [string]$Sha256, [long]$Bytes)
    Assert-NoReparseDirectory (Split-Path -Parent $File)
    if (-not (Test-Path -LiteralPath $File)) {
        if (-not $Download) { throw "Dependency archive is absent. Explicitly use Initialize-TestDependencies.ps1 -Download or Invoke-Tests.ps1 -DownloadDependencies: $File" }
        [IO.Directory]::CreateDirectory((Split-Path -Parent $File)) | Out-Null
        $partial = $File + '.' + [guid]::NewGuid().ToString('N') + '.partial'
        $previousProtocol = [Net.ServicePointManager]::SecurityProtocol
        try {
            # TLS selection is limited to this process and restored after the request.
            [Net.ServicePointManager]::SecurityProtocol = $previousProtocol -bor [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -UseBasicParsing -Uri $Url -OutFile $partial
        }
        finally { [Net.ServicePointManager]::SecurityProtocol = $previousProtocol }
        Assert-VerifiedFile -File $partial -Sha256 $Sha256 -Bytes $Bytes
        [IO.File]::Move($partial, $File) # Never replace an existing archive.
    }
    Assert-VerifiedFile -File $File -Sha256 $Sha256 -Bytes $Bytes
    return $File
}

Assert-NoReparseDirectory $scratchRoot
$pesterArchive = Get-VerifiedArchive -File (Join-Path $scratchRoot ('tools/Pester-' + $specification.pester.version + '/Pester.' + $specification.pester.version + '.nupkg')) -Url $specification.pester.url -Sha256 $specification.pester.sha256 -Bytes $specification.pester.bytes

# Extract from verified bytes on every invocation; do not trust an old extracted module.
$extractRoot = Join-Path $scratchRoot ('test-dependencies/' + [guid]::NewGuid().ToString('N'))
Assert-NoReparseDirectory $extractRoot
if (Test-Path -LiteralPath $extractRoot) { throw 'Development extraction directory already exists.' }
[IO.Directory]::CreateDirectory($extractRoot) | Out-Null
$pesterDirectory = Join-Path $extractRoot 'Pester'
[IO.Directory]::CreateDirectory($pesterDirectory) | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::OpenRead($pesterArchive)
try {
    $prefix = [IO.Path]::GetFullPath($pesterDirectory) + [IO.Path]::DirectorySeparatorChar
    foreach ($entry in $zip.Entries) {
        $destination = [IO.Path]::GetFullPath((Join-Path $pesterDirectory $entry.FullName))
        if (-not $destination.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Dependency package contains an unsafe archive entry.'
        }
    }
}
finally { $zip.Dispose() }
[IO.Compression.ZipFile]::ExtractToDirectory($pesterArchive, $pesterDirectory)
$pesterManifest = Join-Path $pesterDirectory 'Pester.psd1'
Assert-VerifiedFile -File $pesterManifest -Sha256 $specification.pester.manifest_sha256

if ([string]::IsNullOrWhiteSpace($MagickPath)) {
    $imageArchive = Get-VerifiedArchive -File (Join-Path $scratchRoot ('tools/ImageMagick-' + $imageSpecification.version + '-Q16-x64/' + $imageSpecification.asset)) -Url $imageSpecification.url -Sha256 $imageSpecification.sha256 -Bytes $imageSpecification.bytes
    # Use Windows' libarchive tar (7z-capable), not a competing Git GNU tar on PATH.
    $tarPath = Join-Path $env:SystemRoot 'System32\tar.exe'
    $tar = Get-Command $tarPath -CommandType Application -ErrorAction Stop | Select-Object -First 1
    $entries = @(& $tar.Source -tf $imageArchive)
    if ($LASTEXITCODE -ne 0 -or $entries.Count -eq 0) { throw 'Could not inspect the verified ImageMagick archive.' }
    foreach ($entry in $entries) {
        # This reviewed portable archive has only ordinary files at its root.
        if ([string]::IsNullOrWhiteSpace($entry) -or $entry -match '[/\\:]' -or $entry -eq '..' -or $entry -eq '.') {
            throw 'ImageMagick archive layout differs from the reviewed portable package.'
        }
    }
    $imageDirectory = Join-Path $extractRoot 'ImageMagick'
    [IO.Directory]::CreateDirectory($imageDirectory) | Out-Null
    & $tar.Source -xf $imageArchive -C $imageDirectory
    if ($LASTEXITCODE -ne 0) { throw 'ImageMagick extraction failed.' }
    $MagickPath = Join-Path $imageDirectory 'magick.exe'
}
$MagickPath = (Get-Item -LiteralPath $MagickPath).FullName
Assert-NoReparseDirectory (Split-Path -Parent $MagickPath)
Assert-VerifiedFile -File $MagickPath -Sha256 $imageSpecification.executable_sha256

[pscustomobject]@{
    PesterManifest = $pesterManifest
    PesterVersion = $specification.pester.version
    PesterPackageSha256 = $specification.pester.sha256
    MagickPath = $MagickPath
    ImageMagickVersion = $imageSpecification.version
    ImageMagickExecutableSha256 = $imageSpecification.executable_sha256
}
