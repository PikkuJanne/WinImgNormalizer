# Development-only fixture acquisition. No installer execution or global changes.
[CmdletBinding()]
param([switch]$Download)
$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
$root = Join-Path $repository '.scratch/hotfix-dependencies'
$pin = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'hotfix-dependencies.json') -Raw | ConvertFrom-Json
function Get-HotfixVerifiedArchive([string]$Name, $Specification) {
    $path = Join-Path $root $Name
    if (-not [IO.File]::Exists($path)) {
        if (-not $Download) { throw ('Explicit fixture download required: ' + $Name) }
        [IO.Directory]::CreateDirectory($root) | Out-Null
        $partial = $path + '.' + [guid]::NewGuid().ToString('N') + '.partial'
        $protocol = [Net.ServicePointManager]::SecurityProtocol
        try {
            [Net.ServicePointManager]::SecurityProtocol = $protocol -bor [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -UseBasicParsing -Uri $Specification.url -OutFile $partial
        } finally { [Net.ServicePointManager]::SecurityProtocol = $protocol }
        if ((Get-Item -LiteralPath $partial).Length -ne $Specification.bytes -or
            (Get-FileHash -LiteralPath $partial -Algorithm SHA256).Hash -ine $Specification.sha256) { throw 'Hotfix fixture download failed authentication.' }
        [IO.File]::Move($partial, $path)
    }
    $item = Get-Item -LiteralPath $path
    if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint -or $item.Length -ne $Specification.bytes -or
        (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ine $Specification.sha256) { throw 'Hotfix fixture archive integrity failure.' }
    return $path
}
$ancestor = [IO.DirectoryInfo]::new([IO.Path]::GetFullPath($root))
while ($ancestor) {
    if ($ancestor.Exists -and ($ancestor.Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Fixture directory crosses a reparse point.' }
    $ancestor = $ancestor.Parent
}
$installer = Get-HotfixVerifiedArchive 'old-dll-installer.exe' $pin.legacy_installer
$extractorZip = Get-HotfixVerifiedArchive 'innoextract.zip' $pin.innoextract
$baseline = Get-HotfixVerifiedArchive 'baseline.zip' $pin.baseline
$signature = Get-AuthenticodeSignature -LiteralPath $installer
if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Subject -notmatch 'CN=ImageMagick Studio LLC') { throw 'Legacy fixture installer publisher signature is invalid.' }
$extract = Join-Path $root ([guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($extract) | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem
[IO.Compression.ZipFile]::ExtractToDirectory($extractorZip, (Join-Path $extract 'extractor'))
$extractor = Join-Path $extract 'extractor/innoextract.exe'
$legacy = Join-Path $extract 'legacy'
$capture = @(& $extractor --silent -d $legacy $installer 2>&1)
if ($LASTEXITCODE -ne 0) { throw ('Legacy fixture extraction failed: ' + ($capture -join ' ')) }
$magick = Join-Path $legacy 'app/magick.exe'
if ((Get-FileHash -LiteralPath $magick -Algorithm SHA256).Hash -ine $pin.legacy_installer.executable_sha256) { throw 'Extracted legacy executable differs from authenticated fixture.' }
[pscustomobject]@{ LegacyMagickPath = $magick; BaselineZipPath = $baseline; LegacyInstallerSha256 = $pin.legacy_installer.sha256; LegacyExecutableSha256 = $pin.legacy_installer.executable_sha256; LegacySignature = [string]$signature.Status }
