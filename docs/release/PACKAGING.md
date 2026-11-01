# Portable packaging

The standard ZIP includes a verified app-relative ImageMagick engine. No separate
ImageMagick installation is required for the portable ZIP. Application processing
never downloads dependencies or changes PATH, registry, execution policy or an
installed ImageMagick. Normal clones and tagged source ZIPs carry the compressed
`resources/private-runtime.zip` payload for an offline cold first run. The script
embeds its digest and the existing runtime-manifest digest. The expanded runtime
remains Git ignored; the portable ZIP omits the redundant compressed payload.

## Stage the reviewed runtime

`tools/release/runtime-lock.json` pins the exact official upstream portable build,
artifact ZIP digest, inner archive digest, signed executable, configuration files,
SBOM and corresponding-source archive. The chosen engine is ImageMagick 7.1.2-35
**Beta**, Q16-HDRI x64, source `39ad88fc5637b0ed9d2adbf37c3c5f6f174c63a3`, from
[official workflow 41175952230](https://github.com/ImageMagick/ImageMagick/actions/runs/41175952230).
The upstream Beta was selected because it includes libheif 1.23.7's current security
fixes; the latest tagged 7.1.2-34 Windows portable carries older libheif 1.23.6.
The WinImgNormalizer patch is a stable application release after automated gates.

Run at development/build time, using Windows PowerShell 5.1 or PowerShell 7:

```powershell
.\tools\release\Stage-Runtime.ps1 -Download
```

This stages `runtime/` and `portable-runtime.json` beside the existing launcher;
both are Git ignored. `-DestinationRoot <checkout>` stages the identical verified
payload into another maintained checkout. Existing differing runtime files are
refused. The default cache is `.scratch/runtime-downloads`; `-CacheDirectory`
supports an already authenticated shared cache. The public release mirrors exact
upstream portable/SBOM/source assets to preserve future reproducibility after
Actions retention. Before publication the developer-only fallback uses an existing
authenticated GitHub CLI to request the same repository's draft dependency assets.
Read-only CI tokens can receive a missing-release response for a draft; initial
acquisition then fetches the original public upstream
artifact ZIP, verifies the independent GitHub API digest, extracts its single
expected file, and verifies the inner hash. The executable publisher signature is also checked before staging.
No global installation or configuration change occurs.

PE inspection identifies Microsoft OpenMP as the engine's only required non-inbox
DLL. Staging additionally downloads the exact official Microsoft Visual C++ v14
Redistributable 14.51.36247.0 x64, checks its independent Microsoft WinGet digest
and Microsoft signature, and extracts `VCOMP140.DLL` from pinned cabinet bytes.
No installer or MSI is executed. The unchanged signed DLL and original runtime
license plus full plain-text extraction are staged beside `magick.exe`. Its only
PE dependency is Windows `KERNEL32.dll`; actual acceptance checks that the engine
loads the app-local copy, rather than an installed Visual Studio copy. The full
source/digest/license/signature record is `runtime-lock.json`'s `vc_runtime` field
and is preserved in the runtime manifest and outer build provenance.

## Build the final revision once

Review and commit all intended source first. Build from the final clean current
40-character revision, with an absent destination beneath ignored `.scratch`:

```powershell
.\tools\release\Stage-Runtime.ps1 -Download -IncludeDependencySource
$releaseRevision = (git rev-parse HEAD).Trim()
.\tools\release\Build-Release.ps1 -Revision $releaseRevision `
  -OutputDirectory '.scratch\release candidate' -IncludeDependencySource
```

The builder downloads nothing. It validates all staged files against the exact
tracked lock and source notice, rejects extras/reparse points/dirty source, and
uses exact Git blob bytes for application/documentation entries. It never commits
a binary distribution. Runtime-required marker and integrity manifest distinguish
the portable package from a source checkout; missing/damaged release runtime
fails without trying another executable.

The ZIP allowlist consists of the original seven tracked package files, the
runtime-required marker, generated runtime/package manifests, selected upstream
runtime files listed in the lock, upstream SBOM and source-information notice.
Unused ImageMagick command executables and changelog are omitted; all shipped
configuration, license, notices and ICC bytes remain unchanged. No tests, media,
logs, Git metadata, credentials, or developer dependencies enter the ZIP.

Output is `WinImgNormalizer-<version>-portable.zip`, `build-provenance.json`, and
`SHA256SUMS.txt`. With `-IncludeDependencySource`, the output also contains the exact
original upstream portable 7z, SBOM and complete Windows source/build 7z for release
publication. These satisfy source/relink access and preserve dependency identity;
they are optional downloads for recipients, not processing prerequisites.

The builder sorts entries, fixes ZIP timestamps and attributes and requests .NET
NoCompression. Repeated builds on the same writer are checked for equality;
.NET Framework and current .NET can produce different ZIP framing. Promote the
one actually tested build unchanged. Provenance binds application source, upstream
source, files, ZIP and source assets. Checksum files provide byte identity, not a
signature on WinImgNormalizer. The upstream engine's own signature is retained.

## Publication and verification

The scoped 1 November 2026 JPEG colour-repair instruction authorizes normal
merge/tag/stable release publication after the mandatory automated Windows gates,
without a separate owner-testing or discretionary publication gate. Publish the
tested bytes and sidecars plus corresponding-source assets. The same release can
first be a public prerelease candidate for read-only hosted exact-byte checks,
then be promoted to stable/latest without replacing any asset. Fresh-download
the public ZIP and checksum, verify digest/source/tag identity, extract to another
clean directory, and rerun real media smoke. Never replace a published artifact.
Canonical release records are updated after verification; no live website
deployment is part of this repair.

Version 1.0.1 was built once from source/tag commit
`25cd8da78595abd256d96b2ff7d117e2b3a02c18` and promoted without replacing its six
assets. The [publication record](publication-v1.0.1.json) binds fresh anonymous
downloads to the tested hashes; [identity](identity-v1.0.1.json) retains their
source and byte identities. Later repository publication metadata is separate
from that immutable release source. CI prefers an explicit candidate tag, then
downloads the canonical stable public tag for published metadata, and builds
only an unreleased checkout. It does not rebuild a published package.

Version 1.0.0 and its tag/assets remain unchanged. Its original external-engine
requirement and build-time unreleased metadata are historical; see
[the publication record](publication-v1.0.0.json).
