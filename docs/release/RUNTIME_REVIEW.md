# Runtime selection review — 31 October 2026

The selected upstream portable Windows x64 engine is ImageMagick 7.1.2-35
**Beta**, Q16-HDRI, from the successful official Release workflow
[41175952230](https://github.com/ImageMagick/ImageMagick/actions/runs/41175952230),
source `39ad88fc5637b0ed9d2adbf37c3c5f6f174c63a3`. Its exact payload identity,
independent artifact-API digest, signed publisher identity and file allowlist
are in [runtime-lock.json](../../tools/release/runtime-lock.json). This is a
packaging reproducibility pin, not an equality restriction on PATH executables.

The current tagged stable ImageMagick release is 7.1.2-34, published 14 October.
It fixes the newer core TIFF decoder stack overflow
[GHSA-g5j8-ff59-wp88](https://github.com/ImageMagick/ImageMagick/security/advisories/GHSA-g5j8-ff59-wp88).
Core policy/delegate, XMP and bzip2 issues disclosed on 10 October were also
reviewed through the [upstream advisory API](https://api.github.com/repos/ImageMagick/ImageMagick/security-advisories).
The prior JPEG/XMP/jpeg:extent security floor remains historical evidence of a
real security requirement; a feature probe is not a substitute for fixes.

However, the tagged -34 portable SBOM contains libheif 1.23.6. Upstream
[libheif 1.23.7](https://github.com/strukturag/libheif/releases/tag/v1.23.7),
released 26 October, fixes HEIF decoding issues affecting 1.23.6, including
[parallel derived-grid use-after-free](https://github.com/strukturag/libheif/security/advisories/GHSA-35mc-3887-rr37),
[AVC out-of-bounds reads](https://github.com/strukturag/libheif/security/advisories/GHSA-xwcv-cp39-m2w4)
and container/sequence resource exhaustion. The selected -35 upstream build
contains the corrected delegate. Shipping the tagged -34 engine would leave
those relevant fixes out of the advertised HEIC/HEIF input path.

The unchanged signed executable and upstream CycloneDX SBOM were inspected. The
pipeline-relevant inventory is:

| Component | Shipped version | Review basis |
| --- | --- | --- |
| ImageMagick | 7.1.2-35 Beta | Core fixes through -34, plus official current build |
| libheif | 1.23.7 | 26 October security release linked above |
| libde265 | 1.1.3 | [luma-plane disclosure fixed in 1.1.3](https://github.com/strukturag/libde265/security/advisories/GHSA-vm5w-mrw5-q329) |
| libpng | 1.6.59 | [compressed ancillary-chunk use-after-free fixed in 1.6.59](https://github.com/pnggroup/libpng/security/advisories/GHSA-qvg3-h654-xq3j) |
| libjpeg-turbo | 3.2.1 | [27 October fixes and hardening](https://github.com/libjpeg-turbo/libjpeg-turbo/releases/tag/3.2.1) |
| libtiff | 4.7.2 | Exact upstream SBOM and TIFF capabilities; IM core TIFF fix above |
| libwebp | 1.6.0 | [upstream version and change history](https://chromium.googlesource.com/webm/libwebp/+/refs/tags/v1.6.0/NEWS) |
| Little CMS | 2.19.1 | [current stable upstream maintenance release](https://github.com/mm2/Little-CMS/releases/tag/lcms2.19.1) |

The application still invokes bounded single native operations, retains native
diagnostics and policy-denial handling, and never tries a different engine after
denial. No crafted exploit was executed in this review. Actual fixture/codec,
policy, original-preservation and package results are recorded by automated tests;
the SBOM and a version banner alone do not demonstrate those results. This review
does not establish that a build is universally safe or that future advisories
cannot affect it.

The portable EXE is static and brings no side-by-side ImageMagick DLLs. Unused
command EXEs are excluded. All selected upstream config, ICC, LICENSE and NOTICE
files remain byte-identical. Complete upstream Windows/delegate source/build
bytes are distributed alongside the release for source and relink rights; see
[source information](IMAGEMAGICK_SOURCE.md). Original upstream portable/SBOM assets
are mirrored unchanged for reproducibility after Actions retention. Before publication, CI obtains the exact authenticated upstream workflow artifacts
with its read-only build credential. A maintainer can also use the same release in
draft state; the CI read token cannot see drafts. After publication the pinned
public mirror supplies the identical bytes. Media processing has no
network path and receives no publication credential.
