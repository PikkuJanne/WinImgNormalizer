# Website product copy

Prepared visitor copy for a future product page. Apply the metadata and asset rules
in [the integration handoff](INTEGRATION.md). The recorded public release is
available; website integration and deployment remain separate. Application version and release fields
come from [canonical release metadata](../../release-metadata.json).

## Title and introduction

**WinImgNormalizer**

Make a local viewing copy of a photo and video folder on Windows.

Drag one folder onto `WinImgNormalizer.bat`. The tool mirrors its subdirectories
into a new folder under Windows Pictures, creates JPEG derivatives and copies
supported videos unchanged. Your source files remain untouched. JPEGs are lossy
derivatives: keep the originals and a separate backup.

## Current availability

WinImgNormalizer 1.0.0 is published and unsigned.

**[Download WinImgNormalizer 1.0.0](https://github.com/PikkuJanne/WinImgNormalizer/releases/download/v1.0.0/WinImgNormalizer-1.0.0-portable.zip)**

The [release page](https://github.com/PikkuJanne/WinImgNormalizer/releases/tag/v1.0.0)
includes provenance and checksum sidecars. The ZIP is 184659 bytes, with SHA-256
`251828028e144759c919645f423fde08641cabf42e04db7024d7ad17da4ba14d`.
Its embedded unreleased wording records the preparation snapshot; the approved
files were published unchanged. Use **View source** and **Read the quick start**
for source and setup guidance.

## How it works

1. Keep `WinImgNormalizer.ps1` and `WinImgNormalizer.bat` together with their
   original names. Install ImageMagick separately and make the selected
   `magick.exe` available on PATH. Follow the [quick start](../../README.md).
2. Drag exactly one source folder onto the BAT. It uses Windows PowerShell 5.1,
   displays the outcome and pauses before closing. The PS1 also supports direct
   positional invocation in Windows PowerShell 5.1 or PowerShell 7.
3. Open the new run folder under your Windows Pictures location. Subdirectories
   are mirrored; completed outputs remain if another item fails.
4. Review the local log and CSV, including warnings, omissions and skipped files,
   before relying on the viewing copy. Keep the source and a separate backup.

The default target is **1,048,576 bytes (1 MiB) per JPEG**. Decimal 1 MB means
1,000,000 bytes. A custom target uses the PS1's optional second positional argument
as a positive whole number of bytes; for example, `2097152` means 2 MiB. The BAT
accepts one folder only. See the tested direct examples in the quick start.

This is a best-effort size target. Quality search and resizing can lose detail;
images are not enlarged. A valid final JPEG that still exceeds the target is kept
with a warning. The target is not a guaranteed file-size cap.

## Requirements and tested scope

- 64-bit Windows 10 or 11 with Windows PowerShell 5.1 or PowerShell 7. Windows 10
  testing is out of project scope and has not been run; recorded desktop checks
  used Windows 11 and hosted checks
  used Windows Server 2025.
- ImageMagick 7.1.2-32 or a newer supported 7.x build, obtained separately from its
  [official Windows downloads](https://imagemagick.org/download/#windows-binary-release).
  ImageMagick is required even for an empty or video-only source. Image conversion
  needs JPEG write support and the source format's decoder; installed codecs and
  security policy can prevent individual files from working.
- Ordinary read access to the source and write access to Pictures and temporary
  storage, with sufficient free space for copies and scratch files.

The [quick start](../../README.md) records exact tested PowerShell/ImageMagick
builds. Keep the dependency updated and retain its installed security policy. The
application does not install software. Follow organizational execution policy and
the [package verification guide](../release/GETTING_STARTED.md); a trusted checksum
identifies bytes, without establishing a digital signature or publisher identity.

## What is retained and what is omitted

| Source | Viewing copy |
|---|---|
| JPG/JPEG, PNG, BMP, TIF/TIFF, GIF, HEIC/HEIF and WebP | One JPEG derivative per successfully finalized image, subject to decoder support and heuristic skips. |
| MP4, MOV, MKV, AVI, M4V, WMV, WebM, MTS, M2TS, 3GP and 3G2 | An unchanged byte copy, subject to heuristic skips. |
| Other extensions | Counted as ignored; no output file. |

Images are auto-oriented and converted to sRGB, with transparency composited onto
white. GPS/EXIF, XMP, IPTC, comments and ICC profiles are stripped after orientation
and colour handling. Original image bytes and metadata remain untouched. Untagged
ordinary RGB is assumed sRGB; untagged CMYK, malformed profiles and unsupported
colour models fail with an explanation.

GIF/WebP keeps the first displayed logical frame, TIFF the first page, and
HEIC/HEIF the decoder's primary image or first sequence image. Other exposed
frames, pages and images are omitted and reported. The counts describe what the
decoder exposes. The viewing copy does not retain animation or all container
members.

Duplicate skipping is a one-run heuristic based on lowercase filename, UTC
modification-time ticks and input byte length. A later match can be skipped after
an earlier output is validated and finalized. Different content can share those
values; this is not content verification. The log and CSV identify the retained
source and output. No source is deleted by duplicate skipping.

## Local processing and privacy

The application processes files on your computer and does not upload them, add
telemetry or require an account. Windows Pictures may be redirected or managed by
cloud-sync software; that software can synchronize the output.

Copied videos retain their original metadata, including location or identifying
information. Logs and CSV reports can expose private paths, filenames and native
diagnostics. Review and sanitize support material before sharing. Stripping image
metadata does not guarantee anonymity or remove visible identifying content.

## Safe source selection and outcomes

The destination must be outside the source tree. Pictures itself, its parent or
a drive root containing Pictures is rejected before scanning or destination
writes. A source subfolder of Pictures is allowed when the output sits beside the
source. Linked roots and ancestors are rejected; discovered links and junctions
are skipped and reported. Inaccessible subtrees make a run partial while readable
siblings continue. Live UNC/network-share support is outside scope.

Exit 0 means completion; exit 1 means usage, setup or run-level failure; exit 2
means warnings or partial failures that need inspection. Handled cooperative
cancellation returns 130 and retains completed outputs. Closing or force-killing
a terminal can bypass cleanup and reports. An incomplete report is not a complete
inventory.

## Further reading and attribution

[Quick start](../../README.md), [behavior and limits](../BEHAVIOR.md),
[getting started with a trusted prepared package](../release/GETTING_STARTED.md),
[release notes](../../CHANGELOG.md), and [security/reporting policy](../../SECURITY.md).

WinImgNormalizer is a personal tool by Janne Vuorela, provided under the
[MIT license](../../LICENSE) without warranty. The embedded sRGB profile is CC0;
see [third-party notices](../release/THIRD_PARTY_NOTICES.md).
