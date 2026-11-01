# WinImgNormalizer

Prepare a smaller viewing copy of a mixed photo/video folder on Windows. Drag one
folder onto `WinImgNormalizer.bat`; the tool mirrors its subdirectories into a new
folder under Pictures, makes JPEG derivatives and copies supported videos unchanged.
Your source files remain untouched. JPEGs are lossy derivatives: keep the originals
and a separate backup.

**Version 1.0.3 is published and unsigned.**
[Download the portable ZIP](https://github.com/PikkuJanne/WinImgNormalizer/releases/download/v1.0.3/WinImgNormalizer-1.0.3-portable.zip)
or view the [stable release, checksums and provenance](https://github.com/PikkuJanne/WinImgNormalizer/releases/tag/v1.0.3).
The ZIP is 33681996 bytes; SHA-256:
`be3b6d0675eb1b65cbb82e770a12a0dabfd0f9de34a4c5d02f1fb1b4cb4a13da`.
Source and tag commit: `3c266fa6e11285ba74506ad92564bd89a7368532`.
Version 1.0.3 repairs JPEG colour-profile handling, retaining valid wide-gamut
conversion and clearly identified approximate previews when ICC metadata is unusable.
The offline clone/source-ZIP/cache distribution paths fixed in 1.0.2 remain unchanged.
Earlier releases retain their original tags, notes and asset bytes.
No separate ImageMagick installation required. The pinned
upstream engine is ImageMagick 7.1.2-35 Beta Q16-HDRI with patched HEIF support.
Build records retain their preparation-time state; these exact bytes were tested
and promoted unchanged. See [release metadata](release-metadata.json), the
[changelog](CHANGELOG.md) and [package guide](docs/release/GETTING_STARTED.md).
## Requirements

- 64-bit Windows 10 or 11, with Windows PowerShell 5.1 or PowerShell 7.
- The complete official portable ZIP, a normal Git clone, or the tagged GitHub
  source ZIP. No separate ImageMagick installation required for these forms.
- Ordinary read access to the source and write access to Pictures and temporary storage.

The portable package uses one verified app-relative engine for all probes and image
processing, independently of an older or absent system ImageMagick. It does not install
software, download anything while processing, or change PATH, registry or security policy.
Empty and video-only folders do not initialize an image engine. Videos keep their bytes.
A missing optional decoder is a per-file partial result; JPEG/PNG and video siblings continue.

A fresh source clone or source ZIP carries one compressed, digest-pinned private
runtime in `resources`. Its cold first image run verifies those local bytes and
atomically extracts into a versioned per-user cache under Local AppData. No network,
administrator rights or developer staging is needed. Every run authenticates the
selected runtime and pins its absolute executable and compatible configuration for
the entire run. The normal launcher never selects a system ImageMagick.

The full portable ZIP prefers its verified app-relative runtime and omits the
redundant compressed payload. Damaged runtime or payload bytes fail closed. A bare
BAT/PS1 copy can reuse only an authenticated cache for the same version; otherwise
it identifies the incomplete app directory and exact portable asset before any
external decoder runs. Empty/video-only folders still need no engine.

Automated coverage records the actual OS, shell and engine in the hotfix evidence.
Windows 10 has not been executed; Windows 11 and maintained Windows Server CI are distinct
platforms. Compatibility with Windows PowerShell 5.1 remains required.

## Quick start

1. Download the official portable ZIP and checksum sidecars, then follow
   [verification and extraction](docs/release/GETTING_STARTED.md). Keep the complete
   extracted folder together, including `runtime` and `portable-runtime.json`.
2. Drag exactly one source folder onto `WinImgNormalizer.bat`, or onto its shortcut.
   The launcher uses Windows PowerShell 5.1. An interactive console pauses before
   closing; redirected automation exits without a key read or prompt.
3. Read the result and full Output, Report and Log paths. Completed files remain
   available when another item fails. Open the new folder under Pictures.

Pictures means the Windows known folder, which may be redirected or cloud-synced;
it is not always `%USERPROFILE%\Pictures`. The application processes files on your
computer and does not upload them. Cloud-sync software can synchronize the output.

For the two direct positional forms, open PowerShell in the tool's folder and
replace the illustrative source path with your own existing directory:

```powershell
# Default target: 1,048,576 bytes (1 MiB) per JPEG
.\WinImgNormalizer.ps1 'D:\Photos\2024'

# Custom target: 2,097,152 bytes (2 MiB) per JPEG
.\WinImgNormalizer.ps1 'D:\Photos\2024' 2097152
```

Pass one folder and optionally a positive whole-number byte target from 1 through
9223372036854775807. Named options and extra positional arguments are not supported.
The BAT accepts one folder only; use the PS1 form for a custom target. Single quotes
in PowerShell preserve literal `%` and `!` characters; CMD may expand variables
before the BAT receives a path. Relative directory paths are accepted; drive-relative
paths such as `C:` are rejected. No configuration file or runtime account is needed.
If organizational policy blocks execution, follow your organization's process;
selective unblocking does not make an unsigned script satisfy `AllSigned`.

## Source and output safety

The destination must be outside the source tree. Pictures itself, its parent, and
a drive root containing Pictures are rejected before scanning or destination
writes. A source **subfolder of Pictures is allowed** when Pictures is its output
parent: the output then sits beside the source, rather than inside it. The tool
does not silently choose another output location.

Linked source/output roots and ancestors are rejected. Links and junctions found
during traversal are skipped and logged. An inaccessible subtree produces an
incomplete scan while readable siblings continue. Device paths and directory
components ending in a dot or space are unsupported. Local long paths depend on
host, filesystem and installed tools; live UNC/network-share support is outside
the owner-approved scope. Checks cannot guarantee isolation from concurrent hostile
filesystem changes.

Each run creates a new folder with a random 32-character run ID. Final moves refuse
replacement if another file or directory appears at a planned target. Source files
are never deleted, edited or replaced; completed results remain after a partial run.

## What the copy contains

| Source | Output |
|---|---|
| JPG/JPEG, PNG, BMP, TIF/TIFF, GIF, HEIC/HEIF, WebP | One `.jpeg` derivative per successfully retained image, subject to decoder support and heuristic skips. |
| MP4, MOV, MKV, AVI, M4V, WMV, WebM, MTS, M2TS, 3GP, 3G2 | An unchanged byte copy, subject to heuristic skips. |
| Other extensions | Counted as ignored; no output file. |

Images are auto-oriented, converted to sRGB using their source characterization,
composited onto white, then stripped of GPS/EXIF, XMP, IPTC, comments and ICC
profiles. Untagged ordinary RGB is assumed to be sRGB. Valid or safely recoverable
JPEG ICC profiles are applied before conversion, including split profiles. If an
ordinary RGB-family JPEG has unusable ICC metadata, a clean decode with that
profile excluded can produce a viewing copy with a trustworthy alternative
characterization, or with sRGB assumed. The latter reports: “JPEG saved; embedded
colour profile unusable; sRGB assumed, colours may differ.” These previews count
as converted images with warnings and return code 2; they are approximate colour.
Damaged pixel data, policy/resource failures, uncharacterized CMYK/YCCK and
unsupported colour models still fail with an explanation. Original images
keep their embedded metadata. Copied videos **retain their original metadata**,
including any location or identifying information. Reports can expose private
paths and diagnostic details. Metadata stripping is not a guarantee of anonymity.

GIF/WebP retains the first displayed logical frame; TIFF retains its first page;
HEIC/HEIF retains the decoder's primary image or first sequence image. Animation,
other pages and other exposed images are omitted and reported. Counts describe
what the installed decoder exposes, not every hidden item in a container. These
deliberate omissions alone do not turn success into a warning.

Duplicate skipping is a **heuristic**, using lowercase filename, LastWriteTimeUtc
ticks and equal input byte length. A later match is skipped only after an earlier
output has been validated and finalized; the log/CSV identifies that retained
source and output. Different content can share the same key and length and be
skipped. No runtime content hash database is used. This is another reason to keep
the originals and a separate backup.

## Byte target and quality

The default is **1,048,576 bytes = 1 MiB**. Decimal 1 MB is 1,000,000 bytes; the
second argument always means bytes. ImageMagick receives an explicit byte operand,
such as `jpeg:extent=1048576B`. The actual decoded JPEG's file length decides
whether it meets the target.

The encoder searches quality at 100% of the oriented image dimensions, then may
resize to 90%, 80%, 70%, 60% and 50%. Images are not enlarged. If the valid 50% result
still exceeds the target, it is saved as `WARN IMG`, with exact bytes, dimensions
and scale, and the application returns warning code 2. Native/decode failures fail
that item; they do not rescue an earlier above-target candidate. Tiny targets may
be impossible. This is a best-effort size target, rather than a guaranteed cap.

## Reports, outcomes and stopping

Mirrored media sits beside a generated `.WinImgNormalizer` directory (or the first
free suffix when that name is already used by the source). Its `reports` subfolder
holds `WinImgNormalizer_<timestamp>.log` and `.csv`; `work` holds temporary staging
data. Same-stem images receive distinct names, such as `photo__jpg.jpeg` and
`photo__png.jpeg`, with further numeric suffixes when required.

The log shows source-to-output plans, per-file results, omissions, duplicates and
the final accounting. The CSV has one outcome per discovered regular file and
separate scan issues. Text cells have a visible `text:` prefix and escaped controls
to prevent spreadsheet formulas; keep the prefix when viewing in a spreadsheet.
See [behavior details](docs/BEHAVIOR.md) for decoding rules and report limitations.
If a required report fails, completed media remains and code 2 warns that the
inventory is incomplete. Review reports locally and sanitize them before sharing.

| Exit code | Meaning |
|---|---|
| 0 | Completed successfully, including an empty or ignored-only folder. |
| 1 | Usage, setup or run-level failure prevented reliable completion. |
| 2 | Warnings or partial failures; inspect the report and retained results. |
| 130 | The application handled cooperative cancellation; completed outputs remain. |

The BAT displays and returns the child's code after its pause. An unexpected host
exit retains its actual code. Ctrl+C asks the application to stop new work and
terminate owned native processes. A blocked filesystem operation can delay stopping;
closing or force-killing the terminal may bypass cleanup, reports and exit 130.
Leftover work files are incomplete staging data, rather than final JPEGs.

## Troubleshooting

Setup failures begin with `Setup error:` and may occur before a run folder/log
exists. Per-item errors appear in an existing run's log; readable siblings continue.

| Message or outcome | Next step |
|---|---|
| `Usage: WinImgNormalizer.ps1 <sourceFolder> [maxBytes] (one folder only)` | Pass exactly one directory and optionally the byte target. |
| `maxBytes must be a positive whole number from 1 to 9223372036854775807.` | Use digits for a positive integer, such as `2097152`, rather than `2MiB`, zero or a fraction. |
| `The destination is inside or equal to the source, including a directory alias.` | Choose a source subfolder whose tree does not contain Pictures. |
| `ImageMagick magick.exe was not found.` | In a fresh source clone, stage the verified private runtime or select a trusted supported engine on PATH. In a portable ZIP, re-extract the complete official archive. |
| `ImageMagick 7.1.2-35 or newer supported 7.x build is required` | Update the dependency from its official source. |
| `This ImageMagick build has no JPEG encoder` | Choose a build with JPEG write support. |
| `Missing HEIC decoder` (or another named decoder) | Check that format's read support and installed policy; the file was not converted. |
| `WARN IMG` / `Could not reach target; best-effort saved` | Inspect the retained JPEG; choose a larger byte target if needed. |
| `JPEG saved; embedded colour profile unusable; sRGB assumed, colours may differ.` | The viewing copy is usable but its colour is approximate. Keep the original for accurate colour work. |
| Colour/profile, policy, timeout or decode error | Review the item reason and bounded native details. Obtain a correctly characterized/decodable source; retain installed protections. |

Message fragments above identify the relevant error; full messages can include
paths or additional detail. A successful free-space probe/estimate cannot guarantee
later writes. Do not run as administrator, weaken global ImageMagick policy or
disable protection software to bypass a failure.

## More information

[Detailed behavior and limits](docs/BEHAVIOR.md), [package getting started guide](docs/release/GETTING_STARTED.md),
[security and reporting policy](SECURITY.md), [release notes](CHANGELOG.md),
[packaging/provenance](docs/release/PACKAGING.md), and [development checks](tests/README.md).

WinImgNormalizer is a personal tool by Janne Vuorela, provided under the
[MIT license](LICENSE) without warranty. The embedded sRGB profile is CC0;
see [third-party notices](docs/release/THIRD_PARTY_NOTICES.md).

The source-only missing-engine error is: `ImageMagick magick.exe was not found.`
Stage the developer runtime with `tools/release/Stage-Runtime.ps1 -Download`.
