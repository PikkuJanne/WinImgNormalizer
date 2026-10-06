# WinImgNormalizer

Prepare a smaller viewing copy of a mixed photo/video folder on Windows. Drag one
folder onto `WinImgNormalizer.bat`; the tool mirrors its subdirectories into a new
folder under Pictures, makes JPEG derivatives and copies supported videos unchanged.
Your source files remain untouched. JPEGs are lossy derivatives: keep the originals
and a separate backup.

**Version 1.0.0 is unreleased and unsigned.** A portable ZIP has been prepared and
tested locally; no public release or download is established. See the
[changelog](CHANGELOG.md) and [draft release metadata](release-metadata.json).

## Requirements

- 64-bit Windows 10 or 11, with Windows PowerShell 5.1 or PowerShell 7.
- ImageMagick **7.1.2-32 or a newer supported 7.x build**, separately obtained from
  its [official Windows downloads](https://imagemagick.org/download/#windows-binary-release).
  Make the chosen `magick.exe` available through PATH in the launching session.
- Ordinary read access to the source and write access to the Windows Pictures
  location and temporary storage; enough free space for copies and scratch files.

ImageMagick is required even for video-only or empty folders. Image conversion
requires JPEG write support and a decoder for the source format. Codec availability and installed
security policy can prevent individual files from working; a recognized extension
does not guarantee a usable decoder. The tool does not install or update software.
Keep ImageMagick and its delegates updated and retain its security policy.

Desktop checks ran on Windows 11 Pro build 26300 with Windows PowerShell
5.1.26100.9444 and PowerShell 7.6.5. Hosted checks ran on Windows Server 2025
Datacenter 10.0.26100 with PS 5.1.26100.33438 and PS 7.6.6. The maintained full gates
use ImageMagick 7.1.2-32 portable Q16 x64; the same version's Q16-HDRI x64 build has
separate colour and owner workflow checks. Windows 10 has not been tested. These
results establish the recorded builds, rather than every newer codec build.

## Quick start

1. Review the source. Keep `WinImgNormalizer.ps1` and `WinImgNormalizer.bat` together
   with their original names in an ordinary folder, such as Downloads. For a
   trusted prepared ZIP, follow [verification and extraction](docs/release/GETTING_STARTED.md).
2. In PowerShell, check the selected ImageMagick executable:

   ```powershell
   Get-Command magick.exe -CommandType Application
   magick.exe -version
   magick.exe -list format
   ```

3. Drag exactly one source folder onto `WinImgNormalizer.bat`. The launcher uses
   Windows PowerShell 5.1, displays the outcome and pauses before closing. Use an
   ordinary account. It supplies `-ExecutionPolicy Bypass` only to that child
   process; it does not change persistent policy or override Group Policy.
4. Open the new `<source>_WinImgNormalized_<yyyyMMdd_HHmmss>_<runId>` folder in your
   Windows Pictures location. Review its log and CSV before relying on the copy.

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
profiles. Untagged ordinary RGB is assumed to be sRGB; untagged CMYK, malformed
profiles and unsupported colour models fail with an explanation. Original images
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
| `ImageMagick magick.exe was not found.` | Make the selected executable available on PATH; run the three prerequisite checks above. |
| `ImageMagick 7.1.2-32 or newer supported 7.x build is required` | Update the dependency from its official source. |
| `This ImageMagick build has no JPEG encoder` | Choose a build with JPEG write support. |
| `Missing HEIC decoder` (or another named decoder) | Check that format's read support and installed policy; the file was not converted. |
| `WARN IMG` / `Could not reach target; best-effort saved` | Inspect the retained JPEG; choose a larger byte target if needed. |
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
