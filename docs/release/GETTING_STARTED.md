# Getting started with WinImgNormalizer

This portable package is **version 1.0.3, in release preparation and unsigned**.
`package-manifest.json` records its exact source revision; [CHANGELOG.md](CHANGELOG.md)
has release notes. Build records use `unreleased` for their preparation-time state;
the public GitHub release and matching checksum establish publication of these
unchanged bytes. The filename/path examples below are illustrative.

## Verify and extract

Obtain the ZIP and its expected SHA-256 through a source you trust. Compare the
actual hash with the exact filename's entry in `SHA256SUMS.txt` and the value
provided through that trusted channel. Replace this path with your prepared ZIP:

```powershell
Get-FileHash -LiteralPath 'C:\Downloads\WinImgNormalizer-1.0.3-portable.zip' -Algorithm SHA256
```

A matching checksum identifies bytes relative to a trusted value; it is not a
digital signature or publisher authentication. Review the source before running.
If you trust the verified ZIP and Windows marks it as downloaded, selectively
unblock that exact ZIP before extracting:

```powershell
Unblock-File -LiteralPath 'C:\Downloads\WinImgNormalizer-1.0.3-portable.zip'
```

Extract into a new ordinary folder with a name such as `WinImg Normalizer` in
Downloads. Keep `WinImgNormalizer.bat` and `WinImgNormalizer.ps1` together with their
original names. If the PS1 remains blocked, review and selectively unblock that
exact file with `Unblock-File -LiteralPath`. Do not unblock a whole Downloads tree.
Unblocking removes a downloaded-file marker; it does not change execution policy
or make an unsigned script satisfy `AllSigned`. See Microsoft's
[hashing documentation](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/get-filehash)
and [selective unblocking documentation](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/unblock-file).

## Private image engine

No separate ImageMagick installation required for the portable ZIP. Keep the entire
extracted package together: `runtime` contains the verified private engine and its
notices, and `portable-runtime.json` identifies an intact distribution. Every probe,
conversion and final JPEG decode uses that engine. A missing or corrupt runtime
fails before execution: re-extract the official ZIP into a new folder.

The application does not download, install or update dependencies during a run.
An older ImageMagick on PATH is not used by the standard ZIP. Empty and video-only
folders need no image engine. Videos are copied byte-for-byte and keep their metadata.
An unavailable optional decoder produces a named per-file warning/partial result;
other supported files still complete. Explicit local and administrator policy denials
remain effective; the application never tries another engine to evade a denial.

A fresh source clone or tagged source ZIP includes a compressed private runtime.
It verifies those local bytes and initializes a versioned per-user cache offline
on the first image run. No developer staging or system ImageMagick is required.
Keep the complete source copy together, including `resources`. A bare BAT/PS1 pair
uses only an authenticated cache for this version; without it the app identifies
the incomplete directory and names `WinImgNormalizer-1.0.3-portable.zip` before any
external decoder. Damaged recognized runtime files fail closed.
Windows PowerShell 5.1 and PowerShell 7 remain supported. Windows 10 has not been tested.

## Choose a source and run

Use an ordinary account with read access to the source and write access to Pictures
and temporary storage. Keep originals and a separate backup: JPEGs are lossy
derivatives, and this copy can omit pages, animations and heuristic duplicates.

Choose one existing source folder whose tree does not contain the Windows Pictures
location. Pictures itself, its parent, and a drive root containing it are rejected
before scanning or output writes. A **subfolder of Pictures is allowed** when its
output sits alongside it under Pictures. Linked roots/ancestors are rejected; links
found inside the tree are skipped and reported. Live UNC/network-share support is
outside scope. Local long paths still depend on the host, filesystem and tools.

Drag exactly one source folder onto `WinImgNormalizer.bat`. The launcher uses
Windows PowerShell 5.1 with `-ExecutionPolicy Bypass` for that child process only,
then displays the outcome and pauses in an interactive console. Redirected runs exit unattended. It does not change persistent policy or
override Group Policy. See Microsoft's [execution-policy documentation](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies).
Do not change machine-wide policy, disable protection software or weaken ImageMagick
policy. If organizational controls block execution, follow your organization's process.

For the two direct positional forms, open PowerShell in the extracted tool folder
and replace the illustrative directory below:

```powershell
# Default: 1,048,576 bytes (1 MiB) per JPEG
.\WinImgNormalizer.ps1 'D:\Photos\2024'

# Custom: 2,097,152 bytes (2 MiB) per JPEG
.\WinImgNormalizer.ps1 'D:\Photos\2024' 2097152
```

The second argument is a positive whole-number byte target from 1 through
9223372036854775807. Decimal 1 MB is 1,000,000 bytes. Tiny targets may be impossible:
a valid JPEG above the target is retained with `WARN IMG` and exit code 2.
Pass arguments positionally; the BAT accepts one folder only. PowerShell single
quotes preserve literal percent/exclamation characters that CMD can expand before
the BAT receives them. Ordinary relative directories work; drive-relative `C:`
does not.

## Review the results

Find the new `<source>_WinImgNormalized_<yyyyMMdd_HHmmss>_<runId>` folder under the
Windows Pictures known folder, which may be redirected or cloud-synced. Originals
remain unchanged. Mirrored subdirectories contain `.jpeg` images and unchanged video
copies. A generated `.WinImgNormalizer` directory (or a free suffix) contains
`reports` with the log/CSV and `work` with temporary staging data.

Supported image extensions are JPG/JPEG, PNG, BMP, TIF/TIFF, GIF, HEIC/HEIF and WebP,
subject to installed decoders. Supported video extensions are MP4, MOV, MKV, AVI,
M4V, WMV, WebM, MTS, M2TS, 3GP and 3G2. Other extensions are ignored. Image names may
gain extension/numeric suffixes to avoid collisions. Final moves refuse replacement.

Images are auto-oriented, converted to sRGB, composited onto white, then stripped
of GPS/EXIF, XMP, IPTC, comments and ICC profiles. Valid and safely recoverable JPEG
ICC profiles are applied before conversion. An otherwise sound ordinary RGB-family
JPEG with unusable embedded ICC can produce an approximate sRGB viewing copy with
the warning “JPEG saved; embedded colour profile unusable; sRGB assumed, colours
may differ.” It counts as converted with a warning and exit 2. Full detail stays
in the local log/CSV; recurring reasons receive one short console notice and final
counts. Uncharacterized CMYK/YCCK, damaged pixels and policy/resource failures
still fail. Original images retain metadata. Copied
videos **retain their original metadata**, including possible identifying information.

GIF/WebP becomes the first displayed logical frame; TIFF becomes its first page;
HEIC/HEIF uses the decoder's primary or first sequence image. Other exposed images
are omitted and reported; animation is not preserved. Duplicate skipping is a
heuristic: lowercase filename, LastWriteTimeUtc ticks and equal input byte length,
after a prior output has been successfully finalized. Different content can match
and be skipped. Inspect omissions and retained-source mappings in the reports.

| Exit code | Meaning |
|---|---|
| 0 | Completed successfully, including an empty or ignored-only folder. |
| 1 | Usage, setup or run-level failure. A setup error may precede any log. |
| 2 | Warnings/partial completion; review retained media and incomplete reports. |
| 130 | Handled cooperative cancellation; completed outputs remain. |

The BAT returns the child's code after its pause. Ctrl+C requests a cooperative
stop; blocked filesystem operations can delay it. Closing or force-killing the
terminal can bypass cleanup, reports and exit 130. Leftover work files are not final
outputs. A failed/incomplete log or CSV must not be treated as a complete inventory.

`Setup error:` identifies failed prerequisites or source/output safety. Check
`magick.exe` path/version/capabilities for dependency errors, use positive integer
bytes for `maxBytes` errors, and choose a source that does not contain Pictures
for nesting errors. Per-file codec, colour, policy, timeout and decode failures
appear in the run log; accessible siblings continue. Do not weaken protections
to make a failed input convert.

## Privacy and reporting

The application processes files on your computer and does not upload them.
Cloud-sync software can synchronize output from Pictures. Logs/CSV can expose
private paths and diagnostics; inspect and sanitize them before sharing. CSV text
cells retain a visible `text:` prefix to prevent spreadsheet formulas. Metadata
stripping cannot remove visible identifying image content or copied-video metadata.

Report ordinary bugs with a small synthetic reproduction through the
[repository's Issues page](https://github.com/PikkuJanne/WinImgNormalizer/issues).
For a possible vulnerability, use GitHub's private reporting option if the
repository offers it; otherwise request a private contact without posting exploit
details, private media or raw logs. No private reporting address or response
deadline is promised here. A checksum is not a signature.

WinImgNormalizer is authored by Janne Vuorela under [LICENSE](LICENSE).
The embedded colour profile's provenance/license is in
[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md) and [LICENSE-CC0.txt](LICENSE-CC0.txt).
