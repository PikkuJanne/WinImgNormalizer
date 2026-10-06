# Behavior and limits

This supplements the [quick start](../README.md). It describes the existing local
PowerShell application, including behavior that matters when interpreting a run.
The current version is unreleased and unsigned; these are preparation documents.

## Paths and output names

The default output parent is the Windows Pictures known folder, with
`%USERPROFILE%\Pictures` as a fallback only when Windows supplies no location. Each
run creates `<source>_WinImgNormalized_<yyyyMMdd_HHmmss>_<32-character runId>`
exclusively. The public interface accepts a source folder and optional byte target;
internal output/dependency parameters are test seams, rather than user options.

Canonical directory paths and case-insensitive whole path segments determine
containment: `PhotosBackup` is not inside `Photos`. An existing destination ancestor
that aliases the source is rejected through read-only 128-bit directory identities.
If the provider cannot supply the identity, setup fails. Linked roots/ancestors are
rejected; discovered links/junctions are skipped and reported. Unknown files inside
unreadable subtrees are not invented or included in file totals. Such scan omissions
make the run partial, including when no eligible media was readable.

Names are planned for the entire discovered tree in ordinal source-path order.
Unique images keep their stem with `.jpeg`. Same-stem images use extension suffixes:
`photo.jpg` and `photo.png` become `photo__jpg.jpeg` and `photo__png.jpeg`. Existing
unique image names, video names, directories and generated reports/work are
reserved first. A conflicting `photo__png.jpeg` forces the PNG derivative to
`photo__png__2.jpeg`, then the next free suffix as necessary. Comparisons ignore
case. Every eligible source has a `PLAN` entry, even when it later fails or is skipped.

Supported filename characters are treated literally, including spaces, apostrophes,
ampersands, parentheses, brackets, hashes, at-signs, leading hyphens, percent signs,
exclamation marks and Unicode. Neutral private image copies and literal native
output handling prevent ImageMagick filename expressions from changing destinations.
CMD can expand variable expressions before the BAT receives them; use single-quoted
PowerShell paths for those names. Long local paths still depend on the actual host,
filesystem and tools. Live UNC/network-share support is outside scope.

## Image, colour and size policy

Inspection and conversion read the same private image snapshot, with source
length/modification-time checks. One image is deliberately selected: the first
displayed logical GIF/WebP frame including canvas placement, the first TIFF page,
or the decoder's HEIC/HEIF primary image/first sequence image. Only a selected
animation frame is coalesced, then orientation is applied. Logs record exposed
source count, one selected image, omitted count and unit. Counts cannot describe
container items hidden by a decoder. Animation and omitted pages are not preserved.

Source ICC characterization remains attached through the sRGB transform. The
target profile is embedded in the PS1, so no separately installed profile is
required. The requested rendering intent is relative colorimetric with black-point
compensation disabled; a profile supplying only a perceptual mapping may use the
delegate's fallback. This does not certify arbitrary ICC contents or printing
conditions. Untagged ordinary RGB is assumed sRGB; declared linear RGB/grayscale
uses the decoder's transfer semantics. Untagged CMYK and unsupported unprofiled
colour spaces fail instead of guessing. Malformed/incompatible profiles and native
transform diagnostics also fail the item, even with zero native exit and JPEG bytes.

After orientation and colour conversion, alpha is composited onto white in encoded
sRGB channels; fully transparent pixels become white. GPS/EXIF, XMP, IPTC, comments
and source/target ICC profiles are then stripped. The JPEG omits ICC after conversion
to sRGB. Logs record source space, ICC presence and the selected policy. Images remain
lossy derivatives, and source image bytes/embedded metadata remain untouched.

The byte target defaults to 1,048,576 bytes (1 MiB). A second argument is bytes;
ImageMagick receives `jpeg:extent=<decimal bytes>B`, not decimal KB/MB suffixes.
Attempts retain the same colour/alpha policy at 100/90/80/70/60/50% dimensions,
without upscaling, with 4:2:0 sampling and progressive JPEG. Encoder quality search
and resizing can reduce detail. Very large Int64 targets may round inside the
encoder's floating-point search; final file-length comparison remains exact.

Each candidate must have a successful native result, nonempty bytes, one JPEG
frame, positive dimensions and a full pixel decode before a no-replacement move.
Invalid/truncated/stale or unexpected numbered outputs cannot become final results.
Only a valid above-target result advances to the next scale; the valid final 50%
candidate can be retained as `ConvertedWithWarning`/`WARN IMG`. `SizeWarnings` is a
subset of converted images and returns code 2. Failed final attempts do not revive
an earlier above-target candidate.

## Videos and duplicate skips

Videos are byte copies, with no re-encoding or metadata stripping. Copies use owned
partial files; finished length and source length/modification time must match before
finalization. Runtime content hashing is not required. Both image and video output
creation/modified timestamps are restored separately; a failed timestamp setter
keeps valid media and reports a warning/code 2.

The duplicate key is lowercase filename plus LastWriteTimeUtc ticks, with separate
input byte lengths. Sources register only after validated finalization, in ordinal
source-path order. A failed early candidate cannot suppress a later usable source.
Every skip identifies retained source, output and warning status. A finalized image
whose source length/time changed is kept with a warning and excluded from matching.
Different content with the same key/length can still be skipped; this heuristic
does not establish content equality, preserve every source in the derivative tree,
or remove anything from the source. Keep originals and separate backups.

## Failures, resource use and cancellation

Preflight checks arguments, source/output safety and the native executable before
creating a run folder. A temporary exclusive destination write probe is removed.
The free-space estimate includes copied video bytes, per-image budgets, twice the
largest readable image input for scratch and 64 MiB reserve. Each image budget is
the smaller of its target and the larger of 1 MiB or four times its input size.
Missing-decoder images require no conversion space. Measured insufficient space
fails setup; unavailable measurements produce a warning. These checks cannot
guarantee subsequent writes.

A diagnosed sharing/lock violation can retry the same scale twice, with 100/200 ms
backoff and fresh scratch. Generic access denial, damaged input, missing codecs,
ICC/policy errors and exhausted resources do not trigger scale fallback. Nonzero
native exit does not finalize. Only the pinned JPEG writer's recognized zero-exit
lossless-to-lossy notice is accepted after full validation; it still warns/code 2.
Other conversion diagnostics fail. Inspection/validation remains strict.

Each image shares a 120-second budget across inspection, profile extraction,
attempts, backoffs and validation. Native pixel-cache ceilings are 512 MiB memory,
1 GiB map and 2 GiB disk, with at most two worker threads; stricter installed policy
remains effective. These limits are not a total decoder/delegate heap cap or an
exploit-proof sandbox. Synchronous filesystem operations can delay interruption.
Per-item stdout/stderr capture is bounded to 16,384 UTF-16 characters per stream;
overflow fails the item. Preflight queries have a 262,144-character bound and
15-second deadline. Bounded diagnostic details go to local logs.

Native children enter a private Windows job at creation. Timeout/return terminates
owned descendants, preserving unrelated processes; failure to establish ownership
fails start. Private per-image native caches use the existing Windows temporary
root, with TEMP/TMP/MAGICK_TEMPORARY_PATH set for the child. A linked, too-long,
unavailable or source-contained temporary root fails setup. Cleanup removes exact
owned partials and regular recognized cache files only; unknown neighbors remain
with a warning. Unconfirmed process-tree termination preserves that item's scratch.
Concurrent hostile filesystem mutation is outside the isolation guarantee.

Handled Ctrl+C stops new work, terminates owned native work, retains completed
outputs and returns 130. A validated final move admitted before the request may
finish. Terminal closure, pipeline stopping or forced termination can bypass reports,
cleanup and the application/launcher's cancellation outcome. Remaining work files
are incomplete staging data. Neither a derivative folder nor an interrupted report
is a replacement backup.

## Reports and privacy

The generated directory is `.WinImgNormalizer`, or the first free suffix when the
source uses that name. `reports` contains the log and UTF-8 BOM/CRLF CSV; `work`
contains staged media. A discovered regular file has one outcome: Converted,
CopiedVideo, SkippedDuplicate, Ignored, Error, Cancelled or NotStarted. Warnings are
attributes, not extra file counts. Separate ScanIssue rows describe unreadable
locations and skipped links. The ACCOUNTING line reconciles known files, elapsed
time and finalized files/second. Image savings pair processed-image input bytes
with validated JPEG bytes and can be negative; video bytes are separate. Skips,
ignored files and errors contribute no image savings.

CSV text begins with `text:`. Backslashes are doubled; UTF-16 control, format,
line/paragraph separator and surrogate units use `\uNNNN` escapes. To recover an
exact path programmatically, remove exactly one prefix and decode left to right:
`\\` becomes one backslash; `\u` plus four hexadecimal digits becomes that UTF-16
unit. Decode for filesystem/text processing only; retain the prefix in a spreadsheet.
Numeric fields are generated invariant numbers, with blanks for unknown values.

`Import-Csv -LiteralPath 'C:\Example\report.csv' -Encoding UTF8` (replace the illustrative
path) reads string-valued fields in
both shells. For Excel, import UTF-8/comma through Data > From Text/CSV, disable
automatic type detection and set columns to Text before loading. This is guidance;
PowerShell tests do not certify Excel editing/save/reopen behavior.

Required log/CSV failure retains completed media but returns warning code 2 and
marks reporting incomplete. The console log fallback retains at most 8,192 UTF-16
characters, limits lines to 1,024 and reports dropped/truncated lines. A cancelled
run makes a best-effort report attempt and retains code 130. An incomplete report
must not be treated as a complete inventory.

The application does not upload files or add telemetry. Cloud-sync software can
upload output from a managed Pictures folder. Copied videos retain embedded private
metadata; logs/CSV can reveal source/output paths and diagnostics. Inspect and
sanitize support material before sharing. Metadata stripping cannot remove visible
identifying image content. See the [security/reporting policy](../SECURITY.md).
