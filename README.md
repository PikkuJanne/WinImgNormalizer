# WinImgNormalizer — Non-destructive image normalizer for Windows (PowerShell + ImageMagick)
Minimal, no-frills normalizer I use to prep mixed photo/video folders for archiving. It mirrors the tree into a safe copy, converts images to JPEG targeting 1,048,576 bytes (1 MiB), copies videos as-is, and skips simple duplicates. Purpose-built for my workflow, I don’t expect most people to need this. It trades options for speed and repeatability.

**Synopsis**
Recursively mirror a source folder under Pictures, convert images to JPEG (1 MiB best-effort target, auto-orient, convert profiled colour to sRGB, flatten alpha, strip metadata), copy videos unchanged, skip heuristic duplicates by filename/time plus equal byte length, show progress, write a detailed log.
Drag-&-drop workflow: I drop a folder onto the .bat and find the normalized copy in Pictures.

**Requirements**
64-bit Windows 10/11
PowerShell (Windows PowerShell 5.1 or PowerShell 7)
ImageMagick 7.1.2-32 or newer supported 7.x build (magick.exe) in PATH
(Optional) The included .bat wrapper for drag-and-drop

Supported hosts and actually tested hosts are distinct. Desktop regression tests
ran on Windows 11 Pro build 26300 with Windows PowerShell 5.1.26100.9444 and
PowerShell 7.6.5. Hosted CI ran on Windows Server 2025 Datacenter 10.0.26100 with
Windows PowerShell 5.1.26100.33438 and PowerShell 7.6.6. Maintained full gates use
portable ImageMagick 7.1.2-32 Q16 x64; the installed Q16-HDRI build of the same
version has separate colour and owner workflow checks. Windows 10 has not been
tested. Installed codec availability and policy still determine which files work.

**Version and release status**
The first managed application version is **unreleased**. See [CHANGELOG.md](CHANGELOG.md)
for its version, fixes, safeguards, tested environments and limitations, and
[release-metadata.json](release-metadata.json) for machine-readable draft metadata.
No public tag, release date or downloadable package is established by these files.

`Get-WinImgVersion` in `WinImgNormalizer.ps1` is the single application version
source used by usage text and run logs. Maintainers regenerate the changelog header
and release metadata from that source and [the release prose](docs/release/NOTES.md):

```powershell
.\tools\release\Update-ReleaseMetadata.ps1
.\tools\release\Update-ReleaseMetadata.ps1 -Check
```

The check returns failure when generated files are stale. Neither command creates
a tag, publishes a release or builds a download package.

**Installation**
Install a supported ImageMagick build for Windows and ensure magick.exe is in PATH
(`magick -version` should work). The minimum reviewed build is **7.1.2-32**, as of
4 October 2026; update it as upstream publishes security fixes. The tool selects
an application executable rather than a PowerShell function or alias, checks its
version and compiled format capabilities, and records its exact path and identity
in the local log. It never installs or updates dependencies automatically.
Place these two files together (same base name), e.g. in Downloads:
WinImgNormalizer.ps1
WinImgNormalizer.bat (wrapper for double-click + drag-and-drop)
No config files, it just runs.

**Usage**
1) My everyday flow (drag & drop onto .bat)
Drag a folder (with photos/videos) onto WinImgNormalizer.bat.
The normalized copy appears in: %USERPROFILE%\Pictures\<Source>_WinImgNormalized_<yyyyMMdd_HHmmss>_<runId>
A log file is saved inside the output folder.
2) Command line (positional args, avoids PS 5.1 param-set quirks)
#Whole folder (recursive), default 1,048,576-byte (1 MiB) target
.\WinImgNormalizer.ps1 "D:\Photos\2024"
#Whole folder with custom size cap (bytes), for example 2 MiB
.\WinImgNormalizer.ps1 "D:\Photos\2024" 2097152

Supply exactly one source directory and, optionally, a positive whole-number byte
cap from 1 through 9223372036854775807. Small caps remain best-effort and may be
unachievable. Missing folders, files used as folders, non-filesystem sources,
drive-relative paths such as `C:`, and linked source roots or ancestors are rejected
with a setup message. Ordinary relative directory paths are resolved to absolute filesystem
paths; drive and UNC share roots retain their root separator.

The destination must be outside the source tree. Choosing Pictures itself, its
parent, or a drive root that contains Pictures is rejected before enumeration or
any destination write, including probes. Choose a source subfolder instead; the
tool does not silently relocate output. Path comparisons use canonical directories
and case-insensitive whole path segments, so `PhotosBackup` is not inside `Photos`.
Read-only 128-bit directory identities also reject an existing destination ancestor
that aliases the source, including local drive/share aliases. A provider that
cannot supply this identity fails setup because the alias boundary is unverified.
Linked destination paths or ancestors are also rejected. Device paths and directory
components with trailing dots/spaces are unsupported.

Media names and directory components containing spaces, apostrophes, ampersands,
parentheses, exclamation marks, percent signs, brackets, hashes, at-signs, leading
hyphens and Unicode characters are treated literally. Images are read from neutral
private copies; JPEG and extracted-profile writes disable ImageMagick filename
formatting before a literal filesystem move to the planned final name. This also
protects percent expressions in source and destination parent directories.

Long local paths require support from the Windows host, filesystem, permissions
and installed tools. Extended native paths are used for long scratch names;
unsupported or inaccessible paths report a setup or item failure instead of
truncating a path or redirecting output. Live UNC/network-share validation and
support are outside the project's owner-approved scope. Local long-path tests
and lexical drive/UNC-root safety checks remain required; they do not establish
live network-share support.

Traversal skips and logs directory/file reparse points (links and junctions) without
following them. Inaccessible subtrees are reported as an incomplete scan; other
readable media continues. These omissions return exit code 2, including when no
eligible files were readable. Concurrent hostile filesystem changes are outside the
tool's isolation guarantee, though paths are rechecked before reading/writing.

Arguments and ImageMagick are checked before creating Pictures or a run folder.
The destination is checked with a temporary exclusive write probe that is removed,
and a best-effort free-space estimate: copied video bytes, per-image output budgets,
twice the largest readable image input for scratch, and 64 MiB reserve. Each image
budget is the smaller of its byte cap and the larger of 1 MiB or four times its
input size. Missing-decoder files need no conversion space. Insufficient measured
space fails setup; unavailable free-space information (for example, some UNC shares)
is logged as a warning. These checks cannot guarantee later writes succeed.

ImageMagick remains required for **video-only and empty batches**. JPEG write support
is required when there are readable images. An unavailable image decoder is reported
for each affected file without conversion retries; supported images and videos
continue. Such a partial batch returns exit code 2; setup failures return 1.
Compiled capabilities do not override local ImageMagick policy or prove every file
can be decoded. Keep the installed security policy in place.

**What it does**
Non-destructive mirror: exact subfolder structure; image files become .jpeg (same base names when unique).
Images: JPG/JPEG/PNG/BMP/TIF/TIFF/GIF/HEIC/HEIF/WebP → JPEG targeting 1,048,576 bytes (1 MiB).
Auto-orient via EXIF
Strip metadata
Convert to sRGB
Flatten transparency to white
Progressive attempts: scale 100→90→80→70→60→50% while jpeg:extent searches encoder quality toward the byte target.
Videos: mp4/mov/mkv/avi/m4v/wmv/webm/mts/m2ts/3gp/3g2 are copied as-is.
Heuristic duplicates: a later file with the same lowercase filename, LastWriteTimeUtc ticks and byte length as a successfully retained source is skipped and linked to that source/output in the log.
Progress + logs: console progress bar and a timestamped log in the destination.

**Byte targets and best effort**
The default remains **1,048,576 bytes = 1 MiB**, rather than decimal 1 MB
(1,000,000 bytes). A custom second argument also means bytes. The script passes
an invariant byte operand such as `jpeg:extent=1048576B`; ImageMagick's `KB`/`MB`
suffixes use decimal units and are not PowerShell's binary `1KB`/`1MB`.

The fully decoded JPEG's actual file length determines compliance: **bytes ≤ cap**
is `OK IMG`. If all six scales are needed and the valid 50% result still exceeds
the cap, it is retained under the existing best-effort policy as `WARN IMG`, with
exact output/cap bytes, width, height and chosen scale. Invalid or failed final
attempts are errors and cannot resurrect an earlier above-target candidate.
`SizeWarnings` counts retained above-target JPEGs as a subset of converted images,
and the PowerShell application returns warning/partial code 2. The BAT reports
that warning and returns the same code after its pause. Duplicate links retain
`ConvertedWithWarning`.

The scale sequence, no-upscaling behavior, encoder quality search, 4:2:0 sampling,
progressive JPEG, colour and white-alpha policy remain unchanged. Fixing the extent
units can change output bytes/quality because the encoder now receives the intended
budget. JPEGs are lossy derivatives; keep originals. ImageMagick's search uses
floating point, so extremely large Int64 budgets can round internally; the final
file-length comparison remains exact.

Output names are planned for the whole tree before processing, in stable ordinal
source-path order. Unique names keep the familiar basename. Images sharing a stem,
such as `photo.jpg`, `photo.png` and `photo.heic`, receive `photo__jpg.jpeg`,
`photo__png.jpeg` and `photo__heic.jpeg`. Mirrored directories, video names, generated
work/reports and all unique image names are reserved first, ignoring case for
collision checks. An existing `photo__png.jpeg` keeps its name and forces the PNG
derivative to `photo__png__2.jpeg`; further conflicts use the next free number.
Every eligible source has a `PLAN` source-to-output entry in the local log, even
when it is later skipped or fails. The duplicate heuristic still applies.

Duplicate matching follows that ordinal source-path order and registers a source
only after its output has been validated and finalized. A failed conversion, copy
or final move cannot suppress a usable later candidate. Different byte lengths
are processed independently, and each length keeps its first successful source
and output for later matches. A valid above-target JPEG can be retained with its
best-effort warning; later skips report that retained warning status.

Source metadata is read again before duplicate lookup. A finalized image whose
source length or modification time changed during processing is kept with a warning
and excluded from duplicate matching. Video registration uses the verified copy's
length/time snapshot. Every skip records the retained source, retained output and
status. This remains a lightweight heuristic: different content with the same
filename, timestamp and byte length can still be skipped. No runtime content hash
database is added, and source files are never removed.

Every image attempt uses a fresh, exclusively reserved neutral scratch file. A
successful native exit, nonempty JPEG bytes, one frame, positive dimensions and
full pixel decode are required before finalization. Invalid or truncated outputs
and files from earlier attempts cannot become successful derivatives. The current
scale sequence and valid best-effort size fallback remain; explicit multi-frame
selection produces one intended output, and any unexpected numbered multi-output
attempt remains rejected.

GIF and WebP become the first displayed logical frame, including its canvas and
placement. Multipage TIFF becomes its first page; pages are never stacked. HEIC/HEIF
uses the decoder's primary image (or first sequence image). The log records the
decoder-exposed source count, one selected image and the omitted frames/pages/images.
Animation and omitted pages are not preserved in the JPEG. This deliberate still
policy is informational and does not by itself change a successful exit code.

Every image is inspected and converted from the same private copy in owned
scratch, with source length/modification-time stability checks. This temporarily
requires room for the source bytes as well as the JPEG attempts and any inspected
profiles. For sequence-capable formats, the first-image define applies to this
neutral input without appending a selector to a user filename. Only an animation's
selected first image is coalesced onto its canvas, then orientation is applied.
Counts describe images exposed by the installed decoder, not hidden thumbnails or
all items in a HEIF container. Decoder capability and local security policy still
apply; an inspection failure fails that item rather than guessing its properties.

Colour conversion happens while the source ICC profile is still attached. Tagged
RGB and CMYK use that characterization to transform into sRGB through ImageMagick's
colour-management delegate. The target profile is included in the script; no
separately installed ICC file is needed. The requested rendering intent is relative
colorimetric with black-point compensation disabled. A profile that supplies only
a perceptual mapping may use the delegate's fallback; this is not a promise about
all printing conditions or every profile's rendering intents.

Untagged ordinary RGB is assumed to be sRGB. Explicit linear RGB and grayscale
spaces use the decoder's declared transfer semantics when converting to sRGB.
Untagged CMYK and unsupported unprofiled colour spaces fail with an explanation:
the missing source characterization is not guessed. Malformed profiles, unsupported
profile/channel models and reported native transform diagnostics fail the item,
even when ImageMagick writes a decodable JPEG and returns zero. Other readable
items can continue, with exit code 2 for the partial run. Profile checks establish
basic structure/model consistency and an accepted native transform, rather than
certifying arbitrary ICC contents or the source creator's characterization.

After orientation and colour conversion, transparent pixels are composited onto
white in encoded sRGB channels; fully transparent pixels become white. Every scale
attempt retains the same profile, white-background and alpha operations. Only then
are GPS/EXIF, XMP, IPTC, comments and source/target profiles stripped. Final JPEGs
omit ICC after sRGB conversion, retaining the tool's metadata-removal policy. The
log records the inspected source colour space, ICC presence and chosen policy.
Original image bytes and embedded metadata remain untouched.

Videos are copied into owned partial files. After the copy streams finish, the
partial size and source length/modified time must match before finalization. Videos
retain their original bytes and metadata; runtime content hashing is not required.
These checks detect observed changes and cannot guarantee a snapshot against
concurrent hostile filesystem changes.

Both media types use a move on the destination volume that refuses replacement.
If a file or directory appears at a planned target, it is preserved, the item fails
and the batch returns code 2. Cleanup deletes exact owned partials and empty owned
directories only. Unexpected neighboring files remain for diagnosis. A forced
process interruption can leave partials under the run's generated `work` directory;
such files are not final outputs.

**Output location**
Default target is the Windows Pictures folder:
%USERPROFILE%\Pictures\<Source>_WinImgNormalized_<timestamp>_<runId>\
Each run has a random 32-character ID and is created exclusively; an existing file
or folder is never reused as a run. Inside you'll find the mirrored tree, converted
.jpeg images and copied videos. Generated files occupy `.WinImgNormalizer`, or the
first free suffix such as `.WinImgNormalizer__2` when a source entry uses that name.
Its `reports` subfolder contains `WinImgNormalizer_<timestamp>.log`; `work` is
reserved for temporary processing. Source directories with these names are mirrored
normally. If creating or appending the required log fails, valid media work continues
with a visible warning and application code 2. The failed disk sink is disabled once;
a console fallback retains at most 8,192 UTF-16 characters, limits each stored line
to 1,024 characters and reports dropped/truncated lines. The final reporting status
shows whether the disk log is incomplete, including a failure on its last write.

Each successfully discovered regular file has one outcome: Converted, CopiedVideo,
SkippedDuplicate, Ignored, Error, Cancelled or NotStarted. Warnings are attributes,
so they do not increase the file count. Unsupported files are counted as Ignored
at discovery. Unknown files inside unreadable folders are excluded from this known
partition; separate scan issues identify those locations. The final ACCOUNTING
line reconciles the partition, reports observed elapsed time and finalized files
per second, and pairs processed-image input bytes with validated JPEG output bytes.
Image savings are signed and can be negative; video bytes are separate. Skips,
ignored files and errors create no image savings.

The same reports directory contains `WinImgNormalizer_<timestamp>.csv`, written
locally with UTF-8 BOM and CRLF. File rows record original relative paths, planned
and actual output names, warning attributes, attempts, byte totals, omitted frames,
and retained duplicate mappings. ScanIssue rows describe unreadable locations and
skipped links without inventing file totals. A required CSV creation, write or close
failure keeps completed media, prints a warning and returns application code 2;
the incomplete report must not be treated as a complete inventory. Cooperative
cancellation makes one best-effort report attempt and keeps application code 130.

Every CSV text cell starts with the visible safety prefix `text:`. CSV quoting
preserves field boundaries; the prefix prevents a filename from opening as a
formula. Backslashes are doubled, and UTF-16 control, format, line/paragraph
separator and surrogate units are written as `\uNNNN` escapes. To recover an exact
path programmatically, remove exactly one `text:` prefix and decode left to right:
`\\` becomes one backslash and `\u` plus four hexadecimal digits becomes that
UTF-16 unit. Decode only for filesystem or text processing; keep the prefix when
displaying in a spreadsheet. Numeric fields contain generated invariant numbers;
unknown numeric values are blank.

The verified machine import is `Import-Csv -LiteralPath '<report.csv>' -Encoding
UTF8` in Windows PowerShell 5.1 or PowerShell 7. It returns string-valued objects.
For Excel, use Data > From Text/CSV, UTF-8 and comma, disable automatic type detection
and set all columns to Text in Transform Data, then load with the safety prefix
intact. This Excel recipe is documented guidance; actual Excel execution and later
editing/save/reopen behavior are not certified by the PowerShell parser tests.

**Batch wrapper (included)**
The repo includes a minimal wrapper so you can drag a folder onto the .bat.
It runs the .ps1 positionally with Windows PowerShell, `-NoProfile` and a
process-only `-ExecutionPolicy Bypass`. Persistent execution policies and security
settings are not changed; no administrator prompt or automatic installation is added.
Keep the .bat and .ps1 in the same folder and with the same base name.
Keep the BAT's Windows CRLF line endings; the included file-specific Git
attribute preserves its exact bytes in checkouts and source archives.
Drop exactly one folder. Zero or extra arguments, including an empty extra
argument, produce a usage error before processing. To choose a custom byte cap,
use the script's existing second positional argument.
The launcher saves the script's exit code immediately, displays its outcome,
keeps the familiar pause, and returns the saved code:

| Code | Outcome |
|---|---|
| 0 | Completed successfully, including empty or ignored-only input. |
| 1 | Usage, setup or run-level failure prevented reliable completion. |
| 2 | Completed with warnings or partial failures; review the report. |
| 130 | Cooperative cancellation handled by the application; completed outputs retained. |

An unexpected host exit is reported as an error with its actual code and is also
returned unchanged. Closing or force-killing a host may bypass the launcher pause.
The wrapper disables inherited delayed expansion so `!` remains literal and
preserves a trailing folder/root separator when forwarding its quoted argument.
A command prompt can expand `%VARIABLE%` or `!VARIABLE!` before the wrapper receives
the argument. For paths containing those expressions, invoke the PowerShell script
with a single-quoted literal from PowerShell; the wrapper
cannot reconstruct characters already expanded by its caller.

**Technical details**
ImageMagick applies auto-orientation, an embedded-ICC to sRGB transform (or the
explicit untagged RGB/grayscale policy), white alpha compositing, then metadata
stripping. JPEG settings remain -sampling-factor 4:2:0 -interlace Line and
-define jpeg:extent=<exact decimal bytes>B. The six scale attempts remain 100→50%; retries retain
all colour and alpha operations. Strict source inspection and full JPEG decode
precede finalization. Only valid above-target results advance to a smaller scale.
Native failures stop the item; a diagnosed sharing/lock violation can retry the
same scale twice, after 100 and 200 ms, with fresh scratch and identical flags.
Generic access denial, missing codecs, damaged input, ICC errors and exhausted
resources never trigger scale fallback. A nonzero native exit never finalizes.

Stdout and stderr drain independently through fixed buffers on PS5.1 and PS7.
Per-item capture retains at most 16,384 characters per stream (UTF-16 character
storage), preserving the beginning and end plus a truncation marker and total
counts. Overflow fails the item. Preflight format/version queries have a separate
262,144-character bound and retain their initial 15-second deadline. Attempt,
category, native exit and bounded details go to the log; the console shows compact
item results and retry reasons. Only the pinned JPEG writer's zero-exit
lossless-to-lossy notice is accepted, after full JPEG validation; every other
conversion diagnostic fails. An accepted notice counts as NativeWarnings and
returns application warning code 2. Source/ICC inspection and candidate validation
remain strict. The production -quiet/-regard-warnings flags remain enabled.
Each image has one shared 120-second runtime budget across frame/colour inspection,
ICC extraction, conversion attempts, backoffs and full JPEG validation. A timeout
ends that item, discards its candidate and continues accessible siblings with
application code 2. Each ImageMagick invocation sets pixel-cache ceilings of
512 MiB memory, 1 GiB mapped cache and 2 GiB disk, at most two worker threads and
a native time ceiling derived from the remaining budget. Stricter installed
security policies stay effective; the tool does not replace or edit policy files.
These are practical cache limits, not a total decoder/delegate heap cap or an
exploit-proof sandbox. Synchronous filesystem reads/copies are not preempted.

Native processes enter a private Windows job at creation, with only their explicit
stdin/stdout/stderr handles inherited. Timeout or return closes the owned process
tree, including descendants that outlive the parent; unrelated ImageMagick
processes survive. Failure to establish that ownership fails the native start.
Termination and pipe draining have a short bounded grace after the deadline.
TEMP, TMP and MAGICK_TEMPORARY_PATH point to an exclusively allocated per-image
scratch directory under the existing Windows temporary root, separate from deep output
paths. A long, linked, unavailable or source-contained temporary root fails setup
without changing it. Cleanup removes only regular magick-* cache files there;
unexpected or linked entries remain with a warning. The private directory and
reserved cache names establish ownership; concurrent same-pattern filesystem
mutation is outside this check. Unconfirmed tree termination preserves all scratch
for that item with a warning. Completed outputs are retained.
Output creation and modified timestamps are restored separately from the captured
source times. A failed setter retains the valid output, names the failed field and
adds one TimestampWarnings attribute to that finalized image/video; the application
returns code 2. A later heuristic duplicate reports the retained warning status.

Discovery visits folders independently. An unreadable folder or uninspectable entry
makes ScanComplete=False while accessible siblings continue; unknown files inside
that location are not counted as processed. Incomplete directory/entry counts are
separate from skipped links and media totals. Directory and file reparse points are
skipped and reported without traversal; such warnings also return code 2. Linked
roots or ancestors are rejected before output creation; choose the actual directory.
Log messages escape control characters so a path or diagnostic cannot create fake
report lines. Direct stderr reporting is best effort if the host pipeline has stopped;
this does not establish native cancellation, cleanup or a guaranteed Ctrl+C exit.

**Stopping a run**

Ctrl+C asks a command-line run to stop taking new files and terminate its owned native work. When the application handles the request, it keeps completed outputs, removes known incomplete image/video files where safe, records an INTERRUPTED outcome and returns 130. A fully validated final move admitted before the request may finish; that completed output remains. Originals and unrelated files/processes are untouched by cancellation cleanup.

Native waits and chunked video copies check the request repeatedly. A blocked filesystem operation can delay stopping. Closing the terminal or force-killing its process can bypass cleanup and the interrupted record, and does not promise exit 130. Recognizable files in the private work area are incomplete staging data, not completed outputs.

When the PowerShell child returns cooperative code 130, the .bat displays
“Cancelled,” pauses and returns 130. Actual shell and batch interrupt behavior
remains host-specific; terminal closure, a stopped pipeline or forced termination
can bypass application reporting and the launcher's completion handling.

**Tweaks (optional)**
Different size cap: pass a second positional argument in bytes (for example 2097152 for 2 MiB).
Background color for alpha: change -background white in the script (for example to black).
Dimension ceiling: add an explicit -resize rule (for example -resize "1920x1920>") before jpeg:extent if you want a max edge.

**Troubleshooting**
“magick not found” -> Install ImageMagick 7.1.2-32 or newer supported 7.x and ensure magick.exe is in PATH. Test with magick -version.
PS 5.1 “Parameter set cannot be resolved” → Always launch via the provided .bat (positional args) or call the .ps1 with positional arguments only.
Colour/profile errors -> Check the per-file log. Untagged CMYK needs a correctly profiled source; malformed or incompatible profiles are not silently converted. Originals remain intact.
No outputs -> See the log file in the destination for errors (permissions, unreadable files, ...).
Very noisy images may exceed the byte target even at 50% → A fully validated final JPEG is retained as WARN IMG, with its exact bytes/dimensions/scale and a warning outcome.

**Intent & License**
This is a personal tool for a specific workflow (archiving mixed photo folders for interviews/projects). It’s provided as-is, without warranty. Use at your own risk. Feel free to adapt it, just note it intentionally avoids extra features to keep the workflow fast and predictable.
