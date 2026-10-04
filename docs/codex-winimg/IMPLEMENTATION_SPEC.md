# Implementation contract — improve, do not rewrite

This document sets acceptance boundaries, not a replacement source file. Read the
current code before selecting the smallest patch. The product continues to be a
local PowerShell script and batch launcher. External technical references are in
SOURCES.md; implementation decisions below are this project's requirements.

## 1. Public entry points and test isolation

Keep `WinImgNormalizer.ps1 <sourceFolder> [maxBytes]` and the one-folder batch
launcher. Preserve the default cap, output extension, folder mirroring, white
background and video-copy behavior. Do not repeat the old README's blanket claim
that named parameters are inherently unsafe in PowerShell 5.1; the compatibility
requirement is to retain the working positional interface, not to prohibit correct
parameter binding forever. New public options need a specific use case and review.

Wrap orchestration in the smallest practical callable function. Keep an outer
entry point responsible for usage and process exit. Unit-test discovery must not
run the application, write to Pictures or terminate the test host. An internal
injected output-parent/process-runner seam is sufficient for tests; do not build a
configuration framework. Extract only tightly related helpers when useful, preferably
within the existing script or one small supporting file. Do not split into a large
module/class hierarchy. Preserve useful comments, original authorship and the MIT
license. The development Python bundle helper must never become a user runtime need.

Avoid unguarded PowerShell 7/.NET-only APIs in common code: e.g. `$IsWindows`,
`Path.GetRelativePath`, `ProcessStartInfo.ArgumentList`, or newer process-tree kill
methods do not establish Windows PowerShell 5.1 compatibility. Use documented
cross-version paths or isolated version-specific implementations with both tested.
Do not confuse a Linux PowerShell 7 pass with a Windows batch-launcher pass. [P1,P2]

## 2. Preflight, roots and containment

Validate zero/extra arguments, a real filesystem directory and a positive bounded
Int64 limit. Do not silently coerce negative, overflow, noninteger or unusable values.
A practical minimum cap may be proposed only with a documented reason; best-effort
size behavior must remain clear. Normalize roots with root-aware APIs rather than
trimming `C:\` to `C:`. Treat UNC share roots deliberately and compare full path
segments, not naked string prefixes. Keep the actual canonical source path in memory.

Resolve an application executable, not an arbitrary alias/function named magick.
Record exact path, version and supported relevant formats. A version command that
exits nonzero is not a successful dependency check. Use a currently patched supported
ImageMagick build, not a frozen assumption that every 7.x build is suitable. Read
current upstream notices at implementation time; the 2026 `jpeg:extent` hang advisory
is particularly relevant. Do not infer a safe range from inconsistent version
strings in advisory metadata. [I3,I4]

Validate output parent and estimate space for copied videos plus image scratch.
Space checking is best-effort, not a promise that a later write cannot fail. Do not
create application run output on argument/executable failure. Missing optional
codecs should be diagnosed per format; supported images/videos can continue under
the defined partial-run policy. For simplicity the initial product may still require
ImageMagick on video-only input; document and test that rather than changing it
incidentally. Never install a dependency automatically or weaken global policy.

**Initial chosen nesting policy: reject any destination inside or equal to the
source before output creation.** Explain that the user should choose a safe source
subfolder, or a future reviewed output-parent option. This deliberately prevents
processing Pictures itself when the fixed destination is within Pictures. Do not
silently relocate output. Any future exclusion-based approach must prune before
descent, handle aliases/reparse points and have the same safety tests.

Default traversal skips and reports directory/file reparse points rather than
following junctions or links. Resolve an input root alias deliberately or reject
an unresolved one; it is not permission to follow links below the root. Never change
ACLs to make a source readable. A later inaccessible subtree is an incomplete scan,
not a claim that its unknown number of files was processed.

## 3. Deterministic planning and no-overwrite names

Enumerate with stable ordinal path comparisons appropriate to Windows. Reserve the
complete planned namespace before conversion, including mirrored directory names,
video paths, image output names, log/report and scratch paths. Preserve the legacy
name for a nonconflicting file. A suggested colliding image suffix is `__<sourceext>`
with a stable numeric disambiguator when that name also exists. Check the candidate
against ALL reserved file and directory paths case-insensitively for ordinary Windows
filesystems. A source already called `photo__png.jpeg` must not break the scheme.

Create each run root exclusively with a collision-resistant suffix, keeping the
recognizable `<Source>_WinImgNormalized_<timestamp>` prefix where practical. Do not
reuse a directory just because `New-Item -Force` permits it. An existing run folder
must never be adopted or cleaned. Reserve a private scratch namespace that cannot
collide with a mirrored source directory. Preserve an explicit source/output map.
Recheck a final target just before finalization; do not overwrite a new external
arrival even if the initial plan was clear.

## 4. Conversion and finalization transaction

Each attempt has a unique neutral temporary basename within this run's owned area,
preferably on the same volume as the final output. Open new files exclusively.
Never run ImageMagick against a final target. Collect process exit code, diagnostics,
actual attempted settings and timeout/cancellation status. After an acceptable
exit, verify nonzero bytes, JPEG format, one output frame, positive dimensions and a
successful full pixel decode. Header-only identification is not sufficient to catch
truncated data. Measure actual bytes; do not accept an old candidate left by another
attempt. A fatal/nonzero native result fails the attempt even if it left a file.
Known warnings with exit zero may be tolerated after validation and reported.

Finalize using a same-volume no-overwrite rename/move. Do not call the whole batch
atomic or promise crash durability: only the per-file final-name transition is
intended to be indivisible where the filesystem supports it. Do not use `Move-Item
-Force` or another overwrite fallback. Treat cross-volume staging as a design error
or explicitly use another validated safe procedure.

For videos, copy to an owned partial path, wait for stream completion, verify size
and source length/timestamp stability, then finalize without replacing any file.
Runtime content hashing is not required; regression tests verify copied bytes with
hashes. Detected source changes are an error/warning, not a completed stable snapshot.
Track owned temporary paths explicitly. Cleanup may delete only known owned items,
never a wildcard of similarly named files or an entire pre-existing folder.

## 5. Multi-image sources and native filename parsing

A GIF/animated WebP yields one first displayed logical frame, respecting its canvas
placement. A multipage TIFF yields one first page, not stacked pages. HEIC/HEIF
sequences use the decoder's documented primary/first image; inspect fixture results
rather than assuming thumbnails count as intended pages. Log source frame/page
count and omission. Do not rely on the JPEG writer's implicit numbering. [I1]

`-LiteralPath` protects PowerShell filesystem operations, not ImageMagick's input
selectors, output formatting, coder prefixes or percent/bracket syntax. Use safe
native argument construction without `Invoke-Expression` or unescaped shell strings.
Literal-output support or neutral-temp output plus a filesystem rename can protect
final names. For ambiguous input syntax, a literal filesystem copy to a neutral
controlled scratch basename is an acceptable incremental solution; account for its
space/time cost. Never blindly append `[0]` to unexamined user filenames and call
that literal-safe. Validate real Windows argv behavior for spaces, Unicode, percent,
brackets, ampersands, parentheses, exclamation marks and trailing root separators.
Do not strip characters or truncate paths silently. [I1,I2]

## 6. Colour, metadata, orientation and transparency

Auto-orient while orientation information still exists. Inspect the embedded ICC
profile before stripping. Transform tagged colour into sRGB through an appropriate
profile-aware path and only THEN apply the privacy metadata policy. `-colorspace
sRGB` after profile removal cannot recreate an arbitrary lost source profile. Use
trusted redistributable reference profiles with their provenance/license, not an
invented path to a profile presumed installed everywhere. No fonts are needed. [I1]

Define the behavior for already-sRGB, wide-gamut RGB, tagged CMYK, untagged RGB,
untagged CMYK and malformed profiles. Untagged RGB may be treated as sRGB under a
clearly stated assumption. Missing CMYK characterization cannot be promised perfectly
accurate: choose a controlled documented fallback with a warning, or fail the item
with a useful explanation. Do not conceal uncertainty as verified colour accuracy.

Use a documented alpha-compositing order and verify semi-transparent patches against
a reference; white remains the default. No retry may silently remove the alpha or
background operations. After conversion remove privacy metadata (GPS/EXIF/XMP/IPTC
and other unwanted source metadata) and apply the declared output-profile policy.
Default output may omit ICC after known-sRGB conversion to retain the existing
stripping behavior. Test pixels as well as metadata. JPEG comparisons need tolerances,
not a byte-identical cross-build checksum expectation. Originals remain untouched.
Copied videos are unchanged and **retain any metadata they already contain**.

## 7. Size targeting and quality

The existing default is 1,048,576 bytes = 1 MiB. Keep the actual byte limit as the
source of truth. Verify the extent-unit syntax against the installed ImageMagick;
PowerShell `1MB` is not evidence that an ImageMagick `MB` string uses identical
units. Prefer an unambiguous validated byte count. `jpeg:extent` already controls
encoder quality toward a size budget; benchmark before replacing it. [I2]

Initially retain the 100,90,80,70,60,50 percent scale sequence, no upscaling and the
best-effort fallback. Above-target valid files get a distinct warning result, actual
bytes and chosen dimensions/scale. Invalid files cannot use best-effort as an excuse.
Do not silently lower a quality floor, add sharpening, change chroma subsampling or
resize more aggressively. A measured algorithm proposal needs comparative evidence
and owner acceptance, not an unreviewed default change disguised as cleanup.

## 8. Duplicate heuristic

Maintain the lightweight case-normalized filename + LastWriteTimeUtc key and require
matching byte length before a skip. This length guard is an intentional documented
safety change. Register a retained key only after successful finalization; a failed
first candidate cannot suppress a good later one. A valid above-target retained
output may register success-with-warning; record it accurately. Ordering is stable.
Every skip records the retained source and output. No hash database or mandatory
full-file hashing is added. Same-key, same-length different-content inputs can still
be falsely matched; document this and never describe the heuristic as proof.

## 9. Native execution, retry limits and resource budgets

A small process wrapper should return structured exit/output/diagnostic/timeout data.
Drain both output streams without pipe deadlocks; retain a bounded diagnostic tail
or bounded log stream. Do not let different PS5.1/PS7 native-stderr semantics decide
randomly whether a warning is fatal. Capture exit immediately and separately from
PowerShell exceptions. Keep console output quiet but failures diagnosable. [P1,P2]

Limit per-file execution time and reasonable memory/temp-disk usage. Respect stricter
installed ImageMagick security policy; do not replace or relax it. Per-job limits
are not a complete sandbox. Reject unsupported or suspicious input through deliberate
local-file/coder boundaries where feasible without disabling required HEIC/TIFF
support. Do not suggest a web-safe policy that breaks the promised offline formats.
No downloads, URL input, arbitrary indirect reads or server endpoint are part of
this product. Review current dependency advisories. [I3,I4]

Classify failures: missing codec, invalid image, permissions, resource exhaustion,
timeout/cancellation, and truly transient I/O. Retry only justified transient errors
with an explicit low bound and backoff. Do not retry missing codecs at six scales
or remove alpha flags as a universal recovery method. Timeout/cancellation kills
only the owned process and descendants, never all `magick.exe` processes. The
implementation must actually work on Windows PowerShell 5.1; a newer .NET overload
alone is not sufficient.

## 10. Cancellation and ancillary failures

On cooperative cancellation stop starting items, terminate owned in-flight work,
clean known partials where possible, retain verified completed files and emit an
interrupted summary. Real Ctrl+C behavior differs by host; test it through both
PowerShell versions and `.bat`. A forced terminal/process kill may bypass cleanup:
recognizable partials must remain nonfinal, and documentation must not promise an
unconditional cancellation exit code or complete cleanup in that scenario. [P1]

Folder enumeration should continue past inaccessible subtrees and report their
paths safely. Timestamp-setting failures are nonfatal but visible. Log-write failures
need a bounded fallback and console warning without recursive logging failure.
Do not silently ignore failed log/report output. Prefer continuing valid work with
an explicitly degraded run status when feasible.

## 11. Outcome, reporting and exit contract

Use structured outcomes internally. Partition discovered files into converted
images, copied videos, skipped heuristic duplicates, intentionally ignored files,
and failed files. Warning flags (above size target, frames omitted, timestamp
failures) are attributes, not extra files to add to the partition. Record inaccessible
directory count separately because its file count may be unknown. A failed complete
scan cannot be reported as a full success. Show elapsed time and observed throughput;
any ETA is approximate and optional, not a requirement to add complexity.

Report image-only paired input/output bytes, signed savings, and video-copy bytes
separately. Skips/ignored bytes must not become invented compression savings. Keep
the readable log and add a local relative-path CSV including source/output, status,
reason, bytes, scale/dimensions when known, attempts and retained-duplicate mapping.
Prevent CSV formula evaluation for unsafe leading characters (including whitespace
variants); CSV quoting alone does not do so. Escape control/newline text for logs.
Sensitive absolute paths may remain local for troubleshooting but must not be
committed or uploaded as evidence. Progress/report failures must remain visible.

Chosen application exit contract for implementation and Windows tests:

| Code | Meaning |
|---|---|
| 0 | Completed successfully; intentional ignored files/heuristic skips are not errors. Empty/no-eligible-input is explicitly reported. |
| 1 | Usage, setup or run-level failure that prevents reliable normal completion. |
| 2 | Completed with warnings and/or partial failure: e.g. above-target output, inaccessible subtree, file error, required log/report degradation. |
| 130 | Cooperative cancellation handled by the application. Host-force termination is separately documented. |

A normal deliberate first-frame policy can be informational rather than forcing
code 2; visible omission reporting is mandatory. Settle severity centrally and
apply consistently. The batch launcher saves `%ERRORLEVEL%` immediately after the
PowerShell call, chooses an honest message, pauses for drag-and-drop, and returns
the saved code. Reject more than one dropped folder rather than ignoring extras.
Keep `-NoProfile`; existing process-scoped execution-policy behavior is documented,
not expanded into a persistent machine-wide change. [P1,P2]
