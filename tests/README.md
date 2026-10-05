# Windows checks (through M3-T03)

These development tests cover the import boundary, the two existing positional
invocations, setup validation (T008-T012), traversal/run isolation (T013-T016),
deterministic naming/no-overwrite targets (T017-T019), validated conversion and
staged video transactions (T020-T024), and successful-retention duplicate handling
(T025-T027). The combined safety regression (T028) runs one complete mixed tree
twice into separate outputs under the same parent. Frame/page regressions (T029-T031)
check deliberate first-image selection, logical animation canvases and visible omissions.
Colour and privacy regressions (T032-T036) compare managed patches, stripping,
orientation and white alpha composition. Literal native/launcher path checks
(T037-T039) preserve selected images and intended names. Cancellation regressions
(T054-T056) exercise owned native termination, staged-copy interruption and forced
exit boundaries. The characterization evidence in `legacy/` describes the baseline
defects separately.

Run `Initialize-TestDependencies.ps1` explicitly to download hash-pinned Pester
and portable ImageMagick into ignored `.scratch`. It verifies the archives and
extracts private copies without an installer, permanent PATH change, or persistent
execution-policy change. Normal image processing never calls this bootstrap.

```powershell
# Run separately in Windows PowerShell 5.1 and PowerShell 7 on Windows.
.\tests\Initialize-TestDependencies.ps1 -Download
.\tests\Invoke-Tests.ps1
```

The runner records actual host/tool versions, counts and NUnit results in ignored
scratch. It fails for a failed assertion, failed discovery, zero discovered tests
or any skipped test. `-DeliberateFailure` adds a real failing assertion; its expected
nonzero exit is checked by the Windows CI workflow after the normal suite passes.
No passing result should be inferred from merely importing Pester.

The application exposes `Invoke-WinImgNormalizer` and
`Invoke-WinImgNormalizerCommand` when dot-sourced. Internal `OutputParent`,
`MagickPath`, `PreflightRunner` and conversion-only `ProcessRunner` parameters let
tests use owned synthetic scratch trees and controlled process failures. The public
`.ps1 <sourceFolder> [maxBytes]` interface and `.bat` remain unchanged.

The default runner includes `Normalizer.Tests.ps1`, `Preflight.Tests.ps1`,
`Traversal.Tests.ps1`, `Naming.Tests.ps1`, `Transactions.Tests.ps1`,
`Duplicates.Tests.ps1`, `SafetyRegression.Tests.ps1`, `Frames.Tests.ps1`,
`Colour.Tests.ps1`, `Paths.Tests.ps1`, `Sizing.Tests.ps1`, `Diagnostics.Tests.ps1`,
`ConversionRegression.Tests.ps1`, `RecoveryReporting.Tests.ps1`, and `NativeLifetime.Tests.ps1`.
Preflight tests use the verified executable for ordinary JPEG checks and isolated
responses for dependency/version/codec failures. Drive and UNC root tests call only
the lexical path helper; they never normalize a drive root or network share. Denied
writes and disk-space failures use narrow mocks, without ACL or policy changes.
Video-only tests retain the documented ImageMagick prerequisite and verify unchanged
synthetic bytes. Space estimates are advisory; passing them cannot guarantee later
writes succeed.

Invocation tests use a disposable copy with only the outer call given internal
scratch output and dependency paths, then run its real `-File` entry in the current
Windows shell.
They check ordinary JPEG decoding/dimensions, video hashes, mirrored directories,
timestamps, source preservation and logs. This is instrumented invocation evidence;
it does not exercise real Pictures resolution or manual drag-and-drop. Tests never
run normalization against the real Pictures tree, private media or drive roots.

The CI matrix uses Windows Server runners. Desktop results are recorded separately
in `docs/codex-winimg/evidence/M0-T03.json`; Server CI does not prove desktop launcher
or interruption behavior. PSScriptAnalyzer checks belong to later tasks.

Traversal tests verify nested destinations are rejected before probes/enumeration,
case-insensitive segment comparisons, real disposable looping/outside/dangling
junctions, safe missing destination tails and incomplete scans. They force run-name
collisions with existing files/directories, test native allocation beyond MAX_PATH,
and launch two real child applications
at the same fixed timestamp behind a release barrier. Generated namespaces conflict
with source files/directories without replacing source data or logs. Ordinary media
checks retain byte hashes and creation/modified timestamps. Sources and children
remain in owned ignored scratch; loops are pruned in the fixture state checker.

File-symlink tests attempt unprivileged CreateSymbolicLinkW. A host without that
capability uses a mandatory controlled reparse-file entry and records its actual
native error in ignored link-evidence.txt; it does not report a real symlink pass.
UNC/drive-root containment is lexical only, and denied enumeration uses a narrow
mock. No live share, real Pictures, ACL change or manual launcher result is claimed.

Naming tests cover same-stem format plans, complete directory/video/generated
reservations, secondary numeric suffixes and case-only controlled inventories.
Repeated seeded shuffles under five cultures assert identical mapping and row
ordering. Real JPG/PNG/BMP collision outputs are fully decoded; source hashes and
creation/modified timestamps plus unchanged synthetic video bytes are checked.
HEIC naming/processing in the naming suite uses explicitly controlled capabilities
and JPEG candidate bytes, so it proves routing and distinct names. The frame suite
requires separate actual HEVC collection decoding with the selected executable.
External file/directory arrivals are injected during conversion and after the
availability check for both file operations.

Transaction tests require real full JPEG decoding and positive dimensions before
acceptance. A synthetic truncated JPEG retains readable header dimensions but fails
pixel decoding; no-file, empty, malformed, wrong-format and nonzero native results
remain uncommitted. Missing, string, extra-output, timed-out and cancelled controlled
results also fail. Fresh exclusive candidates prevent earlier above-cap results from
being reused after later failed attempts, and a later valid fallback is independently
validated. External final-name arrivals survive the final no-overwrite move.

Opaque synthetic video bytes pass through reserved partials before any final name
appears. Tests check byte hashes, source timestamps, detected source length/timestamp
changes, partial writes, and stream failures. Narrow copy mocks create deterministic
changes or interruptions; they do not claim real Ctrl+C or hard-kill coverage. Exact
owned-candidate cleanup preserves unrelated and numbered neighboring files and
preexisting arrivals. Raw fixtures remain in marked ignored scratch without recursive
test cleanup. No runtime content hashing or universal hostile-filesystem guarantee is
claimed.

Duplicate tests verify that a failed first native result, JPEG validation,
no-overwrite finalization or staged video copy cannot suppress a later usable
same-key source. Same filename and modification timestamp with different source
lengths are retained independently, including a later match to an earlier length.
Retained-source and retained-output links remain deterministic across shuffled
inventories under five cultures. Case-normalized filenames, distinct timestamps,
refreshed inventory metadata and changes during conversion/copying are covered.
A validated above-target finalized JPEG registers its explicit warning status.

The heuristic uses filename, modification timestamp and byte length; it is not
content verification. Tests deliberately show same-key, same-length valid images
and videos with different bytes can still be skipped. Originals remain preserved.
SHA256 checks belong to these regression assertions and are not a runtime hash
database or mandatory content-hashing feature.

The safety regression combines actual colliding JPG/PNG/BMP conversions, reserved
suffixes, a source generated-namespace lookalike, bracket/Unicode paths, hidden and
unsupported source files, mirrored empty directories, duplicate images/videos and
opaque video bytes. After each run it checks every source file's SHA256, creation
and modification times, attributes and directory state. Complete JPEGs receive
full decoding, frame/dimension checks and channel-mean comparisons with a small
tolerance; copied videos retain byte hashes and timestamps. Both runs must emit
the same complete source/output plan and retained duplicate links. The second run
must leave the first run's entire media/log tree and unrelated output-parent entries
unchanged. These fixtures stay in marked ignored scratch without recursive cleanup.
The full default suite retains the earlier mandatory controlled failure regressions;
this combined case uses actual conversion and copying without process mocks.

Frame tests generate actual offset, transparent GIFs with one or two frames, a
two-page TIFF with a real first-page RightTop orientation tag, and animated WebP.
They require exactly one planned JPEG, full JPEG decoding, dimensions and expected
pixel samples with a small JPEG tolerance. Later blue frames/pages must be absent;
white canvas placement, orientation and omission counts must be correct. Literal
bracket source paths, source hashes and creation/modification timestamps are checked.
Invalid inspection and controlled snapshot changes/short-copy metadata must prevent
conversion; unknown snapshot arrivals and neighbors survive exact owned cleanup.

The frame suite also embeds a 1,328-byte self-generated HEVC collection with two
distinct top-level images and tests both HEIC/HEIF extensions against the actual
decoder. It contains no thumbnails or auxiliary images and is not a timed animation.
The decoder-visible count, first primary image and omitted second image are checked
with independent known-color samples. Its recipe, fixture SHA256 and versioned
official generator-wheel provenance are in the test; the generator and its libraries
are not runtime or CI dependencies. Pinned portable ImageMagick can decode HEIC but
cannot encode it, so the fixture was generated once in owned ignored scratch.
These mandatory tests fail if the selected codec cannot decode their real fixtures;
they never count a controlled routing test, unavailable codec or skipped case as a
positive sequence result. Timed HEIC animation remains outside this evidence.

Colour tests use the scoped CC0 ICC profiles, source/license manifest, synthetic
recipe and independent reference tuples in `fixtures/colour/`. The runner hashes
every file in that directory in targeted and complete execution summaries. Optional
reference regeneration uses Pillow 12.3.0/ImageCms with LittleCMS 2.19; mandatory
tests do not require Python or Pillow. The reference requests Relative intent with
black point compensation off. The CMYK profile contains only A2B0, a perceptual
forward mapping, so LittleCMS uses that available fallback for the request. This
verifies the selected fixture's forward transform without claiming general print
accuracy, other rendering intents or reverse CMYK characterization.

Tagged wide RGB, already-sRGB and actual four-channel CMYK fixtures reach the real
runtime argument sequence. A parallel lossless PNG using those same arguments is
compared against independently computed RGB tuples within 3/255 per channel before
JPEG coding; final JPEG patch centers allow 12/255. Actual native probes differed
by at most 1/255 before JPEG. Untagged RGB uses the stated sRGB assumption, known
linear RGB is encoded into sRGB, and untagged CMYK fails before conversion.
Malformed retained ICC profiles, decoder-rejected PNG iCCP, real profile/model
mismatches and zero native exits with diagnostics must not finalize an inaccurate
JPEG. A structurally valid ICC with an invalid curve also exercises the actual
default native runner's rejection. An unsupported ICC version independently
demonstrates a real native zero exit with diagnostics and a fully decoded JPEG;
the default runner must still stop on the first failure and clean its candidate. A genuine ICC remains
active despite spoofed free metadata properties.

Eight actual EXIF orientation tags have synthetic GPS, XMP, Photoshop/IPTC, comments
and ICC attached to asymmetric JPEG fixtures. Displayed corner maps and dimensions
must match; independent JPEG segment inspection verifies privacy/profile removal.
Fully transparent and semi-transparent patches compare white composition after
the managed sRGB transform. A controlled extent override forces six real native
scale attempts and verifies the owned target profile's path and trusted bytes on
every attempt, together with the white alpha and stripping operations. Sources keep byte hashes and creation/modified
timestamps, and each success produces one intended JPEG with exact owned scratch
cleanup. All generated media and raw development history remain in marked ignored
scratch without recursive deletion.

Path tests create synthetic percent, template, bracket, hash, at-sign and leading
hyphen filenames, selected GIF frames, copied opaque video bytes, percent directory ancestors and genuine
tagged ICC inputs. Fully decoded JPEG pixels and exact final file lists distinguish
the intended red image from deceptive blue neighbors. A narrow snapshot observer
delegates the real copy, then creates a differently colored numeric-template
counterpart in owned scratch; the actual native inspection and conversion decide
which image is read. Final-name arrivals exercise no-overwrite behavior.

Launcher tests copy the actual BAT bytes and instrument only the PS1 final call
with an owned output/dependency seam and an argv/host/application-exit receiver. Real Windows
PowerShell 5.1 receives spaces, apostrophes, ampersands, parentheses, exclamation
marks, percent/brackets and Finnish/German/emoji characters through native `-File`
and inherited delayed-expansion CMD. BAT trailing separators are compared as
canonical directory identities. Caller-side CMD expansion of paired percent
variables precedes BAT entry; the direct PowerShell test proves that literal form.
Neither PATH, USERPROFILE, real Pictures nor persistent settings are changed.

Long local paths run against actual owned files beyond MAX_PATH, retaining source
bytes and creation/modified times whether the host supports them or reports an
explicit failure. Drive and UNC root containment checks are lexical and do not
enumerate a root/share. A separate five-second capability probe checks only an
existing localhost administrative alias of the owned fixture. When available,
actual direct UNC input/output and BAT UNC-source conversions verify argv, complete JPEG pixels and source
preservation; unavailable capability is recorded as `not_run` outside the mandatory
Pester count. No share is created, configured or searched, and lexical checks never
establish live UNC support. Raw capability observations remain in owned scratch.

Sizing tests (T040-T042) require exact invariant byte extent arguments for the
default, non-KiB values and Int64 maximum under a hostile numeric culture. Genuine
JPEG comment segments provide a controlled encoder-result length at cap minus one,
equal to cap and cap plus one; actual native full decoding and file bytes decide
status independently of rounded display values. This seam makes no privacy or
quality claim. A real tagged-alpha tiny-cap conversion retains all six existing
scales, managed colour, white alpha composition and stripping operations, then
reports its valid above-target JPEG as a warning. Matching lossless probes require
exact white before JPEG coding; this impossible one-byte cap permits 24/255 corner
error after JPEG compression. Duplicate links retain that status
and count the warning once. A later compliant scale supersedes an above-target trial
without retaining its warning. Truncated, wrong-format or nonzero last attempts cannot
reuse an earlier candidate. Real default conversion preserves geometry without new
quality, sharpening or upscaling flags. Source hashes and creation/modified times,
exact final names and owned scratch cleanup remain checked. Representative size
and quality measurements are recorded separately; these regressions do not imply
owner acceptance of a different resizing or quality algorithm.

The development-only `Measure-SizeQuality.ps1` runs the previous frozen runtime
and the clean implementation on owned, font-free gradient landscape/portrait,
coloured edge/detail, seeded noise and already-lossy JPEG fixtures, plus an opaque
video copy. In a fresh Windows PS5.1/PS7 host, after initializing the pinned test
dependencies:

```powershell
.\tests\Measure-SizeQuality.ps1 -ImplementationCommit <full-SHA> -ResultDirectory .scratch\new-benchmark
```

It records native flags/exits, actual bytes/dimensions/scales, RGB8 MAE/PSNR at
output and original grids, source/video hashes and timestamps, and native/application
timings. Defaults include 1,048,576, 262,144, 65,537 and 1,024-byte targets. The
65,537-byte comparison checks an unchanged native byte budget; binary-divisible
caps may differ because the old decimal suffix understated them. Raw media,
transcripts, JSON and CSV remain ignored. Single-machine synthetic timings include
tracing/validation overhead; these measurements are neither perceptual scores nor
owner approval. No runtime tool installation or algorithm change is introduced.

Diagnostics tests (T043-T045) use an owned native executable compiled with the
Windows .NET Framework compiler and genuine pinned ImageMagick derivatives.
They independently verify separated streams/native exits under both PowerShell
hosts, multi-MiB concurrent output without newlines, fixed retained character
bounds and fail-closed overflow. Genuine no-regard lossless-to-lossy JPEG warning,
unknown-coder, truncated-input and exclusively locked-source probes establish
native diagnostic provenance. The warning also gets explicitly re-emitted by a
controlled child after real production-argv encoding to exercise application
acceptance; this does not claim that the unchanged production quiet/regard flags
naturally produce that zero-exit warning.

Permanent categories must stop once; explicitly diagnosed sharing/lock violations
can retry only twice per image, with logged 100/200 ms backoff, fresh candidates,
the same scale and identical colour/alpha operations. Generic Permission denied
remains permanent even when a held source lock caused it. Tests verify ICC/white
pixel results, native-warning duplicate status/counters, zero-exit ICC/unknown
errors, nonzero valid JPEG rejection, and a combined maximum of eight calls across
six valid size attempts plus two transient retries. Sources and videos retain
hashes/timestamps, and exact owned scratch is checked after every case. These
checks cover stream retention and retry bounds; general conversion timeouts and
cancellation remain later milestone work.

The integrated conversion regression (T046) normalizes the same owned mixed tree
twice with actual native processes: the default 1,048,576-byte target and an exact
65,537-byte target. Tagged RGB/alpha and genuine CMYK, EXIF orientation/privacy,
GIF/WebP displayed canvases, the first oriented TIFF page, and the approved
two-image HEIC/HEIF still collections share literal percent/bracket/Unicode paths.
Assertions cover declared source/output maps, exact per-image cap/scale/status,
full JPEG decoding, progressive coding and RGB 4:2:0 sampling, mirrored directories, duplicate
retention, opaque video bytes, all source hashes/creation/modified times, and the
complete first output tree after the second run. No application process results
are mocked in these two cases.

Colour references reuse the committed independent Pillow/LittleCMS tuples and
their `REFERENCE.json` tolerances: 3/255 before JPEG and 12/255 at uniform patch
centers. Known frame/canvas corners use the same narrowly scoped JPEG coding
tolerance. The CMYK profile's requested Relative intent uses its available A2B0
mapping; this verifies the synthetic fixture rather than general print accuracy.
The seeded noise is genuinely gray; its progressive JPEG has one component and
does not carry RGB chroma. Noise comparisons retain original-to-output-grid lossless sRGB references,
decoded PNG/RGB8 buffers and descriptive MAE/RMSE/PSNR, without a perceptual
quality threshold. Actual recipe, fixture/output/log hashes, source/runtime/test/
profile/reference/tool bindings and native/delegate versions remain in marked
ignored scratch. Seeded fixture bytes and JPEG outputs may vary by codec/build;
these checks do not require cross-build byte identity or certify timed HEIC
animation or all possible decoder implementations.

The recovery/reporting regression (T047-T050) establishes actual denied directory
enumeration on an owned Windows ACL fixture, continues accessible native image
and video siblings, and restores the original DACL in a finally block.
The denied-only case explicitly starts without auto-inheritance. Its fixture-only
native DACL restoration preserves the complete requested owner/group/DACL
descriptor bytes, DACL bytes and control flags exactly; no SDDL text is ignored
or normalized. Original auto-inherited descriptors retain the Set-Acl path. Real
junction loops, outside targets and a root alias are mandatory; a separately
attempted file symlink reports its actual host capability. Missing symlink
privilege does not count as actual file-symlink coverage.

Logging tests distinguish a controlled creation failure from an exclusively held
real log during actual native conversion. They assert one persisted degraded
outcome, bounded fallback/drop counts, quiet diagnostic retention and caught
emergency-stderr failure. A controlled PipelineStoppedException tests helper
classification only; it does not establish Ctrl+C/process lifetime behavior.
Independent timestamp setter failures for images and videos must retain full
validated media, attempt the other field, continue siblings and register a
warning duplicate status without inventing another failed file. Source hashes,
creation/modified times and directory/link state are compared, and source,
runtime, test, dependency/tool bindings and actual observations remain under
marker-owned ignored scratch. No application permission recovery is attempted.

Native lifetime regressions (T051-T053) compile benign Windows process fixtures
only into marked ignored scratch. An independent Toolhelp32 observer records the
owned root, child and grandchild with direct-parent IDs and creation identities.
Actual deadlines must terminate that tree, release held files and inherited output
pipes, and return within a bounded drain allowance. Concurrent stream flooding and
a valid JPEG written before a stall cannot turn timeout into accepted media. An
unrelated actual ImageMagick process waits for three synthetic RGB stdin bytes;
it must survive the owned timeout and then complete its own valid PNG.

Internal lower-only ceilings exercise a shared per-image deadline across source
inspection, ICC extraction, conversion and full candidate validation. A narrowly
controlled observer routes selected phases to actual benign sleepers, or adds
finite native delays; it returns actual wrapper outcomes. Other phases and later
image/video siblings use real ImageMagick and unchanged source bytes. Actual benign
cache exhaustion uses small child-local memory/map/disk limits, with owned temporary
environment paths and no machine resource pressure or installed policy changes.
A private stricter policy copy retains original restrictions and verifies that its lower memory ceiling wins.
Checks retain native exits, job/lifetime details, source hashes/creation/modified
and directory state, full JPEG pixels, opaque video bytes, exact scratch cleanup,
and source/runtime/test/compiler/tool bindings. Unexpected neighboring files survive
nonrecursive cleanup. These checks cover timeout ownership and resource budgets;
real Ctrl+C and host shutdown remain separate lifecycle verification.

The cache uses an exclusive short directory under a validated ordinary Windows
temporary root, separately from the output-volume JPEG/ICC/snapshot candidates.
Tests supply only their marker-owned ignored root, reject source-contained and
overlong roots, and require exact cleanup. A controlled unconfirmed tree flag
retains affected scratch and refuses finalization; it does not claim an actual
failed process kill. Source directory timestamps are fixed after all fixture writes
and freshly read, alongside the unchanged file bytes and timestamps.

## Cancellation and forced termination (T054-T056)

Cancellation regressions (T054-T056) keep actual Windows console delivery separate
from controlled request state and fixture pacing. Benign drivers create a separate
owned console, verify its members, then send real CTRL_C_EVENT with group zero.
Redirected stdin ETX is not a substitute. Retain actual PS5.1/PS7 -File and unchanged
BAT commands, native exits and source/tool bindings; qualify phase observers that
route real native work or control copy pacing. Record existing BAT Done/pause/exit
behavior without asserting universal application exit propagation.

Image/native and chunked-video interruption checks must retain full validated
completed outputs, reject the interrupted candidate/partial, stop subsequent items,
and preserve source state plus unrelated-process identity. A final move admitted
before a request may finish and remain; later incomplete work cannot acquire its
final name. Exact owned cleanup requires confirmed tree completion. Synchronous
I/O may delay a polling boundary; unavailable sinks may prevent a visible summary.

A separately forced disposable worker records actual exit and recognizable owned
scratch. Force exit or terminal closure is not cooperative 130, guaranteed finally
cleanup or a successful interrupted summary. Hash controller/runtime/suite,
transcripts, output/source snapshots and observation JSON under ignored owned roots.
Keep one actual targeted_histories list and qualified cancellation_console_observations/
ConsoleObservationsFile rows; neither a helper result nor event-send return alone
establishes the real-host gate.

## Reconciled outcomes and local CSV (T057-T060)

Reporting.Tests.ps1 checks one outcome per successfully inspected regular file,
ignored discovery, warning attributes, retained mappings and final success before
ancillary failures. Signed image-only paired bytes use actual source/final lengths;
videos, duplicate/ignored/error/cancelled/unstarted bytes create no image savings.
Owned denied-directory fixtures preserve unknown subtree counts separately.
Report I/O faults are distinguished from real media failures, and valid finalized
media must survive required-report degradation.

Hostile legal Windows filenames exercise the visible text: prefix, relative
mapping, CSV quoting and independent reversible decoding through actual Import-Csv
on both supported shells. UTF-16/control cases that Windows cannot store as names
are serializer unit probes. These parser checks do not claim actual Excel or
LibreOffice execution, formula editing or save/reopen behavior. Keep all prefixes
for spreadsheet display; README documents the qualified all-Text import recipe.
Raw CSV/media and detailed source snapshots remain in marker-owned ignored scratch.
