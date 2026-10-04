# Windows checks (through M2-T01)

These development tests cover the import boundary, the two existing positional
invocations, setup validation (T008-T012), traversal/run isolation (T013-T016),
deterministic naming/no-overwrite targets (T017-T019), validated conversion and
staged video transactions (T020-T024), and successful-retention duplicate handling
(T025-T027). The combined safety regression (T028) runs one complete mixed tree
twice into separate outputs under the same parent. Frame/page regressions (T029-T031)
check deliberate first-image selection, logical animation canvases and visible omissions.
They do not certify the later metadata or cancellation fixes. The characterization
evidence in `legacy/` describes those known defects separately.

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
`Duplicates.Tests.ps1`, `SafetyRegression.Tests.ps1` and `Frames.Tests.ps1`.
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
or interruption behavior. PSScriptAnalyzer and the complete corpus belong to later
tasks.

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
