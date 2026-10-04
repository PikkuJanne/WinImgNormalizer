# Windows checks (through M1-T05)

These development tests cover the import boundary, the two existing positional
invocations, setup validation (T008-T012), traversal/run isolation (T013-T016),
deterministic naming/no-overwrite targets (T017-T019), and validated conversion and
staged video transactions (T020-T024), and successful-retention duplicate handling
(T025-T027).
They do not certify the later metadata, source-frame policy or cancellation fixes. The characterization
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
`Traversal.Tests.ps1`, `Naming.Tests.ps1`, `Transactions.Tests.ps1` and
`Duplicates.Tests.ps1`.
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
HEIC naming/processing uses explicitly controlled capabilities and JPEG candidate
bytes, so it proves routing and distinct names rather than an installed HEIC codec.
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
