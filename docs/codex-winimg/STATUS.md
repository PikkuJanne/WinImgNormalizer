# Project status

Updated: 2026-10-04 — M1-T01 completed.

Repository: PikkuJanne/WinImgNormalizer
Actual checkout: C:/projects/WinImgNormalizer
Branch: codex/winimg-hardening
Tested implementation: f57b8adbb6859e8812471a4cc123b73530a2b301
Next task: **M1-T02 — Protect traversal and allocate unique run directories**
Task progress: **4 / 28 accepted**
Specified cases exercised: **T001-T012 (12 / 75); T001-T004 are characterization**
Pester suite: **66 / 66 passed in each actual desktop Windows shell, zero skipped**
Failure controls: **67 total, 66 passed, deliberate assertion failed; exit 1 in each shell**
CI: **push and PR Windows Server PS 5.1/PS 7 matrices passed on tested implementation**
GitHub sync for this containing evidence checkpoint: **pending_verification**
Owner quality acceptance: **not requested / not granted**

## Implementation and evidence

Argument count, FileSystem directory and positive Int64 byte limits are validated
with friendly setup exits before Pictures/run/log creation. Root-aware absolute
paths preserve drive/UNC separators; drive-relative inputs and linked source roots
are rejected. Exact application lookup bypasses aliases/functions. Bounded native
version/format queries enforce the currently reviewed ImageMagick 7.1.2-32 floor
and record executable/version/delegates/relevant capabilities in the local log.
Upstream release/security review and chosen policies are in SOURCES/DECISIONS.

Missing decoders receive one per-file diagnosis and zero conversion attempts;
supported images/videos continue with code 2. HEIF and HEIC use their own advertised
coders. Video-only/empty batches still require ImageMagick; JPEG writing is required
only for readable images. Destination write/delete probes remove themselves; setup
checks a practical overflow-safe estimate and catches actual directory-creation
failures. Unknown free space warns and continues. Large divisible byte caps no
longer overflow the legacy MB/KB quotient; ordinary unit syntax is retained.

Actual Windows 11 Pro build 26300: PS 5.1.26100.9444 and PS 7.6.5, Pester 5.9.1 and
verified portable ImageMagick 7.1.2-32 each passed 66/66. Both real positional forms,
contained child -File setup failures, real dependency identity despite shadowing,
ordinary JPEG full decode, video hashes/timestamps and source preservation passed.
Deliberate controls returned exactly one false-assertion failure and exit 1.
Committed runtime/test hashes match all four runner summaries; persistent execution
policies are unchanged. No installer, persistent PATH change or global policy edit.

Real push/PR CI each passed PS5.1/PS7 Windows Server 2025 jobs: 66 passing tests and
67 total/one identified deliberate failure/exit1. Actual CI versions are PS
5.1.26100.33438 and PS 7.6.6. Job metadata and selected log counts/versions are saved
in evidence/M1-T01-ci.json; desktop evidence is distinct. Initial test-only adapter
expectation failures and review corrections are retained in M1-T01.json.

## Handoff and remaining scope

Started clean at 866aba0 independently equal to the feature remote. Owner-merged
PR #3 main at b153d70 had the same tree; fetch was read-only and no pull/merge/reset
was performed. Successor draft PR #4 contains M1-T01. Implementation f57b8adbb6859e8812471a4cc123b73530a2b301
was pushed and independently synchronized while clean at 2026-10-04T14:59:09.193774+00:00.
This final record-only checkpoint binds that implementation; its own final SHA and
remote observation will be reported externally after commit/push.

Root cases are lexical/owned-path controls; no live UNC share, drive-root traversal,
real Pictures or manual drag-and-drop test. Denied writes/mkdir/space failures are
controlled seams; HEIC/HEIF gap controls do not claim real codec coverage. Existing
containment/traversal and run-root reuse defects remain M1-T02. Collision/frame,
dedupe, invalid outputs, colour/lifecycle/general exits, full corpus/analyzer and
owner acceptance/publication gates remain later tasks. No release readiness claim.
Stop after M1-T01; begin M1-T02 in a fresh thread.
