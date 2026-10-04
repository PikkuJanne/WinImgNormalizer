# Project status

Updated: 2026-10-04 — M1-T06 completed; output-safety milestone closed.

Repository: PikkuJanne/WinImgNormalizer
Actual checkout: C:/projects/WinImgNormalizer
Branch: codex/winimg-hardening
Tested implementation: 4d7f9f2e7681902971b584560211139d555bf011
Next task: **M2-T01 — Implement deliberate first-frame and first-page handling**
Task progress: **9 / 28 accepted**
Specified cases exercised: **T001-T028 (28 / 75); T001-T004 are characterization**
Pester suite: **173 / 173 passed in each actual desktop Windows shell, zero skipped**
Failure controls: **174 total, 173 passed, one intended assertion failed; exit 1 in each shell**
CI: **implementation push and PR Windows Server PS 5.1 / PS 7 matrices passed, including expected controls**
GitHub sync for this containing evidence checkpoint: **pending_verification**
Owner quality acceptance: **not requested / not granted**

## Milestone closure and evidence

M1 output-safety acceptance is complete for the covered mandatory T008-T028 cases.
The combined suites verify preflight containment, pruned traversal, exclusive run
allocation, deterministic collision planning, full JPEG validation, staged no-overwrite
finalization, exact owned cleanup and successful heuristic duplicate retention.
T028 adds the mixed-tree source-preservation and two-run isolation regression.
One actual mixed synthetic source tree is normalized twice into the same
prepopulated output parent. After each run all 15 source files retain SHA256,
length, creation/modification ticks and attributes, and source directory
metadata/paths remain unchanged. Each run contains eight fully decoded JPEGs, two
byte-identical video copies and one log with the declared collision map and two
retained duplicate links. The second run allocates a separate directory and
preserves the first run media/log/tree and existing user entries.

M1-T06 changes test coverage and test documentation; runtime and batch bytes are
unchanged from the starting checkpoint.
The complete milestone diff from 866aba02a7759212a143c1ef2cd52fde3c3f3cdd to the tested
implementation was reviewed. Sanitized evidence binds its exact diff hash/path list
and the audit findings for all finalization/collision routes, public interfaces,
source preservation and runtime dependencies. No blocking audit findings remain.
Existing behavior and limitations from D20-D23 remain documented; closure does not
add a new product decision.

Actual Microsoft Windows 11 Pro build 26300, PS 5.1.26100.9444 and
PS 7.6.5, Pester 5.9.1: 173/173 each, zero skipped.
ImageMagick observation: Version: ImageMagick 7.1.2-32 Q16 x64 ad98b24:20260927 https://imagemagick.org
Controls each have 174 total, 173 passed and exactly T007's deliberately false
assertion with exit 1. All four summaries bind the same clean committed checkout.
Raw file hashes and separate Git blob hashes record any verified CRLF/LF checkout
normalization. Persistent execution policies remain unchanged.
Hosted Windows Server matrices also passed all four normal/control jobs: [push CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37220174594) and [PR CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37220187721).
Evidence: evidence/M1-T06.json and M1-T06-ci.json.

Initial targeted PowerShell 7 development run M1-T06-target-PS7-1 failed 0/1 with
zero skips because the test's timestamp-inspection Get-Item omitted -Force for the
actual hidden source. The fixture and runtime behavior remained unchanged; the
correction adds -Force to the inspection while retaining all original preservation
assertions. Targeted PowerShell 7 rerun PS7-2 and first PowerShell 5.1 run PS51-1
each passed 1/1 with zero skips and exit 0. These three observations retain actual
end-of-run hashes and summary/XML artifacts as uncommitted development history.
Initial implementation push attempt 1 failed its PowerShell 7 development-
dependency bootstrap with observed HTTP 403 before test discovery: zero tests,
exit 1, control skipped. The failing download endpoint was not present in the log,
so neither dependency is identified as the failed request. Its PowerShell 5.1 job
and both PR jobs passed all 173 normal tests and identified
174-total/173-pass/one-failure controls. One full push-workflow retry used the
same implementation, pins and guards without code/test changes; the final recorded
attempt provides the required normal/control acceptance independently of the
retained initial failure. Earlier task failures remain in their original evidence.
Final acceptance uses only the clean committed-revision full desktop
normal/control suites and actual completed hosted push/PR matrices.

## Handoff and remaining scope

Started clean at 91e7f4ee10a86ef9e91560d2cab605b9aa0b8064, independently matching the
feature remote. Owner-merged PR #8 main dc2a2a94c3dd88347dbe5ee7732e7fc17a61ec3e had
the identical tree; read-only fetch did not pull, merge or reset. Successor draft PR
#9 contains M1-T06. The preserved non-Git snapshot remains separate.

Implementation sync is a past exact-SHA observation. This record-only checkpoint
binds that implementation; its own SHA/live remote state is reported externally
after normal commit/push. No raw logs, private media or local tool binaries are tracked.

The duplicate heuristic still permits false matches for different bytes with the
same normalized name, modification time and length. Length/time checks do not prove
content identity; no mandatory runtime hashing was added. Source access times are
excluded. Live UNC, real Pictures, actual HEIC decoding, universal long source/final
paths and denied file-symlink creation remain documented coverage limits.
First-frame/first-page handling starts at M2-T01. Colour/metadata correctness, native
process lifecycle/timeouts, general cancellation/exits, complete corpus/analyzer,
manual launcher, owner quality acceptance and publication remain later gates.
No hostile-filesystem sandbox, batch atomicity, crash durability, merge, release or
deployment is claimed. Stop after M1-T06; start M2-T01 only when next requested.
