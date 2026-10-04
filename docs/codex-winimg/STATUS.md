# Project status

Updated: 2026-10-04 — M0-T03 completed.

Repository: PikkuJanne/WinImgNormalizer
Actual checkout: C:/projects/WinImgNormalizer
Branch: codex/winimg-hardening
Tested implementation: 3b80ee94b8471a22560ea90a7583283c1613195c
Next task: **M1-T01 — Validate inputs, root paths and ImageMagick before work**
Task progress: **3 / 28 accepted**
Specified cases exercised: **T001-T007 (7 / 75); T001-T004 are characterization**
Early Pester suite: **6 / 6 passed in each actual Windows shell, zero skipped**
CI: **PS 5.1/PS 7 Windows matrix passed on the tested implementation**
GitHub sync for this containing evidence checkpoint: **pending_verification**
Owner quality acceptance: **not requested / not granted**

## Implementation and verification

The existing script now defines a callable orchestration function and positional
command adapter. Dot-source import does not normalize, create user folders, alter
caller preferences/log state or exit its host. Internal destination/process-runner
parameters support contained tests; the public invocation forms and batch file
remain unchanged. Existing conversion and known defects are preserved.

Windows 11 Pro build 26300, PS 5.1.26100.9444 and PS 7.6.5 each passed six Pester 5.9.1
tests. Ten runner controls passed their expected outcomes: normal 6/6, exit 0,
deliberate 7 total, 6 passed, 1 failed, exit 1, zero discovery, skipped test and discovery
error each exit 1. Persistent policies and tested source hashes remained unchanged.
A scratch-only exit 0 import mutation was detected before Pester loads the app.

Four independent real ordinary parity runs produced sixteen completely decoded
JPEGs and twelve videos matching M0-T02 hashes, dimensions, tree and timestamps;
source bytes/creation/modified times were unchanged. T006 also executes both real
-File positional forms in both shells, injecting only the outer scratch destination.
Real Pictures resolution and manual batch drag-and-drop are outside this evidence.

Pinned development archives are privately extracted from verified bytes, without
installation or persistent PATH/policy changes. GitHub first failed before tests
because command discovery returned both Windows and Git tar executables. Explicit
System32 libarchive selection fixed it; local default bootstrap with Git tar first
on PATH passed, then real push/PR CI passed both shells and deliberate controls.
The failed initial revision/runs remain recorded. CI runner versions/counts and
job links are in evidence/M0-T03-ci.json; Server CI is distinct from desktop results.

## Handoff and remaining scope

Evidence: M0-T03.json, M0-T03-runner.json, M0-T03-parity.json,
M0-T03-import-control.json and M0-T03-ci.json. Runtime/tests are bound to the tested
implementation above; this checkpoint changes only records/evidence. Its actual
final push and independent remote SHA comparison will be reported externally.

Started clean at a3ef050 independently equal to the feature remote. The owner
merged PR #2 into main at 8e35b6a with the same starting tree; no pull/merge/reset or
default-branch change was performed. Successor draft PR #3 continues this feature.
Original non-Git snapshot, launcher, license and assets remain preserved.

Known collision/frame/dedupe/invalid-output/preflight/colour/lifecycle defects,
full corpus/analyzer/codec coverage, subjective owner acceptance and publication
gates remain assigned to later tasks. No release or website readiness is claimed.
Stop after M0-T03; begin M1-T01 in a subsequent thread.
