# Project status

Updated: 2026-10-04 — M1-T04 completed.

Repository: PikkuJanne/WinImgNormalizer
Actual checkout: C:/projects/WinImgNormalizer
Branch: codex/winimg-hardening
Tested implementation: ed7fb7691eca9dd55082338230657804c1865681
Next task: **M1-T05 — Make heuristic duplicate handling deterministic and success-based**
Task progress: **7 / 28 accepted**
Specified cases exercised: **T001-T024 (24 / 75); T001-T004 are characterization**
Pester suite: **151 / 151 passed in each actual desktop Windows shell, zero skipped**
Failure controls: **152 total, 151 passed, one intended assertion failed; exit 1 in each shell**
CI: **implementation push and PR Windows Server PS 5.1 / PS 7 matrices passed, including expected controls**
GitHub sync for this containing evidence checkpoint: **pending_verification**
Owner quality acceptance: **not requested / not granted**

## Implementation and evidence

Every image conversion attempt now exclusively reserves a fresh neutral path in an
owned GUID directory, including alpha fallback and lower-scale attempts. A successful
native outcome is required even if bytes exist. Candidates must be nonempty regular
files, decode all pixels as one JPEG with positive dimensions, and retain the measured
length and modification time during validation. A valid above-cap attempt is cleaned
before the next attempt; it cannot be relabeled or retained after later failures.
Only a validated final-scale candidate can receive the existing best-effort note.

Videos pass through exclusively reserved partials. Read sharing refuses cooperative
writers during stream copying; streams flush and close before output length and
source length/modification time are rechecked. Image and video finalization require
the same volume and use two-argument File.Move, preserving final-name arrivals.
Cleanup tracks exact owned candidates and removes only empty owned directories;
unknown numbered outputs and unreserved arrivals survive. Item errors return 2.
Public calls, naming plan, conversion settings and byte cap remain unchanged.

Actual Windows 11 Pro build 26300, PS 5.1.26100.9444 and PS 7.6.5, Pester 5.9.1 and
verified ImageMagick 7.1.2-32: 151/151 each, zero skipped. Deliberate controls
each have 152 total, 151 passed and one identified false assertion with exit 1.
All four summaries bind the same clean committed checkout bytes; evidence separately
records Git's CRLF/LF normalization and committed blob hashes. Persistent execution
policy is unchanged. T020-T024 exercise actual full JPEG decoding, a truncated JPEG
with readable header dimensions, >260 native candidate validation/move, wrong-format
and numbered outputs, failed/ambiguous native results, fresh fallbacks, stale attempts,
external arrivals, exact video hashes, controlled copy/source changes and owned cleanup.
Controlled timeout/cancellation outcomes and stream failures are distinct from real
host interruption. Runtime video identity uses length/time, with hashes only in tests.

Evidence: evidence/M1-T04.json and M1-T04-ci.json; prior evidence remains retained.
All four implementation CI jobs passed: [push CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37215530598) and [PR CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37215533848).
Primary API review is recorded in SOURCES.md W3 and decision D22.

## Handoff and remaining scope

Started clean at f73a4d09a966f4600d4c09173225aa5602a47d09, independently equal to the
feature remote. Owner-merged PR #6 main acd2959 had the same tree; read-only fetch
did not pull, merge or reset. Successor draft PR #7 contains M1-T04.
Development observations and corrections are retained in evidence; final results
above alone bind the tested implementation. Initial 8f10472 desktop suites passed,
but all four hosted jobs failed with 150/151 tests passing: inspection of numbered
native siblings used ordinary long paths. Test-only correction ed7fb76 applies the existing internal native
path spelling; runtime bytes are unchanged. Initial failed CI is retained alongside
corrected passing matrices. No private media or raw local logs are tracked; no
installer, persistent PATH/policy or global ImageMagick policy changes.

Implementation synchronization is recorded as a past exact-SHA observation. This
record-only checkpoint binds that implementation; its own SHA and live remote state
are reported externally after commit/push. Length/time stability does not detect a
same-length/same-timestamp source rewrite. No hostile-filesystem sandbox, batch
atomicity, crash durability, real Ctrl+C/hard-kill, live UNC, real Pictures, manual
launcher or owner quality/publication result is claimed.

The inherited duplicate key is still registered before processing and lacks a length
guard; M1-T05 addresses it. Native process lifecycle, frame selection, colour handling,
general exits, complete corpus/analyzer and owner/publication gates remain later tasks.
Stop after M1-T04; begin M1-T05 in a fresh requested thread.
