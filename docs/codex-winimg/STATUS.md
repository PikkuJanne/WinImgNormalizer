# Project status

Updated: 2026-10-04 — M1-T05 completed.

Repository: PikkuJanne/WinImgNormalizer
Actual checkout: C:/projects/WinImgNormalizer
Branch: codex/winimg-hardening
Tested implementation: c9988e204d25bbb13616b6df601d0f5c400a73c1
Next task: **M1-T06 — Close the output-safety milestone**
Task progress: **8 / 28 accepted**
Specified cases exercised: **T001-T027 (27 / 75); T001-T004 are characterization**
Pester suite: **172 / 172 passed in each actual desktop Windows shell, zero skipped**
Failure controls: **173 total, 172 passed, one intended assertion failed; exit 1 in each shell**
CI: **implementation push and PR Windows Server PS 5.1 / PS 7 matrices passed, including expected controls**
GitHub sync for this containing evidence checkpoint: **pending_verification**
Owner quality acceptance: **not requested / not granted**

## Implementation and evidence

Duplicate handling now follows the existing total ordinal output-plan order and
registers retained entries only after successful image or staged-video finalization.
A failed conversion, copy or final move cannot suppress a later usable candidate.
The lightweight case-normalized filename plus modification-time key additionally
requires matching input length before a skip. Every skip records the retained
successful source, output and status, making the relationship visible in local logs.
Valid above-target images may register success with their existing best-effort note.
Each input length keeps its own first success. Live source metadata precedes lookup;
an image whose source length/time changed during processing keeps its valid finalized
output with a warning and is excluded from matching. Video retention uses the stable
copy helper's measured length/time snapshot.

Equal filename, time and length remain a heuristic: different content can still be
falsely matched. This limitation is documented, originals are preserved and no
runtime content hashing or hash database is added. Public invocation forms, byte
cap/scales, JPEG settings, deterministic naming and transaction validation stay
unchanged. Decision D23 records the behavior and its limits.

Actual Windows 11 Pro build 26300, PS 5.1.26100.9444 and PS 7.6.5, Pester 5.9.1 and
verified ImageMagick 7.1.2-32: 172/172 each, zero skipped. Deliberate controls
each have 173 total, 172 passed and one identified false assertion with exit 1.
All four summaries bind the same clean committed checkout bytes; evidence separately
records raw file hashes, Git blob hashes and any CRLF/LF normalization. Persistent
execution policy is unchanged. T025-T027 cover failed-first retention, the input-length
guard, independent length entries, refreshed metadata, actual case variants in separate
directories, 15 seeded shuffle/culture runs per shell, stable retained source/output/
status links and the demonstrated heuristic limitation. Controlled source mutations
verify registration safeguards; ordinary sources retain original bytes/timestamps.
The inherited mandatory suites continue to verify safe sources, collision planning,
full JPEG decoding, staged video bytes and exact owned cleanup.

Evidence: evidence/M1-T05.json and M1-T05-ci.json; prior evidence remains retained.
All four implementation CI jobs passed: [push CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37218484246) and [PR CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37218495898).
Development observations are retained separately from the exact final revision.
Initial committed 4fce2f0 desktop suites passed 171/172 and controls had the same
unexpected T013 diagnostic assertion alongside T007. Source-reparse rejection still
worked. Correction c9988e2 restores its established message and keeps a separate
directory guard; the inherited T013 assertion is unchanged. No CI result is claimed
for the initial revision superseded before implementation CI.

## Handoff and remaining scope

Started clean at 8c9464b7e60709d0baa83cfb21167a26b5027a87, independently equal to the
feature remote. Owner-merged PR #7 main d8936872f3621b40126d5186d5600957d582bd7f had
the identical tree; read-only fetch did not pull, merge or reset. Successor draft PR
#8 contains M1-T05. The preserved non-Git snapshot remains separate.

Implementation synchronization is recorded as a past exact-SHA observation. This
record-only checkpoint binds that implementation; its own SHA and live remote state
are reported externally after commit/push. No private media or raw local logs are
tracked; no installer, persistent PATH/policy or global ImageMagick policy changes.

M1-T06 must still audit the complete output-safety milestone, source bytes/timestamps,
two-run isolation and the combined diff. Native lifecycle, frame selection, colour,
general cancellation/exits, complete corpus/analyzer and owner/publication gates stay
later scope. No hostile-filesystem sandbox, batch atomicity, crash durability, live
UNC, real Pictures, manual launcher or owner quality/publication result is claimed.
Stop after M1-T05; begin M1-T06 in a fresh requested thread.
