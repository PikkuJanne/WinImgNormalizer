# Project status

Updated: 2026-10-05 — M3-T04 completed.

Repository: PikkuJanne/WinImgNormalizer
Actual checkout: C:/projects/WinImgNormalizer
Branch: codex/winimg-hardening
Tested implementation: 158315f1e48e6199c5c18d3e323b5945390789b0
Next task: **M3-T05 — Propagate exit codes and harden the batch launcher**
Task progress: **19 / 28 accepted**
Specified cases exercised: **T001-T060 (60 / 75); T001-T004 are characterization**
Pester suite: **348/348 passed in each desktop shell, zero skipped**
Failure controls: **349 total, 348 passed, one intended failure; native exit 1**
CI: **implementation push/PR Windows matrices passed, including controls**
Containing evidence checkpoint synchronization: **pending_verification**

## Current behavior

Discovery creates one terminal outcome per successfully inspected regular file,
including Ignored unsupported files; inaccessible directories, uninspectable entries and
skipped links remain separate scan issues with unknown file totals.

Validated no-overwrite image and stable video final moves commit their retained outcomes
before ancillary operations; timestamp, cleanup, logging and observer warnings do not
double-count files as errors.

Exact Decimal aggregates pair only finalized image input/output lengths, allow signed
negative savings and keep video bytes separate; elapsed duration and finalized-file
throughput are observed values.

T057 — passed: Both focused hosts observed a balanced eight-file mixed partition with
two finalized images, one copied video, two duplicates, two ignored and one damaged
input. Size/native/timestamp warnings remain attributes; finalized image/video observer
faults retain success without an Error outcome. Empty and ignored-only trees remain
balanced.

T058 — passed: A genuine quality-1 286-byte JPEG becomes a validated 523-byte JPEG,
reporting signed savings -237. The six-file byte cohort pairs only two finalized images
(5032 input,1056 output,3976 saved), while the 1048576-byte copied video is separate and
duplicate/ignored/error input lengths do not enter image totals.

T059 — passed: Thirteen hostile UTF-16 text cases and a comma-decimal numeric case pass
actual BOM UTF-8 Import-Csv with an independent text:/escape decoder. Eleven actual
regular filenames under a literal percent/bracket/Unicode ancestor produce ten JPEGs
plus one duplicate; exact source/planned/actual/retained mappings survive CSV import. No
spreadsheet execution is claimed.

T060 — passed: Actual ACL enumeration denial remains one unknown subtree issue outside
the two-file known partition with exact fixture ACL restoration. Controlled CSV
create/write/close, foreign arrival, logger/console and mirror failures retain media and
degrade coherently. A request during real video copying reports one retained image, one
Cancelled, one NotStarted and two Ignored; a request after the video final move retains
CopiedVideo at130.

## Verification and handoff

On actual Microsoft Windows 11 Pro build 26300, PS 5.1.26100.9444 and PS 7.6.5, Pester
5.9.1, each normal suite passed 348/348 with native exit 0 and zero skipped or
incomplete tests. Each control had 349 total, 348 passed, exactly one T007 deliberate
failure and native exit 1. Summary/XML, parent native records and console hashes bind 31
tested source paths, including the reporting suite and all eight colour/reference
assets, to observed raw bytes and separate Git blobs. Only verified text CRLF-to-LF
normalization is allowed; ICC bytes remain exact. Persistent execution policies were
unchanged.

Implementation Windows Server PS5.1/PS7 push/PR gates also passed: [push CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37311855781); [PR CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37311860891).
Evidence: evidence/M3-T04.json and evidence/M3-T04-ci.json.

Actual development results and their source/artifact qualifications are retained in
evidence/M3-T04.json and SESSION_LOG.md. Clean implementation desktop and hosted gates
are separate acceptance observations; earlier task histories retain their original
revisions.

Started clean at fe79b55e5405521129da27ccf43d20089260a267, equal to the live feature
branch; owner-merged PR #18 main c811391580a727ae0e052e8bef24b8836fe64c22 had the same
tree. Continued without changing checkout/history. Successor draft PR #19 contains this
task. Implementation synchronization is a past exact-SHA observation. This record-only
checkpoint's SHA/live synchronization remains pending_verification until separately
observed after normal commit/push.

Actual Import-Csv and independent reversible-decoder checks establish the documented
machine import contract, not actual Excel/LibreOffice execution or consumer
editing/save/reopen safety. The spreadsheet display prefix must remain intact.

Unknown files inside inaccessible locations are not invented. An incomplete or failed
CSV must not be treated as a complete inventory; valid media remain retained.

Cooperative application130 and best-effort reporting are separate from host closure,
force termination and stopped pipelines. Those boundaries can bypass reporting/cleanup;
synchronous filesystem work remains nonpreemptible.

Native lifetime, source-stability and duplicate matching retain their documented
practical/heuristic boundaries; reporting does not create content identity, mandatory
hashing, a hostile-filesystem sandbox or an unlimited Decimal aggregate.

The BAT and JPEG/colour/frame/size defaults are unchanged in this card. M3-T05 launcher
work, M3-T06/M3-T07 and owner quality/merge/release acceptance remain separate gates.

This session stops at M3-T04. Begin M3-T05 only when requested.
