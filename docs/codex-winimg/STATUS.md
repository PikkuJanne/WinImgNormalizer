# Project status

Updated: 2026-10-05 — M2-T05 completed.

Repository: PikkuJanne/WinImgNormalizer
Actual checkout: C:/projects/WinImgNormalizer
Branch: codex/winimg-hardening
Tested implementation: 5ddbf4156f2c10534a045586a5b5e5e050255c30
Next task: **M2-T06 — Close the conversion-correctness milestone**
Task progress: **14 / 28 accepted**
Specified cases exercised: **T001-T045 (45 / 75); T001-T004 are characterization**
Pester suite: **269/269 passed in each desktop shell, zero skipped**
Failure controls: **270 total, 269 passed, one intended failure; native exit 1**
CI: **implementation push/PR Windows matrices passed, including controls**
Containing evidence checkpoint synchronization: **pending_verification**

## Current behavior

A native failure now stops that image at its current scale even if the encoder leaves a
decodable JPEG. Valid later files continue. Only an exact known JPEG notice with native
exit 0 and full validation can be retained as ConvertedWithWarning; it adds
NativeWarnings within ConvertedImages and returns application code 2. SizeWarnings
remains a separate retained-output attribute.

Stdout and stderr drain independently with bounded retained text; truncation fails. Only
explicit sharing/lock failures can retry twice total per image, after 100/200 ms, at the
same scale with identical flags and fresh owned scratch. Six smaller-size attempts are
reached only from valid above-cap output, giving at most eight conversions including
retries. Source/video preservation, no-overwrite safety, colour/white-alpha/strip
policy, exact byte cap and JPEG settings remain.

T043 — passed: Both actual Windows hosts capture native stdout/stderr/exit independently
under ErrorActionPreference Stop. Genuine pinned lossless-to-lossy warning provenance
distinguishes exit0 without regard-warnings from nonzero with regard-warnings, both with
decoded JPEGs. Controlled native re-emission after actual production conversion
finalizes explicit warning/exit2 with valid white/red JPEG and warning duplicate status;
nonzero, truncated, unknown or ICC-mixed outcomes cannot finalize.

T044 — passed: Actual owned unknown-coder, truncated-JPEG and exclusive-lock diagnostics
classify permanently; the native Permission denied text remains nonretryable. Controlled
native missing-codec, damage, denied, resource and unknown failures stop once despite
valid JPEG bytes. Only explicit sharing text or trusted Win32 HRESULT 32/33 retries;
another facility with low word 32 stops. Two 100/200 ms-backoff retries per file
preserve scale/ICC/alpha arguments and fresh candidates. The actual eight-call case
contains two controlled sharing failures and six real valid above-cap size conversions,
ending in a 541-byte 32x24 JPEG/size warning/exit2 with source/video preservation.

T045 — passed: Two actual concurrent newline-free child streams each emitted 4194330
characters and native exit9; retained prefix/tail were bounded 1024 characters per
stream, counts exact and streams complete without deadlock. At production capture
limits, a controlled zero-exit native process that wrote a valid JPEG while overflowing
streams is rejected once as OutputLimit; useful tail details remain in the bounded log
and console output stays below 4096 characters.

## Verification and handoff

Actual Microsoft Windows 11 Pro build 26300, PS 5.1.26100.9444 and PS 7.6.5, Pester
5.9.1: each normal run passed 269/269 with native exit 0 and zero skipped or incomplete
tests. Each control had 270 total, 269 passed, one exact T007 deliberate failure and
native exit 1. Summary/XML, parent native records and console hashes bind 26 tested
source paths to raw checkout and separate Git blobs, with only verified CRLF-to-LF
normalization where needed. Persistent execution policies were unchanged.

Implementation Windows Server PS5.1/PS7 push/PR gates also passed: [push CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37267178588); [PR CI](https://github.com/PikkuJanne/WinImgNormalizer/actions/runs/37267182413).
Evidence: evidence/M2-T05.json and evidence/M2-T05-ci.json.

Preserved development failures and the initial committed T005 failure are documented in
evidence/M2-T05.json and SESSION_LOG.md. The next revision passed full desktop and
hosted gates, then a separate probe found a diagnostic-field omission in the development
measurement trace. The final revision corrects that consumer, passed fresh acceptance
gates, and preserves both earlier committed histories.

Started clean at 2e30cf9948669f7af32533efc9d842472693be6d, equal to the live feature
branch; owner-merged PR #13 main 590241944033c1332105e21a9dab49578a939941 had the same
tree. Continued without changing checkout/history. Successor draft PR #14 contains this
task. Implementation synchronization is a past exact-SHA observation; this record-only
checkpoint's SHA/live synchronization remains pending_verification until observed
externally after normal commit/push.

General per-file deadlines, cancellation, native descendant lifecycle and final
reporting/BAT exit propagation remain later M3 gates. Fixed character capture is not a
CPU/disk/lifetime sandbox. Controlled warning, retry and process-local resource fixtures
remain qualified; prior codec/colour/provider limits and owner quality/default approval
are unchanged. Full limitations and raw artifact bindings are in evidence/M2-T05.json.

Stop after M2-T05; begin M2-T06 only when requested.
