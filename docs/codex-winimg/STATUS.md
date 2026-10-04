# Project status

Updated: 2026-10-04 — M1-T02 completed.

Repository: PikkuJanne/WinImgNormalizer
Actual checkout: C:/projects/WinImgNormalizer
Branch: codex/winimg-hardening
Tested implementation: 8c9b157338419958a14c662ca379ccb8c50e8825
Next task: **M1-T03 — Plan collision-free output names before conversion**
Task progress: **5 / 28 accepted**
Specified cases exercised: **T001-T016 (16 / 75); T001-T004 are characterization**
Pester suite: **95 / 95 passed in each actual desktop Windows shell, zero skipped**
Failure controls: **96 total, 95 passed, one intended assertion failed; exit 1 in each shell**
CI: **corrected push and PR Windows Server PS5.1/PS7 matrices passed (95/95 plus expected control)**
GitHub sync for this containing evidence checkpoint: **pending_verification**
Owner quality acceptance: **not requested / not granted**

## Implementation and evidence

Canonical source/destination paths and OrdinalIgnoreCase whole-segment comparisons
reject destinations inside/equal to source before enumeration or writes, including
probes. All source/output reparse ancestors, including dangling links, are rejected.
Device paths and components ending in dots/spaces are unsupported. Ordinary missing
output tails use a verified existing ancestor.

One nonrecursive inventory skips/reports file/directory reparse points and unreadable
subtrees as incomplete scans. Mirrors use that inventory; paths are rechecked before
media processing. Scan/safety omissions return code 2 even without eligible files.

Timestamp/GUID run roots are created exclusively with Win32 CreateDirectoryW and
eight bounded collision retries, never adopting existing objects. Internal extended
Windows prefixes avoid native path limits for generated paths. Every observed
source top-level name reserves a free `.WinImgNormalizer[__N]` namespace for work
and reports. The log is opened CreateNew under reports; arriving logs are preserved.
Log setup failure stops without a shared TEMP fallback. Ordinary mirrored media
paths, positional calls and conversion defaults are retained.

Actual Windows 11 Pro build 26300: PS 5.1.26100.9444 and PS 7.6.5, Pester 5.9.1,
verified portable ImageMagick 7.1.2-32 passed 95/95. Both controls produced 96 total,
95 passed, one false assertion and exit 1. All four summaries match committed
runtime/test hashes; persistent policy is unchanged. Ordinary JPEG full decode,
source/video byte hashes and timestamps pass.

T013-T016 cover real disposable looping/outside/dangling directory junctions,
controlled file reparse/post-inventory replacement, lexical drive/UNC boundaries,
incomplete scan controls, existing run file/directory collisions, simultaneous
same-stamp child applications, a 272-character native path and namespace/log conflicts.
Desktop file-symlink creation was denied (Win32 1314); controlled coverage is distinct.
No live UNC, real Pictures, ACL exhaustion or manual launcher check. This is not a
hostile-filesystem sandbox or universal long-path guarantee for source/ImageMagick.

## Handoff and remaining scope

Started clean at 95cbc7956897794bcebd558f8395fdc0391a3baa, independently equal to the
feature remote. Owner-merged PR #4 main at c604ade had the same tree; read-only fetch
did not pull/merge/reset. Successor draft PR #5 contains M1-T02. Initial c6157fb
desktop passed 94/94, but push/PR pwsh CI failed six setup tests. An owned 272-character
native path reproduced the limit; corrected 8c9b157 adds internal extended prefixes
and a regression. Initial failures remain in the evidence.

Corrected implementation was pushed and independently matched clean feature HEAD
at 2026-10-04T15:17:42.8747086Z. This record-only checkpoint binds that implementation;
its own final SHA and remote observation are reported externally after commit/push.
Evidence: M1-T02.json and M1-T02-ci.json. Previous preflight evidence remains M1-T01.

User output-name collisions/finalization, duplicate safeguards, frames/colour,
validation/lifecycle/general exits, full corpus/analyzer and owner quality/publication
gates remain later tasks. No release readiness claim.
Stop after M1-T02; begin M1-T03 in a fresh thread.
