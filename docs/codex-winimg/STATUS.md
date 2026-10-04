# Project status

Updated: 2026-10-04 — M0-T02 partial characterization checkpoint.

Repository: PikkuJanne/WinImgNormalizer
Actual checkout: C:/projects/WinImgNormalizer
Branch: codex/winimg-hardening
Reviewed starting HEAD: d323be72cbd61fe207af6b7876f52b2d809a2311
Advertised main: 2a51c0b6fa36370ee8341ff0d6c9b0b638a27a73
Current/next task: **M0-T02 — blocked on real ImageMagick tooling**
Task progress: **1 / 28 accepted; M0-T02 is incomplete**
Coverage: **T004 completed; T002/T003 partial; T001 not run**
Real ImageMagick application runs: **0**
Instrumented legacy fake-process observations: **6 across PS5.1/PS7**
GitHub sync for this containing checkpoint: **pending_verification**
CI: **unconfigured; no passing CI claimed**
Owner quality acceptance: **not requested / not granted**

## Work and blocker

Added a synthetic fixture generator, environment inventory, contained legacy
snapshot wrapper, bounded controller and native-process fakes. The tracked app,
launcher and assets are unchanged. The snapshot replaces only the Pictures lookup;
all execution uses a separate -File process. No dot-sourcing, real Pictures writes,
unsafe nested recursion, persistent policy change or software installation occurred.

The formal harness check captured six fake-process observations, passed ten
containment rejection controls and exercised six Unicode destinations. Source
hashes/creation/modified times and persistent shell policies stayed unchanged.
Failed-first duplicate suppression and zero/invalid output acceptance were
reproduced in both shells. These are defects, not passing corrected regressions.
Mocked nested planning is recorded without unsafe filesystem execution.

Real ordinary-image/video, same-stem/directory collision and GIF/TIFF observations
are not_run. The full generator is implemented but its ImageMagick encodes are
unverified. HEIC/HEIF and external profile provenance remain deferred coverage.
No ImageMagick was found on PATH, likely package locations or bundled tools.
The owner tooling choice is pending: verified official portable ImageMagick
7.1.2-32 Q16 x64 inside ignored scratch, or an existing verified executable.
No download or installation occurred. Keep M0-T02 incomplete until real runs
meet acceptance; do not advance to M0-T03.

Environment: Windows 11 Pro 10.0.26300 x64, fi-FI/UI en-US; Windows PowerShell
5.1.26100.9444; PowerShell 7.6.5; bundled Python3.12.14 with Pillow12.3.0.
PATH Python3.14.6 lacks Pillow. Pester3.4.0 is discoverable but unused; analyzer
and ExifTool absent. Workers remove inherited PSModulePath only in their child
environment so each edition can resolve its own default modules.

## Reconciliation and synchronization

The clean feature branch matched independently advertised remote HEAD at the
starting SHA above. The owner merged previous PR #1; main's merge tree equals
the selected feature tree. Fetch updated refs only; no pull, merge, reset or
history rewrite was performed. Origin fetch/push identity remains canonical.
Use one successor open draft PR because the preceding PR is already merged.

Evidence: evidence/M0-T02.json and M0-T02-observations.json. The checkpoint binds
tested files by precommit hashes. Its actual push/remote comparison belongs in
the response/PR after commit; no prospective sync or CI pass is recorded here.
The original non-Git snapshot remains preserved. Publication gates remain intact.
