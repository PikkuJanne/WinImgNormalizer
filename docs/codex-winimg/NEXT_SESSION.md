# Next session handoff

Resume **M0-T02 — Characterize the current application with safe fixtures**.
It is incomplete; **do not start M0-T03** yet.

Use C:/projects/WinImgNormalizer on codex/winimg-hardening; the original configured
folder remains a preserved non-Git snapshot. Read AGENTS.md, STATUS.md, TASKS.json,
GIT_WORKFLOW.md, tasks/M0-T02.md and tests/legacy/README.md. Verify clean state,
canonical fetch/push URLs and live feature SHA independently before resuming.
Final post-commit sync is in the preceding response/successor draft PR.

The synthetic generator and real-run controller are prepared. check_harness.py
captured six fake-process observations and passed ten containment controls on
actual PS5.1/PS7, including Unicode destinations and source bytes/timestamps.
Failed-first duplicate and zero/invalid output acceptance are instrumented legacy
defects. Raw artifacts remain ignored; sanitized actual commands/transcripts and
hashes are in evidence/M0-T02.json and M0-T02-observations.json.

Blocker: real ImageMagick is unavailable. The owner tooling choice is pending:
use pinned official portable ImageMagick7.1.2-32 Q16 x64 in ignored scratch
(asset/digest/size in tests/legacy/toolchain.json), or provide an existing verified
magick.exe. No download/installation occurred. Do not treat elapsed time as approval.
After the choice, verify provenance, run the full generator/controller on both
shells and append independent fixture/output inspections and transcripts.

Use bundled Python3.12.14 with Pillow12.3.0; PATH Python3.14.6 lacks Pillow.
Workers remove inherited PSModulePath case-insensitively only in their environment.
No Pester/analyzer or CI pass exists. HEIC/HEIF and external profiles remain deferred
coverage without support claims. T001 not_run; T002/T003 partial; T004 completed.
Advance to M0-T03 only after M0-T02 acceptance.

The tracked app/launcher are unchanged. Snapshot only the exact Pictures block,
pin reviewed source hash and use ASCII Base64 injection for PS5.1 Unicode safety.
Never dot-source the app, redefine USERPROFILE or run unsafe recursion against
real Pictures/drive roots. The owner merged PR #1; continue the same feature branch
with one successor open draft PR. No default-branch change or rewrite is needed.
